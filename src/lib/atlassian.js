const CLAUDE_API = 'https://api.anthropic.com/v1/messages'
const MODEL      = 'claude-sonnet-4-20250514'
const MCP_URL    = 'https://mcp.atlassian.com/v1/sse'

const SYSTEM = `You are a Jira/Confluence data extractor for Pandora A/S.
Jira hierarchy: Project → Initiative → Epic → Story/Task.
Return ONLY valid JSON. No markdown, no preamble, no explanation.`

export async function fetchProductLineData(plName, jiraProjectKey) {
  const prompt = `
Fetch all delivery data for Product Line: "${plName}"
${jiraProjectKey ? `Jira project key: ${jiraProjectKey}` : 'Search by product line name.'}

Return JSON matching this exact schema:
{
  "productLine": string,
  "jiraProject": string,
  "fetchedAt": string (ISO),
  "initiatives": [{
    "key": string,
    "name": string,
    "goal": string,
    "percentComplete": number (0-100),
    "ragStatus": "green"|"amber"|"red",
    "epics": [{
      "key": string,
      "name": string,
      "percentComplete": number,
      "ragStatus": "green"|"amber"|"red",
      "team": string,
      "dueDate": string|null
    }]
  }],
  "teams": [{
    "name": string,
    "velocityAvg": number|null,
    "sprintCompletionRate": number|null,
    "capacity": number|null
  }],
  "impediments": [{
    "key": string,
    "summary": string,
    "team": string,
    "blockedSince": string|null,
    "stakeholderAsk": string|null
  }]
}

Rules:
- ragStatus: green ≥70%, amber 40-69%, red <40%
- velocityAvg: avg story points/sprint over last 3 sprints
- sprintCompletionRate: % committed items completed, avg last 3 sprints
- capacity: % of planned capacity currently available
- impediments: issues with impediment or blocker label only
- Use null for any field that cannot be determined
`

  const res = await fetch(CLAUDE_API, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 1000,
      system: SYSTEM,
      messages: [{ role: 'user', content: prompt }],
      mcp_servers: [{ type: 'url', url: MCP_URL, name: 'atlassian' }]
    })
  })

  if (!res.ok) throw new Error(`API error: ${res.status}`)
  const data = await res.json()
  const text = (data.content || [])
    .filter(b => b.type === 'text')
    .map(b => b.text)
    .join('')
    .replace(/```json|```/g, '')
    .trim()

  return JSON.parse(text)
}
