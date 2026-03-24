// ─── Keys ────────────────────────────────────────────────────────────────────
const key = (plName) => `pl-report:${plName.toLowerCase().replace(/\s+/g, '-')}`

// ─── Shape ───────────────────────────────────────────────────────────────────
// {
//   lastGenerated: ISO string,
//   teams: {
//     [teamName]: {
//       prev:    { happiness, teamSpirit, belonging, respondents, date } | null
//       current: { happiness, teamSpirit, belonging, respondents, date } | null
//       surveyResponses: [ { happiness, teamSpirit, belonging, submittedAt } ]
//     }
//   }
// }

export function loadState(plName) {
  try {
    const raw = localStorage.getItem(key(plName))
    return raw ? JSON.parse(raw) : null
  } catch { return null }
}

export function saveState(plName, state) {
  try {
    localStorage.setItem(key(plName), JSON.stringify(state))
  } catch (e) {
    console.warn('Storage write failed', e)
  }
}

// Called when DL hits Generate — promotes current → prev, clears current responses
export function archiveScores(plName) {
  const state = loadState(plName)
  if (!state) return
  const teams = { ...state.teams }
  for (const name of Object.keys(teams)) {
    const t = teams[name]
    if (t.current) {
      teams[name] = { ...t, prev: t.current, current: null, surveyResponses: [] }
    }
  }
  saveState(plName, { ...state, teams, lastGenerated: new Date().toISOString() })
}

// Initialise a team slot if it doesn't exist
export function ensureTeam(plName, teamName) {
  const state = loadState(plName) || { teams: {} }
  if (!state.teams[teamName]) {
    state.teams[teamName] = { prev: null, current: null, surveyResponses: [] }
    saveState(plName, state)
  }
}

// Add a survey response for a team
export function addSurveyResponse(plName, teamName, response) {
  const state = loadState(plName) || { teams: {} }
  if (!state.teams[teamName]) {
    state.teams[teamName] = { prev: null, current: null, surveyResponses: [] }
  }
  const responses = [...(state.teams[teamName].surveyResponses || []), {
    ...response,
    submittedAt: new Date().toISOString()
  }]
  // Recalculate current average
  const avg = (field) => {
    const vals = responses.map(r => r[field]).filter(v => v != null)
    return vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10 : null
  }
  const current = responses.length >= 1 ? {
    happiness:   avg('happiness'),
    teamSpirit:  avg('teamSpirit'),
    belonging:   avg('belonging'),
    respondents: responses.length,
    date: new Date().toISOString()
  } : null

  state.teams[teamName] = { ...state.teams[teamName], surveyResponses: responses, current }
  saveState(plName, state)
  return responses.length
}

// Get scores for a team (for display in report tool)
export function getTeamScores(plName, teamName) {
  const state = loadState(plName)
  if (!state?.teams?.[teamName]) return { prev: null, current: null, respondents: 0 }
  const t = state.teams[teamName]
  return {
    prev:        t.prev,
    current:     t.current,
    respondents: (t.surveyResponses || []).length
  }
}

// Get all teams' scores for the report
export function getAllTeamScores(plName) {
  const state = loadState(plName)
  if (!state?.teams) return {}
  return state.teams
}

// Build survey URL for a team
export function buildSurveyUrl(teamName, plName) {
  const base = window.location.origin
  const params = new URLSearchParams({ team: teamName, pl: plName })
  return `${base}/survey?${params.toString()}`
}

// Trend direction
export function trend(prev, current) {
  if (prev == null || current == null) return 'flat'
  const delta = current - prev
  if (delta > 0.2) return 'up'
  if (delta < -0.2) return 'down'
  return 'flat'
}

export function trendLabel(prev, current) {
  const t = trend(prev, current)
  if (t === 'up')   return '↑'
  if (t === 'down') return '↓'
  return '→'
}

export function trendClass(prev, current) {
  const t = trend(prev, current)
  if (t === 'up')   return 'trend-up'
  if (t === 'down') return 'trend-down'
  return 'trend-flat'
}
