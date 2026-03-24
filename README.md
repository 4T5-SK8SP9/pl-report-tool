# PL Report Tool — v0.2

Stakeholder report generator for the Agile Center of Excellence (and any Product Line).
Built with React + Vite. Exports to PowerPoint (.pptx). Hosted on GitHub Pages.

---

## Deploy to GitHub Pages (no Node.js needed on your machine)

### Step 1 — Create repo on GitHub
1. Go to https://github.com/new
2. Name it `pl-report-tool`
3. Set to **Private**
4. Click **Create repository**

### Step 2 — Upload files via browser
1. On the new repo page click **"uploading an existing file"**
2. Unzip `pl-report-tool.zip`
3. Drag ALL files and folders into the upload area (including the hidden `.github` folder)
4. Write commit message: `Initial commit v0.2`
5. Click **Commit changes**

### Step 3 — Enable GitHub Pages
1. Repo **Settings** → **Pages** → Source: **GitHub Actions** → Save

### Step 4 — Wait ~2 minutes
Your app is live at:
```
https://YOUR-USERNAME.github.io/pl-report-tool/
```

### Survey links (mobile)
```
https://YOUR-USERNAME.github.io/pl-report-tool/?member=Jacob%20Kingo
```

---

## Local dev (requires Node.js)
```bash
npm install && npm run dev
```

---

## Slide structure

| # | Slide | Source |
|---|---|---|
| 1 | Cover | Auto |
| 2 | Portfolio overview | Live Jira + commentary |
| 3 | Active epics | Live Jira |
| 4 | Flow metrics | Live cycle time + empty velocity/reliability |
| 5 | Dependencies | Live Jira |
| 6 | Throughput + Capacity | Live + manual |
| 7 | Wellness | Pulse survey |
| 8 | Risks | Manual entry |
| 9 | Stakeholder asks | Manual entry |
| 10 | Closing | Auto |
