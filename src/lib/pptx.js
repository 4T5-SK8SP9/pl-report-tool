// ─── Colour constants (no # prefix — PptxGenJS requirement) ──────────────────
const C = {
  navyDark: '0D1440', navy: '1E2761', ice: 'CADCFC', accent: '4FC3F7',
  white: 'FFFFFF', muted: '8A9BC4', slate: '475569', light: 'F7F9FF',
  border: 'E2E8F0', hint: '94A3B8',
  green: '16A34A', greenBg: 'F0FDF4', greenBorder: 'BBF7D0',
  amber: 'D97706', amberBg: 'FFFBEB', amberBorder: 'FDE68A',
  red:   'DC2626', redBg:   'FEF2F2', redBorder:   'FECACA',
  teal:  '0F766E', tealBg:  'F0FDFA', tealBorder:  '99F6E4',
}

const RAG_MAP = {
  green: { fill: C.greenBg, border: C.greenBorder, text: C.green,  label: 'On track'  },
  amber: { fill: C.amberBg, border: C.amberBorder, text: C.amber,  label: 'At risk'   },
  red:   { fill: C.redBg,   border: C.redBorder,   text: C.red,    label: 'Off track' },
}

function ragFromPct(pct) {
  if (pct == null) return 'amber'
  if (pct >= 70)   return 'green'
  if (pct >= 40)   return 'amber'
  return 'red'
}

function trendArrow(prev, curr) {
  if (prev == null || curr == null) return '–'
  const d = curr - prev
  if (d > 0.2)  return '↑'
  if (d < -0.2) return '↓'
  return '→'
}

function trendColor(prev, curr) {
  if (prev == null || curr == null) return C.hint
  const d = curr - prev
  if (d > 0.2)  return C.green
  if (d < -0.2) return C.red
  return C.hint
}

const makeShadow = () => ({
  type: 'outer', blur: 6, offset: 2, angle: 135, color: '000000', opacity: 0.08
})

function header(pptx, slide, title, sub, bgColor = C.navy) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.7,
    fill: { color: bgColor }, line: { color: bgColor }
  })
  slide.addText(title, {
    x: 0.4, y: 0, w: 7.8, h: 0.7,
    fontSize: 19, color: C.white, bold: true, fontFace: 'Calibri', valign: 'middle', margin: 0
  })
  if (sub) slide.addText(sub, {
    x: 7.8, y: 0, w: 2.0, h: 0.7,
    fontSize: 9, color: C.ice, fontFace: 'Calibri', align: 'right', valign: 'middle', margin: 0
  })
}

function progressBar(slide, x, y, w, pct, color) {
  slide.addShape(slide._pptx?.shapes?.RECTANGLE || 'rect', { x, y, w, h: 0.14, fill: { color: C.border }, line: { color: C.border } })
  if (pct > 0) {
    slide.addShape(slide._pptx?.shapes?.RECTANGLE || 'rect', {
      x, y, w: Math.max(0.04, w * pct / 100), h: 0.14,
      fill: { color: color }, line: { color: color }
    })
  }
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function buildPptx(liveData, manual, wellnessData, timestamp) {
  const pptx   = new window.PptxGenJS()
  pptx.layout  = 'LAYOUT_16x9'
  pptx.title   = `${liveData.productLine} — PL Report`

  const PL    = liveData.productLine || 'Product Line'
  const inits = liveData.initiatives || []
  const teams = liveData.teams       || []
  const imps  = liveData.impediments || []

  // Helper: attach pptx ref for shapes inside loops
  const addBar = (slide, x, y, w, pct, color) => {
    slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.14, fill: { color: C.border }, line: { color: C.border } })
    if (pct > 0) slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w: Math.max(0.04, w * pct / 100), h: 0.14,
      fill: { color: color }, line: { color: color }
    })
  }

  // ── Slide 1: Cover ──────────────────────────────────────────────────────────
  const s1 = pptx.addSlide()
  s1.background = { color: C.navyDark }
  s1.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } })
  s1.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 4.7, w: 10, h: 0.925, fill: { color: C.navy }, line: { color: C.navy } })
  s1.addText(PL.toUpperCase(), { x: 0.42, y: 1.05, w: 9, h: 0.45, fontSize: 11, color: C.accent, bold: true, charSpacing: 6, fontFace: 'Calibri', margin: 0 })
  s1.addText('Product Line Report', { x: 0.42, y: 1.55, w: 9, h: 1.1, fontSize: 42, color: C.white, bold: true, fontFace: 'Calibri', margin: 0 })
  s1.addText('Strategic progress  ·  Team health  ·  Impediments', { x: 0.42, y: 2.75, w: 8.5, h: 0.45, fontSize: 14, color: C.ice, fontFace: 'Calibri', italic: true, margin: 0 })
  s1.addText(`Generated: ${timestamp}`, { x: 0.42, y: 4.8, w: 6, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri', margin: 0 })
  s1.addText(`Delivery Lead: ${manual.dlName || '—'}`, { x: 0.42, y: 5.13, w: 6, h: 0.28, fontSize: 10, color: C.muted, fontFace: 'Calibri', margin: 0 })

  // ── Slide 2: Initiatives overview ───────────────────────────────────────────
  const s2 = pptx.addSlide()
  s2.background = { color: C.light }
  header(pptx, s2, 'Initiative progress', PL)

  const initW = inits.length <= 2 ? 4.5 : inits.length === 3 ? 2.9 : 2.15
  const initGap = 0.15

  inits.slice(0, 6).forEach((ini, i) => {
    const col = i % (inits.length <= 2 ? 2 : inits.length === 3 ? 3 : 4)
    const row = Math.floor(i / (inits.length <= 2 ? 2 : inits.length === 3 ? 3 : 4))
    const x   = 0.4 + col * (initW + initGap)
    const y   = 0.85 + row * 2.2
    const rag = ini.ragStatus || ragFromPct(ini.percentComplete)
    const rc  = RAG_MAP[rag]
    const pct = ini.percentComplete ?? 0
    const ec  = (ini.epics || []).length
    const ed  = (ini.epics || []).filter(e => (e.percentComplete || 0) >= 100).length

    s2.addShape(pptx.shapes.RECTANGLE, { x, y, w: initW, h: 1.95, fill: { color: rc.fill }, line: { color: rc.border }, shadow: makeShadow() })
    s2.addShape(pptx.shapes.RECTANGLE, { x, y, w: initW, h: 0.055, fill: { color: rc.text }, line: { color: rc.text } })
    s2.addText(ini.key || '', { x: x + 0.1, y: y + 0.1, w: initW - 0.2, h: 0.22, fontSize: 9, color: rc.text, bold: true, fontFace: 'Calibri', margin: 0 })
    s2.addText(ini.name || 'Unnamed', { x: x + 0.1, y: y + 0.32, w: initW - 0.2, h: 0.42, fontSize: 12, color: '1E293B', bold: true, fontFace: 'Calibri', margin: 0 })
    if (ini.goal) s2.addText(ini.goal.substring(0, 90), { x: x + 0.1, y: y + 0.75, w: initW - 0.2, h: 0.38, fontSize: 9, color: C.slate, fontFace: 'Calibri', margin: 0 })
    addBar(s2, x + 0.1, y + 1.22, initW - 0.2, pct, rc.text)
    s2.addText(`${pct}%`, { x: x + 0.1, y: y + 1.42, w: (initW - 0.2) / 2, h: 0.24, fontSize: 11, color: rc.text, bold: true, fontFace: 'Calibri', margin: 0 })
    s2.addText(rc.label, { x: x + initW - 1.1, y: y + 1.42, w: 0.95, h: 0.24, fontSize: 10, color: rc.text, align: 'right', fontFace: 'Calibri', margin: 0 })
    s2.addText(`${ed}/${ec} epics complete`, { x: x + 0.1, y: y + 1.68, w: initW - 0.2, h: 0.2, fontSize: 9, color: C.hint, fontFace: 'Calibri', margin: 0 })
  })

  if (manual.initiativeCommentary) {
    const usedRows = Math.ceil(Math.min(inits.length, 6) / (inits.length <= 3 ? inits.length : 4))
    const cy = 0.85 + usedRows * 2.2
    if (cy < 5.1) {
      s2.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: cy, w: 9.2, h: 5.3 - cy, fill: { color: C.white }, line: { color: C.border }, shadow: makeShadow() })
      s2.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: cy, w: 0.055, h: 5.3 - cy, fill: { color: C.accent }, line: { color: C.accent } })
      s2.addText('Delivery Lead commentary', { x: 0.55, y: cy + 0.1, w: 8.9, h: 0.24, fontSize: 9, color: C.muted, bold: true, fontFace: 'Calibri', margin: 0 })
      s2.addText(manual.initiativeCommentary, { x: 0.55, y: cy + 0.36, w: 8.9, h: 5.3 - cy - 0.48, fontSize: 12, color: '1E293B', fontFace: 'Calibri', valign: 'top', margin: 0 })
    }
  }

  // ── Slides 3+: Epic breakdown per initiative ────────────────────────────────
  inits.slice(0, 5).forEach((ini) => {
    const s = pptx.addSlide()
    s.background = { color: C.light }
    const rag = ini.ragStatus || ragFromPct(ini.percentComplete)
    const hc  = rag === 'green' ? C.teal : rag === 'red' ? 'B91C1C' : '92400E'
    header(pptx, s, ini.name, `${ini.key}  ·  Epic breakdown`, hc)

    const epics = (ini.epics || []).slice(0, 8)
    const colW  = 4.55

    epics.forEach((ep, ei) => {
      const col = ei % 2
      const row = Math.floor(ei / 2)
      const x   = 0.4 + col * (colW + 0.1)
      const y   = 0.88 + row * 1.1
      const er  = ep.ragStatus || ragFromPct(ep.percentComplete)
      const ec2 = RAG_MAP[er]
      const pct = ep.percentComplete ?? 0

      s.addShape(pptx.shapes.RECTANGLE, { x, y, w: colW, h: 0.95, fill: { color: ec2.fill }, line: { color: ec2.border }, shadow: makeShadow() })
      s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 0.055, h: 0.95, fill: { color: ec2.text }, line: { color: ec2.text } })
      s.addText(`${ep.key || ''}  ·  ${ep.team || ''}`, { x: x + 0.13, y: y + 0.07, w: colW - 0.22, h: 0.22, fontSize: 9, color: ec2.text, bold: true, fontFace: 'Calibri', margin: 0 })
      s.addText(ep.name || '', { x: x + 0.13, y: y + 0.27, w: colW - 0.22, h: 0.28, fontSize: 12, color: '1E293B', fontFace: 'Calibri', margin: 0 })
      addBar(s, x + 0.13, y + 0.63, colW - 0.95, pct, ec2.text)
      s.addText(`${pct}%`, { x: x + colW - 0.72, y: y + 0.56, w: 0.58, h: 0.25, fontSize: 11, color: ec2.text, bold: true, align: 'right', fontFace: 'Calibri', margin: 0 })
      if (ep.dueDate) s.addText(`Due: ${ep.dueDate}`, { x: x + 0.13, y: y + 0.76, w: colW - 0.22, h: 0.17, fontSize: 9, color: C.hint, fontFace: 'Calibri', margin: 0 })
    })
  })

  // ── Slide n: Team health heatmap ────────────────────────────────────────────
  const sTH = pptx.addSlide()
  sTH.background = { color: C.light }
  header(pptx, sTH, 'Team health overview', PL, C.teal)

  // Columns: Team | Velocity | Reliability | Capacity | Happiness | Spirit | Belonging | Overall
  const thCols = ['Velocity', 'Reliability', 'Capacity', 'Happiness', 'Spirit', 'Belonging', 'Overall']
  const thXs   = [2.05, 3.05, 4.05, 5.18, 6.18, 7.18, 8.55]
  const thW    = 0.92

  thCols.forEach((c, ci) => {
    sTH.addText(c, { x: thXs[ci], y: 0.75, w: thW, h: 0.3, fontSize: 9, color: C.muted, bold: true, align: 'center', fontFace: 'Calibri', margin: 0 })
  })
  sTH.addText('Team', { x: 0.4, y: 0.75, w: 1.55, h: 0.3, fontSize: 9, color: C.muted, bold: true, fontFace: 'Calibri', margin: 0 })

  const minR = (ppt) => ragFromPct(ppt)

  const rowH2 = Math.min(0.44, 4.3 / Math.max(teams.length, 1))

  teams.slice(0, 11).forEach((t, ti) => {
    const y   = 1.1 + ti * (rowH2 + 0.03)
    const bg  = ti % 2 === 0 ? C.white : 'F8FAFC'
    const ws  = wellnessData[t.name] || {}
    const prevW = ws.prev
    const currW = ws.current
    const resp  = ws.respondents || 0
    const MIN_RESP = 5

    sTH.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 9.1, h: rowH2, fill: { color: bg }, line: { color: C.border } })
    sTH.addText(t.name, { x: 0.48, y, w: 1.5, h: rowH2, fontSize: 11, color: '1E293B', fontFace: 'Calibri', valign: 'middle', margin: 0 })

    // Jira-sourced cells
    const jiraCells = [
      { val: t.velocityAvg   ? `${t.velocityAvg}pts`       : '—', rag: t.velocityAvg   ? ragFromPct(Math.min(100, t.velocityAvg / 80 * 100)) : 'amber' },
      { val: t.sprintCompletionRate ? `${t.sprintCompletionRate}%` : '—', rag: t.sprintCompletionRate ? ragFromPct(t.sprintCompletionRate) : 'amber' },
      { val: t.capacity      ? `${t.capacity}%`            : '—', rag: t.capacity      ? ragFromPct(t.capacity) : 'amber' },
    ]
    jiraCells.forEach((cell, ci) => {
      const rc = RAG_MAP[cell.rag]
      sTH.addShape(pptx.shapes.RECTANGLE, { x: thXs[ci] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fill: { color: rc.fill }, line: { color: rc.border } })
      sTH.addText(cell.val, { x: thXs[ci] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fontSize: 10, color: rc.text, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
    })

    // Wellness cells (happiness, spirit, belonging) — with trend
    const dims = ['happiness', 'teamSpirit', 'belonging']
    dims.forEach((dim, di) => {
      const ci    = 3 + di
      const prev  = prevW?.[dim]
      const curr  = resp >= MIN_RESP ? currW?.[dim] : null
      const arrow = trendArrow(prev, curr)
      const arrowC = trendColor(prev, curr)
      const displayVal = curr != null ? `${curr}` : resp > 0 && resp < MIN_RESP ? `(${resp})` : '—'
      const ragVal = curr != null ? ragFromPct(curr * 20) : 'amber'
      const rc     = RAG_MAP[ragVal]

      sTH.addShape(pptx.shapes.RECTANGLE, { x: thXs[ci] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fill: { color: rc.fill }, line: { color: rc.border } })
      if (prev != null && curr != null) {
        sTH.addText(displayVal, { x: thXs[ci] + 0.03, y: y + 0.03, w: (thW - 0.06) * 0.6, h: rowH2 - 0.06, fontSize: 10, color: rc.text, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
        sTH.addText(arrow, { x: thXs[ci] + (thW - 0.06) * 0.6 + 0.03, y: y + 0.03, w: (thW - 0.06) * 0.38, h: rowH2 - 0.06, fontSize: 11, color: arrowC, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
      } else {
        sTH.addText(displayVal, { x: thXs[ci] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fontSize: 10, color: curr != null ? rc.text : C.hint, bold: curr != null, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
      }
    })

    // Overall RAG
    const allRags = [
      jiraCells[0].rag, jiraCells[1].rag, jiraCells[2].rag,
      resp >= MIN_RESP && currW?.happiness  ? ragFromPct(currW.happiness  * 20) : 'amber',
      resp >= MIN_RESP && currW?.teamSpirit ? ragFromPct(currW.teamSpirit * 20) : 'amber',
      resp >= MIN_RESP && currW?.belonging  ? ragFromPct(currW.belonging  * 20) : 'amber',
    ]
    const reds   = allRags.filter(r => r === 'red').length
    const ambers = allRags.filter(r => r === 'amber').length
    const overallRag = reds >= 2 ? 'red' : reds >= 1 || ambers >= 3 ? 'amber' : 'green'
    const orc = RAG_MAP[overallRag]
    sTH.addShape(pptx.shapes.RECTANGLE, { x: thXs[6] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fill: { color: orc.fill }, line: { color: orc.border } })
    sTH.addText(orc.label, { x: thXs[6] + 0.03, y: y + 0.03, w: thW - 0.06, h: rowH2 - 0.06, fontSize: 9, color: orc.text, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
  })

  // Legend
  const ly = 5.18
  ;[['green', 'On track'], ['amber', 'Watch'], ['red', 'At risk']].forEach(([r, lbl], li) => {
    const rc = RAG_MAP[r]
    sTH.addShape(pptx.shapes.RECTANGLE, { x: 0.4 + li * 1.1, y: ly, w: 0.95, h: 0.24, fill: { color: rc.fill }, line: { color: rc.border } })
    sTH.addText(lbl, { x: 0.4 + li * 1.1, y: ly, w: 0.95, h: 0.24, fontSize: 9, color: rc.text, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
  })
  sTH.addText('Wellness scores: prev → current  ·  (n) = awaiting min 5 responses  ·  → flat  ↑ improving  ↓ declining', {
    x: 3.8, y: ly + 0.02, w: 5.8, h: 0.22, fontSize: 8, color: C.hint, fontFace: 'Calibri', italic: true, margin: 0
  })

  // ── Slide n+1: Capacity chart ──────────────────────────────────────────────
  const sCap = pptx.addSlide()
  sCap.background = { color: C.light }
  header(pptx, sCap, 'Capacity overview', PL, C.teal)

  if (teams.length > 0) {
    const capVals   = teams.slice(0, 12).map(t => t.capacity ?? 75)
    const capLabels = teams.slice(0, 12).map(t => t.name.length > 11 ? t.name.slice(0, 11) + '…' : t.name)
    const capColors = teams.slice(0, 12).map(t => ragFromPct(t.capacity) === 'green' ? '22C55E' : ragFromPct(t.capacity) === 'amber' ? 'F59E0B' : 'EF4444')
    sCap.addChart(pptx.charts.BAR,
      [{ name: 'Capacity %', labels: capLabels, values: capVals }],
      {
        x: 0.4, y: 0.85, w: 6.3, h: 4.5, barDir: 'col',
        chartColors: capColors,
        chartArea: { fill: { color: C.white }, roundedCorners: true },
        catAxisLabelColor: '64748B', valAxisLabelColor: '64748B', valAxisMaxVal: 100,
        valGridLine: { color: C.border, size: 0.5 }, catGridLine: { style: 'none' },
        showValue: true, dataLabelColor: '1E293B', dataLabelFontSize: 10,
        showLegend: false, showTitle: false,
      }
    )
  }
  if (manual.capacityCommentary) {
    sCap.addShape(pptx.shapes.RECTANGLE, { x: 6.9, y: 0.9, w: 2.7, h: 4.3, fill: { color: C.white }, line: { color: C.border }, shadow: makeShadow() })
    sCap.addShape(pptx.shapes.RECTANGLE, { x: 6.9, y: 0.9, w: 0.055, h: 4.3, fill: { color: C.accent }, line: { color: C.accent } })
    sCap.addText('Capacity notes', { x: 7.05, y: 1.0, w: 2.5, h: 0.26, fontSize: 9, color: C.muted, bold: true, fontFace: 'Calibri', margin: 0 })
    sCap.addText(manual.capacityCommentary, { x: 7.05, y: 1.28, w: 2.5, h: 3.8, fontSize: 11, color: '1E293B', fontFace: 'Calibri', valign: 'top', margin: 0 })
  }

  // ── Slide n+2: Impediments ─────────────────────────────────────────────────
  const sImp = pptx.addSlide()
  sImp.background = { color: C.light }
  header(pptx, sImp, "What's holding us back", PL, 'B91C1C')

  imps.slice(0, 6).forEach((imp, ii) => {
    const col = ii % 2
    const row = Math.floor(ii / 2)
    const x   = 0.4 + col * 4.75
    const y   = 0.88 + row * 1.55

    sImp.addShape(pptx.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.4, fill: { color: C.redBg }, line: { color: C.redBorder }, shadow: makeShadow() })
    sImp.addShape(pptx.shapes.RECTANGLE, { x, y, w: 0.055, h: 1.4, fill: { color: C.red }, line: { color: C.red } })
    sImp.addText(`${imp.key || 'IMP'}  ·  ${imp.team || ''}`, { x: x + 0.13, y: y + 0.09, w: 4.2, h: 0.22, fontSize: 9, color: C.red, bold: true, fontFace: 'Calibri', margin: 0 })
    sImp.addText(imp.summary || '', { x: x + 0.13, y: y + 0.3, w: 4.2, h: 0.42, fontSize: 12, color: '1E293B', fontFace: 'Calibri', margin: 0 })
    if (imp.blockedSince) sImp.addText(`Blocked since: ${imp.blockedSince}`, { x: x + 0.13, y: y + 0.76, w: 3.0, h: 0.2, fontSize: 9, color: C.hint, fontFace: 'Calibri', margin: 0 })
    if (imp.stakeholderAsk) {
      sImp.addShape(pptx.shapes.RECTANGLE, { x: x + 0.13, y: y + 1.0, w: 4.2, h: 0.3, fill: { color: C.amberBg }, line: { color: C.amberBorder } })
      sImp.addText(`Ask: ${imp.stakeholderAsk}`, { x: x + 0.22, y: y + 1.0, w: 4.0, h: 0.3, fontSize: 9, color: C.amber, fontFace: 'Calibri', valign: 'middle', margin: 0 })
    }
  })

  // ── Slide n+3: Where you can help ──────────────────────────────────────────
  const sHelp = pptx.addSlide()
  sHelp.background = { color: C.light }
  header(pptx, sHelp, 'Where you can help', PL, C.navy)

  const asks = (manual.stakeholderAsks || '').split('\n').filter(Boolean).slice(0, 5)
  asks.forEach((ask, ai) => {
    const y = 0.88 + ai * 0.88
    sHelp.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.72, fill: { color: 'EFF6FF' }, line: { color: 'BFDBFE' }, shadow: makeShadow() })
    sHelp.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 0.055, h: 0.72, fill: { color: C.accent }, line: { color: C.accent } })
    sHelp.addText(`${ai + 1}`, { x: 0.55, y, w: 0.4, h: 0.72, fontSize: 16, color: C.accent, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0 })
    sHelp.addText(ask, { x: 1.1, y, w: 8.3, h: 0.72, fontSize: 13, color: '1E293B', fontFace: 'Calibri', valign: 'middle', margin: 0 })
  })
  if (asks.length === 0) {
    sHelp.addText('No specific asks this period.', { x: 0.4, y: 1.6, w: 9.2, h: 0.5, fontSize: 14, color: C.hint, italic: true, fontFace: 'Calibri', margin: 0 })
  }
  if (manual.generalNotes) {
    const ny = 0.88 + asks.length * 0.88 + 0.2
    if (ny < 5.0) {
      sHelp.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: ny, w: 9.2, h: 5.25 - ny, fill: { color: C.white }, line: { color: C.border }, shadow: makeShadow() })
      sHelp.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: ny, w: 0.055, h: 5.25 - ny, fill: { color: C.muted }, line: { color: C.muted } })
      sHelp.addText('Additional context', { x: 0.55, y: ny + 0.1, w: 8.9, h: 0.24, fontSize: 9, color: C.muted, bold: true, fontFace: 'Calibri', margin: 0 })
      sHelp.addText(manual.generalNotes, { x: 0.55, y: ny + 0.36, w: 8.9, h: 5.25 - ny - 0.48, fontSize: 12, color: '1E293B', fontFace: 'Calibri', valign: 'top', margin: 0 })
    }
  }

  // ── Slide last: Closing ────────────────────────────────────────────────────
  const sEnd = pptx.addSlide()
  sEnd.background = { color: C.navyDark }
  sEnd.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } })
  sEnd.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 4.7, w: 10, h: 0.925, fill: { color: C.navy }, line: { color: C.navy } })
  sEnd.addText('Thank you', { x: 0.42, y: 1.6, w: 9, h: 1.0, fontSize: 42, color: C.white, bold: true, fontFace: 'Calibri', margin: 0 })
  sEnd.addText('Questions & discussion', { x: 0.42, y: 2.7, w: 9, h: 0.5, fontSize: 18, color: C.ice, italic: true, fontFace: 'Calibri', margin: 0 })
  sEnd.addText(`${PL}  ·  ${manual.dlName || 'Delivery Lead'}`, { x: 0.42, y: 3.5, w: 9, h: 0.35, fontSize: 12, color: C.accent, fontFace: 'Calibri', margin: 0 })
  sEnd.addText(`Report generated: ${timestamp}`, { x: 0.42, y: 4.8, w: 6, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri', margin: 0 })

  const fn = `${PL.replace(/\s+/g, '-')}-PL-Report-${timestamp.replace(/[: ,/]/g, '-')}.pptx`
  return pptx.writeFile({ fileName: fn })
}
