import ReportGenerator from './views/ReportGenerator.jsx'

// Survey routing handled inside ReportGenerator via ?member=Name URL param
// Single-page app — compatible with GitHub Pages static hosting
export default function App() {
  return <ReportGenerator />
}
