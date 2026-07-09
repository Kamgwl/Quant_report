import Dashboard from './Dashboard.jsx'

// Dashboard is self-contained: it manages its own light/dark theme (persisted to
// localStorage under "qd-theme") and renders its own theme toggle in the header.
export default function App() {
  return <Dashboard />
}
