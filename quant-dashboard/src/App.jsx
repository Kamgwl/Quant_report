import { useEffect, useState } from 'react'
import Dashboard from './Dashboard.jsx'

const STORAGE_KEY = 'qd-theme'

export default function App() {
  const [theme, setTheme] = useState(
    () => localStorage.getItem(STORAGE_KEY) || 'dark',
  )

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme)
    localStorage.setItem(STORAGE_KEY, theme)
  }, [theme])

  return (
    <>
      <button
        className="theme-toggle"
        onClick={() => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))}
        aria-label="Toggle light / dark theme"
        title="Toggle light / dark theme"
      >
        {theme === 'dark' ? '☀️ Light' : '🌙 Dark'}
      </button>
      <Dashboard theme={theme} />
    </>
  )
}
