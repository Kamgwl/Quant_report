import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Dev + preview servers bind all interfaces on 8445 so the dashboard is
// reachable from other machines on the network (e.g. http://<this-ip>:8445).
export default defineConfig({
  plugins: [react()],
  server: { host: '0.0.0.0', port: 8445, strictPort: true },
  preview: { host: '0.0.0.0', port: 8445, strictPort: true },
})
