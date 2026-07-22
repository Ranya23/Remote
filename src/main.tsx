import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './index.css'
import { supabaseConfigError } from './supabaseClient'

// ==========================================
// THE MOBILE WI-FI BYPASS
// Chrome disables 'crypto' on local IP addresses, which instantly crashes Supabase.
// When Supabase crashes, it stops Vite from loading your CSS and buttons!
// This creates a local fallback so your phone portal can connect perfectly.
// ==========================================
if (typeof window !== 'undefined' && !window.crypto) {
  Object.defineProperty(window, 'crypto', {
    value: {
      getRandomValues: (arr: any) => {
        for (let i = 0; i < arr.length; i++) {
          arr[i] = Math.floor(Math.random() * 256);
        }
        return arr;
      }
    }
  });
}

class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { error: Error | null }> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('🚨 App crashed during render:', error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ background: '#111', color: '#fff', minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: 'sans-serif', textAlign: 'center' }}>
          <h1 style={{ color: '#ef4444', marginBottom: 12 }}>App failed to start</h1>
          <p style={{ color: '#9ca3af', maxWidth: 480 }}>{this.state.error.message}</p>
          <p style={{ color: '#6b7280', fontSize: 12, marginTop: 16 }}>Check the browser console for details.</p>
        </div>
      );
    }
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <ErrorBoundary>
      {supabaseConfigError && (
        <div style={{ background: '#7f1d1d', color: '#fff', padding: '8px 16px', textAlign: 'center', fontSize: 13, fontWeight: 700 }}>
          🚨 Supabase not configured — VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY missing at build time.
        </div>
      )}
      <App />
    </ErrorBoundary>
  </React.StrictMode>,
)