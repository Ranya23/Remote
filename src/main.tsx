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
// When a Google sign-in (or an email confirmation / password-reset link)
// fails on Supabase's side, it redirects back with
// "#error=...&error_code=...&error_description=..." appended to the URL -
// this is separate from the "?code=..." success case AuthRedirectHandler.tsx
// handles, and Supabase always uses this hash format for errors regardless
// of the PKCE flowType set in supabaseClient.ts. Since this app uses
// HashRouter, that hash is exactly what react-router reads as the route to
// match, so an error redirect showed up as "No routes matched location
// '/error=server_error&...'" in the console and just left the page blank -
// the actual error (e.g. "Database error saving new user") never reached
// the person trying to sign in. This runs before the router ever mounts,
// rewrites the hash to a real route (#/account), and stashes the message
// so Account.tsx can show it once it loads.
if (typeof window !== 'undefined') {
  const rawHash = window.location.hash.startsWith('#') ? window.location.hash.slice(1) : window.location.hash;
  if (rawHash.startsWith('error=') || rawHash.includes('&error=')) {
    const params = new URLSearchParams(rawHash);
    const description = params.get('error_description');
    try {
      sessionStorage.setItem('nextslide_auth_error', description || 'Sign-in failed. Please try again.');
    } catch { /* sessionStorage unavailable (e.g. private mode) - the redirect still gets cleaned up below */ }
    window.location.hash = '#/account';
  }
}

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