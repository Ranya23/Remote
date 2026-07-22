import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { supabase } from './supabaseClient';
import { useAuth } from './AuthContext';

export interface SavedItem {
  id: string;
  user_id: string;
  kind: 'lesson' | 'quiz';
  title: string;
  file_id?: string | null;
  file_type?: string | null;
  size_bytes?: number;
  questions?: any[] | null;
  created_at?: string;
}

// Where Supabase should send people back to after clicking an email-
// verification, "Continue with Google", or password-reset link. Deliberately
// the plain origin (no #hash) - see AuthRedirectHandler.tsx for why.
function authRedirectUrl() {
  return `${window.location.origin}${window.location.pathname}`;
}

export function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// Called from FileUpload.tsx (after a lesson is saved) and Present.tsx
// (after saving a quiz) - a thin, best-effort wrapper so neither flow has to
// know or care about auth details. Returns a result object (rather than
// throwing) so FileUpload.tsx can show a friendly message if the server-side
// limit trigger ends up rejecting the insert.
export async function recordSavedItem(
  item: Omit<SavedItem, 'id' | 'user_id'>
): Promise<{ ok: boolean; reason?: 'not_logged_in' | 'presentation_limit' | 'storage_limit' | 'unknown' }> {
  const { data } = await supabase.auth.getUser();
  const user = data.user;
  if (!user) return { ok: false, reason: 'not_logged_in' };
  try {
    const { error } = await supabase.from('saved_items').insert({
      id: `si_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`,
      user_id: user.id,
      ...item,
    });
    if (error) {
      const msg = error.message || '';
      if (msg.includes('PRESENTATION_LIMIT_REACHED')) return { ok: false, reason: 'presentation_limit' };
      if (msg.includes('STORAGE_LIMIT_REACHED')) return { ok: false, reason: 'storage_limit' };
      throw error;
    }
    return { ok: true };
  } catch (err) {
    console.warn('⚠️ Could not save to account (has supabase_migration_auth.sql been run yet?):', err);
    return { ok: false, reason: 'unknown' };
  }
}

export default function AccountPage() {
  const navigate = useNavigate();
  const { user, profile, loading: authLoading, usage, refreshUsage, refreshProfile, signOut } = useAuth();

  const [items, setItems] = useState<SavedItem[]>([]);
  const [loadingItems, setLoadingItems] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  // --- Auth form state -----------------------------------------------------
  const [mode, setMode] = useState<'login' | 'signup' | 'forgot'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [fullName, setFullName] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [verifyEmailSent, setVerifyEmailSent] = useState<string | null>(null); // email just signed up with
  const [resetEmailSent, setResetEmailSent] = useState(false);

  // --- Password-recovery state ----------------------------------------------
  const [inRecovery, setInRecovery] = useState(false);
  const [newPassword, setNewPassword] = useState('');

  useEffect(() => {
    const { data: sub } = supabase.auth.onAuthStateChange((event) => {
      if (event === 'PASSWORD_RECOVERY') setInRecovery(true);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  const loadItems = async (userId: string) => {
    setLoadingItems(true);
    setError('');
    try {
      const { data, error: err } = await supabase
        .from('saved_items')
        .select('*')
        .eq('user_id', userId)
        .order('created_at', { ascending: false });
      if (err) throw err;
      setItems(data || []);
    } catch (err) {
      setError("Couldn't load your saved items - has supabase_migration_auth.sql been run yet?");
      console.warn(err);
    } finally {
      setLoadingItems(false);
    }
  };

  useEffect(() => {
    if (user) loadItems(user.id);
    else setItems([]);
  }, [user]);

  // --- Auth actions ----------------------------------------------------------
  const handleSignup = async () => {
    if (!email.trim() || !password) return;
    setSubmitting(true);
    setError('');
    try {
      const { data, error: err } = await supabase.auth.signUp({
        email: email.trim(),
        password,
        options: {
          emailRedirectTo: authRedirectUrl(),
          data: fullName.trim() ? { full_name: fullName.trim() } : undefined,
        },
      });
      if (err) throw err;
      if (!data.session) {
        // Confirm-email is on (recommended) - no session until they click the link.
        setVerifyEmailSent(email.trim());
      }
    } catch (err: any) {
      setError(err.message || 'Could not create your account.');
    } finally {
      setSubmitting(false);
    }
  };

  const handleLogin = async () => {
    if (!email.trim() || !password) return;
    setSubmitting(true);
    setError('');
    try {
      const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      if (err) throw err;
    } catch (err: any) {
      setError(
        err.message === 'Email not confirmed'
          ? 'Please verify your email first - check your inbox for the confirmation link.'
          : err.message || 'Could not log in.'
      );
    } finally {
      setSubmitting(false);
    }
  };

  const handleGoogleLogin = async () => {
    setError('');
    const { error: err } = await supabase.auth.signInWithOAuth({
      provider: 'google',
      options: { redirectTo: authRedirectUrl() },
    });
    if (err) setError(err.message || 'Could not start Google sign-in.');
  };

  const handleForgotPassword = async () => {
    if (!email.trim()) return;
    setSubmitting(true);
    setError('');
    try {
      const { error: err } = await supabase.auth.resetPasswordForEmail(email.trim(), {
        redirectTo: authRedirectUrl(),
      });
      if (err) throw err;
      setResetEmailSent(true);
    } catch (err: any) {
      setError(err.message || 'Could not send reset email.');
    } finally {
      setSubmitting(false);
    }
  };

  const handleSetNewPassword = async () => {
    if (newPassword.length < 6) {
      setError('Password must be at least 6 characters.');
      return;
    }
    setSubmitting(true);
    setError('');
    try {
      const { error: err } = await supabase.auth.updateUser({ password: newPassword });
      if (err) throw err;
      setInRecovery(false);
      setNewPassword('');
      setInfo('Password updated - you\u2019re all set.');
    } catch (err: any) {
      setError(err.message || 'Could not update your password.');
    } finally {
      setSubmitting(false);
    }
  };

  const handleLogout = async () => {
    await signOut();
    setItems([]);
  };

  const openLesson = (item: SavedItem) => {
    if (!item.file_id) return;
    const url = `${window.location.origin}${window.location.pathname}#/present/${item.file_id}`;
    window.open(url, '_blank');
  };

  const openQuiz = (item: SavedItem) => {
    const lessons = items.filter((i) => i.kind === 'lesson' && i.file_id);
    if (!lessons.length) {
      setError('Save a presentation first - a quiz needs one to run inside.');
      return;
    }
    const target = lessons.length === 1 ? lessons[0] : lessons.find((l) => confirm(`Attach "${item.title}" to "${l.title}"?`));
    const chosen = target || lessons[0];
    const url = `${window.location.origin}${window.location.pathname}#/present/${chosen.file_id}`;
    const win = window.open(url, '_blank');
    if (win) {
      const token = `preset_${Date.now().toString(36)}`;
      sessionStorage.setItem(token, JSON.stringify({ title: item.title, questions: item.questions || [] }));
      win.location.href = `${url}?presetQuiz=${token}`;
    }
  };

  const deleteItem = async (item: SavedItem) => {
    const label = item.kind === 'lesson' ? 'presentation' : 'quiz';
    if (!confirm(`Remove "${item.title}" (${label})?`)) return;
    setItems((prev) => prev.filter((i) => i.id !== item.id));
    try {
      const { error: err } = await supabase.from('saved_items').delete().eq('id', item.id);
      if (err) throw err;
      refreshUsage(); // frees up quota immediately in the UI
    } catch (err) {
      console.warn('⚠️ Could not delete saved item:', err);
    }
  };

  // --- Loading ---------------------------------------------------------------
  if (authLoading) {
    return <div className="flex items-center justify-center w-full mt-24 text-gray-500 text-sm">Loading...</div>;
  }

  // --- Password recovery: show "set a new password" instead of anything else --
  if (inRecovery) {
    return (
      <div className="flex flex-col items-center justify-center w-full max-w-sm mx-auto p-6 mt-20 text-white">
        <h1 className="text-2xl font-bold mb-1 text-center">Set a new password</h1>
        <p className="text-sm text-gray-400 mb-6 text-center">Choose a new password for your account.</p>
        <div className="w-full flex flex-col gap-3">
          <input
            type="password"
            value={newPassword}
            onChange={(e) => setNewPassword(e.target.value)}
            placeholder="New password"
            className="bg-gray-900 border border-gray-700 rounded-lg px-4 py-3 text-sm"
          />
          <button
            onClick={handleSetNewPassword}
            disabled={submitting || newPassword.length < 6}
            className="bg-blue-600 disabled:opacity-30 rounded-lg py-3 text-sm font-bold"
          >
            {submitting ? 'Saving...' : 'Save new password'}
          </button>
        </div>
        {error && <p className="text-xs text-red-400 mt-4 text-center">{error}</p>}
      </div>
    );
  }

  // --- Not logged in: login / signup / forgot --------------------------------
  if (!user) {
    return (
      <div className="flex flex-col items-center justify-center w-full max-w-sm mx-auto p-6 mt-20 text-white">
        <h1 className="text-2xl font-bold mb-1 text-center">My Account</h1>
        <p className="text-sm text-gray-400 mb-6 text-center">Save presentations and quizzes, and open them any time - on any device.</p>

        {verifyEmailSent ? (
          <div className="w-full p-4 rounded-xl bg-emerald-950/40 border border-emerald-700 text-center">
            <p className="text-sm text-emerald-300 mb-1">Almost there!</p>
            <p className="text-xs text-gray-300">
              We sent a verification link to <span className="font-semibold">{verifyEmailSent}</span>. Click it, then come back and log in.
            </p>
            <button onClick={() => { setVerifyEmailSent(null); setMode('login'); }} className="text-xs text-gray-400 mt-4 underline">
              Back to log in
            </button>
          </div>
        ) : mode === 'forgot' ? (
          resetEmailSent ? (
            <div className="w-full p-4 rounded-xl bg-emerald-950/40 border border-emerald-700 text-center">
              <p className="text-sm text-emerald-300 mb-1">Check your email</p>
              <p className="text-xs text-gray-300">We sent a password reset link to <span className="font-semibold">{email}</span>.</p>
              <button onClick={() => { setResetEmailSent(false); setMode('login'); }} className="text-xs text-gray-400 mt-4 underline">
                Back to log in
              </button>
            </div>
          ) : (
            <div className="w-full flex flex-col gap-3">
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Your email"
                className="bg-gray-900 border border-gray-700 rounded-lg px-4 py-3 text-sm"
              />
              <button onClick={handleForgotPassword} disabled={submitting || !email.trim()} className="bg-blue-600 disabled:opacity-30 rounded-lg py-3 text-sm font-bold">
                {submitting ? 'Sending...' : 'Send reset link'}
              </button>
              <button onClick={() => setMode('login')} className="text-xs text-gray-500">← Back to log in</button>
            </div>
          )
        ) : (
          <>
            <div className="flex w-full mb-4 rounded-lg overflow-hidden border border-gray-700">
              <button onClick={() => setMode('login')} className={`flex-1 py-2 text-sm font-bold ${mode === 'login' ? 'bg-blue-600' : 'bg-gray-900 text-gray-400'}`}>Log in</button>
              <button onClick={() => setMode('signup')} className={`flex-1 py-2 text-sm font-bold ${mode === 'signup' ? 'bg-blue-600' : 'bg-gray-900 text-gray-400'}`}>Sign up</button>
            </div>

            <button
              onClick={handleGoogleLogin}
              className="w-full flex items-center justify-center gap-2 bg-white text-gray-900 rounded-lg py-3 text-sm font-bold mb-4"
            >
              <svg width="16" height="16" viewBox="0 0 48 48"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3c-1.6 4.6-6 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.1 8 3l6-6C34.6 5.1 29.6 3 24 3 12.4 3 3 12.4 3 24s9.4 21 21 21 21-9.4 21-21c0-1.4-.1-2.7-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.6 15.9 18.9 13 24 13c3.1 0 5.8 1.1 8 3l6-6C34.6 5.1 29.6 3 24 3c-7.5 0-14 4.2-17.7 10.7z"/><path fill="#4CAF50" d="M24 45c5.5 0 10.4-1.9 14.2-5.1l-6.6-5.4C29.6 36.4 27 37 24 37c-5.3 0-9.7-3.4-11.3-8l-6.6 5.1C9.9 40.7 16.4 45 24 45z"/><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.3-2.2 4.2-4.1 5.6l6.6 5.4C41.5 35.6 45 30.3 45 24c0-1.4-.1-2.7-.4-3.5z"/></svg>
              Continue with Google
            </button>

            <div className="w-full flex items-center gap-3 mb-4">
              <div className="flex-1 h-px bg-gray-700" />
              <span className="text-xs text-gray-500">or</span>
              <div className="flex-1 h-px bg-gray-700" />
            </div>

            <div className="w-full flex flex-col gap-3">
              {mode === 'signup' && (
                <input
                  value={fullName}
                  onChange={(e) => setFullName(e.target.value)}
                  placeholder="Your name (optional)"
                  className="bg-gray-900 border border-gray-700 rounded-lg px-4 py-3 text-sm"
                />
              )}
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Email"
                className="bg-gray-900 border border-gray-700 rounded-lg px-4 py-3 text-sm"
              />
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Password"
                className="bg-gray-900 border border-gray-700 rounded-lg px-4 py-3 text-sm"
              />
              {mode === 'login' ? (
                <button onClick={handleLogin} disabled={submitting || !email.trim() || !password} className="bg-blue-600 disabled:opacity-30 rounded-lg py-3 text-sm font-bold">
                  {submitting ? 'Logging in...' : 'Log in'}
                </button>
              ) : (
                <button onClick={handleSignup} disabled={submitting || !email.trim() || password.length < 6} className="bg-blue-600 disabled:opacity-30 rounded-lg py-3 text-sm font-bold">
                  {submitting ? 'Creating...' : 'Create account'}
                </button>
              )}
              {mode === 'login' && (
                <button onClick={() => setMode('forgot')} className="text-xs text-gray-500 text-center">Forgot password?</button>
              )}
              {mode === 'signup' && password.length > 0 && password.length < 6 && (
                <p className="text-xs text-gray-500 text-center">Password must be at least 6 characters.</p>
              )}
            </div>
          </>
        )}

        {error && <p className="text-xs text-red-400 mt-4 text-center">{error}</p>}
        <button onClick={() => navigate('/')} className="text-xs text-gray-500 mt-6">← Back to upload</button>
      </div>
    );
  }

  // --- Logged in: dashboard ------------------------------------------------
  const lessons = items.filter((i) => i.kind === 'lesson');
  const quizzes = items.filter((i) => i.kind === 'quiz');
  const displayName = profile?.display_name || user.email;

  const pctPresentations = usage ? Math.min(100, (usage.presentationCount / usage.presentationLimit) * 100) : 0;
  const pctStorage = usage ? Math.min(100, (usage.storageBytes / usage.storageLimitBytes) * 100) : 0;
  const nearLimit = (pct: number) => pct >= 90;

  return (
    <div className="flex flex-col w-full max-w-2xl mx-auto p-6 mt-12 text-white">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-2xl font-bold">👋 {displayName}</h1>
          <p className="text-xs text-gray-500">{user.email}</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => navigate('/')} className="text-xs bg-blue-600 hover:bg-blue-500 px-3 py-2 rounded-lg font-bold">+ New presentation</button>
          <button onClick={handleLogout} className="text-xs text-gray-500 hover:text-red-400 px-3 py-2">Log out</button>
        </div>
      </div>

      {info && <p className="text-xs text-emerald-400 mb-4">{info}</p>}
      {error && <p className="text-xs text-red-400 mb-4">{error}</p>}

      {usage && (
        <div className="w-full mb-6 grid grid-cols-2 gap-3">
          <div className="bg-gray-900 border border-gray-700 rounded-lg p-3">
            <div className="flex justify-between text-xs text-gray-400 mb-1">
              <span>Presentations</span>
              <span>{usage.presentationCount} / {usage.presentationLimit}</span>
            </div>
            <div className="h-1.5 bg-gray-800 rounded-full overflow-hidden">
              <div className={`h-full rounded-full ${nearLimit(pctPresentations) ? 'bg-red-500' : 'bg-blue-500'}`} style={{ width: `${pctPresentations}%` }} />
            </div>
          </div>
          <div className="bg-gray-900 border border-gray-700 rounded-lg p-3">
            <div className="flex justify-between text-xs text-gray-400 mb-1">
              <span>Storage</span>
              <span>{formatBytes(usage.storageBytes)} / {formatBytes(usage.storageLimitBytes)}</span>
            </div>
            <div className="h-1.5 bg-gray-800 rounded-full overflow-hidden">
              <div className={`h-full rounded-full ${nearLimit(pctStorage) ? 'bg-red-500' : 'bg-blue-500'}`} style={{ width: `${pctStorage}%` }} />
            </div>
          </div>
        </div>
      )}

      {loadingItems && <p className="text-sm text-gray-500">Loading...</p>}

      <h2 className="text-sm font-bold text-gray-400 uppercase mb-2">📚 Presentations</h2>
      {!loadingItems && lessons.length === 0 && <p className="text-sm text-gray-500 mb-6">No saved presentations yet - start one from the upload page.</p>}
      <div className="flex flex-col gap-2 mb-6">
        {lessons.map((item) => (
          <div key={item.id} className="flex items-center gap-3 bg-gray-900 border border-gray-700 rounded-lg px-4 py-3">
            <span className="flex-1 text-sm truncate">{item.title}</span>
            {typeof item.size_bytes === 'number' && item.size_bytes > 0 && (
              <span className="text-xs text-gray-500 shrink-0">{formatBytes(item.size_bytes)}</span>
            )}
            <button onClick={() => openLesson(item)} className="text-xs bg-emerald-600 hover:bg-emerald-500 px-3 py-1.5 rounded-lg font-bold shrink-0">▶ Open</button>
            <button onClick={() => deleteItem(item)} className="text-gray-500 hover:text-red-400 text-sm shrink-0">✕</button>
          </div>
        ))}
      </div>

      <h2 className="text-sm font-bold text-gray-400 uppercase mb-2">🧠 Quizzes</h2>
      {!loadingItems && quizzes.length === 0 && <p className="text-sm text-gray-500 mb-6">No saved quizzes yet - build one from inside a presentation and save it there.</p>}
      <div className="flex flex-col gap-2">
        {quizzes.map((item) => (
          <div key={item.id} className="flex items-center gap-3 bg-gray-900 border border-gray-700 rounded-lg px-4 py-3">
            <span className="flex-1 text-sm truncate">{item.title} <span className="text-gray-500">({(item.questions || []).length} q)</span></span>
            <button onClick={() => openQuiz(item)} className="text-xs bg-indigo-600 hover:bg-indigo-500 px-3 py-1.5 rounded-lg font-bold shrink-0">▶ Open</button>
            <button onClick={() => deleteItem(item)} className="text-gray-500 hover:text-red-400 text-sm shrink-0">✕</button>
          </div>
        ))}
      </div>

      {!profile && (
        <button onClick={refreshProfile} className="text-xs text-gray-600 mt-8 self-start">retry loading profile</button>
      )}
    </div>
  );
}
