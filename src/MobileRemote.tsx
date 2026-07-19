import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Document, Page, pdfjs } from 'react-pdf';
import { supabase } from './supabaseClient';

pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

const GAS_URL = 'https://script.google.com/macros/s/AKfycbx48s5aNamkERYuvJ-BE7-RBF2zt15mFZ-C-SXL_UIGZkG46RdyuPYIOlO6o0HZcr3N/exec';

const texts = {
  ku: {
    session: 'سێشن', slide: 'سلاید', next: 'دواتر', prev: 'پێشتر', controller: 'پادی کۆنترۆڵ', laser: 'لێزەر',
    draw: 'کێشان', highlight: 'هایلایت', erase: 'پاکەرەوە', color: 'ڕەنگ', clear: 'سڕینەوە', timer: 'کات',
    switchLang: 'EN', connecting: 'پەیوەندی دەکرێت...', loadingPdf: 'خوێندنەوەی سلاید...', error: 'هەڵە',
    spotlight: 'تیشک', zoom: 'زووم', black: 'ڕەشکردنەوە', white: 'سپیکردنەوە', notes: 'تێبینی', hideNotes: 'شاردنەوە',
    play: 'لێدان', pause: 'وەستان', mute: 'بێدەنگ', unmute: 'دەنگ', reset: 'ڕێکخستنەوە', video: 'ڤیدیۆ',
    start: 'دەستپێکردن', minutesLabel: 'خولەک', undo: 'گەڕانەوە', first: 'یەکەم', last: 'کۆتایی',
  },
  en: {
    session: 'Session', slide: 'Slide', next: 'NEXT', prev: 'PREV', controller: 'Controller Area', laser: 'Laser',
    draw: 'Draw', highlight: 'Highlight', erase: 'Erase', color: 'Color', clear: 'Clear', timer: 'Time',
    switchLang: 'کوردی', connecting: 'Connecting...', loadingPdf: 'Loading slide...', error: 'Error',
    spotlight: 'Spotlight', zoom: 'Zoom', black: 'Black screen', white: 'White screen', notes: 'Notes', hideNotes: 'Hide',
    play: 'Play', pause: 'Pause', mute: 'Mute', unmute: 'Unmute', reset: 'Reset', video: 'Video',
    start: 'Start', minutesLabel: 'min', undo: 'Undo', first: 'First', last: 'Last',
  },
};

interface Point { x: number; y: number; }
type DrawMode = 'draw' | 'highlight' | 'erase';
interface Stroke { points: Point[]; color: string; width: number; mode: DrawMode; }
type CanvasDataMap = Record<number, Stroke[]>;

// One entry per visible slide (mirrors the type in Present.tsx). Built and
// shared by the host - the remote no longer guesses a fixed total.
interface FlatSlide {
  itemIndex: number;
  pageInItem: number;
  fileType: string;
  name?: string;
  notes?: string;
  thumbnail?: string;
}

type ScreenMode = 'normal' | 'black' | 'white';
interface ZoomState { scale: number; x: number; y: number; }
interface VideoTime { playing: boolean; time: number; duration: number; volume: number; }

type ResolvedPreview =
  | { fileType: 'pdf'; data: Uint8Array }
  | { fileType: 'image'; url: string }
  | { fileType: 'video-link'; name?: string }
  | { fileType: 'other'; name?: string }
  | null;

// Preset swatches shown as quick-tap buttons. The native color input next to
// them lets you pick literally any color, so this list is just convenience.
const PRESET_COLORS = ['#eab308', '#ef4444', '#3b82f6', '#22c55e', '#f97316', '#ffffff', '#000000'];

// Highlight is a thick, semi-transparent stroke (marker style). Erase uses
// destination-out compositing, so its "color" is irrelevant — only the
// stroke shape matters — but it still needs a width.
const STROKE_WIDTHS: Record<DrawMode, number> = { draw: 4, highlight: 22, erase: 28 };

type ToolMode = 'none' | 'laser' | DrawMode | 'spotlight' | 'zoom';

const TYPE_ICON: Record<string, string> = { pdf: '📄', image: '🖼️', 'video-link': '▶️', other: '📁' };

export default function MobileRemote() {
  const [sessionId, setSessionId] = useState<string | null>(null);
  const [currentSlide, setCurrentSlide] = useState(1);
  const [flatSlides, setFlatSlides] = useState<FlatSlide[]>([]);

  const [ready, setReady] = useState(false);
  const channelRef = useRef<any>(null);

  const [activeMode, setActiveMode] = useState<ToolMode>('none');
  const [selectedColor, setSelectedColor] = useState('#eab308');
  const [lang, setLang] = useState<'ku' | 'en'>('ku');
  const t = texts[lang];

  // PIN gate (soft client-side lock) - if the host set a PIN when the
  // session was created, the remote must enter it before it can control
  // anything. This deters someone else grabbing the QR code and taking
  // over; it is NOT cryptographic security (a determined person could
  // still read it out of network traffic) - real access control needs a
  // Supabase RLS policy, which lives outside this file.
  const [requiredPin, setRequiredPin] = useState<string | null>(null);
  const [pinInput, setPinInput] = useState('');
  const [pinUnlocked, setPinUnlocked] = useState(false);
  const [pinError, setPinError] = useState(false);

  // "Another controller" awareness via Supabase Presence on the same
  // channel - no extra connection needed. Lets two people in the room
  // realize they're both holding the remote before they start fighting
  // over Next/Prev.
  const [otherRemoteCount, setOtherRemoteCount] = useState(0);
  const myClientId = useMemo(() => Math.random().toString(36).slice(2), []);

  const readyRef = useRef(false);
  useEffect(() => { readyRef.current = ready; }, [ready]);

  // Countdown timer (⌛) - replaces the old always-running stopwatch.
  // null secondsLeft means "no timer set" (shows --:--); 0 means it just
  // finished (shows a solid red pulsing 00:00).
  const [timerMinutesInput, setTimerMinutesInput] = useState('10');
  const [timerSecondsLeft, setTimerSecondsLeft] = useState<number | null>(null);
  const [timerRunning, setTimerRunning] = useState(false);
  const [timerPanelOpen, setTimerPanelOpen] = useState(false);
  const trackpadRef = useRef<HTMLDivElement>(null);
  const lastSentTime = useRef<number>(0);
  const isDrawing = useRef(false);

  // Preview state - now generalized to whatever type the active slide is,
  // not just pdf.
  const [fileId, setFileId] = useState<string | null>(null);
  const [preview, setPreview] = useState<ResolvedPreview>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [trackpadWidth, setTrackpadWidth] = useState(0);
  const lastFetchedFileId = useRef<string | null>(null);
  // Mirrors `fileId` for use inside the mount-once broadcast handler below,
  // which would otherwise only ever see the fileId from the very first render.
  const fileIdRef = useRef<string | null>(null);
  useEffect(() => { fileIdRef.current = fileId; }, [fileId]);

  // Screen mode / zoom / video state, mirrored from the host.
  const [screenMode, setScreenMode] = useState<ScreenMode>('normal');
  const [zoom, setZoom] = useState<ZoomState>({ scale: 1, x: 0, y: 0 });
  const [videoTime, setVideoTime] = useState<VideoTime>({ playing: false, time: 0, duration: 0, volume: 100 });
  const [notesOpen, setNotesOpen] = useState(true);

  // Local laser preview — mirrors what gets broadcast, so the phone shows
  // the exact same dot (position + on/off) as the projector.
  const [myLaser, setMyLaser] = useState({ x: 0.5, y: 0.5, active: false });
  const zoomDrag = useRef<{ lastX: number; lastY: number } | null>(null);

  // Keeps a reliable "current slide" value for use inside refs/effects
  // that shouldn't go stale (canvas resize callbacks, etc).
  const currentSlideRef = useRef(1);

  // Local drawing overlay — same idea as the one in Present.tsx, so strokes
  // you draw show up on your own phone screen too, not just the projector.
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const ctxRef = useRef<CanvasRenderingContext2D | null>(null);
  const allDrawingsRef = useRef<CanvasDataMap>({});
  const currentLineRef = useRef<Point[]>([]);

  const activeFlat = flatSlides[currentSlide - 1];

  // Replays every stroke for a slide in order, each with its own
  // color/width/mode. Erase strokes use destination-out compositing, so
  // replaying them in their original position in the sequence correctly
  // "cuts out" whatever was drawn before them — this is what makes the
  // eraser work for both live drawing and reloading a saved session.
  const redrawCanvasForSlide = (slideNum: number, drawingsMap: CanvasDataMap) => {
    const canvas = canvasRef.current;
    const ctx = ctxRef.current;
    if (!canvas || !ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const strokes = drawingsMap[slideNum] || [];
    strokes.forEach((stroke) => {
      if (stroke.points.length === 0) return;
      ctx.save();
      ctx.globalCompositeOperation = stroke.mode === 'erase' ? 'destination-out' : 'source-over';
      ctx.globalAlpha = stroke.mode === 'highlight' ? 0.35 : 1;
      ctx.strokeStyle = stroke.color;
      ctx.lineWidth = stroke.width;
      ctx.beginPath();
      ctx.moveTo(stroke.points[0].x * canvas.width, stroke.points[0].y * canvas.height);
      for (let i = 1; i < stroke.points.length; i++) {
        ctx.lineTo(stroke.points[i].x * canvas.width, stroke.points[i].y * canvas.height);
      }
      ctx.stroke();
      ctx.closePath();
      ctx.restore();
    });
  };

  useEffect(() => {
    const currentUrl = window.location.href;
    const match = currentUrl.match(/[?&]session=([^&]+)/);
    const session = match ? match[1].trim().replace(/[/#\s]+$/, '') : null;
    if (!session) return;
    setSessionId(session);

    let cancelled = false;
    let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
    let reconnectAttempt = 0;

    // Critical path: current_slide + file_id + slide_map + the session PIN
    // (if any). Re-run on every (re)connect, not just on first mount, so a
    // phone that reconnects after a drop catches anything it missed while
    // disconnected.
    const fetchCurrentSessionState = async () => {
      const { data, error } = await supabase
        .from('sessions')
        .select('current_slide, file_id, slide_map')
        .eq('id', session)
        .maybeSingle();
      if (error) console.error('🚨 DB Fetch Error:', error.message);
      if (data) {
        if (data.current_slide) {
          setCurrentSlide(data.current_slide);
          currentSlideRef.current = data.current_slide;
        }
        if (data.file_id) setFileId(data.file_id);
        if (Array.isArray(data.slide_map)) setFlatSlides(data.slide_map);
      } else if (!cancelled) {
        // The presenter may not have finished preparing/writing the slide
        // list yet (e.g. the phone scanned the QR a moment too early).
        // Retry briefly instead of leaving the remote permanently stuck
        // with an empty slide list.
        setTimeout(() => { if (!cancelled) fetchCurrentSessionState(); }, 1500);
      }

      // Non-critical: existing drawings + screen/zoom/PIN state for this
      // session. Wrapped separately so any issue here (missing column, bad
      // JSON, etc) can never block the slide from displaying.
      try {
        const { data: extra } = await supabase
          .from('sessions')
          .select('canvas_data, session_state')
          .eq('id', session)
          .maybeSingle();
        if (extra?.canvas_data) {
          allDrawingsRef.current = extra.canvas_data as CanvasDataMap;
          setTimeout(() => redrawCanvasForSlide(currentSlideRef.current, allDrawingsRef.current), 500);
        }
        if (extra?.session_state) {
          const s = extra.session_state as any;
          if (s.screenMode) setScreenMode(s.screenMode);
          if (s.zoom) setZoom(s.zoom);
          if (s.videoState) setVideoTime(s.videoState);
          // Once unlocked, stay unlocked across reconnects instead of
          // re-prompting - only set requiredPin the first time we see it.
          if (s.pin) setRequiredPin((prev) => prev ?? s.pin);
        }
      } catch (drawErr) {
        console.error('🚨 Extra session data fetch error (non-blocking):', drawErr);
      }
    };

    const scheduleReconnect = () => {
      if (cancelled || reconnectTimer) return;
      const delay = Math.min(1000 * 2 ** reconnectAttempt, 15000);
      reconnectAttempt += 1;
      reconnectTimer = setTimeout(() => {
        reconnectTimer = null;
        if (cancelled) return;
        if (channelRef.current) { supabase.removeChannel(channelRef.current); channelRef.current = null; }
        fetchCurrentSessionState();
        connect();
      }, delay);
    };

    const connect = () => {
      if (cancelled) return;
      const room = supabase.channel(`session_${session}`, {
        config: { broadcast: { ack: true }, presence: { key: myClientId } },
      });

      room.on('broadcast', { event: 'slide_change' }, (payload: any) => {
        const n = payload.payload?.slide;
        if (typeof n !== 'number') return;
        setCurrentSlide(n);
        currentSlideRef.current = n;
        redrawCanvasForSlide(n, allDrawingsRef.current);

        // The presenter may have moved to a different lesson item entirely,
        // not just a different page of the same one. If the fileId that
        // came with this broadcast differs from what we last loaded, the
        // preview effect below needs to refetch - this is what keeps the
        // phone showing exactly what the projector shows.
        const incomingFileId = payload.payload?.fileId;
        if (typeof incomingFileId === 'string' && incomingFileId && incomingFileId !== fileIdRef.current) {
          setFileId(incomingFileId);
        }
      });

      room.on('broadcast', { event: 'slide_map_update' }, (payload: any) => {
        if (Array.isArray(payload.payload?.slideMap)) setFlatSlides(payload.payload.slideMap);
      });

      room.on('broadcast', { event: 'screen_mode' }, (payload: any) => {
        if (payload.payload?.mode) setScreenMode(payload.payload.mode);
      });

      room.on('broadcast', { event: 'zoom_change' }, (payload: any) => {
        const { scale, x, y } = payload.payload || {};
        if (typeof scale === 'number') setZoom({ scale, x: x || 0, y: y || 0 });
      });

      room.on('broadcast', { event: 'video_time_update' }, (payload: any) => {
        const { playing, time, duration, volume } = payload.payload || {};
        if (typeof time === 'number') setVideoTime({ playing: !!playing, time, duration: duration || 0, volume: volume ?? 100 });
      });

      // Presence: who else is currently connected to this session's
      // channel, so we can flag "another device is also controlling this".
      room.on('presence', { event: 'sync' }, () => {
        const state = room.presenceState();
        const otherKeys = Object.keys(state).filter((k) => {
          if (k === myClientId) return false;
          const entries = (state as Record<string, any[]>)[k] || [];
          return entries.some((entry) => entry?.role === 'remote');
        });
        setOtherRemoteCount(otherKeys.length);
      });

      room.subscribe(async (status: string) => {
        if (cancelled) return;
        if (status === 'SUBSCRIBED') {
          setReady(true);
          reconnectAttempt = 0;
          await room.track({ role: 'remote', joinedAt: Date.now() });
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT' || status === 'CLOSED') {
          setReady(false);
          console.error(`🚨 WebSocket disconnected (${status}) - reconnecting...`);
          scheduleReconnect();
        }
      });

      channelRef.current = room;
    };

    fetchCurrentSessionState();
    connect();

    // Phones suspend the socket when the screen locks or the tab is
    // backgrounded, which doesn't always surface as a clean TIMED_OUT
    // event. Reconnecting the instant the app is visible again beats
    // waiting for that to eventually fire on its own.
    const onVisibilityChange = () => {
      if (document.visibilityState !== 'visible') return;
      if (!readyRef.current) {
        if (reconnectTimer) { clearTimeout(reconnectTimer); reconnectTimer = null; }
        if (channelRef.current) { supabase.removeChannel(channelRef.current); channelRef.current = null; }
        reconnectAttempt = 0;
        fetchCurrentSessionState();
        connect();
      } else {
        fetchCurrentSessionState();
      }
    };
    document.addEventListener('visibilitychange', onVisibilityChange);

    return () => {
      cancelled = true;
      if (reconnectTimer) clearTimeout(reconnectTimer);
      document.removeEventListener('visibilitychange', onVisibilityChange);
      if (channelRef.current) supabase.removeChannel(channelRef.current);
    };
  }, [myClientId]);

  // Keep-awake: prevents the phone from locking/dimming mid-session, which
  // is what silently kills the realtime connection in the first place.
  useEffect(() => {
    let cancelled = false;
    let wakeLock: any = null;

    const requestWakeLock = async () => {
      try {
        if ('wakeLock' in navigator) {
          wakeLock = await (navigator as any).wakeLock.request('screen');
        }
      } catch (err) {
        console.warn('⚠️ Wake Lock unavailable:', err);
      }
    };
    requestWakeLock();

    // The Wake Lock is automatically released whenever the tab is hidden,
    // so it has to be re-acquired every time the phone comes back.
    const onVisibility = () => {
      if (document.visibilityState === 'visible' && !cancelled) requestWakeLock();
    };
    document.addEventListener('visibilitychange', onVisibility);

    return () => {
      cancelled = true;
      document.removeEventListener('visibilitychange', onVisibility);
      wakeLock?.release?.().catch(() => {});
    };
  }, []);

  // Countdown timer tick - only runs while a timer is actively counting
  // down, unlike the old stopwatch which ran for the whole session.
  useEffect(() => {
    if (!timerRunning) return;
    const interval = setInterval(() => {
      setTimerSecondsLeft((prev) => {
        if (prev === null) return null;
        if (prev <= 1) {
          setTimerRunning(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
    return () => clearInterval(interval);
  }, [timerRunning]);

  // Fires once when the countdown reaches zero: a vibration pattern plus a
  // short tone, so it's noticeable even if the phone's face-down or the
  // presenter isn't looking at it right at that second.
  const timerAlertFiredRef = useRef(false);
  useEffect(() => {
    if (timerSecondsLeft !== 0) { timerAlertFiredRef.current = false; return; }
    if (timerAlertFiredRef.current) return;
    timerAlertFiredRef.current = true;

    if (navigator.vibrate) navigator.vibrate([200, 100, 200, 100, 200]);

    try {
      const AudioCtx = (window as any).AudioContext || (window as any).webkitAudioContext;
      const ctx = new AudioCtx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.type = 'sine';
      osc.frequency.value = 880;
      gain.gain.setValueAtTime(0.2, ctx.currentTime);
      osc.start();
      osc.stop(ctx.currentTime + 0.5);
      osc.onended = () => ctx.close();
    } catch (err) {
      console.warn('⚠️ Timer beep unavailable:', err);
    }
  }, [timerSecondsLeft]);

  // Fetches whatever the active slide actually is - pdf, image, or just
  // metadata for video/other (video is controlled remotely, not mirrored,
  // to avoid double audio playback from both the phone and the projector).
  useEffect(() => {
    if (!fileId) return;
    if (lastFetchedFileId.current === fileId) return; // same item as before - don't refetch on every page flip within it
    lastFetchedFileId.current = fileId;

    const fetchFile = async () => {
      setPreviewError(null);
      setPreview(null);

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 15000);

      try {
        const response = await fetch(`${GAS_URL}?action=getPdf&fileId=${fileId}`, { signal: controller.signal });
        clearTimeout(timeoutId);

        const text = await response.text();
        let result: any;
        try {
          result = JSON.parse(text);
        } catch {
          // We got something back, but it wasn't JSON - most likely a
          // Google sign-in redirect page instead of the actual API response.
          setPreviewError('Server did not return valid data (possibly a Google sign-in redirect)');
          return;
        }

        if (result.status !== 'success') {
          setPreviewError(result.message || 'Failed to load slide');
          return;
        }
        if (result.embedUrl) {
          setPreview({ fileType: 'video-link', name: result.name });
          return;
        }
        if (result.data && result.mimeType) {
          const mimeType: string = result.mimeType;
          const fileType: string = result.fileType || (mimeType === 'application/pdf' ? 'pdf' : mimeType.indexOf('image/') === 0 ? 'image' : 'other');
          const binaryString = atob(result.data);
          const bytes = new Uint8Array(binaryString.length);
          for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);

          if (fileType === 'pdf') {
            setPreview({ fileType: 'pdf', data: bytes });
          } else if (fileType === 'image') {
            const blob = new Blob([bytes], { type: mimeType });
            setPreview({ fileType: 'image', url: URL.createObjectURL(blob) });
          } else {
            setPreview({ fileType: 'other', name: result.name });
          }
        } else {
          setPreviewError('Unexpected response from server');
        }
      } catch (error: any) {
        clearTimeout(timeoutId);
        setPreviewError(error.name === 'AbortError' ? 'Request timed out — check your connection' : error.message || 'Network error');
      }
    };
    fetchFile();
  }, [fileId]);

  // pdfFile is memoized so react-pdf sees a stable object identity and
  // doesn't re-parse on every re-render.
  const pdfFile = useMemo(() => {
    if (preview?.fileType !== 'pdf') return null;
    return { data: preview.data };
  }, [preview]);

  // pageInItem tells us which page of the active pdf to show - this is the
  // piece that used to just be "currentSlide" before flat indexing existed.
  const pageInItem = activeFlat?.pageInItem || 1;

  // Keeps the preview sized to the trackpad box. This is the critical path
  // (the loading screen waits on trackpadWidth), so it always runs first
  // and on its own — the drawing canvas setup below can never block it.
  useEffect(() => {
    const trackpad = trackpadRef.current;
    if (!trackpad) return;

    const resizeCanvas = () => {
      const w = trackpad.clientWidth;
      const h = trackpad.clientHeight;
      setTrackpadWidth(w);

      try {
        const canvas = canvasRef.current;
        if (canvas && w > 0 && h > 0) {
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');
          if (ctx) {
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
            ctxRef.current = ctx;
          }
          redrawCanvasForSlide(currentSlideRef.current, allDrawingsRef.current);
        }
      } catch (canvasErr) {
        console.error('🚨 Drawing canvas setup error (non-blocking):', canvasErr);
      }
    };

    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);

    let observer: ResizeObserver | undefined;
    if (typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver(resizeCanvas);
      observer.observe(trackpad);
    }

    return () => {
      observer?.disconnect();
      window.removeEventListener('resize', resizeCanvas);
    };
  }, []);

  const formatTime = (totalSeconds: number) => {
    const m = Math.floor(totalSeconds / 60).toString().padStart(2, '0');
    const s = Math.floor(totalSeconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
  };

  // Starts a fresh countdown from whatever's in the minutes input (e.g. 10 -> 10:00).
  const startTimer = () => {
    const mins = parseFloat(timerMinutesInput);
    if (!mins || mins <= 0) return;
    setTimerSecondsLeft(Math.round(mins * 60));
    setTimerRunning(true);
    setTimerPanelOpen(false);
  };

  const togglePauseTimer = () => {
    if (timerSecondsLeft === null || timerSecondsLeft <= 0) return;
    setTimerRunning((r) => !r);
  };

  // Clears the timer back to "not set" (--:--) so a new duration can be entered.
  const resetTimer = () => {
    setTimerRunning(false);
    setTimerSecondsLeft(null);
    setTimerPanelOpen(false);
  };

  // Jumps straight to any global slide number - this now actually works
  // across item boundaries (a PDF's pages, then a link, then images, ...)
  // instead of only paging through whatever item happened to be on screen.
  const updateSlide = async (newSlideNumber: number) => {
    if (!ready || !channelRef.current || !flatSlides.length) return;
    if (newSlideNumber < 1 || newSlideNumber > flatSlides.length) return;
    setCurrentSlide(newSlideNumber);
    currentSlideRef.current = newSlideNumber;
    redrawCanvasForSlide(newSlideNumber, allDrawingsRef.current);

    await channelRef.current.send({ type: 'broadcast', event: 'slide_change', payload: { slide: newSlideNumber } });
    await supabase.from('sessions').update({ current_slide: newSlideNumber }).eq('id', sessionId);
  };

  const sendPointerData = (e: React.TouchEvent | React.MouseEvent, type: 'start' | 'move' | 'end') => {
    if (!ready || !channelRef.current || activeMode === 'none' || !trackpadRef.current) return;

    const rect = trackpadRef.current.getBoundingClientRect();
    const touchEvent = e as React.TouchEvent;
    const touch = touchEvent.touches?.[0] ?? touchEvent.changedTouches?.[0];
    const clientX = 'touches' in e ? touch?.clientX : (e as React.MouseEvent).clientX;
    const clientY = 'touches' in e ? touch?.clientY : (e as React.MouseEvent).clientY;

    if (clientX === undefined || clientY === undefined) return;

    const x = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
    const y = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));

    if (activeMode === 'laser') {
      const now = Date.now();
      if (now - lastSentTime.current < 16 && type === 'move') return;
      lastSentTime.current = now;

      const isActive = type !== 'end';
      setMyLaser({ x, y, active: isActive });
      channelRef.current.send({ type: 'broadcast', event: 'laser_move', payload: { x, y, active: isActive } });
    } else if (activeMode === 'spotlight') {
      const now = Date.now();
      if (now - lastSentTime.current < 16 && type === 'move') return;
      lastSentTime.current = now;
      const isActive = type !== 'end';
      channelRef.current.send({ type: 'broadcast', event: 'spotlight_move', payload: { x, y, active: isActive, radius: 160 } });
    } else if (activeMode === 'zoom') {
      // Dragging while in zoom mode pans; the +/- buttons (rendered below)
      // handle scale. Kept separate from laser/spotlight since it doesn't
      // need per-point broadcasting, just relative deltas.
      if (type === 'start') {
        zoomDrag.current = { lastX: clientX, lastY: clientY };
        return;
      }
      if (type === 'move' && zoomDrag.current) {
        const dxPct = ((clientX - zoomDrag.current.lastX) / rect.width) * 100;
        const dyPct = ((clientY - zoomDrag.current.lastY) / rect.height) * 100;
        zoomDrag.current = { lastX: clientX, lastY: clientY };
        setZoom((prev) => {
          const next = { ...prev, x: prev.x + dxPct / prev.scale, y: prev.y + dyPct / prev.scale };
          channelRef.current?.send({ type: 'broadcast', event: 'zoom_change', payload: next });
          return next;
        });
      }
      if (type === 'end') zoomDrag.current = null;
    } else if (activeMode === 'draw' || activeMode === 'highlight' || activeMode === 'erase') {
      const mode = activeMode as DrawMode;
      const width = STROKE_WIDTHS[mode];
      const color = selectedColor;

      channelRef.current.send({ type: 'broadcast', event: 'draw_stroke', payload: { x, y, type, mode, color, width } });

      const ctx = ctxRef.current;
      const canvas = canvasRef.current;
      if (ctx && canvas) {
        const pxX = x * canvas.width;
        const pxY = y * canvas.height;

        if (type === 'start') {
          ctx.save();
          ctx.globalCompositeOperation = mode === 'erase' ? 'destination-out' : 'source-over';
          ctx.globalAlpha = mode === 'highlight' ? 0.35 : 1;
          ctx.strokeStyle = color;
          ctx.lineWidth = width;
          ctx.beginPath();
          ctx.moveTo(pxX, pxY);
          currentLineRef.current = [{ x, y }];
        } else if (type === 'move') {
          ctx.lineTo(pxX, pxY);
          ctx.stroke();
          currentLineRef.current.push({ x, y });
        } else if (type === 'end') {
          ctx.closePath();
          ctx.restore();
          if (currentLineRef.current.length > 0) {
            const slideNum = currentSlideRef.current;
            const newStroke: Stroke = { points: currentLineRef.current, color, width, mode };
            allDrawingsRef.current[slideNum] = [...(allDrawingsRef.current[slideNum] || []), newStroke];
            currentLineRef.current = [];
          }
        }
      }
    }
  };

  const handleClear = async () => {
    if (!ready || !channelRef.current) return;
    await channelRef.current.send({ type: 'broadcast', event: 'draw_clear', payload: {} });

    const slideNum = currentSlideRef.current;
    allDrawingsRef.current[slideNum] = [];
    if (ctxRef.current && canvasRef.current) {
      ctxRef.current.clearRect(0, 0, canvasRef.current.width, canvasRef.current.height);
    }
  };

  // Removes just the most recent stroke, rather than wiping the whole
  // slide - much friendlier for "oops, one wrong line" mid-explanation.
  const handleUndo = async () => {
    if (!ready || !channelRef.current) return;
    const slideNum = currentSlideRef.current;
    const strokes = allDrawingsRef.current[slideNum] || [];
    if (!strokes.length) return;
    allDrawingsRef.current[slideNum] = strokes.slice(0, -1);
    redrawCanvasForSlide(slideNum, allDrawingsRef.current);
    await channelRef.current.send({ type: 'broadcast', event: 'draw_undo', payload: {} });
  };

  // Switching modes (or tapping the active mode off) force-hides the laser
  // immediately, so if you forget to lift your finger and just hit the
  // button instead, the dot still disappears on both screens.
  const handleModeChange = (mode: ToolMode) => {
    const newMode = activeMode === mode ? 'none' : mode;

    if (activeMode === 'laser' && newMode !== 'laser') {
      setMyLaser((prev) => ({ ...prev, active: false }));
      channelRef.current?.send({ type: 'broadcast', event: 'laser_move', payload: { x: myLaser.x, y: myLaser.y, active: false } });
    }
    if (activeMode === 'spotlight' && newMode !== 'spotlight') {
      channelRef.current?.send({ type: 'broadcast', event: 'spotlight_move', payload: { x: 0.5, y: 0.5, active: false, radius: 160 } });
    }

    setActiveMode(newMode as ToolMode);
  };

  const setScreenModeRemote = (mode: ScreenMode) => {
    const next = screenMode === mode ? 'normal' : mode;
    setScreenMode(next);
    channelRef.current?.send({ type: 'broadcast', event: 'screen_mode', payload: { mode: next } });
  };

  const adjustZoom = (delta: number) => {
    setZoom((prev) => {
      const next = { ...prev, scale: Math.max(1, Math.min(4, +(prev.scale + delta).toFixed(2))) };
      channelRef.current?.send({ type: 'broadcast', event: 'zoom_change', payload: next });
      return next;
    });
  };

  const resetZoom = () => {
    const next = { scale: 1, x: 0, y: 0 };
    setZoom(next);
    channelRef.current?.send({ type: 'broadcast', event: 'zoom_change', payload: next });
  };

  const sendVideoControl = (action: string, value?: number) => {
    channelRef.current?.send({ type: 'broadcast', event: 'video_control', payload: { action, value } });
  };

  const handleTouchStart = (e: React.TouchEvent | React.MouseEvent) => { isDrawing.current = true; sendPointerData(e, 'start'); };
  const handleTouchMove = (e: React.TouchEvent | React.MouseEvent) => { if (!isDrawing.current) return; sendPointerData(e, 'move'); };
  const handleTouchEnd = (e: React.TouchEvent | React.MouseEvent) => { isDrawing.current = false; sendPointerData(e, 'end'); };

  if (!sessionId) return <div className="bg-black text-white p-6 text-xl font-bold text-center mt-20">Invalid Session Connection Link.</div>;

  if (requiredPin && !pinUnlocked) {
    return (
      <div className="flex flex-col items-center justify-center h-screen w-full bg-black text-white p-6 gap-4">
        <span className="text-5xl">🔒</span>
        <p className="text-sm text-gray-400 text-center">Enter the session PIN shown on the projector screen</p>
        <input
          type="text"
          inputMode="numeric"
          maxLength={6}
          autoFocus
          value={pinInput}
          onChange={(e) => { setPinInput(e.target.value.replace(/\D/g, '')); setPinError(false); }}
          className={`w-36 text-center text-3xl tracking-[0.4em] bg-gray-900 border rounded-lg py-3 ${pinError ? 'border-red-500' : 'border-gray-700'}`}
          placeholder="••••"
        />
        {pinError && <p className="text-red-400 text-xs font-bold">Incorrect PIN, try again.</p>}
        <button
          onClick={() => { if (pinInput === requiredPin) { setPinUnlocked(true); setPinError(false); } else setPinError(true); }}
          className="px-8 py-2.5 rounded-lg bg-blue-600 font-bold text-lg"
        >
          Unlock
        </button>
      </div>
    );
  }

  const isVideoActive = activeFlat?.fileType === 'video-link';

  return (
    <div dir={lang === 'ku' ? 'rtl' : 'ltr'} className="flex flex-col h-screen w-full bg-black text-white font-sans select-none overflow-hidden">

      {!ready && (
        <div className="w-full bg-red-600 text-white text-center py-1 text-xs font-bold animate-pulse">
          {t.connecting}
        </div>
      )}

      {ready && otherRemoteCount > 0 && (
        <div className="w-full bg-amber-600 text-black text-center py-1 text-xs font-bold">
          ⚠ {otherRemoteCount === 1 ? 'Another device is also controlling this session' : `${otherRemoteCount} other devices are also controlling this session`}
        </div>
      )}

      <div className="flex justify-between items-center p-4 bg-gray-900 border-b border-gray-800">
        <div className="flex items-center gap-3 relative">
          <button
            onClick={() => (timerSecondsLeft === null ? setTimerPanelOpen((o) => !o) : togglePauseTimer())}
            className={`flex items-center gap-1.5 px-3 py-1 rounded font-mono text-xl shadow-inner ${
              timerSecondsLeft === 0
                ? 'bg-red-600 text-white animate-pulse'
                : timerRunning && timerSecondsLeft !== null && timerSecondsLeft <= 60
                ? 'bg-amber-600 text-white animate-pulse'
                : timerRunning
                ? 'bg-gray-800 text-green-400'
                : 'bg-gray-800 text-yellow-400'
            }`}
          >
            <span>⌛</span>
            <span>{timerSecondsLeft === null ? '--:--' : formatTime(timerSecondsLeft)}</span>
          </button>

          {timerSecondsLeft !== null && (
            <button
              onClick={resetTimer}
              aria-label="Clear timer"
              className="text-xs text-gray-400 bg-gray-800 w-6 h-6 rounded-full flex items-center justify-center"
            >
              ✕
            </button>
          )}

          {timerPanelOpen && timerSecondsLeft === null && (
            <div
              className="absolute top-full mt-2 right-0 bg-gray-900 border border-gray-700 rounded-lg p-2 flex flex-col gap-2 z-50 shadow-xl"
              style={{ direction: 'ltr' }}
            >
              <div className="flex items-center gap-1">
                {[5, 10, 15, 20].map((m) => (
                  <button
                    key={m}
                    onClick={() => setTimerMinutesInput(String(m))}
                    className={`px-2 py-1 rounded text-xs font-bold ${timerMinutesInput === String(m) ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300'}`}
                  >
                    {m}
                  </button>
                ))}
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  inputMode="numeric"
                  min={1}
                  value={timerMinutesInput}
                  onChange={(e) => setTimerMinutesInput(e.target.value)}
                  className="w-16 bg-gray-800 text-white text-center rounded px-2 py-1 text-sm"
                  placeholder="10"
                />
                <span className="text-xs text-gray-400 shrink-0">{t.minutesLabel}</span>
                <button onClick={startTimer} className="px-3 py-1 rounded bg-blue-600 text-white text-xs font-bold shrink-0">
                  {t.start}
                </button>
              </div>
            </div>
          )}
        </div>
        <div className="flex items-center gap-3">
          <button onClick={() => setLang(lang === 'ku' ? 'en' : 'ku')} className="bg-gray-800 px-3 py-1 rounded text-sm font-bold">{t.switchLang}</button>
          <div className="bg-blue-600 px-3 py-1 rounded-full font-bold" style={{ direction: 'ltr' }}>{t.slide} {currentSlide}{flatSlides.length ? ` / ${flatSlides.length}` : ''}</div>
        </div>
      </div>

      {/* Slide thumbnails - one per global slide number, sourced from the
          host's flatSlides list so it always matches the real slide count
          (fixes "slide 6 doesn't show" for good, since this can no longer
          drift from what's actually on screen). */}
      <div className="w-full bg-gray-900 border-b border-gray-800 p-2 overflow-x-auto flex gap-2" style={{ direction: 'ltr' }}>
        {flatSlides.map((slide, i) => (
          <button
            key={i}
            onClick={() => updateSlide(i + 1)}
            disabled={!ready}
            className={`min-w-[56px] h-14 rounded flex flex-col items-center justify-center font-bold text-xs gap-0.5 transition-all overflow-hidden ${currentSlide === i + 1 ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50 cursor-not-allowed' : ''}`}
          >
            {slide.thumbnail ? (
              <img src={slide.thumbnail} alt="" className="w-8 h-6 object-cover rounded-sm" />
            ) : (
              <span className="text-base leading-none">{TYPE_ICON[slide.fileType] || '📄'}</span>
            )}
            <span>{i + 1}</span>
          </button>
        ))}
      </div>

      <div className="px-4 pt-3 flex gap-2">
        <button disabled={!ready || currentSlide <= 1} onClick={() => updateSlide(1)} className="flex-1 h-9 rounded-lg bg-gray-800 text-white text-xs font-bold disabled:opacity-40">
          ⏮ {t.first}
        </button>
        <button disabled={!ready || currentSlide >= flatSlides.length} onClick={() => updateSlide(flatSlides.length)} className="flex-1 h-9 rounded-lg bg-gray-800 text-white text-xs font-bold disabled:opacity-40">
          {t.last} ⏭
        </button>
      </div>

      <div className="p-4 grid grid-cols-2 gap-4">
        <button disabled={!ready} onClick={() => updateSlide(currentSlide - 1)} className={`h-20 rounded-xl bg-gray-800 text-white text-xl font-bold ${!ready ? 'opacity-50' : 'active:bg-gray-700'}`}>{t.prev}</button>
        <button disabled={!ready} onClick={() => updateSlide(currentSlide + 1)} className={`h-20 rounded-xl bg-blue-600 text-white text-xl font-bold shadow-lg ${!ready ? 'opacity-50' : 'active:bg-blue-700'}`}>{t.next}</button>
      </div>

      {/* Presenter notes - only shown when the active slide actually has
          some (populate `notes` on each slide object from your backend). */}
      {activeFlat?.notes && (
        <div className="mx-4 mb-2 bg-gray-900 border border-gray-800 rounded-lg p-3">
          <div className="flex justify-between items-center mb-1">
            <span className="text-xs text-gray-400 font-bold uppercase">{t.notes}</span>
            <button onClick={() => setNotesOpen((o) => !o)} className="text-xs text-blue-400">{notesOpen ? t.hideNotes : t.notes}</button>
          </div>
          {notesOpen && <p className="text-sm text-gray-200 whitespace-pre-wrap">{activeFlat.notes}</p>}
        </div>
      )}

      {/* Screen controls: black/white screen restore-to-color, spotlight,
          zoom - separate row from the draw tools since they apply to the
          whole screen rather than being a drawing mode. */}
      <div className="px-4 mb-2 flex gap-2 flex-wrap">
        <button disabled={!ready} onClick={() => setScreenModeRemote('black')} className={`px-3 py-1.5 rounded-full text-xs font-bold ${screenMode === 'black' ? 'bg-white text-black' : 'bg-gray-800 text-gray-400'}`}>⬛ {t.black}</button>
        <button disabled={!ready} onClick={() => setScreenModeRemote('white')} className={`px-3 py-1.5 rounded-full text-xs font-bold ${screenMode === 'white' ? 'bg-white text-black' : 'bg-gray-800 text-gray-400'}`}>⬜ {t.white}</button>
        {screenMode !== 'normal' && (
          <button onClick={() => setScreenModeRemote('normal')} className="px-3 py-1.5 rounded-full text-xs font-bold bg-blue-600 text-white">{t.reset}</button>
        )}
      </div>

      {/* Video controls - shown only when the active slide is a video link.
          Playback isn't mirrored on the phone itself (avoids double audio
          from both screens); these send remote commands to the projector's
          player instead. Currently wired for YouTube embeds. */}
      {isVideoActive && (
        <div className="mx-4 mb-2 bg-gray-900 border border-gray-800 rounded-lg p-3 flex flex-col gap-2">
          <span className="text-xs text-gray-400 font-bold uppercase">{t.video}</span>
          <div className="flex items-center gap-2">
            <button onClick={() => sendVideoControl(videoTime.playing ? 'pause' : 'play')} className="px-4 py-2 rounded-lg bg-blue-600 text-sm font-bold">
              {videoTime.playing ? `⏸ ${t.pause}` : `▶ ${t.play}`}
            </button>
            <button onClick={() => sendVideoControl('seek', Math.max(0, videoTime.time - 10))} className="px-3 py-2 rounded-lg bg-gray-800 text-sm">-10s</button>
            <button onClick={() => sendVideoControl('seek', videoTime.time + 10)} className="px-3 py-2 rounded-lg bg-gray-800 text-sm">+10s</button>
          </div>
          {videoTime.duration > 0 && (
            <input
              type="range" min={0} max={videoTime.duration} value={videoTime.time}
              onChange={(e) => sendVideoControl('seek', Number(e.target.value))}
              className="w-full"
            />
          )}
          <div className="flex items-center gap-2">
            <span className="text-xs text-gray-400">🔊</span>
            <input
              type="range" min={0} max={100} value={videoTime.volume}
              onChange={(e) => sendVideoControl('volume', Number(e.target.value))}
              className="w-full"
            />
          </div>
        </div>
      )}

      <div className="flex-1 px-4 pb-4 flex flex-col min-h-0">
        <div className="flex justify-between items-center mb-2 flex-wrap gap-y-2">
          <span className="text-xs text-gray-400 font-bold uppercase">{t.controller}</span>
          <div className="flex gap-2 flex-wrap justify-end">
            <button disabled={!ready} onClick={() => handleModeChange('laser')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'laser' ? 'bg-red-600 text-white' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.laser}</button>
            <button disabled={!ready} onClick={() => handleModeChange('spotlight')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'spotlight' ? 'bg-purple-600 text-white' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.spotlight}</button>
            <button disabled={!ready} onClick={() => handleModeChange('draw')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'draw' ? 'bg-yellow-500 text-black' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.draw}</button>
            <button disabled={!ready} onClick={() => handleModeChange('highlight')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'highlight' ? 'bg-yellow-500 text-black' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.highlight}</button>
            <button disabled={!ready} onClick={() => handleModeChange('erase')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'erase' ? 'bg-yellow-500 text-black' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.erase}</button>
            <button disabled={!ready} onClick={() => handleModeChange('zoom')} className={`px-3 py-1 rounded-full text-xs font-bold ${activeMode === 'zoom' ? 'bg-green-600 text-white' : 'bg-gray-800 text-gray-400'} ${!ready ? 'opacity-50' : ''}`}>{t.zoom}</button>
            {(activeMode === 'draw' || activeMode === 'highlight' || activeMode === 'erase') && (
              <>
                <button onClick={handleUndo} className="px-3 py-1 rounded-full bg-gray-700 text-white text-xs font-bold">↶ {t.undo}</button>
                <button onClick={handleClear} className="px-3 py-1 rounded-full bg-gray-700 text-white text-xs font-bold">{t.clear}</button>
              </>
            )}
          </div>
        </div>

        {/* Zoom controls - only shown while zoom mode is active; drag on the trackpad below to pan. */}
        {activeMode === 'zoom' && (
          <div className="flex items-center gap-2 mb-2">
            <button onClick={() => adjustZoom(-0.25)} className="w-9 h-9 rounded-full bg-gray-800 text-lg font-bold">-</button>
            <span className="text-sm text-gray-300 font-mono w-14 text-center">{zoom.scale.toFixed(2)}x</span>
            <button onClick={() => adjustZoom(0.25)} className="w-9 h-9 rounded-full bg-gray-800 text-lg font-bold">+</button>
            <button onClick={resetZoom} className="px-3 py-1.5 rounded-full bg-gray-700 text-xs font-bold">{t.reset}</button>
          </div>
        )}

        {/* Color picker — only meaningful for draw/highlight (erase ignores color) */}
        {(activeMode === 'draw' || activeMode === 'highlight') && (
          <div className="flex items-center gap-2 mb-2 overflow-x-auto pb-1">
            <span className="text-xs text-gray-400 font-bold shrink-0">{t.color}:</span>
            {PRESET_COLORS.map((c) => (
              <button
                key={c}
                onClick={() => setSelectedColor(c)}
                aria-label={c}
                className={`shrink-0 w-7 h-7 rounded-full border-2 ${selectedColor === c ? 'border-blue-400 scale-110' : 'border-gray-700'} transition-transform`}
                style={{ backgroundColor: c }}
              />
            ))}
            {/* Native color input — lets you pick literally any color on the phone */}
            <input
              type="color"
              value={selectedColor}
              onChange={(e) => setSelectedColor(e.target.value)}
              className="shrink-0 w-7 h-7 rounded-full border-2 border-gray-700 bg-transparent p-0 overflow-hidden"
            />
          </div>
        )}

        {/* Trackpad doubles as the live slide preview, drawing canvas, and laser/spotlight preview */}
        <div
          ref={trackpadRef}
          onTouchStart={handleTouchStart} onTouchMove={handleTouchMove} onTouchEnd={handleTouchEnd}
          onMouseDown={handleTouchStart} onMouseMove={handleTouchMove} onMouseUp={handleTouchEnd} onMouseLeave={handleTouchEnd}
          className={`flex-1 w-full min-h-0 rounded-2xl border-2 flex items-center justify-center transition-colors relative touch-none overflow-hidden bg-white ${activeMode !== 'none' && ready ? 'border-yellow-500' : 'border-gray-800'}`}
        >
          {isVideoActive ? (
            <div className="flex flex-col items-center gap-2 pointer-events-none text-black">
              <span className="text-5xl">▶️</span>
              <p className="text-xs font-bold text-center px-4">{activeFlat?.name || 'Video'}</p>
              <p className="text-[10px] text-gray-500 text-center px-6">Preview not shown here to avoid double audio — use the controls above.</p>
            </div>
          ) : previewError ? (
            <p className="text-red-500 text-xs font-bold p-4 text-center pointer-events-none">{t.error}: {previewError}</p>
          ) : preview?.fileType === 'pdf' && pdfFile && trackpadWidth > 0 ? (
            <div className="pointer-events-none">
              <Document
                file={pdfFile}
                loading={null}
                onLoadError={(err) => { console.error('🚨 PDF load error:', err); setPreviewError(err?.message || 'Failed to load PDF'); }}
              >
                <Page
                  pageNumber={pageInItem}
                  renderTextLayer={false}
                  renderAnnotationLayer={false}
                  width={trackpadWidth}
                  onRenderError={(err) => console.error('🚨 PDF page render error:', err)}
                />
              </Document>
            </div>
          ) : preview?.fileType === 'image' ? (
            <img src={preview.url} alt="" className="max-w-full max-h-full object-contain pointer-events-none" />
          ) : preview?.fileType === 'other' ? (
            <div className="flex flex-col items-center gap-2 pointer-events-none text-black">
              <span className="text-4xl">📁</span>
              <p className="text-xs font-bold text-center px-4">{preview.name || 'File'}</p>
            </div>
          ) : (
            <div className="flex flex-col items-center pointer-events-none">
              <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-500 mb-2"></div>
              <p className="text-black text-xs font-bold">{t.loadingPdf}</p>
            </div>
          )}

          {/* Local drawing overlay, mirrors what's drawn via Present.tsx's canvas */}
          <canvas ref={canvasRef} className="absolute top-0 left-0 w-full h-full pointer-events-none z-30" />

          {/* Local laser dot — same visual as the projector, mirrors myLaser state */}
          {myLaser.active && (
            <div
              className="absolute pointer-events-none z-40 rounded-full bg-red-500 shadow-[0_0_12px_#ef4444]"
              style={{ width: '14px', height: '14px', left: `${myLaser.x * 100}%`, top: `${myLaser.y * 100}%`, transform: 'translate(-50%, -50%)' }}
            />
          )}
        </div>
      </div>
    </div>
  );
}
