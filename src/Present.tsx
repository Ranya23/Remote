import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { QRCodeSVG } from 'qrcode.react';
import { Document, Page, pdfjs } from 'react-pdf';
import { supabase } from './supabaseClient';

pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

// Same deployment URL as FileUpload.tsx - keep these in sync.
const GAS_URL = 'https://script.google.com/macros/s/AKfycbx48s5aNamkERYuvJ-BE7-RBF2zt15mFZ-C-SXL_UIGZkG46RdyuPYIOlO6o0HZcr3N/exec';

interface LessonSlideRef {
  fileId: string;
  fileType: string;
  name: string;
  notes?: string; // Optional presenter notes - populate this from the backend to show them on the phone remote.
}

type ResolvedSlide =
  | { fileType: 'pdf'; blobUrl: string; name?: string }
  | { fileType: 'image'; blobUrl: string; name?: string }
  | { fileType: 'video-link'; embedUrl: string; platform?: string; name?: string }
  | { fileType: 'other'; blobUrl: string; name?: string };

// One entry per *visible slide*, not per lesson item. A 5-page PDF item
// produces 5 entries; an image or video-link item produces exactly 1.
// This is what "auto sort" numbering (1-5 for the PDF, 6 for a link, 7-8
// for two images, ...) actually means under the hood, and it's the piece
// that was missing before - the remote only ever knew about "PDF pages of
// whatever item happens to be on screen right now", not a global count.
interface FlatSlide {
  itemIndex: number;
  pageInItem: number; // 1-based. Always 1 for non-pdf items.
  fileType: string;
  name?: string;
  notes?: string;
  thumbnail?: string; // small data-URL preview, shown on the remote's thumbnail strip
}

type ScreenMode = 'normal' | 'black' | 'white';
interface ZoomState { scale: number; x: number; y: number; } // x/y are pan offsets in % of container size
interface SpotlightState { x: number; y: number; active: boolean; radius: number; }
interface VideoState { playing: boolean; time: number; duration: number; volume: number; }
interface SessionState { screenMode: ScreenMode; zoom: ZoomState; videoState: VideoState; pin?: string; }

const DEFAULT_ZOOM: ZoomState = { scale: 1, x: 0, y: 0 };
const DEFAULT_VIDEO_STATE: VideoState = { playing: false, time: 0, duration: 0, volume: 100 };
const DEFAULT_SESSION_STATE: SessionState = { screenMode: 'normal', zoom: DEFAULT_ZOOM, videoState: DEFAULT_VIDEO_STATE };

interface Point { x: number; y: number; }
type DrawMode = 'draw' | 'highlight' | 'erase';
interface Stroke { points: Point[]; color: string; width: number; mode: DrawMode; }
type CanvasDataMap = Record<number, Stroke[]>; // keyed by flat slide number (1-based)

// Turns a getPdf response into something we know how to render.
// Doesn't do any network calls itself - pure data shaping.
function normalizeResponse(json: any, nameHint?: string): ResolvedSlide {
  if (json.embedUrl) {
    return { fileType: 'video-link', embedUrl: json.embedUrl, platform: json.platform, name: nameHint || json.name };
  }

  if (json.data && json.mimeType) {
    const mimeType: string = json.mimeType;
    const fileType: string =
      json.fileType || (mimeType === 'application/pdf' ? 'pdf' : mimeType.indexOf('image/') === 0 ? 'image' : 'other');

    // base64 -> Blob -> object URL (works for pdf, image, or anything else)
    const byteChars = atob(json.data);
    const byteNumbers = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: mimeType });
    const blobUrl = URL.createObjectURL(blob);

    if (fileType === 'pdf') return { fileType: 'pdf', blobUrl, name: nameHint || json.name };
    if (fileType === 'image') return { fileType: 'image', blobUrl, name: nameHint || json.name };
    return { fileType: 'other', blobUrl, name: nameHint || json.name };
  }

  throw new Error('Unrecognized response from server');
}

async function fetchGetPdf(fileId: string) {
  const url = `${GAS_URL}?action=getPdf&fileId=${encodeURIComponent(fileId)}`;
  const response = await fetch(url);
  const text = await response.text();

  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error('Server sent an invalid response (expected JSON).');
  }

  if (json.status !== 'success') {
    throw new Error(json.message || 'Failed to load this slide.');
  }

  return json;
}

// Lightweight page-count lookup - parses the PDF structure without
// rendering anything, just so we know how many flat slide slots it needs.
async function getPdfPageCount(blobUrl: string): Promise<number> {
  const doc = await pdfjs.getDocument(blobUrl).promise;
  const n = doc.numPages;
  try { doc.destroy(); } catch { /* noop */ }
  return n;
}

// Renders a small, low-quality JPEG of one PDF page for the remote's
// thumbnail strip. Deliberately tiny (120px wide) - it only needs to be
// recognizable at ~56px on a phone, not sharp, and keeping it small keeps
// the slide_map broadcast/payload light.
async function renderPdfThumbnail(blobUrl: string, pageNumber: number): Promise<string | undefined> {
  try {
    const doc = await pdfjs.getDocument(blobUrl).promise;
    const page = await doc.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = 120 / baseViewport.width;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    if (!ctx) return undefined;
    await page.render({ canvas, canvasContext: ctx, viewport }).promise;
    const dataUrl = canvas.toDataURL('image/jpeg', 0.5);
    try { doc.destroy(); } catch { /* noop */ }
    return dataUrl;
  } catch {
    // Thumbnails are a nice-to-have - never let a failure here block slide numbering.
    return undefined;
  }
}

// Same idea for a plain image slide - downscaled so its thumbnail is a
// few KB instead of shipping the full-resolution image to every remote.
async function renderImageThumbnail(blobUrl: string): Promise<string | undefined> {
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = reject;
      el.src = blobUrl;
    });
    const scale = 120 / img.width;
    const canvas = document.createElement('canvas');
    canvas.width = 120;
    canvas.height = Math.max(1, Math.round(img.height * scale));
    const ctx = canvas.getContext('2d');
    if (!ctx) return undefined;
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/jpeg', 0.5);
  } catch {
    return undefined;
  }
}

// Adds autoplay / API params for platforms we know how to remote-control.
// YouTube's lightweight postMessage protocol only works once enablejsapi=1
// is present; other platforms are left untouched (see the note in the
// video-control broadcast handler below).
function withPlaybackParams(embedUrl: string, platform?: string): string {
  try {
    const url = new URL(embedUrl);
    if (platform === 'youtube' || url.hostname.includes('youtube.com') || url.hostname.includes('youtu.be')) {
      url.searchParams.set('enablejsapi', '1');
      url.searchParams.set('autoplay', '1');
      url.searchParams.set('playsinline', '1');
      return url.toString();
    }
    return embedUrl;
  } catch {
    return embedUrl;
  }
}

function postYouTubeCommand(iframe: HTMLIFrameElement | null, func: string, args: any[] = []) {
  if (!iframe || !iframe.contentWindow) return;
  iframe.contentWindow.postMessage(JSON.stringify({ event: 'command', func, args }), '*');
}

export default function Present() {
  const { fileId } = useParams<{ fileId: string }>();

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [lessonSlides, setLessonSlides] = useState<LessonSlideRef[] | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [resolved, setResolved] = useState<ResolvedSlide | null>(null);

  // Page-within-item state. Only meaningful when resolved.fileType === 'pdf'.
  const [currentPage, setCurrentPage] = useState(1);
  const [numPages, setNumPages] = useState<number | null>(null);
  // Replaces the old "landOnLastPageRef boolean" - now that flatSlides tells
  // us the exact target page up front, we can land on it directly instead
  // of guessing "first or last page" once the PDF finishes loading.
  const landOnPageRef = useRef<number | null>(null);

  // The flattened, global slide list - the thing that actually fixes "click
  // slide 6 on the phone and nothing happens". Built once per lesson (or
  // once numPages is known for a single-file presentation) and shared with
  // the remote via the `slide_map` column + `slide_map_update` broadcast.
  const [flatSlides, setFlatSlides] = useState<FlatSlide[]>([]);
  const flatSlidesRef = useRef<FlatSlide[]>([]);
  useEffect(() => { flatSlidesRef.current = flatSlides; }, [flatSlides]);

  // Caches every lesson item we've already downloaded (built up while
  // preparing flatSlides) so switching to an already-visited item is
  // instant instead of re-fetching.
  const itemCacheRef = useRef<Map<number, ResolvedSlide>>(new Map());

  // Refs so the (mount-once) realtime listener always sees fresh values
  // without re-subscribing the channel every render.
  const currentIndexRef = useRef(0);
  useEffect(() => { currentIndexRef.current = currentIndex; }, [currentIndex]);

  // Stale-closure-free copy of the current global slide number, for use
  // inside the mount-once broadcast handlers below (declared ahead of them
  // on purpose, even though currentFlatIndex itself is computed further down).
  const currentFlatIndexRef = useRef(0);

  // --- New presentation-tools state (laser/draw already existed on the
  // remote; everything else here is new) ---------------------------------
  const [screenMode, setScreenMode] = useState<ScreenMode>('normal');
  const [zoom, setZoom] = useState<ZoomState>(DEFAULT_ZOOM);
  const [spotlight, setSpotlight] = useState<SpotlightState>({ x: 0.5, y: 0.5, active: false, radius: 160 });
  const [laser, setLaser] = useState({ x: 0.5, y: 0.5, active: false });
  const [videoState, setVideoState] = useState<VideoState>(DEFAULT_VIDEO_STATE);

  const presentCanvasRef = useRef<HTMLCanvasElement>(null);
  const presentCtxRef = useRef<CanvasRenderingContext2D | null>(null);
  const allDrawingsRef = useRef<CanvasDataMap>({});
  const currentLineRef = useRef<Point[]>([]);
  const videoIframeRef = useRef<HTMLIFrameElement>(null);

  const sessionStateSaveTimer = useRef<any>(null);

  // --- Remote-control session (QR code + Supabase sync) -------------------
  const [sessionId] = useState(() => {
    const key = `nextslide_session_${fileId || 'default'}`;
    const saved = localStorage.getItem(key);
    if (saved) return saved;
    const newId = Math.random().toString(36).substring(2, 9);
    localStorage.setItem(key, newId);
    return newId;
  });

  const remoteUrl = `${window.location.origin}${window.location.pathname}#/remote?session=${sessionId}`;
  const wrapperRef = useRef<HTMLDivElement>(null);
  const channelRef = useRef<any>(null);

  // Session PIN (soft lock, see the matching comment in MobileRemote.tsx) -
  // generated once when the session row is first created, then persisted
  // in session_state so refreshing this tab doesn't mint a new one.
  const [sessionPin, setSessionPin] = useState<string | null>(null);
  // Presence: how many phones are currently connected & controlling this
  // session, so the host can spot "two people are fighting over Next".
  const [connectedRemoteCount, setConnectedRemoteCount] = useState(0);
  const hostClientId = useMemo(() => `host_${Math.random().toString(36).slice(2)}`, []);

  const toggleFullscreen = () => {
    if (!wrapperRef.current) return;
    if (!document.fullscreenElement) {
      wrapperRef.current.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen();
    }
  };

  // The file actually on screen right now - for a lesson this is the
  // active item's fileId, not the lesson's own id. MobileRemote.tsx fetches
  // and pages through exactly this file, so keeping the session row pointed
  // at it is what lets the remote's own preview stay correct.
  const activeFileId = lessonSlides ? lessonSlides[currentIndex]?.fileId : fileId;

  // The single global slide number (1-based) that both screens now agree
  // on. This replaces the old "PDF page of whatever's currently loaded"
  // notion that couldn't represent "slide 6" once a lesson had more than
  // one item.
  const currentFlatIndex = useMemo(() => {
    if (!flatSlides.length) return 0;
    const idx = flatSlides.findIndex((s) => s.itemIndex === currentIndex && s.pageInItem === currentPage);
    return idx === -1 ? 0 : idx;
  }, [flatSlides, currentIndex, currentPage]);

  // Redraws every stroke recorded for a given flat slide, in order, so
  // erase strokes (destination-out) correctly cut out whatever was drawn
  // before them - same algorithm MobileRemote.tsx uses for its own mirror.
  const redrawCanvasForSlide = useCallback((flatSlideNum: number, drawingsMap: CanvasDataMap) => {
    const canvas = presentCanvasRef.current;
    const ctx = presentCtxRef.current;
    if (!canvas || !ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const strokes = drawingsMap[flatSlideNum] || [];
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
  }, []);

  // Debounced write-back for the "slow moving" bits of shared state
  // (screen mode, zoom, video state) so a drag gesture doesn't hammer
  // Supabase with a write per pixel - broadcasts still go out immediately,
  // only the persisted copy (used to restore state on refresh / late join)
  // is throttled.
  const persistSessionState = useCallback((patch: Partial<SessionState>) => {
    if (sessionStateSaveTimer.current) clearTimeout(sessionStateSaveTimer.current);
    sessionStateSaveTimer.current = setTimeout(() => {
      supabase.from('sessions').upsert({
        id: sessionId,
        session_state: { screenMode, zoom, videoState, ...patch },
      });
    }, 300);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionId, screenMode, zoom, videoState]);

  // Creates the session row (if missing), hydrates local state from
  // whatever was last saved, and subscribes to remote commands.
  useEffect(() => {
    if (!fileId) return;

    let cancelled = false;
    let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
    let reconnectAttempt = 0;

    const setupSession = async () => {
      const { data } = await supabase
        .from('sessions')
        .select('id, canvas_data, session_state')
        .eq('id', sessionId)
        .single();

      if (!data) {
        // New session: mint a 4-digit PIN so the presenter can require it
        // before anyone else's phone can take over Next/Prev.
        const pin = String(Math.floor(1000 + Math.random() * 9000));
        setSessionPin(pin);
        await supabase.from('sessions').insert([{
          id: sessionId, file_id: fileId, current_slide: 1, canvas_data: {},
          session_state: { ...DEFAULT_SESSION_STATE, pin },
        }]);
      } else {
        await supabase.from('sessions').update({ file_id: fileId }).eq('id', sessionId);
        if (data.canvas_data) allDrawingsRef.current = data.canvas_data as CanvasDataMap;
        if (data.session_state) {
          const s = data.session_state as SessionState;
          if (s.screenMode) setScreenMode(s.screenMode);
          if (s.zoom) setZoom(s.zoom);
          if (s.pin) setSessionPin(s.pin);
        }
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
        connect();
      }, delay);
    };

    const connect = () => {
      if (cancelled) return;
      const channel = supabase.channel(`session_${sessionId}`, {
        config: { broadcast: { ack: true }, presence: { key: hostClientId } },
      });

      channel.on('broadcast', { event: 'slide_change' }, (payload) => {
        const flatNum = payload.payload?.slide;
        if (typeof flatNum !== 'number') return;
        const list = flatSlidesRef.current;
        if (!list.length) return;
        const clamped = Math.min(Math.max(0, flatNum - 1), list.length - 1);
        const target = list[clamped];
        if (!target) return;
        if (target.itemIndex !== currentIndexRef.current) {
          landOnPageRef.current = target.pageInItem;
          setCurrentIndex(target.itemIndex);
        } else {
          setCurrentPage(target.pageInItem);
        }
      });

      channel.on('broadcast', { event: 'laser_move' }, (payload) => {
        const { x, y, active } = payload.payload || {};
        if (typeof x === 'number' && typeof y === 'number') setLaser({ x, y, active: !!active });
      });

      channel.on('broadcast', { event: 'spotlight_move' }, (payload) => {
        const { x, y, active, radius } = payload.payload || {};
        if (typeof x === 'number' && typeof y === 'number') {
          setSpotlight((prev) => ({ x, y, active: !!active, radius: radius || prev.radius }));
        }
      });

      channel.on('broadcast', { event: 'zoom_change' }, (payload) => {
        const { scale, x, y } = payload.payload || {};
        if (typeof scale === 'number') {
          const next = { scale, x: x || 0, y: y || 0 };
          setZoom(next);
          persistSessionState({ zoom: next });
        }
      });

      channel.on('broadcast', { event: 'screen_mode' }, (payload) => {
        const mode = payload.payload?.mode as ScreenMode | undefined;
        if (mode) {
          setScreenMode(mode);
          persistSessionState({ screenMode: mode });
        }
      });

      channel.on('broadcast', { event: 'video_control' }, (payload) => {
        const { action, value } = payload.payload || {};
        const iframe = videoIframeRef.current;
        if (action === 'play') postYouTubeCommand(iframe, 'playVideo');
        else if (action === 'pause') postYouTubeCommand(iframe, 'pauseVideo');
        else if (action === 'seek' && typeof value === 'number') postYouTubeCommand(iframe, 'seekTo', [value, true]);
        else if (action === 'volume' && typeof value === 'number') postYouTubeCommand(iframe, 'setVolume', [value]);
        else if (action === 'mute') postYouTubeCommand(iframe, 'mute');
        else if (action === 'unmute') postYouTubeCommand(iframe, 'unMute');
      });

      channel.on('broadcast', { event: 'draw_stroke' }, (payload) => {
        const { x, y, type, mode, color, width } = payload.payload || {};
        const ctx = presentCtxRef.current;
        const canvas = presentCanvasRef.current;
        if (!ctx || !canvas || typeof x !== 'number' || typeof y !== 'number') return;
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
            const flatNum = currentFlatIndexRef.current + 1;
            const stroke: Stroke = { points: currentLineRef.current, color, width, mode };
            allDrawingsRef.current[flatNum] = [...(allDrawingsRef.current[flatNum] || []), stroke];
            currentLineRef.current = [];
            supabase.from('sessions').upsert({ id: sessionId, canvas_data: allDrawingsRef.current });
          }
        }
      });

      channel.on('broadcast', { event: 'draw_clear' }, () => {
        const flatNum = currentFlatIndexRef.current + 1;
        allDrawingsRef.current[flatNum] = [];
        const canvas = presentCanvasRef.current;
        const ctx = presentCtxRef.current;
        if (canvas && ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
        supabase.from('sessions').upsert({ id: sessionId, canvas_data: allDrawingsRef.current });
      });

      // Undo: drop just the most recent stroke for the current slide,
      // rather than clearing everything - mirrors handleUndo on the remote.
      channel.on('broadcast', { event: 'draw_undo' }, () => {
        const flatNum = currentFlatIndexRef.current + 1;
        const strokes = allDrawingsRef.current[flatNum] || [];
        if (!strokes.length) return;
        allDrawingsRef.current[flatNum] = strokes.slice(0, -1);
        redrawCanvasForSlide(flatNum, allDrawingsRef.current);
        supabase.from('sessions').upsert({ id: sessionId, canvas_data: allDrawingsRef.current });
      });

      // Presence: count phones connected to this session (excluding this
      // host tab itself) so the sidebar can flag more than one controller.
      channel.on('presence', { event: 'sync' }, () => {
        const state = channel.presenceState();
        const remoteKeys = Object.keys(state).filter((k) => k !== hostClientId);
        setConnectedRemoteCount(remoteKeys.length);
      });

      channel.subscribe(async (status: string) => {
        if (cancelled) return;
        if (status === 'SUBSCRIBED') {
          reconnectAttempt = 0;
          await channel.track({ role: 'host', joinedAt: Date.now() });
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT' || status === 'CLOSED') {
          console.error(`🚨 WebSocket disconnected (${status}) - reconnecting...`);
          scheduleReconnect();
        }
      });

      channelRef.current = channel;
    };

    setupSession();
    connect();

    // The projector's tab can also get backgrounded (switching windows to
    // pull up a different file, OS notification, etc). Reconnect the
    // instant it's foregrounded again instead of waiting on a timeout.
    const onVisibilityChange = () => {
      if (document.visibilityState !== 'visible' || channelRef.current) return;
      if (reconnectTimer) { clearTimeout(reconnectTimer); reconnectTimer = null; }
      reconnectAttempt = 0;
      connect();
    };
    document.addEventListener('visibilitychange', onVisibilityChange);

    return () => {
      cancelled = true;
      if (reconnectTimer) clearTimeout(reconnectTimer);
      document.removeEventListener('visibilitychange', onVisibilityChange);
      if (channelRef.current) { supabase.removeChannel(channelRef.current); channelRef.current = null; }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionId, fileId, hostClientId]);
  // -------------------------------------------------------------------------

  // Keep-awake: prevents the projector's screen/computer from sleeping or
  // dimming mid-presentation.
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

  // Keeps currentFlatIndexRef (declared above, near currentIndexRef) in sync
  // for the mount-once broadcast handlers.
  useEffect(() => { currentFlatIndexRef.current = currentFlatIndex; }, [currentFlatIndex]);

  // Persist + broadcast the current global slide number whenever it
  // changes - this is what lets the phone's thumbnail strip and the main
  // screen agree on "slide 6", regardless of which lesson item that is.
  useEffect(() => {
    if (!flatSlides.length || !activeFileId) return;
    supabase.from('sessions').upsert({ id: sessionId, file_id: activeFileId, current_slide: currentFlatIndex + 1 });
    // fileId is included here (not just the slide number) so the remote can
    // tell when the presenter has switched to a different lesson item -
    // without this, the phone has no way to know its cached preview is now
    // pointing at the wrong file, and just keeps showing whatever it loaded
    // first (or nothing, if that first load happened before this item's
    // real fileId was known).
    channelRef.current?.send({
      type: 'broadcast',
      event: 'slide_change',
      payload: { slide: currentFlatIndex + 1, fileId: activeFileId },
    });
  }, [currentFlatIndex, activeFileId, flatSlides.length, sessionId]);

  // Share the flat slide list itself with the remote (thumbnail strip,
  // total count, per-slide type/name/notes) whenever it's built or changes.
  useEffect(() => {
    if (!flatSlides.length) return;
    supabase.from('sessions').upsert({ id: sessionId, slide_map: flatSlides }).then(({ error }) => {
      if (error) console.error('🚨 slide_map upsert failed:', error.message, error);
    });
    channelRef.current?.send({ type: 'broadcast', event: 'slide_map_update', payload: { slideMap: flatSlides } });
  }, [flatSlides, sessionId]);

  // Resize + redraw the annotation canvas to match the stage, and redraw
  // whenever the visible slide changes.
  useEffect(() => {
    const stage = wrapperRef.current;
    if (!stage) return;

    const resizeCanvas = () => {
      const canvas = presentCanvasRef.current;
      if (!canvas) return;
      const w = stage.clientWidth;
      const h = stage.clientHeight;
      if (w === 0 || h === 0) return;
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        presentCtxRef.current = ctx;
      }
      redrawCanvasForSlide(currentFlatIndex + 1, allDrawingsRef.current);
    };

    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);
    let observer: ResizeObserver | undefined;
    if (typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver(resizeCanvas);
      observer.observe(stage);
    }
    return () => {
      observer?.disconnect();
      window.removeEventListener('resize', resizeCanvas);
    };
  }, [currentFlatIndex, redrawCanvasForSlide]);

  // Listens for YouTube's postMessage player updates so we can relay
  // play/pause/time/duration back to the phone remote's progress bar.
  // Only fires for YouTube embeds with enablejsapi=1 (see withPlaybackParams).
  useEffect(() => {
    let lastBroadcast = 0;
    function handleMessage(e: MessageEvent) {
      if (typeof e.data !== 'string') return;
      let data: any;
      try { data = JSON.parse(e.data); } catch { return; }
      if (data.event !== 'infoDelivery' || !data.info) return;
      const now = Date.now();
      if (now - lastBroadcast < 900) return;
      lastBroadcast = now;
      const next: VideoState = {
        playing: data.info.playerState === 1,
        time: data.info.currentTime || 0,
        duration: data.info.duration || 0,
        volume: typeof data.info.volume === 'number' ? data.info.volume : videoState.volume,
      };
      setVideoState(next);
      channelRef.current?.send({ type: 'broadcast', event: 'video_time_update', payload: next });
    }
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Establishes the YouTube "listening" handshake once a video-link slide
  // with jsapi enabled is on screen, so infoDelivery messages start flowing.
  useEffect(() => {
    if (resolved?.fileType !== 'video-link') return;
    const iframe = videoIframeRef.current;
    if (!iframe) return;
    const t = setTimeout(() => {
      iframe.contentWindow?.postMessage(JSON.stringify({ event: 'listening', id: 'present' }), '*');
    }, 1000);
    return () => clearTimeout(t);
  }, [resolved]);

  // Initial load: figure out whether this id is a lesson or a single slide.
  useEffect(() => {
    if (!fileId) return;
    let cancelled = false;

    async function load() {
      if (!fileId) return;
      setLoading(true);
      setError(null);
      setLessonSlides(null);
      setResolved(null);
      setFlatSlides([]);
      itemCacheRef.current.clear();

      try {
        const json = await fetchGetPdf(fileId);
        if (cancelled) return;

        if (Array.isArray(json.slides)) {
          setLessonSlides(json.slides);
          setCurrentIndex(0);
          setLoading(false);
        } else {
          const slide = normalizeResponse(json);
          itemCacheRef.current.set(0, slide);
          setResolved(slide);
          setLoading(false);
        }
      } catch (err: any) {
        if (!cancelled) {
          setError(err.message || 'Something went wrong loading this file.');
          setLoading(false);
        }
      }
    }

    load();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileId]);

  // When in lesson mode, load whichever slide is currently active - reusing
  // the cache built while preparing flatSlides so this is usually instant.
  useEffect(() => {
    if (!lessonSlides) return;
    let cancelled = false;

    async function loadCurrentSlide() {
      const cached = itemCacheRef.current.get(currentIndex);
      if (cached) {
        setResolved(cached);
        setLoading(false);
        return;
      }

      setLoading(true);
      setError(null);
      try {
        const ref = lessonSlides![currentIndex];
        const json = await fetchGetPdf(ref.fileId);
        if (cancelled) return;
        const slide = normalizeResponse(json, ref.name);
        itemCacheRef.current.set(currentIndex, slide);
        setResolved(slide);
        setLoading(false);
      } catch (err: any) {
        if (!cancelled) {
          setError(err.message || 'Something went wrong loading this slide.');
          setLoading(false);
        }
      }
    }

    loadCurrentSlide();
    return () => {
      cancelled = true;
    };
  }, [lessonSlides, currentIndex]);

  // Background pass: walk every lesson item, resolve it (cached for reuse
  // above), and for PDFs work out the real page count - this is what builds
  // the flat, globally-numbered slide list the remote needs. Runs
  // progressively so the remote's thumbnail strip fills in as it goes,
  // rather than waiting for every item to finish.
  useEffect(() => {
    if (!lessonSlides) return;
    let cancelled = false;

    (async () => {
      const result: FlatSlide[] = [];
      for (let i = 0; i < lessonSlides.length; i++) {
        if (cancelled) return;
        const ref = lessonSlides[i];
        try {
          let entry = itemCacheRef.current.get(i);
          if (!entry) {
            const json = await fetchGetPdf(ref.fileId);
            if (cancelled) return;
            entry = normalizeResponse(json, ref.name);
            itemCacheRef.current.set(i, entry);
          }
          if (entry.fileType === 'pdf') {
            const n = await getPdfPageCount(entry.blobUrl);
            for (let p = 1; p <= n; p++) {
              const thumbnail = await renderPdfThumbnail(entry.blobUrl, p);
              result.push({ itemIndex: i, pageInItem: p, fileType: 'pdf', name: ref.name, notes: ref.notes, thumbnail });
            }
          } else {
            const thumbnail = entry.fileType === 'image' ? await renderImageThumbnail(entry.blobUrl) : undefined;
            result.push({ itemIndex: i, pageInItem: 1, fileType: entry.fileType, name: ref.name, notes: ref.notes, thumbnail });
          }
        } catch {
          // Couldn't preload this item - still reserve one flat slot for it
          // so the numbering of everything after it doesn't collapse.
          result.push({ itemIndex: i, pageInItem: 1, fileType: ref.fileType || 'other', name: ref.name, notes: ref.notes });
        }
        if (!cancelled) setFlatSlides([...result]);
      }
    })();

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lessonSlides]);

  // Single-file (non-lesson) mode: build the flat list directly from
  // numPages once known, so remote thumbnails work here too.
  useEffect(() => {
    if (lessonSlides) return;
    if (!resolved) return;
    let cancelled = false;

    (async () => {
      if (resolved.fileType === 'pdf') {
        if (!numPages) return;
        const slides: FlatSlide[] = [];
        for (let i = 0; i < numPages; i++) {
          if (cancelled) return;
          const thumbnail = await renderPdfThumbnail(resolved.blobUrl, i + 1);
          slides.push({ itemIndex: 0, pageInItem: i + 1, fileType: 'pdf', thumbnail });
          if (!cancelled) setFlatSlides([...slides]);
        }
      } else if (resolved.fileType === 'image') {
        const thumbnail = await renderImageThumbnail(resolved.blobUrl);
        if (!cancelled) setFlatSlides([{ itemIndex: 0, pageInItem: 1, fileType: resolved.fileType, name: resolved.name, thumbnail }]);
      } else {
        if (!cancelled) setFlatSlides([{ itemIndex: 0, pageInItem: 1, fileType: resolved.fileType, name: resolved.name }]);
      }
    })();

    return () => { cancelled = true; };
  }, [lessonSlides, resolved, numPages]);

  // Whenever a new item finishes resolving: non-pdf items are always a
  // single "page", so reset immediately. PDFs wait for onPdfLoadSuccess to
  // know their real page count before landing on the requested page.
  useEffect(() => {
    if (!resolved) return;
    if (resolved.fileType !== 'pdf') {
      setCurrentPage(1);
      setNumPages(null);
      landOnPageRef.current = null;
    } else {
      setNumPages(null);
      if (landOnPageRef.current == null) setCurrentPage(1);
    }
  }, [resolved]);

  const onPdfLoadSuccess = useCallback(({ numPages: n }: { numPages: number }) => {
    setNumPages(n);
    if (landOnPageRef.current != null) {
      const target = Math.min(landOnPageRef.current, n);
      landOnPageRef.current = null;
      setCurrentPage(target);
    }
  }, []);

  // Clean up every cached object URL (not just the last one - items are
  // now preloaded and kept around) when this file changes or we unmount.
  useEffect(() => {
    return () => {
      itemCacheRef.current.forEach((slide) => {
        if ('blobUrl' in slide) URL.revokeObjectURL(slide.blobUrl);
      });
      itemCacheRef.current.clear();
    };
  }, [fileId]);

  // Jump straight to any global slide number - used by both the Prev/Next
  // buttons below and incoming remote slide_change events.
  const goToFlatIndex = useCallback((flatIdx: number) => {
    const list = flatSlides;
    if (!list.length) return;
    const clamped = Math.min(Math.max(0, flatIdx), list.length - 1);
    const target = list[clamped];
    if (target.itemIndex !== currentIndex) {
      landOnPageRef.current = target.pageInItem;
      setCurrentIndex(target.itemIndex);
    } else {
      setCurrentPage(target.pageInItem);
    }
  }, [flatSlides, currentIndex]);

  const goPrev = useCallback(() => {
    if (flatSlides.length) { goToFlatIndex(currentFlatIndex - 1); return; }
    // Fallback for the brief window before flatSlides has been prepared.
    if (resolved?.fileType === 'pdf' && currentPage > 1) { setCurrentPage(currentPage - 1); return; }
    if (!lessonSlides || currentIndex === 0) return;
    setCurrentIndex((i) => i - 1);
  }, [flatSlides, currentFlatIndex, goToFlatIndex, resolved, currentPage, lessonSlides, currentIndex]);

  const goNext = useCallback(() => {
    if (flatSlides.length) { goToFlatIndex(currentFlatIndex + 1); return; }
    if (resolved?.fileType === 'pdf' && numPages && currentPage < numPages) { setCurrentPage(currentPage + 1); return; }
    if (!lessonSlides || currentIndex >= lessonSlides.length - 1) return;
    setCurrentIndex((i) => i + 1);
  }, [flatSlides, currentFlatIndex, goToFlatIndex, resolved, numPages, currentPage, lessonSlides, currentIndex]);

  // Keyboard shortcuts on the main screen: arrows to navigate, B/W for
  // black/white screen, Escape to restore to color.
  useEffect(() => {
    function handleKeyDown(e: KeyboardEvent) {
      if (e.key === 'ArrowLeft') goPrev();
      if (e.key === 'ArrowRight') goNext();
      if (e.key.toLowerCase() === 'b') setScreenMode((m) => (m === 'black' ? 'normal' : 'black'));
      if (e.key.toLowerCase() === 'w') setScreenMode((m) => (m === 'white' ? 'normal' : 'white'));
      if (e.key === 'Escape') setScreenMode('normal');
    }
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [goPrev, goNext]);

  const retry = () => {
    setError(null);
    setLoading(true);
    if (lessonSlides) {
      itemCacheRef.current.delete(currentIndex);
      setCurrentIndex((i) => i);
      setLessonSlides((slides) => (slides ? [...slides] : slides));
    } else if (fileId) {
      fetchGetPdf(fileId)
        .then((json) => {
          if (Array.isArray(json.slides)) {
            setLessonSlides(json.slides);
            setCurrentIndex(0);
          } else {
            const slide = normalizeResponse(json);
            itemCacheRef.current.set(0, slide);
            setResolved(slide);
          }
          setLoading(false);
        })
        .catch((err) => {
          setError(err.message || 'Something went wrong.');
          setLoading(false);
        });
    }
  };

  const showNav = flatSlides.length > 1;
  const isFirstSlide = flatSlides.length ? currentFlatIndex === 0 : currentIndex === 0;
  const isLastSlide = flatSlides.length ? currentFlatIndex === flatSlides.length - 1 : true;

  let navLabel = '';
  if (flatSlides.length) {
    navLabel = `Slide ${currentFlatIndex + 1} of ${flatSlides.length}`;
    if (resolved?.name) navLabel += ` - ${resolved.name}`;
  }

  const zoomTransform = `scale(${zoom.scale}) translate(${zoom.x}%, ${zoom.y}%)`;

  return (
    <div className="flex h-screen w-full bg-black text-white overflow-hidden">
      {/* Sidebar: QR code for the phone remote + session info */}
      <div className="w-80 bg-gray-900 border-r border-gray-800 p-6 flex flex-col items-center justify-between shrink-0">
        <div className="w-full flex justify-end">
          <button onClick={toggleFullscreen} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1 rounded">
            Full Screen
          </button>
        </div>

        <div className="w-full flex flex-col items-center">
          <h2 className="text-3xl font-bold mb-2">NextSlide</h2>
          <p className="text-sm text-gray-400 mb-8">Scan to control</p>

          <div className="bg-white p-4 rounded-xl mb-2">
            <QRCodeSVG value={remoteUrl} size={180} />
          </div>

          <a href={remoteUrl} target="_blank" rel="noreferrer" className="text-xs text-blue-400 underline hover:text-blue-300 mb-6">
            Click here to test Remote on PC
          </a>
        </div>

        <div className="w-full text-center bg-gray-800 rounded p-4 border border-gray-700 flex flex-col gap-3">
          <div>
            <p className="text-xs text-gray-500 mb-1">Session</p>
            <p className="font-mono text-2xl text-blue-400">{sessionId}</p>
          </div>
          {sessionPin && (
            <div className="pt-3 border-t border-gray-700">
              <p className="text-xs text-gray-500 mb-1">🔒 Control PIN</p>
              <p className="font-mono text-2xl tracking-[0.3em] text-amber-400">{sessionPin}</p>
              <p className="text-[10px] text-gray-500 mt-1">Share this only with whoever should be able to control slides</p>
            </div>
          )}
          <div className="pt-3 border-t border-gray-700 text-xs">
            {connectedRemoteCount === 0 && <span className="text-gray-500">No remote connected yet</span>}
            {connectedRemoteCount === 1 && <span className="text-green-400">● 1 remote connected</span>}
            {connectedRemoteCount > 1 && (
              <span className="text-amber-400 font-bold">⚠ {connectedRemoteCount} remotes connected</span>
            )}
          </div>
        </div>
      </div>

      {/* Main slide area */}
      <div className="flex-1 flex flex-col min-w-0">
        {showNav && (
          <div className="flex items-center justify-between px-4 py-3 bg-gray-900 border-b border-gray-800">
            <button
              onClick={goPrev}
              disabled={isFirstSlide}
              className="px-4 py-2 rounded-lg bg-gray-800 hover:bg-gray-700 disabled:opacity-30 text-sm font-semibold"
            >
              Prev
            </button>
            <span className="text-sm text-gray-300">{navLabel}</span>
            <button
              onClick={goNext}
              disabled={isLastSlide}
              className="px-4 py-2 rounded-lg bg-gray-800 hover:bg-gray-700 disabled:opacity-30 text-sm font-semibold"
            >
              Next
            </button>
          </div>
        )}

        <div ref={wrapperRef} className="flex-1 flex items-center justify-center overflow-hidden relative bg-black">
          {loading && (
            <div className="flex flex-col items-center gap-3 text-gray-400">
              <div className="w-10 h-10 border-4 border-gray-700 border-t-blue-500 rounded-full animate-spin" />
              <span className="text-sm">Loading slide...</span>
            </div>
          )}

          {!loading && error && (
            <div className="flex flex-col items-center gap-4 text-center px-6">
              <span className="text-red-400 font-semibold">{error}</span>
              <button
                onClick={retry}
                className="px-5 py-2 rounded-lg bg-blue-600 hover:bg-blue-500 text-sm font-semibold"
              >
                Try again
              </button>
            </div>
          )}

          {/* Zoomable/pannable content layer - annotations zoom together with the slide. */}
          <div
            className="w-full h-full flex items-center justify-center"
            style={{ transform: zoomTransform, transformOrigin: 'center center', transition: 'transform 0.15s ease-out' }}
          >
            {!loading && !error && resolved?.fileType === 'pdf' && (
              <div className="w-full h-full flex items-center justify-center overflow-auto bg-white">
                <Document
                  file={resolved.blobUrl}
                  loading={<div className="p-12 text-black">Loading PDF...</div>}
                  onLoadSuccess={onPdfLoadSuccess}
                  onLoadError={(err) => setError(`PDF error: ${err.message}`)}
                >
                  <Page
                    pageNumber={currentPage}
                    renderTextLayer={false}
                    renderAnnotationLayer={false}
                    height={window.innerHeight * 0.85}
                  />
                </Document>
              </div>
            )}

            {!loading && !error && resolved?.fileType === 'image' && (
              <img
                src={resolved.blobUrl}
                alt={resolved.name || 'Slide image'}
                className="max-w-full max-h-full object-contain"
              />
            )}

            {!loading && !error && resolved?.fileType === 'video-link' && (
              <iframe
                ref={videoIframeRef}
                src={withPlaybackParams(resolved.embedUrl, resolved.platform)}
                title={resolved.name || 'Video'}
                className="w-full h-full border-0"
                allow="autoplay; fullscreen; picture-in-picture"
                allowFullScreen
              />
            )}

            {!loading && !error && resolved?.fileType === 'other' && (
              <div className="flex flex-col items-center gap-4 text-center px-6">
                <span className="text-gray-300">This file type can't be previewed inline.</span>
                <a href={resolved.blobUrl} download={resolved.name || 'download'} className="px-5 py-2 rounded-lg bg-blue-600 hover:bg-blue-500 text-sm font-semibold">Download file</a>
              </div>
            )}

            {/* Annotation layer - laser/draw/highlight/erase strokes broadcast from the phone now render here too. */}
            <canvas ref={presentCanvasRef} className="absolute top-0 left-0 w-full h-full pointer-events-none z-10" />
          </div>

          {/* Spotlight/focus mode - darkens everything outside a circle around the presenter's pointer. */}
          {spotlight.active && (
            <div
              className="absolute inset-0 pointer-events-none z-20"
              style={{
                background: `radial-gradient(circle at ${spotlight.x * 100}% ${spotlight.y * 100}%, transparent ${spotlight.radius}px, rgba(0,0,0,0.78) ${spotlight.radius + 70}px)`,
              }}
            />
          )}

          {/* Laser pointer dot, screen-fixed regardless of zoom. */}
          {laser.active && (
            <div
              className="absolute pointer-events-none z-30 rounded-full bg-red-500 shadow-[0_0_16px_#ef4444]"
              style={{ width: 16, height: 16, left: `${laser.x * 100}%`, top: `${laser.y * 100}%`, transform: 'translate(-50%, -50%)' }}
            />
          )}

          {/* Black/White screen - fully covers the stage, topmost layer. */}
          {screenMode !== 'normal' && (
            <div className={`absolute inset-0 z-50 ${screenMode === 'black' ? 'bg-black' : 'bg-white'}`} />
          )}
        </div>
      </div>
    </div>
  );
}
