import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { useParams, useSearchParams } from 'react-router-dom';
import { QRCodeSVG } from 'qrcode.react';
import { Document, Page, pdfjs } from 'react-pdf';
import { supabase } from './supabaseClient';
import { QuizReportCard, exportReportPDF, exportReportPNG, type QuizReportData } from './quizReport';
import { recordSavedItem } from './Account';
import type { SlideTransition } from './pptxParse';

// Best-effort lookup of whatever extractPptxMeta (see FileUpload.tsx) saved
// for a given uploaded file - speaker notes and transition info, keyed by
// slide/page number. Returns empty maps (never throws) if the table
// doesn't exist yet, the file wasn't a pptx, or nothing was ever saved for
// it - none of that should ever block a slide from loading.
async function fetchPptxMeta(fileId: string): Promise<{ notesByPage: Record<number, string>; transitionsByPage: Record<number, SlideTransition> }> {
  const empty = { notesByPage: {}, transitionsByPage: {} };
  try {
    const { data, error } = await supabase.from('pptx_meta').select('notes, transitions').eq('file_id', fileId).maybeSingle();
    if (error || !data) return empty;
    return { notesByPage: (data.notes as Record<number, string>) || {}, transitionsByPage: (data.transitions as Record<number, SlideTransition>) || {} };
  } catch {
    return empty;
  }
}

// --- Tiny synthesized beeps for the quiz countdown - Web Audio API only,
// no sound files to bundle/host. Only the presenter's screen plays these
// (not every student's phone), so a room full of phones doesn't all beep
// at once.
let sharedAudioCtx: AudioContext | null = null;
function getAudioCtx(): AudioContext | null {
  if (typeof window === 'undefined') return null;
  const AC = window.AudioContext || (window as any).webkitAudioContext;
  if (!AC) return null;
  if (!sharedAudioCtx) sharedAudioCtx = new AC();
  if (sharedAudioCtx.state === 'suspended') sharedAudioCtx.resume().catch(() => {});
  return sharedAudioCtx;
}
function playBeep(frequency: number, durationMs: number, volume = 0.15) {
  const ctx = getAudioCtx();
  if (!ctx) return;
  const osc = ctx.createOscillator();
  const gain = ctx.createGain();
  osc.type = 'sine';
  osc.frequency.value = frequency;
  gain.gain.value = volume;
  osc.connect(gain);
  gain.connect(ctx.destination);
  osc.start();
  gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + durationMs / 1000);
  osc.stop(ctx.currentTime + durationMs / 1000);
}
function playQuestionStartChime() { playBeep(523, 100, 0.12); setTimeout(() => playBeep(784, 140, 0.12), 110); }
function playTickBeep(urgent: boolean) { playBeep(urgent ? 660 : 440, 110, urgent ? 0.18 : 0.1); }
function playTimesUpChime() { playBeep(880, 150, 0.16); setTimeout(() => playBeep(392, 320, 0.16), 160); }

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
  transition?: SlideTransition; // this slide's own PPTX transition - how it should animate IN when navigated to
  thumbnail?: string; // small data-URL preview, shown on the remote's thumbnail strip
}

type ScreenMode = 'normal' | 'black' | 'white';
interface ZoomState { scale: number; x: number; y: number; } // x/y are pan offsets in % of container size
interface SpotlightState { x: number; y: number; active: boolean; radius: number; }
interface VideoState { playing: boolean; time: number; duration: number; volume: number; }
interface SessionState { screenMode: ScreenMode; zoom: ZoomState; videoState: VideoState; pin?: string; }

// Mirrors the countdown timer MobileRemote.tsx owns - Present.tsx never
// starts/stops it, just displays whatever the phone broadcasts (see the
// `timer_state` / `timer_alert` listeners below).
interface ProjectorTimerState { secondsLeft: number | null; running: boolean; visible: boolean; }
const DEFAULT_PROJECTOR_TIMER: ProjectorTimerState = { secondsLeft: null, running: false, visible: false };

const DEFAULT_ZOOM: ZoomState = { scale: 1, x: 0, y: 0 };
const DEFAULT_VIDEO_STATE: VideoState = { playing: false, time: 0, duration: 0, volume: 100 };
const DEFAULT_SESSION_STATE: SessionState = { screenMode: 'normal', zoom: DEFAULT_ZOOM, videoState: DEFAULT_VIDEO_STATE };

// --- Audience-facing live quiz ---------------------------------------------
// These shapes must stay in sync with AudienceJoin.tsx (voting/answer UI,
// leaderboard, Q&A, reactions) and MobileRemote.tsx (phone-side quiz
// builder/controls) - all three read/write the same `audience_state`
// column + `audience_state_update` broadcast, same pattern as session_state.
interface QuizOption { id: string; text: string; imageUrl?: string; }
// 'mcq' behaves exactly as before. 'short'/'long' skip options/scoring
// entirely - the audience types an answer instead of picking one, and
// there's no auto-grading (there's no way to auto-grade free text), just a
// flat participation score once they submit.
type QuizQuestionType = 'mcq' | 'short' | 'long';
interface QuizQuestion {
  id: string;
  type: QuizQuestionType;
  question: string;
  options: QuizOption[];
  correctOptionId: string;
  source?: string;          // reading material / citation shown after reveal
  timeLimitSeconds: number;
}
interface QuizAnswerRecord {
  optionId: string;        // mcq
  text?: string;            // short/long
  answeredAt: number;
  correct: boolean;
  points: number;
}
interface QuizParticipant {
  id: string;
  name: string;
  emoji?: string;                     // optional, picked at join; auto-assigned for top 3 if absent
  joinedAt: number;
  totalScore: number;
  answers: Record<string, QuizAnswerRecord>; // keyed by questionId
}
type QuizStatus = 'building' | 'lobby' | 'question' | 'reveal' | 'finished';
interface QuizState {
  questions: QuizQuestion[];
  currentIndex: number;               // -1 = lobby, not yet on a question
  status: QuizStatus;
  questionStartedAt: number | null;
  participants: Record<string, QuizParticipant>;
  // Which participant's short/long answer is currently featured big on the
  // projector, for the current question - either the teacher tapped their
  // card, or "🎲 Spotlight" picked one at random. Reset on every question
  // change.
  spotlightParticipantId?: string | null;
}
const DEFAULT_QUIZ_STATE: QuizState = { questions: [], currentIndex: -1, status: 'building', questionStartedAt: null, participants: {}, spotlightParticipantId: null };
interface SavedQuiz { id: string; title: string; questions: QuizQuestion[]; createdAt: number; }

// A palette of visually-distinct colors for the free-text answer cards, so
// each participant's card is easy to tell apart at a glance. Deterministic
// per participant (same idea as autoEmojiFor below).
const ANSWER_CARD_COLORS = ['#f87171', '#fb923c', '#fbbf24', '#a3e635', '#34d399', '#22d3ee', '#60a5fa', '#a78bfa', '#f472b6', '#fb7185'];
function answerCardColorFor(participantId: string): string {
  let hash = 0;
  for (let i = 0; i < participantId.length; i++) hash = (hash * 31 + participantId.charCodeAt(i)) >>> 0;
  return ANSWER_CARD_COLORS[hash % ANSWER_CARD_COLORS.length];
}

const CELEBRATION_EMOJIS = ['🎉', '🥳', '🌟', '🔥', '🚀', '⭐', '🎊', '💫'];
// Deterministic so the same participant always gets the same fallback emoji
// within a session, instead of it changing on every re-render.
function autoEmojiFor(participantId: string): string {
  let hash = 0;
  for (let i = 0; i < participantId.length; i++) hash = (hash * 31 + participantId.charCodeAt(i)) >>> 0;
  return CELEBRATION_EMOJIS[hash % CELEBRATION_EMOJIS.length];
}

function scoreAnswer(correct: boolean, answeredAt: number, questionStartedAt: number, timeLimitSeconds: number): number {
  if (!correct) return 0;
  const elapsed = Math.max(0, (answeredAt - questionStartedAt) / 1000);
  const timeLeftFraction = Math.max(0, 1 - elapsed / Math.max(1, timeLimitSeconds));
  return Math.round(500 + 500 * timeLeftFraction); // 500-1000 pts: correctness always worth something, speed adds a bonus
}
function rankParticipants(participants: Record<string, QuizParticipant>): QuizParticipant[] {
  return Object.values(participants).sort((a, b) => b.totalScore - a.totalScore);
}

interface AudienceQuestion { id: string; text: string; upvotes: number; answered: boolean; createdAt: number; }
type FeedbackKind = '👍' | '❤️' | '👏' | '🤔' | '🐢' | '🚀';
type FeedbackCounts = Record<FeedbackKind, number>;
const EMPTY_FEEDBACK: FeedbackCounts = { '👍': 0, '❤️': 0, '👏': 0, '🤔': 0, '🐢': 0, '🚀': 0 };
interface AudienceState {
  joinCount: number;
  quiz: QuizState;
  savedQuizzes: SavedQuiz[];
  questions: AudienceQuestion[]; // student-submitted questions - doubles as the "question bank" the teacher builds quizzes from
  feedback: FeedbackCounts;
  qnaOpen: boolean;
}
const DEFAULT_AUDIENCE_STATE: AudienceState = { joinCount: 0, quiz: DEFAULT_QUIZ_STATE, savedQuizzes: [], questions: [], feedback: EMPTY_FEEDBACK, qnaOpen: true };


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
// Maps a slide's extracted <p:transition> (see pptxParse.ts) to an actual
// CSS animation played when that slide comes into view. This is a
// deliberately simple, honest approximation of what PowerPoint itself does
// - a real crossfade/push between the outgoing and incoming slide would mean
// keeping both rendered and layered during the swap, which is a bigger
// rendering-pipeline change; this instead plays a one-shot "entrance"
// animation on the incoming slide, using the real type and duration read
// out of the file. Keyframes are defined once, in TRANSITION_KEYFRAMES_CSS
// below, and injected via a single <style> tag.
function transitionAnimationStyle(transition?: SlideTransition): CSSProperties {
  if (!transition || transition.kind === 'cut') return {};
  const durationMs = Math.max(80, transition.durationMs || 500);
  if (transition.kind === 'fade') {
    return { animation: `nextslide-fade-in ${durationMs}ms ease-out` };
  }
  const dir = transition.direction || 'l';
  const animName =
    dir === 'l' ? 'nextslide-slide-in-l' : dir === 'r' ? 'nextslide-slide-in-r' : dir === 'u' ? 'nextslide-slide-in-u' : 'nextslide-slide-in-d';
  return { animation: `${animName} ${durationMs}ms cubic-bezier(0.22,0.61,0.36,1)` };
}
const TRANSITION_KEYFRAMES_CSS = `
@keyframes nextslide-fade-in { from { opacity: 0; } to { opacity: 1; } }
@keyframes nextslide-slide-in-l { from { transform: translateX(6%); opacity: 0.3; } to { transform: translateX(0); opacity: 1; } }
@keyframes nextslide-slide-in-r { from { transform: translateX(-6%); opacity: 0.3; } to { transform: translateX(0); opacity: 1; } }
@keyframes nextslide-slide-in-u { from { transform: translateY(6%); opacity: 0.3; } to { transform: translateY(0); opacity: 1; } }
@keyframes nextslide-slide-in-d { from { transform: translateY(-6%); opacity: 0.3; } to { transform: translateY(0); opacity: 1; } }
`;

function withPlaybackParams(embedUrl: string, platform?: string): string {
  try {
    const url = new URL(embedUrl);
    if (isYouTubeEmbed(embedUrl, platform)) {
      url.searchParams.set('enablejsapi', '1');
      url.searchParams.set('autoplay', '1');
      url.searchParams.set('playsinline', '1');
      url.searchParams.set('origin', window.location.origin);
      return url.toString();
    }
    return embedUrl;
  } catch {
    return embedUrl;
  }
}

function isYouTubeEmbed(embedUrl: string, platform?: string): boolean {
  try {
    const url = new URL(embedUrl);
    return platform === 'youtube' || url.hostname.includes('youtube.com') || url.hostname.includes('youtu.be');
  } catch {
    return false;
  }
}

// Loads YouTube's official IFrame Player API script once (safe to call
// repeatedly/concurrently - later callers just await the same promise) and
// resolves once window.YT.Player is actually usable. This replaces the old
// approach of hand-rolling postMessage({event:'listening'}) pings and
// sniffing raw 'infoDelivery' messages, which only worked once the player
// happened to volunteer one on its own - in practice that meant the phone's
// progress bar stayed empty until something like a seek nudged it into
// talking. The real API fires a proper onReady event as soon as metadata
// is available, and getCurrentTime()/getDuration()/getVolume() can just be
// asked for directly instead of guessed at from whatever last flew by.
let ytApiLoadPromise: Promise<void> | null = null;
function loadYouTubeIframeAPI(): Promise<void> {
  if (typeof window === 'undefined') return Promise.resolve();
  const w = window as any;
  if (w.YT && w.YT.Player) return Promise.resolve();
  if (ytApiLoadPromise) return ytApiLoadPromise;
  ytApiLoadPromise = new Promise((resolve) => {
    const previous = w.onYouTubeIframeAPIReady;
    w.onYouTubeIframeAPIReady = () => {
      previous?.();
      resolve();
    };
    if (!document.querySelector('script[src="https://www.youtube.com/iframe_api"]')) {
      const tag = document.createElement('script');
      tag.src = 'https://www.youtube.com/iframe_api';
      document.head.appendChild(tag);
    }
  });
  return ytApiLoadPromise;
}

export default function Present() {
  const { fileId } = useParams<{ fileId: string }>();
  const [searchParams] = useSearchParams();

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
  const [projectorTimer, setProjectorTimer] = useState<ProjectorTimerState>(DEFAULT_PROJECTOR_TIMER);
  const [timerAlert, setTimerAlert] = useState<{ label: string; key: number; flashMs: number } | null>(null);
  useEffect(() => {
    if (!timerAlert) return;
    const id = setTimeout(() => setTimerAlert((cur) => (cur?.key === timerAlert.key ? null : cur)), timerAlert.flashMs);
    return () => clearTimeout(id);
  }, [timerAlert]);
  const [laser, setLaser] = useState({ x: 0.5, y: 0.5, active: false });
  const [videoState, setVideoState] = useState<VideoState>(DEFAULT_VIDEO_STATE);

  const presentCanvasRef = useRef<HTMLCanvasElement>(null);
  const presentCtxRef = useRef<CanvasRenderingContext2D | null>(null);
  const allDrawingsRef = useRef<CanvasDataMap>({});
  const currentLineRef = useRef<Point[]>([]);
  // Pixel coords of the last point drawn for the in-progress remote stroke.
  // See the draw_stroke handler below for why this exists (fixes highlighter
  // strokes rendering much darker than their final, redrawn appearance).
  const lastStrokePxRef = useRef<{ x: number; y: number } | null>(null);
  const videoIframeRef = useRef<HTMLIFrameElement>(null);
  // The real YT.Player instance bound to videoIframeRef - see the effect
  // further down that creates/destroys it as video slides come and go.
  const ytPlayerRef = useRef<any>(null);

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
  const audienceUrl = `${window.location.origin}${window.location.pathname}#/audience?session=${sessionId}`;
  const wrapperRef = useRef<HTMLDivElement>(null);
  // Separate from wrapperRef: this is the actual target for the real
  // Fullscreen API (see the comment above the JSX that uses it). wrapperRef
  // itself stays scoped to just the slide area, since it also drives the
  // annotation canvas's size (see the resize effect below) - widening it to
  // include the sidebar/quiz-stage would throw off every stroke's x/y math.
  const fullscreenTargetRef = useRef<HTMLDivElement>(null);
  const channelRef = useRef<any>(null);

  // Tracks real OS-level fullscreen (via the Fullscreen API), separate from
  // focusMode. Used to hide the Prev/Next/slide-name bar so real fullscreen
  // shows only the slide - the bar is still useful in the normal windowed
  // view, just not once the browser has taken over the whole screen.
  const [isFullscreen, setIsFullscreen] = useState(false);
  useEffect(() => {
    const handleFullscreenChange = () => setIsFullscreen(!!document.fullscreenElement);
    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  // Session PIN (soft lock, see the matching comment in MobileRemote.tsx) -
  // generated once when the session row is first created, then persisted
  // in session_state so refreshing this tab doesn't mint a new one.
  const [sessionPin, setSessionPin] = useState<string | null>(null);
  // Presence: how many phones are currently connected & controlling this
  // session, so the host can spot "two people are fighting over Next".
  const [connectedRemoteCount, setConnectedRemoteCount] = useState(0);
  const hostClientId = useMemo(() => `host_${Math.random().toString(36).slice(2)}`, []);

  // --- Audience live quiz ----------------------------------------------
  const [audienceState, setAudienceState] = useState<AudienceState>(DEFAULT_AUDIENCE_STATE);
  const audienceStateRef = useRef<AudienceState>(DEFAULT_AUDIENCE_STATE);
  useEffect(() => { audienceStateRef.current = audienceState; }, [audienceState]);
  const [quizPanelOpen, setQuizPanelOpen] = useState(false);
  const [presenterLang, setPresenterLang] = useState<'ku' | 'en'>('ku');
  const [quizStageMinimized, setQuizStageMinimized] = useState(false);

  // Draft questions, built up before "Start Quiz" is pressed.
  const [draftQuestions, setDraftQuestions] = useState<QuizQuestion[]>([]);
  const [qType, setQType] = useState<QuizQuestionType>('mcq');
  const [qText, setQText] = useState('');
  const [qOptions, setQOptions] = useState<{ text: string; imageUrl: string }[]>([{ text: '', imageUrl: '' }, { text: '', imageUrl: '' }]);
  const [qCorrectIndex, setQCorrectIndex] = useState<number | null>(null);
  const [qSource, setQSource] = useState('');
  const [qTimeLimit, setQTimeLimit] = useState(20);
  const [saveTitle, setSaveTitle] = useState('');

  const persistAudienceState = useCallback((next: AudienceState) => {
    audienceStateRef.current = next;
    setAudienceState(next);
    supabase.from('sessions').upsert({ id: sessionId, audience_state: next }).then(({ error }) => {
      if (error) console.error('🚨 audience_state upsert failed:', error.message, error);
    });
    channelRef.current?.send({ type: 'broadcast', event: 'audience_state_update', payload: { audienceState: next } });
  }, [sessionId]);

  const addDraftQuestion = () => {
    const question = qText.trim();
    if (!question) return;
    let newQ: QuizQuestion;
    if (qType === 'mcq') {
      const options = qOptions.map((o) => ({ text: o.text.trim(), imageUrl: o.imageUrl.trim() })).filter((o) => o.text || o.imageUrl);
      if (options.length < 2 || qCorrectIndex === null) return;
      newQ = {
        id: `q_${Date.now().toString(36)}`,
        type: 'mcq',
        question,
        options: options.map((o, i) => ({ id: `opt_${i}`, text: o.text, imageUrl: o.imageUrl || undefined })),
        correctOptionId: `opt_${qCorrectIndex}`,
        source: qSource.trim() || undefined,
        timeLimitSeconds: qTimeLimit,
      };
    } else {
      // Short/long answer: no options, nothing to auto-grade - the source
      // note is still useful context to show after reveal.
      newQ = {
        id: `q_${Date.now().toString(36)}`,
        type: qType,
        question,
        options: [],
        correctOptionId: '',
        source: qSource.trim() || undefined,
        timeLimitSeconds: qTimeLimit,
      };
    }
    setDraftQuestions((prev) => [...prev, newQ]);
    setQText(''); setQOptions([{ text: '', imageUrl: '' }, { text: '', imageUrl: '' }]); setQCorrectIndex(null); setQSource(''); setQTimeLimit(20);
  };
  const removeDraftQuestion = (id: string) => setDraftQuestions((prev) => prev.filter((q) => q.id !== id));

  // All quiz-flow transitions live here so both the presenter's own buttons
  // AND a `quiz_control` command arriving from the phone call the exact
  // same logic - single source of truth, presenter stays authoritative.
  const startQuizFlow = useCallback((questions: QuizQuestion[]) => {
    if (!questions.length) return;
    setQuizStageMinimized(false);
    persistAudienceState({ ...audienceStateRef.current, quiz: { questions, currentIndex: -1, status: 'lobby', questionStartedAt: null, participants: {}, spotlightParticipantId: null } });
  }, [persistAudienceState]);

  const advanceQuiz = useCallback(() => {
    const quiz = audienceStateRef.current.quiz;
    if (quiz.status === 'lobby') {
      persistAudienceState({ ...audienceStateRef.current, quiz: { ...quiz, currentIndex: 0, status: 'question', questionStartedAt: Date.now(), spotlightParticipantId: null } });
    } else if (quiz.status === 'reveal') {
      const nextIndex = quiz.currentIndex + 1;
      if (nextIndex >= quiz.questions.length) {
        persistAudienceState({ ...audienceStateRef.current, quiz: { ...quiz, status: 'finished' } });
      } else {
        persistAudienceState({ ...audienceStateRef.current, quiz: { ...quiz, currentIndex: nextIndex, status: 'question', questionStartedAt: Date.now(), spotlightParticipantId: null } });
      }
    }
  }, [persistAudienceState]);

  const revealQuizNow = useCallback(() => {
    const quiz = audienceStateRef.current.quiz;
    if (quiz.status !== 'question') return;
    persistAudienceState({ ...audienceStateRef.current, quiz: { ...quiz, status: 'reveal' } });
  }, [persistAudienceState]);

  // Featured a specific short/long answer big on the projector - or, if no
  // participantId is given, picks one at random from whoever's answered
  // the current question so far. Either the teacher tapping a card, or the
  // "🎲 Spotlight" button (on the projector or the remote), lands here.
  const spotlightAnswer = useCallback((participantId?: string) => {
    const quiz = audienceStateRef.current.quiz;
    const question = quiz.currentIndex >= 0 ? quiz.questions[quiz.currentIndex] : null;
    if (!question) return;
    let target = participantId;
    if (!target) {
      const answeredIds = Object.values(quiz.participants).filter((p) => p.answers[question.id]?.text).map((p) => p.id);
      if (!answeredIds.length) return;
      target = answeredIds[Math.floor(Math.random() * answeredIds.length)];
    }
    persistAudienceState({ ...audienceStateRef.current, quiz: { ...quiz, spotlightParticipantId: target } });
  }, [persistAudienceState]);

  const resetQuiz = useCallback(() => {
    setDraftQuestions([]);
    persistAudienceState({ ...audienceStateRef.current, quiz: DEFAULT_QUIZ_STATE });
  }, [persistAudienceState]);

  // Opened from the account dashboard's "▶ Open" on a saved quiz (see
  // Account.tsx's openQuiz): the dashboard can't reach into this specific
  // tab's React state directly (it's a brand new tab), so it hands the quiz
  // off through a one-time sessionStorage token referenced in the URL
  // instead. Auto-starts the quiz and opens the panel so it's ready to go
  // the moment the presenter has an audience.
  const presetQuizToken = searchParams.get('presetQuiz');
  const presetQuizAppliedRef = useRef(false);
  useEffect(() => {
    if (!presetQuizToken || presetQuizAppliedRef.current) return;
    presetQuizAppliedRef.current = true;
    try {
      const raw = sessionStorage.getItem(presetQuizToken);
      if (!raw) return;
      sessionStorage.removeItem(presetQuizToken);
      const { questions } = JSON.parse(raw);
      if (Array.isArray(questions) && questions.length) {
        startQuizFlow(questions);
        setQuizPanelOpen(true);
      }
    } catch (err) {
      console.warn('⚠️ Could not load the preset quiz:', err);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [presetQuizToken]);

  // Saved quizzes - prepared once, started whenever ("any day, same lesson
  // link"). Lives in the same audience_state row as everything else, so no
  // extra Supabase setup is needed - just note this ties saved quizzes to
  // whichever browser/device is used to open the presenter link, since
  // that's what session_id itself is tied to (see the sessionId comment
  // above). Same device each time -> saved quizzes are always there.
  const saveQuiz = useCallback((title: string, questions: QuizQuestion[]) => {
    if (!title.trim() || !questions.length) return;
    const saved: SavedQuiz = { id: `sv_${Date.now().toString(36)}`, title: title.trim(), questions, createdAt: Date.now() };
    persistAudienceState({ ...audienceStateRef.current, savedQuizzes: [...audienceStateRef.current.savedQuizzes, saved] });
    // Also lands in the presenter's account dashboard (if logged in), so
    // it's reachable from anywhere, not just from inside this one file's
    // saved-quizzes list.
    recordSavedItem({ kind: 'quiz', title: title.trim(), questions });
  }, [persistAudienceState]);
  const deleteSavedQuiz = useCallback((id: string) => {
    persistAudienceState({ ...audienceStateRef.current, savedQuizzes: audienceStateRef.current.savedQuizzes.filter((s) => s.id !== id) });
  }, [persistAudienceState]);

  // Auto-reveal when the countdown for the current question runs out -
  // the presenter tab is the single timer authority so everyone's clock
  // agrees, regardless of individual device clock drift.
  useEffect(() => {
    const quiz = audienceState.quiz;
    if (quiz.status !== 'question' || !quiz.questionStartedAt) return;
    const q = quiz.questions[quiz.currentIndex];
    if (!q) return;
    const msLeft = quiz.questionStartedAt + q.timeLimitSeconds * 1000 - Date.now();
    if (msLeft <= 0) { revealQuizNow(); return; }
    const t = setTimeout(revealQuizNow, msLeft);
    return () => clearTimeout(t);
  }, [audienceState.quiz.status, audienceState.quiz.questionStartedAt, audienceState.quiz.currentIndex, revealQuizNow, audienceState.quiz]);

  const quiz = audienceState.quiz;
  const currentQuestion = quiz.currentIndex >= 0 ? quiz.questions[quiz.currentIndex] : null;
  const leaderboard = rankParticipants(quiz.participants);

  const reportData: QuizReportData = useMemo(() => ({
    title: 'Quiz Results',
    dateLabel: new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }),
    leaderboard: leaderboard.map((p, i) => ({
      rank: i + 1,
      name: p.name,
      emoji: p.emoji || (i < 3 ? autoEmojiFor(p.id) : ''),
      score: p.totalScore,
      correctCount: Object.values(p.answers).filter((a) => a.correct).length,
      totalQuestions: quiz.questions.length,
    })),
    questions: quiz.questions.map((q) => {
      const answersForQ = leaderboard.map((p) => p.answers[q.id]).filter(Boolean) as QuizAnswerRecord[];
      return {
        id: q.id,
        question: q.question,
        correctText: q.options.find((o) => o.id === q.correctOptionId)?.text || '',
        source: q.source,
        correctCount: answersForQ.filter((a) => a.correct).length,
        incorrectCount: answersForQ.filter((a) => !a.correct).length,
      };
    }),
  }), [leaderboard, quiz.questions]);

  const reportNodeRef = useRef<HTMLDivElement>(null);
  const [exporting, setExporting] = useState<'pdf' | 'png' | null>(null);
  const downloadReport = async (format: 'pdf' | 'png') => {
    if (!reportNodeRef.current || exporting) return;
    setExporting(format);
    try {
      if (format === 'pdf') await exportReportPDF(reportNodeRef.current, `quiz-results-${sessionId}.pdf`);
      else await exportReportPNG(reportNodeRef.current, `quiz-results-${sessionId}.png`);
    } finally {
      setExporting(null);
    }
  };

  const toggleFullscreen = () => {
    if (!fullscreenTargetRef.current) return;
    if (!document.fullscreenElement) {
      fullscreenTargetRef.current.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen();
    }
  };

  // Theater/focus mode - hides the sidebar so the slide fills the screen.
  // This is what actually responds to a remote request from the phone:
  // the browser's real Fullscreen API can only be started by a direct user
  // gesture on THIS page (a click here works; a command arriving over the
  // network from the phone does not - browsers block that for security).
  // Focus mode gets you the same practical result (no sidebar, slide fills
  // the screen) without that restriction, so it's what the phone controls.
  const [focusMode, setFocusMode] = useState(false);
  const handleFullscreenRequest = useCallback(() => {
    setFocusMode((prev) => {
      const next = !prev;
      channelRef.current?.send({ type: 'broadcast', event: 'fullscreen_state', payload: { active: next } });
      return next;
    });
    // Also attempt the real OS-level fullscreen in case this happens to run
    // in a context where it's allowed - silently ignored if blocked.
    if (fullscreenTargetRef.current && !document.fullscreenElement) {
      fullscreenTargetRef.current.requestFullscreen().catch(() => {});
    }
  }, []);

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
        .select('id, canvas_data, session_state, audience_state')
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
          audience_state: DEFAULT_AUDIENCE_STATE,
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
        if (data.audience_state) {
          audienceStateRef.current = data.audience_state as AudienceState;
          setAudienceState(data.audience_state as AudienceState);
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
        const player = ytPlayerRef.current;
        if (!player) return;
        try {
          if (action === 'play') player.playVideo();
          else if (action === 'pause') player.pauseVideo();
          else if (action === 'seek' && typeof value === 'number') player.seekTo(value, true);
          else if (action === 'volume' && typeof value === 'number') player.setVolume(value);
          else if (action === 'mute') player.mute();
          else if (action === 'unmute') player.unMute();
        } catch {
          // Player not fully initialized yet - command is dropped, same
          // best-effort behavior as before.
        }
      });

      channel.on('broadcast', { event: 'fullscreen_toggle' }, () => {
        handleFullscreenRequest();
      });

      // Countdown timer mirror - MobileRemote.tsx is authoritative, this
      // only ever displays what it's told (see ProjectorTimerState above).
      channel.on('broadcast', { event: 'timer_state' }, (payload) => {
        const { secondsLeft, running, visible } = payload.payload || {};
        setProjectorTimer({
          secondsLeft: typeof secondsLeft === 'number' ? secondsLeft : null,
          running: !!running,
          visible: !!visible,
        });
      });
      channel.on('broadcast', { event: 'timer_alert' }, (payload) => {
        const { label, flashMs } = payload.payload || {};
        if (typeof label !== 'string') return;
        setTimerAlert({ label, key: Date.now(), flashMs: typeof flashMs === 'number' ? flashMs : 900 });
      });

      channel.on('broadcast', { event: 'quiz_join' }, (payload) => {
        const { participantId, name, emoji } = payload.payload || {};
        if (!participantId || !name) return;
        const current = audienceStateRef.current;
        if (current.quiz.participants[participantId]) return; // already joined
        const participants = { ...current.quiz.participants, [participantId]: { id: participantId, name, emoji: emoji || undefined, joinedAt: Date.now(), totalScore: 0, answers: {} } };
        persistAudienceState({ ...current, quiz: { ...current.quiz, participants } });
      });

      channel.on('broadcast', { event: 'quiz_answer' }, (payload) => {
        const { participantId, questionId, optionId, text, answeredAt } = payload.payload || {};
        if (!participantId || !questionId) return;
        const current = audienceStateRef.current;
        const quiz = current.quiz;
        const participant = quiz.participants[participantId];
        const question = quiz.questions.find((q) => q.id === questionId);
        if (!participant || !question || participant.answers[questionId]) return; // no double-answers

        let answerRecord: QuizAnswerRecord;
        if (question.type === 'short' || question.type === 'long') {
          if (!text || !String(text).trim()) return;
          // Free text isn't auto-gradable - flat participation credit for
          // submitting, same for everyone regardless of what they wrote.
          const FREE_TEXT_POINTS = 50;
          answerRecord = { optionId: '', text: String(text).trim(), answeredAt: answeredAt || Date.now(), correct: false, points: FREE_TEXT_POINTS };
        } else {
          if (!optionId) return;
          const correct = question.correctOptionId === optionId;
          const points = scoreAnswer(correct, answeredAt || Date.now(), quiz.questionStartedAt || Date.now(), question.timeLimitSeconds);
          answerRecord = { optionId, answeredAt: answeredAt || Date.now(), correct, points };
        }
        const updatedParticipant: QuizParticipant = {
          ...participant,
          totalScore: participant.totalScore + answerRecord.points,
          answers: { ...participant.answers, [questionId]: answerRecord },
        };
        persistAudienceState({ ...current, quiz: { ...quiz, participants: { ...quiz.participants, [participantId]: updatedParticipant } } });
      });

      channel.on('broadcast', { event: 'audience_join' }, () => {
        const current = audienceStateRef.current;
        persistAudienceState({ ...current, joinCount: current.joinCount + 1 });
      });

      // Student-submitted questions - these feed the "question bank" the
      // quiz builder pulls from (see the Quiz panel's "From student
      // questions" section), they're not just a live Q&A feed.
      channel.on('broadcast', { event: 'audience_question' }, (payload) => {
        const text = (payload.payload?.text || '').trim();
        if (!text) return;
        const current = audienceStateRef.current;
        const newQ: AudienceQuestion = { id: `sq_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`, text, upvotes: 0, answered: false, createdAt: Date.now() };
        persistAudienceState({ ...current, questions: [...current.questions, newQ] });
      });
      channel.on('broadcast', { event: 'audience_upvote' }, (payload) => {
        const { questionId } = payload.payload || {};
        if (!questionId) return;
        const current = audienceStateRef.current;
        const questions = current.questions.map((q) => (q.id === questionId ? { ...q, upvotes: q.upvotes + 1 } : q));
        persistAudienceState({ ...current, questions });
      });

      // Lets the phone remote fully drive quiz creation/flow too.
      channel.on('broadcast', { event: 'quiz_control' }, (payload) => {
        const { action, questions, participantId: spotlightId } = payload.payload || {};
        if (action === 'start_quiz' && Array.isArray(questions)) startQuizFlow(questions);
        else if (action === 'advance') advanceQuiz();
        else if (action === 'reveal_now') revealQuizNow();
        else if (action === 'spotlight_answer') spotlightAnswer(spotlightId || undefined);
        else if (action === 'reset') resetQuiz();
        else if (action === 'save_quiz' && Array.isArray(questions) && payload.payload?.title) saveQuiz(payload.payload.title, questions);
        else if (action === 'delete_saved' && payload.payload?.id) deleteSavedQuiz(payload.payload.id);
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
          currentLineRef.current = [{ x, y }];
          lastStrokePxRef.current = { x: pxX, y: pxY };
        } else if (type === 'move') {
          // Stroke only the new short segment (last point -> this point),
          // each in its own beginPath(). Re-stroking the whole accumulated
          // path on every move (the old behaviour) re-composites every
          // earlier segment's alpha again and again, so a translucent
          // highlighter stroke gets darker and darker as you draw - by the
          // time you lift your finger it's nearly opaque and can blot out
          // whatever's underneath. A fresh short segment per move keeps
          // each pixel's alpha basically constant with the live stroke,
          // matching how redrawCanvasForSlide renders it afterwards.
          const last = lastStrokePxRef.current;
          if (last) {
            ctx.beginPath();
            ctx.moveTo(last.x, last.y);
            ctx.lineTo(pxX, pxY);
            ctx.stroke();
          }
          lastStrokePxRef.current = { x: pxX, y: pxY };
          currentLineRef.current.push({ x, y });
        } else if (type === 'end') {
          ctx.restore();
          lastStrokePxRef.current = null;
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

  // Creates a real YT.Player bound to the video iframe whenever a YouTube
  // slide is on screen, and tears it down when that slide goes away. Once
  // it's ready, onReady/onStateChange plus a 400ms poll keep videoState
  // (playing/time/duration/volume) accurate and broadcast to the phone -
  // this is what makes the remote's progress bar show up reliably and its
  // volume slider actually track the real level, instead of both sitting
  // frozen until some other action happened to jostle the old postMessage
  // listener into life.
  useEffect(() => {
    if (resolved?.fileType !== 'video-link') return;
    if (!isYouTubeEmbed(resolved.embedUrl, resolved.platform)) return;
    const iframe = videoIframeRef.current;
    if (!iframe) return;

    let cancelled = false;
    let pollInterval: ReturnType<typeof setInterval> | null = null;

    const syncFromPlayer = () => {
      const player = ytPlayerRef.current;
      if (!player || typeof player.getPlayerState !== 'function') return;
      try {
        const next: VideoState = {
          playing: player.getPlayerState() === 1,
          time: player.getCurrentTime() || 0,
          duration: player.getDuration() || 0,
          volume: player.getVolume(),
        };
        setVideoState(next);
        channelRef.current?.send({ type: 'broadcast', event: 'video_time_update', payload: next });
      } catch {
        // Player exists but isn't fully initialized yet - next poll retries.
      }
    };

    // Reset immediately so the remote doesn't keep showing the previous
    // video's progress while this one's player spins up.
    setVideoState(DEFAULT_VIDEO_STATE);
    channelRef.current?.send({ type: 'broadcast', event: 'video_time_update', payload: DEFAULT_VIDEO_STATE });

    loadYouTubeIframeAPI().then(() => {
      if (cancelled) return;
      const YT = (window as any).YT;
      ytPlayerRef.current = new YT.Player(iframe, {
        events: { onReady: syncFromPlayer, onStateChange: syncFromPlayer },
      });
      pollInterval = setInterval(syncFromPlayer, 400);
    });

    return () => {
      cancelled = true;
      if (pollInterval) clearInterval(pollInterval);
      try { ytPlayerRef.current?.destroy?.(); } catch { /* iframe already gone */ }
      ytPlayerRef.current = null;
    };
  }, [resolved, currentFlatIndex]);

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
            const { notesByPage, transitionsByPage } = await fetchPptxMeta(ref.fileId);
            for (let p = 1; p <= n; p++) {
              const thumbnail = await renderPdfThumbnail(entry.blobUrl, p);
              result.push({ itemIndex: i, pageInItem: p, fileType: 'pdf', name: ref.name, notes: notesByPage[p] || ref.notes, transition: transitionsByPage[p], thumbnail });
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
        const { notesByPage, transitionsByPage } = fileId ? await fetchPptxMeta(fileId) : { notesByPage: {}, transitionsByPage: {} };
        const slides: FlatSlide[] = [];
        for (let i = 0; i < numPages; i++) {
          if (cancelled) return;
          const thumbnail = await renderPdfThumbnail(resolved.blobUrl, i + 1);
          slides.push({ itemIndex: 0, pageInItem: i + 1, fileType: 'pdf', notes: notesByPage[i + 1], transition: transitionsByPage[i + 1], thumbnail });
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
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

  // Hidden once the browser is truly fullscreen or the phone has put us in
  // focus mode - both cases want the slide to fill the whole screen with no
  // chrome. Prev/Next still work via keyboard arrows and the phone remote.
  const showNav = flatSlides.length > 1 && !isFullscreen && !focusMode;
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
      {!focusMode && (
      <div className="w-80 bg-gray-900 border-r border-gray-800 p-6 flex flex-col items-center justify-between shrink-0">
        <div className="w-full flex justify-end gap-2">
          <button onClick={() => setQuizPanelOpen(true)} className="text-xs bg-emerald-600 hover:bg-emerald-700 px-3 py-1 rounded relative">
            🧠 Quiz
            {quiz.status !== 'building' && quiz.status !== 'finished' && <span className="absolute -top-1 -right-1 w-2.5 h-2.5 rounded-full bg-red-500 animate-pulse" />}
          </button>
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
      )}

      {focusMode && (
        <button
          onClick={handleFullscreenRequest}
          className="fixed top-3 right-3 z-[200] bg-gray-900/80 hover:bg-gray-800 text-white text-xs px-3 py-1.5 rounded-full border border-gray-700"
        >
          🗗 Exit focus mode
        </button>
      )}

      {/* Main slide area, quiz stage, and timer overlay all live inside
          fullscreenTargetRef. The real Fullscreen API only renders
          descendants of whatever element it was called on - anything
          rendered as a *sibling* (even position:fixed) simply doesn't
          appear on screen once fullscreen is active. The quiz QR stage and
          projector timer both used to be siblings of the slide area, which
          is exactly why starting a quiz while in full screen used to show
          nothing until you exited full screen. */}
      <div ref={fullscreenTargetRef} className="flex-1 flex flex-col min-w-0 relative">
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
            {/* Transition layer - remounts (and replays its entrance
                animation) every time the flat slide index changes, using
                whatever transition pptxParse.ts read off this specific
                slide. flatSlides[currentFlatIndex] is undefined outside
                pptx/lesson mode (e.g. plain PDFs with no extracted meta),
                in which case this is just a plain instant cut, same as
                before. */}
            <div key={currentFlatIndex} className="w-full h-full flex items-center justify-center" style={transitionAnimationStyle(flatSlides[currentFlatIndex]?.transition)}>
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
                key={`video-${currentFlatIndex}`}
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
          </div>
          <style>{TRANSITION_KEYFRAMES_CSS}</style>

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

      {/* Fullscreen quiz stage - takes over the whole projector screen for
          every phase except building the quiz. This is what the audience
          watches; their phones only show the tappable options. */}
      {quiz.status !== 'building' && !(quiz.status === 'finished' && quizStageMinimized) && (
        <div className="fixed inset-0 z-[150] bg-gradient-to-br from-indigo-950 via-gray-950 to-black flex flex-col items-center justify-center text-white p-8 gap-6 overflow-y-auto">
          <button
            onClick={() => setPresenterLang((l) => (l === 'ku' ? 'en' : 'ku'))}
            className="fixed top-4 left-4 bg-white/10 hover:bg-white/20 px-3 py-1.5 rounded-full text-xs font-bold"
          >
            {presenterLang === 'ku' ? 'English' : 'کوردی'}
          </button>
          {quiz.status === 'finished' && (
            <button onClick={() => setQuizStageMinimized(true)} className="fixed top-4 right-4 bg-white/10 hover:bg-white/20 px-3 py-1.5 rounded-full text-xs font-bold">
              ✕ {ST(presenterLang, 'close')}
            </button>
          )}

          {quiz.status === 'lobby' && (
            <div className="flex flex-col items-center gap-6 text-center">
              <h1 className="text-3xl font-bold">{ST(presenterLang, 'scanQr')}</h1>
              <div className="bg-white p-6 rounded-3xl shadow-2xl shadow-indigo-500/30">
                <QRCodeSVG value={audienceUrl} size={340} />
              </div>
              <p className="text-indigo-300 font-mono text-sm break-all">{audienceUrl}</p>
              <div className="flex flex-wrap gap-2 justify-center max-w-2xl mt-2">
                {Object.values(quiz.participants).length === 0 && <span className="text-gray-400 text-sm">{ST(presenterLang, 'waitingJoin')}</span>}
                {Object.values(quiz.participants).map((p) => (
                  <span key={p.id} className="bg-indigo-600/30 border border-indigo-500 px-3 py-1 rounded-full text-sm">{p.name}</span>
                ))}
              </div>
              <button
                onClick={advanceQuiz}
                disabled={Object.values(quiz.participants).length === 0}
                className="mt-4 bg-emerald-600 disabled:opacity-30 hover:bg-emerald-500 px-8 py-3 rounded-full font-bold text-lg"
              >
                ▶ {ST(presenterLang, 'beginQuiz')}
              </button>
            </div>
          )}

          {(quiz.status === 'question' || quiz.status === 'reveal') && currentQuestion && (
            <QuizLiveStage
              quiz={quiz}
              question={currentQuestion}
              lang={presenterLang}
              onReveal={revealQuizNow}
              onAdvance={advanceQuiz}
              onSpotlight={spotlightAnswer}
              isLast={quiz.currentIndex >= quiz.questions.length - 1}
            />
          )}

          {quiz.status === 'finished' && (
            <div className="flex flex-col items-center gap-6 w-full max-w-2xl">
              <h1 className="text-3xl font-bold">🏆 {ST(presenterLang, 'leaderboard')}</h1>
              <LeaderboardList leaderboard={leaderboard} questionCount={quiz.questions.length} celebrate />
              <div className="flex gap-3">
                <button onClick={() => downloadReport('pdf')} disabled={!!exporting} className="bg-blue-600 hover:bg-blue-500 disabled:opacity-50 px-5 py-2.5 rounded-full font-bold text-sm">
                  {exporting === 'pdf' ? '…' : '⬇ PDF'}
                </button>
                <button onClick={() => downloadReport('png')} disabled={!!exporting} className="bg-blue-600 hover:bg-blue-500 disabled:opacity-50 px-5 py-2.5 rounded-full font-bold text-sm">
                  {exporting === 'png' ? '…' : '⬇ PNG'}
                </button>
                <button onClick={resetQuiz} className="bg-gray-700 hover:bg-gray-600 px-5 py-2.5 rounded-full font-bold text-sm">
                  🔄 {ST(presenterLang, 'newQuiz')}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Projector timer overlay - shown when the phone remote has "show on
          projector" enabled for the countdown timer. Rendered inside
          fullscreenTargetRef (see the comment above) so it stays visible
          while the presenter is in real full screen. */}
      <ProjectorTimer state={projectorTimer} alert={timerAlert} />
      </div>

      {/* Off-screen (not visually visible, but fully rendered so html2canvas
          can capture it) - shared PDF/PNG report node. */}
      <div style={{ position: 'fixed', top: 0, left: -9999, pointerEvents: 'none' }}>
        <div ref={reportNodeRef}><QuizReportCard data={reportData} /></div>
      </div>

      {quizPanelOpen && (
        <div className="fixed inset-0 z-[100] bg-black/70 flex items-center justify-center p-4" onClick={() => setQuizPanelOpen(false)}>
          <div
            className="bg-gray-900 border border-gray-700 rounded-2xl w-full max-w-lg max-h-[85vh] overflow-y-auto p-6 flex flex-col gap-5"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between">
              <h3 className="text-lg font-bold">🧠 Live Quiz</h3>
              <button onClick={() => setQuizPanelOpen(false)} className="text-gray-400 hover:text-white text-xl leading-none">✕</button>
            </div>

            {quiz.status !== 'building' ? (
              <div className="flex flex-col gap-3">
                <p className="text-sm text-gray-300">
                  Quiz in progress ({quiz.status}) - controls are on the big screen{quiz.status === 'finished' && quizStageMinimized ? ', or ' : ''}
                  {quiz.status === 'finished' && quizStageMinimized && (
                    <button onClick={() => setQuizStageMinimized(false)} className="text-blue-400 underline">reopen it</button>
                  )}
                  . You can also drive it entirely from the phone remote.
                </p>
                {quiz.status === 'finished' && (
                  <div className="flex gap-2">
                    <button onClick={() => downloadReport('pdf')} disabled={!!exporting} className="flex-1 bg-blue-600 hover:bg-blue-500 disabled:opacity-50 rounded-lg py-2 text-sm font-bold">{exporting === 'pdf' ? '…' : '⬇ PDF'}</button>
                    <button onClick={() => downloadReport('png')} disabled={!!exporting} className="flex-1 bg-blue-600 hover:bg-blue-500 disabled:opacity-50 rounded-lg py-2 text-sm font-bold">{exporting === 'png' ? '…' : '⬇ PNG'}</button>
                    <button onClick={resetQuiz} className="flex-1 bg-gray-700 hover:bg-gray-600 rounded-lg py-2 text-sm font-bold">🔄 New quiz</button>
                  </div>
                )}
              </div>
            ) : (
              <div className="flex flex-col gap-4">
                <div className="flex items-center gap-4 bg-gray-800 rounded-xl p-4">
                  <div className="bg-white p-2 rounded-lg shrink-0">
                    <QRCodeSVG value={audienceUrl} size={72} />
                  </div>
                  <p className="text-xs text-gray-400">Shown big on the projector once the quiz starts - audience scans it there, no need to share it separately.</p>
                </div>

                {audienceState.savedQuizzes.length > 0 && (
                  <div className="flex flex-col gap-2 bg-gray-800/40 rounded-xl p-3">
                    <p className="text-xs font-bold text-gray-400 uppercase">📚 Saved quizzes</p>
                    {audienceState.savedQuizzes.map((sq) => (
                      <div key={sq.id} className="flex items-center gap-2 bg-gray-800/70 rounded-lg px-3 py-2">
                        <span className="flex-1 text-xs truncate">{sq.title} <span className="text-gray-500">({sq.questions.length}q)</span></span>
                        <button onClick={() => startQuizFlow(sq.questions)} className="text-[11px] bg-emerald-600 hover:bg-emerald-500 px-2.5 py-1 rounded-full font-bold shrink-0">▶ Start</button>
                        <button onClick={() => deleteSavedQuiz(sq.id)} className="text-gray-500 hover:text-red-400 text-xs shrink-0">✕</button>
                      </div>
                    ))}
                  </div>
                )}

                {audienceState.questions.length > 0 && (
                  <div className="flex flex-col gap-2 bg-gray-800/40 rounded-xl p-3">
                    <p className="text-xs font-bold text-gray-400 uppercase">📥 From student questions</p>
                    {[...audienceState.questions].sort((a, b) => b.upvotes - a.upvotes).map((sq) => (
                      <div key={sq.id} className="flex items-center gap-2 bg-gray-800/70 rounded-lg px-3 py-2">
                        <span className="flex-1 text-xs truncate">{sq.text} <span className="text-gray-500">({sq.upvotes} 👍)</span></span>
                        <button onClick={() => setQText(sq.text)} className="text-[11px] bg-blue-600 hover:bg-blue-500 px-2.5 py-1 rounded-full font-bold shrink-0">Use</button>
                      </div>
                    ))}
                    <p className="text-[10px] text-gray-500">Fills the question box below - add options and mark the correct answer, then "Add to quiz".</p>
                  </div>
                )}

                {draftQuestions.length > 0 && (
                  <div className="flex flex-col gap-2">
                    <p className="text-xs font-bold text-gray-400 uppercase">{draftQuestions.length} question{draftQuestions.length === 1 ? '' : 's'} added</p>
                    {draftQuestions.map((q, i) => (
                      <div key={q.id} className="bg-gray-800/60 rounded-lg p-2.5 flex items-center justify-between gap-2">
                        <span className="text-xs truncate">{i + 1}. {q.question}</span>
                        <button onClick={() => removeDraftQuestion(q.id)} className="text-gray-500 hover:text-red-400 text-xs shrink-0">✕</button>
                      </div>
                    ))}
                  </div>
                )}

                <div className="flex flex-col gap-3 border-t border-gray-800 pt-4">
                  <p className="text-xs font-bold text-gray-400 uppercase">Add question</p>

                  <div className="flex gap-1.5">
                    {(['mcq', 'short', 'long'] as QuizQuestionType[]).map((t) => (
                      <button
                        key={t}
                        onClick={() => setQType(t)}
                        className={`flex-1 rounded-lg py-1.5 text-xs font-bold ${qType === t ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-400'}`}
                      >
                        {t === 'mcq' ? 'Multiple choice' : t === 'short' ? 'Short answer' : 'Long answer'}
                      </button>
                    ))}
                  </div>

                  <input value={qText} onChange={(e) => setQText(e.target.value)} placeholder="Ask a question..." className="bg-gray-800 border border-gray-700 rounded-lg px-3 py-2 text-sm" />

                  {qType === 'mcq' && (
                  <div className="flex flex-col gap-2">
                    {qOptions.map((opt, i) => (
                      <div key={i} className="flex items-center gap-2">
                        <button
                          onClick={() => setQCorrectIndex(i)}
                          title="Mark as correct answer"
                          className={`shrink-0 w-6 h-6 rounded-full border-2 flex items-center justify-center text-xs ${qCorrectIndex === i ? 'bg-emerald-600 border-emerald-500' : 'border-gray-600'}`}
                        >
                          {qCorrectIndex === i ? '✓' : ''}
                        </button>
                        <input
                          value={opt.text}
                          onChange={(e) => setQOptions((prev) => prev.map((o, j) => (j === i ? { ...o, text: e.target.value } : o)))}
                          placeholder={`Option ${i + 1}`}
                          className="flex-1 bg-gray-800 border border-gray-700 rounded-lg px-3 py-1.5 text-sm"
                        />
                        {qOptions.length > 2 && (
                          <button onClick={() => setQOptions((prev) => prev.filter((_, j) => j !== i))} className="shrink-0 text-gray-500 hover:text-red-400 text-sm">✕</button>
                        )}
                      </div>
                    ))}
                    {qOptions.length < 6 && (
                      <button onClick={() => setQOptions((prev) => [...prev, { text: '', imageUrl: '' }])} className="text-xs text-blue-400 self-start">+ Add option</button>
                    )}
                    <p className="text-[10px] text-gray-500">Tap the circle to mark the correct answer. Options can optionally have an image too:</p>
                    {qOptions.map((opt, i) => (
                      <input
                        key={i}
                        value={opt.imageUrl}
                        onChange={(e) => setQOptions((prev) => prev.map((o, j) => (j === i ? { ...o, imageUrl: e.target.value } : o)))}
                        placeholder={`Image URL for option ${i + 1} (optional)`}
                        className="bg-gray-800/60 border border-gray-800 rounded-lg px-3 py-1 text-[11px] text-gray-400"
                      />
                    ))}
                  </div>
                  )}
                  {qType !== 'mcq' && (
                    <p className="text-[10px] text-gray-500">
                      Everyone types their own {qType === 'short' ? 'short' : 'long'} answer - no options, nothing auto-graded. Their name and answer show up as a card on the projector once they submit.
                    </p>
                  )}

                  <input value={qSource} onChange={(e) => setQSource(e.target.value)} placeholder="Source / further reading (optional)" className="bg-gray-800 border border-gray-700 rounded-lg px-3 py-2 text-xs" />

                  <div className="flex items-center gap-2">
                    <span className="text-xs text-gray-400">⏱ Time limit</span>
                    <input type="range" min={5} max={60} step={5} value={qTimeLimit} onChange={(e) => setQTimeLimit(Number(e.target.value))} className="flex-1" />
                    <span className="text-xs font-mono w-10 text-right">{qTimeLimit}s</span>
                  </div>

                  <button
                    onClick={addDraftQuestion}
                    disabled={!qText.trim() || (qType === 'mcq' && qCorrectIndex === null)}
                    className="bg-gray-700 disabled:opacity-30 hover:bg-gray-600 rounded-lg py-2 text-sm font-bold"
                  >
                    + Add to quiz
                  </button>
                </div>

                <div className="flex gap-2">
                  <input
                    value={saveTitle}
                    onChange={(e) => setSaveTitle(e.target.value)}
                    placeholder="Quiz title (to save for later)"
                    className="flex-1 bg-gray-800 border border-gray-700 rounded-lg px-3 py-2 text-xs"
                  />
                  <button
                    onClick={() => { saveQuiz(saveTitle, draftQuestions); setSaveTitle(''); }}
                    disabled={draftQuestions.length === 0 || !saveTitle.trim()}
                    className="bg-gray-700 disabled:opacity-30 hover:bg-gray-600 px-4 rounded-lg text-xs font-bold shrink-0"
                  >
                    💾 Save
                  </button>
                </div>

                <button
                  onClick={() => startQuizFlow(draftQuestions)}
                  disabled={draftQuestions.length === 0}
                  className="bg-emerald-600 disabled:opacity-30 rounded-lg py-2.5 text-sm font-bold"
                >
                  🚀 Start quiz ({draftQuestions.length} question{draftQuestions.length === 1 ? '' : 's'})
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// --- Small shared text dictionary for the audience-facing big-screen stage.
const STAGE_TEXT: Record<'ku' | 'en', Record<string, string>> = {
  ku: {
    scanQr: 'تکایە کۆدی QR بسکان بکە بۆ بەشداریکردن', waitingJoin: 'چاوەڕێی بەشداربووان...', beginQuiz: 'دەستپێکردنی کویز',
    question: 'پرسیار', timeLeft: 'کاتی ماوە', correctAnswer: 'وەڵامی ڕاست', leaderboard: 'پێشەنگەکان',
    downloadResults: 'داگرتنی ئەنجامەکان', newQuiz: 'کویزی نوێ', next: 'دواتر', revealNow: 'دەرخستنی ئێستا',
    source: 'سەرچاوە', answered: 'وەڵامیان دایەوە', close: 'داخستن',
  },
  en: {
    scanQr: 'Please scan the QR code to join', waitingJoin: 'Waiting for participants...', beginQuiz: 'Begin quiz',
    question: 'Question', timeLeft: 'Time left', correctAnswer: 'Correct answer', leaderboard: 'Leaderboard',
    downloadResults: 'Download results', newQuiz: 'New quiz', next: 'Next', revealNow: 'Reveal now',
    source: 'Source', answered: 'answered', close: 'Close',
  },
};
function ST(lang: 'ku' | 'en', key: string) { return STAGE_TEXT[lang][key] || key; }

function QuizLiveStage({ quiz, question, lang, onReveal, onAdvance, onSpotlight, isLast }: {
  quiz: QuizState; question: QuizQuestion; lang: 'ku' | 'en'; onReveal: () => void; onAdvance: () => void; onSpotlight: (participantId?: string) => void; isLast: boolean;
}) {
  const [, tick] = useState(0);
  useEffect(() => {
    if (quiz.status !== 'question') return;
    const i = setInterval(() => tick((n) => n + 1), 250);
    return () => clearInterval(i);
  }, [quiz.status]);

  const secondsLeft = quiz.questionStartedAt
    ? Math.max(0, Math.ceil((quiz.questionStartedAt + question.timeLimitSeconds * 1000 - Date.now()) / 1000))
    : question.timeLimitSeconds;

  const lastBeepSecondRef = useRef<number | null>(null);
  const timesUpPlayedRef = useRef(false);
  useEffect(() => {
    if (quiz.status === 'question') playQuestionStartChime();
    lastBeepSecondRef.current = null;
    timesUpPlayedRef.current = false;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [question.id]);
  useEffect(() => {
    if (quiz.status !== 'question') return;
    if (secondsLeft >= 1 && secondsLeft <= 5 && lastBeepSecondRef.current !== secondsLeft) {
      lastBeepSecondRef.current = secondsLeft;
      playTickBeep(secondsLeft <= 3);
    }
    if (secondsLeft === 0 && !timesUpPlayedRef.current) {
      timesUpPlayedRef.current = true;
      playTimesUpChime();
    }
  }, [secondsLeft, quiz.status]);

  const answeredCount = Object.values(quiz.participants).filter((p) => p.answers[question.id]).length;
  const totalParticipants = Object.values(quiz.participants).length;
  const pct = Math.max(0, Math.min(100, (secondsLeft / question.timeLimitSeconds) * 100));

  const votesByOption: Record<string, number> = {};
  question.options.forEach((o) => { votesByOption[o.id] = 0; });
  Object.values(quiz.participants).forEach((p) => {
    const a = p.answers[question.id];
    if (a) votesByOption[a.optionId] = (votesByOption[a.optionId] || 0) + 1;
  });
  const totalAnswers = Object.values(votesByOption).reduce((s, n) => s + n, 0);
  const palette = ['bg-red-600', 'bg-blue-600', 'bg-amber-500', 'bg-emerald-600', 'bg-purple-600', 'bg-pink-600'];

  return (
    <div className="w-full max-w-3xl flex flex-col items-center gap-6">
      <div className="flex items-center gap-3 text-sm text-indigo-300">
        <span>{ST(lang, 'question')} {quiz.currentIndex + 1} / {quiz.questions.length}</span>
        <span>•</span>
        <span>{answeredCount}/{totalParticipants} {ST(lang, 'answered')}</span>
      </div>
      <h1 className="text-3xl font-bold text-center">{question.question}</h1>

      {quiz.status === 'question' && (
        <div className="w-full max-w-md h-3 bg-white/10 rounded-full overflow-hidden">
          <div className={`h-full transition-all duration-300 ${secondsLeft <= 5 ? 'bg-red-500' : 'bg-emerald-500'}`} style={{ width: `${pct}%` }} />
        </div>
      )}
      {quiz.status === 'question' && <span className="text-4xl font-mono font-bold">{secondsLeft}s</span>}

      {question.type !== 'mcq' ? (
        <FreeTextAnswerStage quiz={quiz} question={question} onSpotlight={onSpotlight} lang={lang} />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 w-full">
          {question.options.map((opt, i) => {
            const isCorrect = quiz.status === 'reveal' && opt.id === question.correctOptionId;
            const votes = votesByOption[opt.id] || 0;
            const optPct = totalAnswers ? Math.round((votes / totalAnswers) * 100) : 0;
            return (
              <div
                key={opt.id}
                className={`relative rounded-2xl p-4 overflow-hidden font-semibold text-left ${palette[i % palette.length]} ${quiz.status === 'reveal' && !isCorrect ? 'opacity-40' : ''} ${isCorrect ? 'ring-4 ring-white' : ''}`}
              >
                {quiz.status === 'reveal' && (
                  <div className="absolute inset-0 bg-black/30" style={{ width: `${optPct}%` }} />
                )}
                <div className="relative flex items-center gap-3">
                  {opt.imageUrl && <img src={opt.imageUrl} alt="" className="w-12 h-12 object-cover rounded-lg" />}
                  <span className="flex-1">{opt.text}</span>
                  {isCorrect && <span className="text-xl">✓</span>}
                  {quiz.status === 'reveal' && <span className="text-xs font-mono">{votes} · {optPct}%</span>}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {quiz.status === 'reveal' && question.source && (
        <p className="text-xs text-indigo-300 max-w-lg text-center">📚 {ST(lang, 'source')}: {question.source}</p>
      )}

      <div className="flex gap-3 mt-2">
        {quiz.status === 'question' && (
          <button onClick={onReveal} className="bg-white/10 hover:bg-white/20 px-6 py-2.5 rounded-full font-bold text-sm">
            {ST(lang, 'revealNow')}
          </button>
        )}
        {quiz.status === 'reveal' && (
          <button onClick={onAdvance} className="bg-emerald-600 hover:bg-emerald-500 px-8 py-3 rounded-full font-bold text-lg">
            {isLast ? `🏆 ${ST(lang, 'leaderboard')}` : `${ST(lang, 'next')} ▶`}
          </button>
        )}
      </div>
    </div>
  );
}

// Grid of every submitted short/long answer - name on top of the card, each
// participant in their own consistent color (see answerCardColorFor above),
// so a room full of answers is still easy to scan and tell apart. Clicking
// any card - or the "🎲 Spotlight" button for a random pick - features it
// big in the middle with the rest dimmed behind, per the "select one to
// show big, others behind" request.
function FreeTextAnswerStage({ quiz, question, onSpotlight, lang }: {
  quiz: QuizState; question: QuizQuestion; onSpotlight: (participantId?: string) => void; lang: 'ku' | 'en';
}) {
  const answered = Object.values(quiz.participants)
    .filter((p) => p.answers[question.id]?.text)
    .sort((a, b) => (a.answers[question.id]?.answeredAt || 0) - (b.answers[question.id]?.answeredAt || 0));
  const spotlighted = quiz.spotlightParticipantId ? quiz.participants[quiz.spotlightParticipantId] : null;
  const spotlightedAnswer = spotlighted?.answers[question.id]?.text;

  if (!answered.length) {
    return <p className="text-indigo-300 text-sm py-8">{lang === 'ku' ? 'چاوەڕوانی وەڵام...' : 'Waiting for answers...'}</p>;
  }

  return (
    <div className="w-full flex flex-col items-center gap-5">
      <button
        onClick={() => onSpotlight()}
        className="bg-white/10 hover:bg-white/20 px-5 py-2 rounded-full font-bold text-sm flex items-center gap-2"
      >
        🎲 {lang === 'ku' ? 'وەڵامێک هەڵبژێرە' : 'Spotlight a random answer'}
      </button>

      {spotlighted && spotlightedAnswer && (
        <div className="w-full max-w-2xl rounded-3xl p-6 shadow-2xl border-4" style={{ borderColor: answerCardColorFor(spotlighted.id), background: `${answerCardColorFor(spotlighted.id)}22` }}>
          <div className="flex items-center gap-2 mb-3">
            <span className="text-2xl">{spotlighted.emoji || autoEmojiFor(spotlighted.id)}</span>
            <span className="font-bold text-lg">{spotlighted.name}</span>
          </div>
          <p className="text-xl leading-relaxed whitespace-pre-wrap">{spotlightedAnswer}</p>
        </div>
      )}

      <div className={`grid grid-cols-2 sm:grid-cols-3 gap-3 w-full transition-opacity ${spotlighted ? 'opacity-50' : ''}`}>
        {answered.map((p) => {
          const text = p.answers[question.id]?.text || '';
          const color = answerCardColorFor(p.id);
          const isSpotlighted = quiz.spotlightParticipantId === p.id;
          return (
            <button
              key={p.id}
              onClick={() => onSpotlight(p.id)}
              className={`text-left rounded-xl p-3 border-2 hover:scale-[1.02] transition-transform ${isSpotlighted ? 'ring-2 ring-white' : ''}`}
              style={{ borderColor: color, background: `${color}18` }}
            >
              <p className="text-xs font-bold truncate mb-1">{p.emoji || autoEmojiFor(p.id)} {p.name}</p>
              <p className="text-sm line-clamp-3 whitespace-pre-wrap">{text}</p>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function LeaderboardList({ leaderboard, questionCount, celebrate }: { leaderboard: QuizParticipant[]; questionCount: number; celebrate?: boolean }) {
  const medals = ['🥇', '🥈', '🥉'];
  return (
    <div className="w-full flex flex-col gap-2">
      {leaderboard.length === 0 && <p className="text-gray-400 text-center text-sm">No participants.</p>}
      {leaderboard.map((p, i) => {
        const correctCount = Object.values(p.answers).filter((a) => a.correct).length;
        const winnerEmoji = i < 3 ? (p.emoji || autoEmojiFor(p.id)) : null;
        return (
          <div
            key={p.id}
            className={`flex items-center gap-3 rounded-xl px-4 py-3 ${i === 0 ? 'bg-gradient-to-r from-amber-500/30 to-amber-600/10 border border-amber-500' : i < 3 ? 'bg-white/10 border border-white/20' : 'bg-white/5'}`}
          >
            <span className="text-xl w-8 text-center shrink-0">{medals[i] || `#${i + 1}`}</span>
            {winnerEmoji && celebrate && (
              <span className="text-2xl shrink-0" style={{ animation: `bounce-emoji 0.9s ease-in-out ${i * 0.15}s infinite` }}>{winnerEmoji}</span>
            )}
            <span className="flex-1 font-semibold truncate">{p.name}</span>
            <span className="text-xs text-gray-400">{correctCount}/{questionCount} ✓</span>
            <span className="font-mono font-bold text-lg w-16 text-right">{p.totalScore}</span>
          </div>
        );
      })}
      {celebrate && (
        <style>{`@keyframes bounce-emoji { 0%,100% { transform: translateY(0) scale(1); } 50% { transform: translateY(-6px) scale(1.2); } }`}</style>
      )}
    </div>
  );
}

// Small persistent countdown chip (bottom corner) plus, at the 59s/3/2/1/0
// thresholds, a big centered flash - both purely a mirror of whatever
// MobileRemote.tsx's timer is doing (see ProjectorTimerState above). Renders
// nothing at all unless the presenter has turned "show on projector" on.
function formatProjectorTime(totalSeconds: number) {
  const m = Math.floor(totalSeconds / 60).toString().padStart(2, '0');
  const s = Math.floor(totalSeconds % 60).toString().padStart(2, '0');
  return `${m}:${s}`;
}
function ProjectorTimer({ state, alert }: { state: ProjectorTimerState; alert: { label: string; key: number; flashMs: number } | null }) {
  if (!state.visible || state.secondsLeft === null) return null;
  const urgent = state.secondsLeft <= 60;
  return (
    <>
      <div
        className={`fixed bottom-5 right-5 z-[250] font-mono font-bold rounded-2xl px-5 py-2.5 text-3xl shadow-2xl border ${
          state.secondsLeft === 0
            ? 'bg-red-600/90 text-white border-red-400 animate-pulse'
            : urgent
            ? 'bg-amber-500/90 text-black border-amber-300 animate-pulse'
            : 'bg-black/70 text-white border-white/20'
        }`}
      >
        ⌛ {formatProjectorTime(state.secondsLeft)}
      </div>
      {alert && (
        <div key={alert.key} className="fixed inset-0 z-[280] flex items-center justify-center pointer-events-none bg-black/50">
          <span
            className={`font-mono font-black drop-shadow-2xl ${alert.label === '0' ? 'text-red-500' : 'text-amber-400'}`}
            style={{ fontSize: '22vw', lineHeight: 1, animation: 'timer-flash-pop 0.35s ease-out' }}
          >
            {alert.label === '0' ? '⏰' : alert.label}
          </span>
          <style>{`@keyframes timer-flash-pop { 0% { transform: scale(0.5); opacity: 0; } 60% { transform: scale(1.15); opacity: 1; } 100% { transform: scale(1); opacity: 1; } }`}</style>
        </div>
      )}
    </>
  );
}
