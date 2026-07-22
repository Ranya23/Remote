import { useState, useEffect, useRef } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import { Document, Page, pdfjs } from 'react-pdf';
import { supabase } from './supabaseClient';

pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

const texts = {
  ku: { scan: 'سکان بکە', session: 'سێشن', slide: 'سلاید', switchLang: 'EN', loading: 'خوێندنەوەی فایلی PDF...', error: 'هەڵە', fullscreen: 'شاشەی تەواو' },
  en: { scan: 'Scan to control', session: 'Session', slide: 'Slide', switchLang: 'کوردی', loading: 'Downloading secure PDF...', error: 'Error', fullscreen: 'Full Screen' }
};

interface Point { x: number; y: number; }
type Line = Point[];
type CanvasDataMap = Record<number, Line[]>;

const GAS_URL = 'https://script.google.com/macros/s/AKfycbx48s5aNamkERYuvJ-BE7-RBF2zt15mFZ-C-SXL_UIGZkG46RdyuPYIOlO6o0HZcr3N/exec';

export default function PresentationView() {
  const currentUrl = window.location.href;
  const match = currentUrl.match(/\/present\/([a-zA-Z0-9_-]+)/);
  const fileId = match ? match[1] : 'default';
    
  const [sessionId] = useState(() => {
    const saved = localStorage.getItem(`nextslide_session_${fileId}`);
    if (saved) return saved;
    const newId = Math.random().toString(36).substring(2, 9);
    localStorage.setItem(`nextslide_session_${fileId}`, newId);
    return newId;
  });

  const [currentSlide, setCurrentSlide] = useState(1);
  const [laser, setLaser] = useState({ x: 0, y: 0, active: false });
  const [lang, setLang] = useState<'ku' | 'en'>('ku');
  const t = texts[lang];

  const currentSlideRef = useRef(1); 
  const [pdfData, setPdfData] = useState<string | null>(null);
  const [pdfError, setPdfError] = useState<string | null>(null);

  const allDrawingsRef = useRef<CanvasDataMap>({});
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const wrapperRef = useRef<HTMLDivElement>(null);
  const ctxRef = useRef<CanvasRenderingContext2D | null>(null);
  const currentLineRef = useRef<Line>([]);

  const updateSlide = (newSlide: number) => {
    setCurrentSlide(newSlide);
    currentSlideRef.current = newSlide;
  };

  useEffect(() => {
    const fetchPdf = async () => {
      if (!fileId || fileId === 'default') return setPdfError(`Cannot find ID in URL`);
      try {
        const response = await fetch(`${GAS_URL}?action=getPdf&fileId=${fileId}`);
        const result = await response.json();
        if (result.status === 'success' && result.data) {
          setPdfData(`data:application/pdf;base64,${result.data}`);
        } else {
          setPdfError(result.message || 'Failed to download PDF');
        }
      } catch (error: any) {
        setPdfError(error.message || 'Network error');
      }
    };
    fetchPdf();
  }, [fileId]);

  const redrawCanvasForSlide = (slideNum: number, drawingsMap: CanvasDataMap) => {
    const canvas = canvasRef.current;
    const ctx = ctxRef.current;
    if (!canvas || !ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const lines = drawingsMap[slideNum] || [];
    lines.forEach((line) => {
      if (line.length === 0) return;
      ctx.beginPath();
      ctx.moveTo(line[0].x * canvas.width, line[0].y * canvas.height);
      for (let i = 1; i < line.length; i++) {
        ctx.lineTo(line[i].x * canvas.width, line[i].y * canvas.height);
      }
      ctx.stroke();
      ctx.closePath();
    });
  };

  const saveCanvasToDatabase = async (updatedDrawings: CanvasDataMap) => {
    await supabase.from('sessions').upsert({ id: sessionId, file_id: fileId, current_slide: currentSlideRef.current, canvas_data: updatedDrawings });
  };

  const toggleFullscreen = () => {
    if (!wrapperRef.current) return;
    if (!document.fullscreenElement) {
      wrapperRef.current.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen();
    }
  };

  useEffect(() => {
    if (canvasRef.current && wrapperRef.current) {
      const canvas = canvasRef.current;
      canvas.width = wrapperRef.current.clientWidth;
      canvas.height = wrapperRef.current.clientHeight;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.lineWidth = 4;
        ctx.strokeStyle = '#eab308';
        ctxRef.current = ctx;
      }
    }

    const fetchOrInitSession = async () => {
      const { data } = await supabase.from('sessions').select('*').eq('id', sessionId).single();
      if (data) {
        updateSlide(data.current_slide || 1);
        const loadedDrawings = (data.canvas_data as CanvasDataMap) || {};
        allDrawingsRef.current = loadedDrawings;
        setTimeout(() => redrawCanvasForSlide(data.current_slide || 1, loadedDrawings), 500);
      } else {
        await supabase.from('sessions').insert([{ id: sessionId, file_id: fileId, current_slide: 1, canvas_data: {} }]);
      }
    };

    fetchOrInitSession();

    console.log("🖥️ [Presentation] Connecting to Supabase Channel:", `session_${sessionId}`);
    const channel = supabase.channel(`session_${sessionId}`);
    
    channel.on('broadcast', { event: 'slide_change' }, (payload) => {
      console.log("🔥 [Presentation] RECEIVED slide_change signal!", payload);
      const nextSlide = payload.payload.slide;
      updateSlide(nextSlide);
      redrawCanvasForSlide(nextSlide, allDrawingsRef.current);
    });
    
    channel.on('broadcast', { event: 'laser_move' }, (payload) => setLaser(payload.payload));
    
    channel.on('broadcast', { event: 'draw_stroke' }, (payload) => {
      const ctx = ctxRef.current;
      const canvas = canvasRef.current;
      if (!ctx || !canvas) return;
      const { x, y, type } = payload.payload;
      const pxX = x * canvas.width;
      const pxY = y * canvas.height;

      if (type === 'start') {
        ctx.beginPath();
        ctx.moveTo(pxX, pxY);
        currentLineRef.current = [{ x, y }];
      } else if (type === 'move') {
        ctx.lineTo(pxX, pxY);
        ctx.stroke();
        currentLineRef.current.push({ x, y });
      } else if (type === 'end') {
        ctx.closePath();
        if (currentLineRef.current.length > 0) {
          const slideNum = currentSlideRef.current;
          allDrawingsRef.current[slideNum] = [...(allDrawingsRef.current[slideNum] || []), currentLineRef.current];
          saveCanvasToDatabase(allDrawingsRef.current);
          currentLineRef.current = [];
        }
      }
    });
    
    channel.on('broadcast', { event: 'draw_clear' }, () => {
      if (ctxRef.current && canvasRef.current) ctxRef.current.clearRect(0, 0, canvasRef.current.width, canvasRef.current.height);
      const slideNum = currentSlideRef.current;
      allDrawingsRef.current[slideNum] = [];
      saveCanvasToDatabase(allDrawingsRef.current);
    });

    channel.subscribe((status) => {
      console.log("🖥️ [Presentation] Channel Status:", status);
    });
    
    return () => { supabase.removeChannel(channel); };
  }, [sessionId, fileId]);

  return (
    <div dir={lang === 'ku' ? 'rtl' : 'ltr'} className="flex h-screen w-full bg-gray-900 text-white overflow-hidden relative">
      <div className="w-80 bg-gray-800 border-x border-gray-700 p-6 flex flex-col items-center justify-between z-20">
        <div className="w-full flex justify-between items-center">
          <button onClick={toggleFullscreen} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1 rounded">{t.fullscreen}</button>
          <button onClick={() => setLang(lang === 'ku' ? 'en' : 'ku')} className="text-xs bg-gray-700 px-3 py-1 rounded">{t.switchLang}</button>
        </div>
        <div className="w-full flex flex-col items-center">
          <h2 className="text-3xl font-bold mb-2">NextSlide</h2>
          <p className="text-sm text-gray-400 mb-8">{t.scan}</p>
          
          <div className="bg-white p-4 rounded-xl mb-2">
            <QRCodeSVG value={`${window.location.origin}${window.location.pathname}#/remote?session=${sessionId}`} size={180} />
          </div>
          
          {/* THE NEW CLICKABLE TEST LINK */}
          <a 
            href={`${window.location.origin}${window.location.pathname}#/remote?session=${sessionId}`} 
            target="_blank" 
            rel="noreferrer"
            className="text-xs text-blue-400 underline hover:text-blue-300 mb-6"
          >
            🖥️ Click here to test Remote on PC
          </a>

        </div>
        <div className="w-full text-center bg-gray-900 rounded p-4 border border-gray-700">
          <p className="text-xs text-gray-500 mb-1">{t.session}</p>
          <p className="font-mono text-2xl text-blue-400">{sessionId}</p>
        </div>
      </div>

      <div className="flex-1 flex flex-col items-center justify-center p-8 bg-black">
        <div className="absolute top-4 right-4 bg-gray-800/80 px-4 py-2 rounded z-10">
          {t.slide} {currentSlide}
        </div>

        <div ref={wrapperRef} className="relative flex items-center justify-center bg-white rounded-xl shadow-2xl overflow-hidden aspect-[4/3] max-h-full max-w-full" style={{ height: window.innerHeight * 0.85 }}>
          {pdfError ? (
            <div className="p-12 text-red-500 font-bold text-center">
              <p className="text-2xl mb-2">{t.error}</p>
              <p className="text-sm">{pdfError}</p>
            </div>
          ) : !pdfData ? (
            <div className="p-12 text-black flex flex-col items-center justify-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-500 mb-4"></div>
              <p className="font-bold">{t.loading}</p>
            </div>
          ) : (
            <Document file={pdfData} loading={<div className="p-12 text-black">{t.loading}</div>} onLoadError={(error) => setPdfError(`PDF Error: ${error.message}`)}>
              <Page pageNumber={currentSlide} renderTextLayer={false} renderAnnotationLayer={false} height={window.innerHeight * 0.85} />
            </Document>
          )}
          <canvas ref={canvasRef} className="absolute top-0 left-0 w-full h-full pointer-events-none z-30" />
          {laser.active && (
            <div className="absolute pointer-events-none z-40 rounded-full bg-red-500 shadow-[0_0_12px_#ef4444]" style={{ width: '14px', height: '14px', left: `${laser.x * 100}%`, top: `${laser.y * 100}%`, transform: 'translate(-50%, -50%)' }} />
          )}
        </div>
      </div>
    </div>
  );
}