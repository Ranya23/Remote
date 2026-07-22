import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { supabase } from './supabaseClient';
import { extractPptxMeta } from './pptxParse';
import { recordSavedItem, formatBytes } from './Account';
import { useAuth } from './AuthContext';

// Kept in sync with Code.gs's supported types
const ACCEPTED_MIME_TYPES = [
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/webp',
  'image/svg+xml',
];

// Fallback check by extension, since some browsers/OSes report an empty
// or generic mimeType (e.g. application/octet-stream) for certain files.
const ACCEPTED_EXTENSIONS = ['.pdf', '.pptx', '.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg'];

const MAX_FILE_SIZE_MB = 25; // Apps Script requests top out around ~50MB total; stay well under that

function isAcceptedFile(file: File) {
  const nameLower = file.name.toLowerCase();
  const extOk = ACCEPTED_EXTENSIONS.some((ext) => nameLower.endsWith(ext));
  const mimeOk = ACCEPTED_MIME_TYPES.includes(file.type);
  return extOk || mimeOk;
}

// Rough client-side check just to give a fast, friendly error before
// hitting the network. Code.gs does the authoritative parsing/validation.
function looksLikeSupportedVideoLink(url: string) {
  return /youtube\.com|youtu\.be|vimeo\.com|drive\.google\.com/i.test(url);
}

function guessMimeFromName(filename: string): string {
  const lower = filename.toLowerCase();
  if (lower.endsWith('.pdf')) return 'application/pdf';
  if (lower.endsWith('.pptx')) return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  return 'application/octet-stream';
}

function typeBadge(fileType?: string) {
  switch (fileType) {
    case 'pdf':
      return { label: 'PDF', className: 'bg-red-900 text-red-300' };
    case 'image':
      return { label: 'IMAGE', className: 'bg-green-900 text-green-300' };
    case 'video-link':
      return { label: 'VIDEO', className: 'bg-purple-900 text-purple-300' };
    default:
      return { label: 'FILE', className: 'bg-gray-700 text-gray-300' };
  }
}

interface SlideItem {
  localId: string;
  name: string;
  fileType?: string; // 'pdf' | 'image' | 'video-link' | 'other', set once upload finishes
  fileId?: string;
  embedUrl?: string;
  platform?: string;
  sizeBytes?: number;
  status: 'uploading' | 'done' | 'error';
  errorMessage?: string;
}

export default function FileUpload() {
  const [slides, setSlides] = useState<SlideItem[]>([]);
  const [videoUrl, setVideoUrl] = useState('');
  const [isStarting, setIsStarting] = useState(false);
  const [message, setMessage] = useState('');

  const navigate = useNavigate();
  const { user, profile, usage, refreshUsage } = useAuth();

  // Your newest deployment URL
  const GAS_URL = 'https://script.google.com/macros/s/AKfycbx48s5aNamkERYuvJ-BE7-RBF2zt15mFZ-C-SXL_UIGZkG46RdyuPYIOlO6o0HZcr3N/exec';

  const updateSlide = (localId: string, patch: Partial<SlideItem>) => {
    setSlides((prev) => prev.map((s) => (s.localId === localId ? { ...s, ...patch } : s)));
  };

  const removeSlide = (localId: string) => {
    setSlides((prev) => prev.filter((s) => s.localId !== localId));
  };

  const moveSlide = (localId: string, direction: -1 | 1) => {
    setSlides((prev) => {
      const index = prev.findIndex((s) => s.localId === localId);
      const targetIndex = index + direction;
      if (index === -1 || targetIndex < 0 || targetIndex >= prev.length) return prev;
      const next = [...prev];
      [next[index], next[targetIndex]] = [next[targetIndex], next[index]];
      return next;
    });
  };

  const uploadOneFile = (file: File) => {
    const localId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    setSlides((prev) => [...prev, { localId, name: file.name, status: 'uploading', sizeBytes: file.size }]);

    // For .pptx specifically: start parsing the file's own XML for speaker
    // notes + transition info right away, in parallel with the upload
    // itself (they both read the same File, no conflict). This never
    // touches Code.gs / the PDF conversion Present.tsx actually renders -
    // it's a separate, best-effort pass that only feeds Supabase, so a
    // parsing hiccup here can never break the upload or the presentation.
    const isPptx = file.name.toLowerCase().endsWith('.pptx');
    const metaPromise = isPptx ? extractPptxMeta(file) : null;

    const reader = new FileReader();
    reader.onload = async () => {
      const base64Data = (reader.result as string).split(',')[1];
      try {
        const params = new URLSearchParams();
        params.append('action', 'upload');
        params.append('filename', file.name);
        params.append('mimeType', file.type || guessMimeFromName(file.name));
        params.append('data', base64Data);

        const response = await fetch(GAS_URL, { method: 'POST', body: params });
        const textResponse = await response.text();

        try {
          const result = JSON.parse(textResponse);
          if (result.status === 'success') {
            updateSlide(localId, { status: 'done', fileId: result.fileId, fileType: result.fileType });

            if (metaPromise && result.fileId) {
              metaPromise
                .then((meta) => {
                  if (
                    !Object.keys(meta.notesByPage).length &&
                    !Object.keys(meta.transitionsByPage).length &&
                    !Object.keys(meta.buildsByPage).length
                  ) return;
                  return supabase.from('pptx_meta').upsert({
                    file_id: result.fileId,
                    notes: meta.notesByPage,
                    transitions: meta.transitionsByPage,
                    builds: meta.buildsByPage,
                  });
                })
                .catch((err) => {
                  // Notes/transitions/builds are a nice-to-have layered on top
                  // of a working upload - never surface this as an upload error.
                  console.warn('⚠️ Could not save PPTX notes/transitions/builds (does the pptx_meta table have the builds column yet?):', err);
                });
            }
          } else {
            updateSlide(localId, { status: 'error', errorMessage: result.message || 'Upload failed' });
          }
        } catch (parseError) {
          updateSlide(localId, { status: 'error', errorMessage: 'Google sent an invalid response' });
        }
      } catch (error: any) {
        updateSlide(localId, { status: 'error', errorMessage: error.message || 'Network error' });
      }
    };
    reader.readAsDataURL(file);
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach((file) => {
      if (!isAcceptedFile(file)) {
        setMessage(`"${file.name}" is an unsupported type — skipped.`);
        return;
      }
      if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) {
        setMessage(`"${file.name}" is over ${MAX_FILE_SIZE_MB}MB — skipped.`);
        return;
      }
      uploadOneFile(file);
    });

    event.target.value = '';
  };

  const handleVideoLinkSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    const trimmedUrl = videoUrl.trim();
    if (!trimmedUrl) return;

    if (!looksLikeSupportedVideoLink(trimmedUrl)) {
      setMessage('Please paste a YouTube, Vimeo, or Google Drive link.');
      return;
    }

    const localId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    setSlides((prev) => [...prev, { localId, name: trimmedUrl, status: 'uploading' }]);
    setVideoUrl('');

    try {
      const params = new URLSearchParams();
      params.append('action', 'addVideoLink');
      params.append('url', trimmedUrl);

      const response = await fetch(GAS_URL, { method: 'POST', body: params });
      const textResponse = await response.text();

      try {
        const result = JSON.parse(textResponse);
        if (result.status === 'success') {
          updateSlide(localId, {
            status: 'done',
            fileId: result.fileId,
            fileType: result.fileType,
            embedUrl: result.embedUrl,
            platform: result.platform,
            name: result.platform ? `${result.platform} video` : trimmedUrl,
          });
        } else {
          updateSlide(localId, { status: 'error', errorMessage: result.message || 'Could not add link' });
        }
      } catch (parseError) {
        updateSlide(localId, { status: 'error', errorMessage: 'Google sent an invalid response' });
      }
    } catch (error: any) {
      updateSlide(localId, { status: 'error', errorMessage: error.message || 'Network error' });
    }
  };

  const readySlides = slides.filter((s) => s.status === 'done');
  const isAnyUploading = slides.some((s) => s.status === 'uploading');

  const handleStartLesson = async () => {
    if (readySlides.length === 0) return;

    if (!user) {
      setMessage('Log in or create a free account to save and open your presentation later.');
      navigate('/account');
      return;
    }

    const totalBytes = readySlides.reduce((sum, s) => sum + (s.sizeBytes || 0), 0);
    if (usage) {
      if (usage.presentationCount >= usage.presentationLimit) {
        setMessage(`You've reached your limit of ${usage.presentationLimit} presentations. Delete one from My Account to add a new one.`);
        return;
      }
      if (usage.storageBytes + totalBytes > usage.storageLimitBytes) {
        setMessage(
          `That would put you over your ${formatBytes(usage.storageLimitBytes)} storage limit ` +
          `(currently using ${formatBytes(usage.storageBytes)}). Delete something from My Account first.`
        );
        return;
      }
    }

    // Single slide: skip the lesson wrapper entirely and behave exactly like
    // the original single-file flow. This keeps plain PDF/PPTX/image uploads
    // working today, independent of whether the viewer understands lessons yet.
    if (readySlides.length === 1) {
      const only = readySlides[0];
      const result = await recordSavedItem({
        kind: 'lesson',
        title: only.name,
        file_id: only.fileId,
        file_type: only.fileType,
        size_bytes: only.sizeBytes || 0,
      });
      if (!result.ok && (result.reason === 'presentation_limit' || result.reason === 'storage_limit')) {
        setMessage(
          result.reason === 'presentation_limit'
            ? `You've reached your limit of ${usage?.presentationLimit ?? 5} presentations.`
            : `That would put you over your storage limit.`
        );
        return;
      }
      refreshUsage();
      navigate(`/present/${only.fileId}`, {
        state: { fileType: only.fileType, embedUrl: only.embedUrl, platform: only.platform },
      });
      return;
    }

    setIsStarting(true);
    setMessage('Saving lesson...');

    try {
      const params = new URLSearchParams();
      params.append('action', 'saveLesson');
      params.append(
        'slides',
        JSON.stringify(readySlides.map((s) => ({ fileId: s.fileId, fileType: s.fileType, name: s.name })))
      );

      const response = await fetch(GAS_URL, { method: 'POST', body: params });
      const textResponse = await response.text();

      try {
        const result = JSON.parse(textResponse);
        if (result.status === 'success') {
          setMessage('Success! Starting lesson...');
          const title = readySlides.length <= 2
            ? readySlides.map((s) => s.name).join(' + ')
            : `${readySlides[0].name} + ${readySlides.length - 1} more`;
          const saveResult = await recordSavedItem({
            kind: 'lesson',
            title,
            file_id: result.fileId,
            file_type: 'lesson',
            size_bytes: totalBytes,
          });
          if (!saveResult.ok && (saveResult.reason === 'presentation_limit' || saveResult.reason === 'storage_limit')) {
            setMessage(
              saveResult.reason === 'presentation_limit'
                ? `You've reached your limit of ${usage?.presentationLimit ?? 5} presentations - this lesson wasn't saved to My Account.`
                : `That would put you over your storage limit - this lesson wasn't saved to My Account.`
            );
          }
          refreshUsage();
          navigate(`/present/${result.fileId}`, { state: { fileType: 'lesson' } });
        } else {
          setMessage('Server returned: ' + JSON.stringify(result));
        }
      } catch (parseError) {
        setMessage('Google sent invalid format: ' + textResponse.substring(0, 100));
      }
    } catch (error: any) {
      setMessage('Network Error: ' + error.message);
    } finally {
      setIsStarting(false);
    }
  };

  return (
    <div className="flex flex-col items-center justify-center w-full max-w-2xl mx-auto p-6 mt-20">
      <div className="w-full flex justify-end mb-2">
        <button
          type="button"
          onClick={() => navigate('/account')}
          className="text-xs text-gray-400 hover:text-white bg-gray-900 border border-gray-700 rounded-full px-3 py-1.5"
        >
          {user ? `👤 ${profile?.display_name || user.email} · My account` : '👤 Log in / Create account'}
        </button>
      </div>
      <h1 className="text-3xl font-bold text-white mb-2 text-center">NextSlide Uploader</h1>
      <p className="text-sm text-gray-400 mb-8 text-center">
        Add as many files and video links as you need, then start the lesson.
      </p>

      <div className="w-full flex justify-center items-center h-56 border-2 border-dashed rounded-lg bg-gray-900 border-gray-600 hover:bg-gray-800 transition-colors">
        <label className="flex flex-col items-center justify-center w-full h-full cursor-pointer">
          <div className="flex flex-col items-center justify-center pt-5 pb-6">
            <svg className="w-10 h-10 mb-3 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
            </svg>
            <p className="mb-2 text-sm text-gray-400">
              <span className="font-semibold text-white">Click to upload</span> or drag and drop — pick multiple at once
            </p>
            <p className="text-xs text-gray-500">PDF, PPTX, or images (PNG, JPG, GIF, WEBP, SVG)</p>
          </div>
          <input
            type="file"
            className="hidden"
            multiple
            onChange={handleFileSelect}
            accept=".pdf,.pptx,.png,.jpg,.jpeg,.gif,.webp,.svg,application/pdf,application/vnd.openxmlformats-officedocument.presentationml.presentation,image/*"
          />
        </label>
      </div>

      <div className="w-full flex items-center gap-3 my-6">
        <div className="flex-1 h-px bg-gray-700" />
        <span className="text-xs text-gray-500 uppercase tracking-wide">or add a video link</span>
        <div className="flex-1 h-px bg-gray-700" />
      </div>

      <form onSubmit={handleVideoLinkSubmit} className="w-full flex flex-col sm:flex-row gap-3">
        <input
          type="url"
          value={videoUrl}
          onChange={(e) => setVideoUrl(e.target.value)}
          placeholder="Paste a YouTube, Vimeo, or Google Drive link"
          className="flex-1 rounded-lg bg-gray-900 border border-gray-600 text-white text-sm px-4 py-3 placeholder-gray-500 focus:outline-none focus:border-blue-500"
        />
        <button
          type="submit"
          disabled={!videoUrl.trim()}
          className="rounded-lg bg-blue-600 hover:bg-blue-500 disabled:bg-gray-700 disabled:text-gray-500 text-white text-sm font-semibold px-6 py-3 transition-colors"
        >
          Add video
        </button>
      </form>

      {slides.length > 0 && (
        <div className="w-full mt-8">
          <h2 className="text-sm font-semibold text-gray-300 mb-3">
            Lesson slides ({readySlides.length}/{slides.length} ready)
          </h2>
          <ul className="w-full flex flex-col gap-2">
            {slides.map((slide, index) => {
              const badge = typeBadge(slide.fileType);
              return (
                <li
                  key={slide.localId}
                  className="flex items-center gap-3 bg-gray-900 border border-gray-700 rounded-lg px-4 py-3"
                >
                  <span className="text-xs text-gray-500 w-5 text-center">{index + 1}</span>

                  {slide.status === 'uploading' && (
                    <span className="text-xs text-yellow-400 whitespace-nowrap">Uploading...</span>
                  )}
                  {slide.status === 'done' && (
                    <span className={`text-xs font-bold px-2 py-1 rounded ${badge.className}`}>{badge.label}</span>
                  )}
                  {slide.status === 'error' && (
                    <span className="text-xs font-bold px-2 py-1 rounded bg-red-900 text-red-300">ERROR</span>
                  )}

                  <span className="flex-1 text-sm text-white truncate">{slide.name}</span>
                  {slide.status === 'error' && slide.errorMessage && (
                    <span className="text-xs text-red-400 truncate max-w-[10rem]">{slide.errorMessage}</span>
                  )}

                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      type="button"
                      onClick={() => moveSlide(slide.localId, -1)}
                      disabled={index === 0}
                      className="text-gray-400 hover:text-white disabled:opacity-30 disabled:hover:text-gray-400 px-1"
                      aria-label="Move up"
                    >
                      ↑
                    </button>
                    <button
                      type="button"
                      onClick={() => moveSlide(slide.localId, 1)}
                      disabled={index === slides.length - 1}
                      className="text-gray-400 hover:text-white disabled:opacity-30 disabled:hover:text-gray-400 px-1"
                      aria-label="Move down"
                    >
                      ↓
                    </button>
                    <button
                      type="button"
                      onClick={() => removeSlide(slide.localId)}
                      className="text-gray-400 hover:text-red-400 px-1"
                      aria-label="Remove"
                    >
                      ✕
                    </button>
                  </div>
                </li>
              );
            })}
          </ul>
        </div>
      )}

      <button
        type="button"
        onClick={handleStartLesson}
        disabled={readySlides.length === 0 || isAnyUploading || isStarting}
        className="w-full mt-6 rounded-lg bg-blue-600 hover:bg-blue-500 disabled:bg-gray-700 disabled:text-gray-500 text-white font-bold py-4 transition-colors"
      >
        {isStarting
          ? 'Starting...'
          : isAnyUploading
          ? 'Waiting for uploads to finish...'
          : `Start Lesson (${readySlides.length} slide${readySlides.length === 1 ? '' : 's'})`}
      </button>

      {message && (
        <div className="mt-6 p-4 w-full text-center rounded-xl font-bold text-sm bg-gray-800 text-yellow-400">
          {message}
        </div>
      )}
    </div>
  );
}
