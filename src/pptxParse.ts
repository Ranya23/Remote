// --- Client-side PPTX metadata extraction ----------------------------------
//
// PPTX is just a ZIP of XML parts. This runs once, in the browser, right at
// upload time (see FileUpload.tsx), completely independent of the existing
// Code.gs pipeline that converts the file to a PDF for on-screen rendering.
// It never touches how slides are displayed - it only pulls out two things
// Code.gs doesn't give us today:
//
//   1. Speaker notes per slide (ppt/notesSlides/notesSlideN.xml)
//   2. The slide's <p:transition> effect + duration, for a CSS-based replay
//      on "next slide" that approximates what PowerPoint itself would show
//
// Both are keyed by slide position (1-based, matching the PDF page numbers
// Code.gs produces), and both are best-effort: if anything about a given
// slide can't be parsed, that slide just ends up with no notes / no
// transition (falls back to a plain cut) instead of blowing up the upload.
//
// NOTE ON SCOPE: this deliberately does NOT attempt full per-object,
// per-build animation playback (individual bullet reveals, motion paths,
// Morph, SmartArt) - that would mean replacing the PDF-page rendering
// pipeline with a from-scratch DrawingML-to-HTML/SVG renderer, which is a
// much larger project on its own and out of scope for this pass. What's
// here is real: actual transition type/duration read straight out of the
// file, applied as a real CSS transition when the presenter advances.

import JSZip from 'jszip';

export type TransitionKind = 'fade' | 'slide' | 'cut';
export interface SlideTransition {
  kind: TransitionKind;
  durationMs: number;
  direction?: 'l' | 'r' | 'u' | 'd'; // only meaningful for 'slide'
}
export interface PptxMeta {
  notesByPage: Record<number, string>;
  transitionsByPage: Record<number, SlideTransition>;
  slideCount: number;
}

const parser = new DOMParser();
function parseXml(text: string): Document {
  return parser.parseFromString(text, 'application/xml');
}

// Every real OOXML relationship/content-type lookup we need, done with
// plain tag-name matching (not namespace-aware) - the prefixes PowerPoint
// itself writes (p:, a:, r:, mc:...) are consistent enough in practice that
// this is the same pragmatic approach most lightweight browser-side pptx
// readers use, and it keeps this file dependency-free besides JSZip.
function firstEl(doc: Document | Element, tag: string): Element | null {
  return doc.getElementsByTagName(tag)[0] || null;
}
function allEls(doc: Document | Element, tag: string): Element[] {
  return Array.from(doc.getElementsByTagName(tag));
}

// Resolves "../slides/slide3.xml" (relative to ppt/_rels/) etc. into a
// normalized zip-entry path.
function resolveRelPath(basePath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseDir = basePath.split('/').slice(0, -1); // drop the file itself
  const parts = target.split('/');
  const stack = [...baseDir];
  for (const part of parts) {
    if (part === '..') stack.pop();
    else if (part === '.') continue;
    else stack.push(part);
  }
  return stack.join('/');
}

async function readXml(zip: JSZip, path: string): Promise<Document | null> {
  const entry = zip.file(path);
  if (!entry) return null;
  const text = await entry.async('text');
  return parseXml(text);
}

// The ordered list of slideN.xml paths, in actual presentation order (NOT
// necessarily numeric filename order - PowerPoint doesn't guarantee those
// match, especially after slides have been reordered/duplicated).
async function getOrderedSlidePaths(zip: JSZip): Promise<string[]> {
  const presDoc = await readXml(zip, 'ppt/presentation.xml');
  const relsDoc = await readXml(zip, 'ppt/_rels/presentation.xml.rels');
  if (!presDoc || !relsDoc) return [];

  const relIdToTarget = new Map<string, string>();
  for (const rel of allEls(relsDoc, 'Relationship')) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) relIdToTarget.set(id, resolveRelPath('ppt/_rels/presentation.xml.rels', target));
  }

  const sldIdLst = firstEl(presDoc, 'p:sldIdLst');
  if (!sldIdLst) return [];
  const paths: string[] = [];
  for (const sldId of allEls(sldIdLst, 'p:sldId')) {
    // r:id - attribute lookup by local name since the namespace prefix on
    // the *attribute itself* is always 'r' in practice for this element.
    const rid = sldId.getAttribute('r:id');
    if (!rid) continue;
    const target = relIdToTarget.get(rid);
    if (target) paths.push(target);
  }
  return paths;
}

// --- Speaker notes -----------------------------------------------------
async function extractNotesForSlide(zip: JSZip, slidePath: string): Promise<string | undefined> {
  try {
    const slideDir = slidePath.split('/').slice(0, -1).join('/');
    const slideFile = slidePath.split('/').pop()!;
    const relsPath = `${slideDir}/_rels/${slideFile}.rels`;
    const relsDoc = await readXml(zip, relsPath);
    if (!relsDoc) return undefined;

    let notesPath: string | null = null;
    for (const rel of allEls(relsDoc, 'Relationship')) {
      const type = rel.getAttribute('Type') || '';
      if (type.endsWith('/notesSlide')) {
        const target = rel.getAttribute('Target');
        if (target) notesPath = resolveRelPath(relsPath, target);
        break;
      }
    }
    if (!notesPath) return undefined;

    const notesDoc = await readXml(zip, notesPath);
    if (!notesDoc) return undefined;

    // Prefer the "body" placeholder (the actual notes text box, as opposed
    // to the slide-thumbnail placeholder or slide-number/date/footer
    // placeholders that also live on a notes page).
    const shapes = allEls(notesDoc, 'p:sp');
    let bodyShape: Element | null = null;
    for (const sp of shapes) {
      const ph = firstEl(sp, 'p:ph');
      const phType = ph?.getAttribute('type');
      if (phType === 'body') { bodyShape = sp; break; }
    }
    // Fallback: whichever shape has the most actual text - in practice
    // that's always the notes body, even when its placeholder type isn't
    // explicitly "body" (some templates omit it).
    if (!bodyShape) {
      let best = { shape: null as Element | null, len: 0 };
      for (const sp of shapes) {
        const ph = firstEl(sp, 'p:ph');
        const phType = ph?.getAttribute('type');
        if (phType === 'sldImg') continue; // the embedded slide thumbnail, never text
        const text = allEls(sp, 'a:t').map((t) => t.textContent || '').join('');
        if (text.length > best.len) best = { shape: sp, len: text.length };
      }
      bodyShape = best.shape;
    }
    if (!bodyShape) return undefined;

    const paragraphs = allEls(bodyShape, 'a:p').map((p) =>
      allEls(p, 'a:t').map((t) => t.textContent || '').join('')
    );
    const text = paragraphs.join('\n').trim();
    return text || undefined;
  } catch {
    return undefined; // best-effort - one bad slide shouldn't break the rest
  }
}

// --- Transitions ---------------------------------------------------------
// Effect tag -> our simplified kind. Anything not listed here (Morph,
// SmartArt-driven effects, the newer p159:* "gallery" transitions, honeycomb,
// ripple, etc.) falls through to the 'fade' fallback below, per the graceful-
// degradation requirement - never 'skip or break the presentation'.
const SLIDE_LIKE = new Set(['push', 'cover', 'pull', 'comb']);
const FADE_LIKE = new Set(['fade']);
const CUT_LIKE = new Set(['cut']);

function parseTransitionEl(transEl: Element): SlideTransition {
  // Duration: modern files use dur="500" (milliseconds); older ones use a
  // spd="slow|med|fast" enum instead.
  let durationMs = 500;
  const dur = transEl.getAttribute('dur');
  const spd = transEl.getAttribute('spd');
  if (dur && !isNaN(Number(dur))) durationMs = Number(dur);
  else if (spd === 'slow') durationMs = 1000;
  else if (spd === 'fast') durationMs = 250;

  // The effect is whichever child element is present - <p:fade/>, <p:push
  // dir="l"/>, <p:cut/>, <p:wipe .../>, etc. Just look at the first
  // element child's local tag name.
  const child = Array.from(transEl.children).find((c) => !c.tagName.includes(':timing'));
  const rawTag = child?.tagName || '';
  const local = rawTag.includes(':') ? rawTag.split(':')[1] : rawTag;

  if (FADE_LIKE.has(local)) return { kind: 'fade', durationMs };
  if (CUT_LIKE.has(local)) return { kind: 'cut', durationMs: 0 };
  if (SLIDE_LIKE.has(local)) {
    const dir = (child?.getAttribute('dir') || 'l').slice(0, 1) as 'l' | 'r' | 'u' | 'd';
    const validDir = (['l', 'r', 'u', 'd'] as const).includes(dir) ? dir : 'l';
    return { kind: 'slide', durationMs, direction: validDir };
  }
  // Graceful fallback for wipe/wheel/blinds/checker/circle/diamond/random/
  // morph/gallery/anything-we-don't-specifically-model.
  return { kind: 'fade', durationMs };
}

async function extractTransitionForSlide(zip: JSZip, slidePath: string): Promise<SlideTransition | undefined> {
  try {
    const slideDoc = await readXml(zip, slidePath);
    if (!slideDoc) return undefined;
    const transEl = firstEl(slideDoc, 'p:transition');
    if (!transEl) return undefined; // no explicit transition set -> caller defaults to 'cut'
    return parseTransitionEl(transEl);
  } catch {
    return undefined;
  }
}

// --- Entry point -----------------------------------------------------------
// Runs the whole extraction pass once per upload. Never throws - always
// resolves to *something* usable (possibly empty maps), so a parsing
// hiccup never blocks the actual upload/conversion flow in FileUpload.tsx.
export async function extractPptxMeta(file: File): Promise<PptxMeta> {
  const empty: PptxMeta = { notesByPage: {}, transitionsByPage: {}, slideCount: 0 };
  try {
    const zip = await JSZip.loadAsync(file);
    const slidePaths = await getOrderedSlidePaths(zip);
    if (!slidePaths.length) return empty;

    const notesByPage: Record<number, string> = {};
    const transitionsByPage: Record<number, SlideTransition> = {};

    for (let i = 0; i < slidePaths.length; i++) {
      const page = i + 1;
      const [notes, transition] = await Promise.all([
        extractNotesForSlide(zip, slidePaths[i]),
        extractTransitionForSlide(zip, slidePaths[i]),
      ]);
      if (notes) notesByPage[page] = notes;
      if (transition) transitionsByPage[page] = transition;
    }

    return { notesByPage, transitionsByPage, slideCount: slidePaths.length };
  } catch (err) {
    console.warn('⚠️ PPTX metadata extraction skipped:', err);
    return empty;
  }
}
