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

// --- Build (bullet/object entrance animation) steps ------------------------
//
// One "step" is everything that appears on a single click - usually one
// bullet, occasionally a small group of shapes set to animate together.
// Each step is a list of regions (percent-of-slide rectangles) that should
// be masked until that step fires. Percent coordinates, not pixels, because
// they're computed straight from each shape's real position in EMUs
// (English Metric Units - PowerPoint's native unit) against the slide's
// fixed canvas size, so they line up with the rendered PDF page regardless
// of how large it's displayed.
//
// This deliberately only covers shape-level and paragraph-level ENTRANCE
// builds (Appear/Fade/Fly/Wipe applied to a whole placeholder or one
// bullet at a time) - by far the most common real-world use of PPTX
// animation. Anything this can't confidently place - motion paths, exit/
// emphasis effects, SmartArt/chart-internal animation - causes the whole
// slide to be skipped (no `builds` entry at all), same graceful-degradation
// philosophy as the transition parsing above: a slide we can't model
// correctly just renders fully built, like it does today, rather than
// revealing something in the wrong place.
export type BuildEffect = 'fade' | 'fly' | 'wipe' | 'appear';
export interface BuildRegion {
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  effect: BuildEffect;
}
export interface SlideBuildStep {
  regions: BuildRegion[];
}
export interface SlideBuilds {
  steps: SlideBuildStep[];
}

export interface PptxMeta {
  notesByPage: Record<number, string>;
  transitionsByPage: Record<number, SlideTransition>;
  buildsByPage: Record<number, SlideBuilds>;
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

// --- Build steps -----------------------------------------------------------

// The slide canvas size in EMUs (<p:sldSz cx cy> on ppt/presentation.xml) -
// every shape's percent position is computed against this. Read once per
// upload and reused for every slide, not per-slide (it's a document-wide
// setting).
async function getSlideSize(zip: JSZip): Promise<{ cx: number; cy: number } | null> {
  try {
    const presDoc = await readXml(zip, 'ppt/presentation.xml');
    if (!presDoc) return null;
    const sldSz = firstEl(presDoc, 'p:sldSz');
    if (!sldSz) return null;
    const cx = Number(sldSz.getAttribute('cx'));
    const cy = Number(sldSz.getAttribute('cy'));
    if (!cx || !cy || isNaN(cx) || isNaN(cy)) return null;
    return { cx, cy };
  } catch {
    return null;
  }
}

interface ShapeGeom {
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  // Character count of each paragraph in this shape's text body, in order -
  // used to proportionally split the shape's box when separate build steps
  // target individual paragraphs (the classic "reveal one bullet at a
  // time" case). Empty for non-text shapes (pictures, etc).
  paragraphChars: number[];
}

// Top-level shapes only (direct children of the slide's <p:spTree>) - deep
// enough to cover placeholders, text boxes, pictures and the occasional
// grouped shape as one unit, which is what the vast majority of real decks
// animate. A shape with no explicit <a:xfrm> (position inherited from the
// layout/master) is left out entirely - we can't place it reliably, so it's
// simply never masked, rather than guessed at.
function collectShapeGeometry(slideDoc: Document, slideSize: { cx: number; cy: number }): Map<string, ShapeGeom> {
  const map = new Map<string, ShapeGeom>();
  const spTree = firstEl(slideDoc, 'p:spTree');
  if (!spTree) return map;

  for (const child of Array.from(spTree.children)) {
    const tag = child.tagName;
    if (!['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp'].includes(tag)) continue;

    const cNvPr = firstEl(child, 'p:cNvPr');
    const id = cNvPr?.getAttribute('id');
    if (!id) continue;

    const xfrm = firstEl(child, 'a:xfrm') || firstEl(child, 'p:xfrm');
    if (!xfrm) continue;
    const off = firstEl(xfrm, 'a:off');
    const ext = firstEl(xfrm, 'a:ext');
    if (!off || !ext) continue;

    const xEmu = Number(off.getAttribute('x'));
    const yEmu = Number(off.getAttribute('y'));
    const cxEmu = Number(ext.getAttribute('cx'));
    const cyEmu = Number(ext.getAttribute('cy'));
    if ([xEmu, yEmu, cxEmu, cyEmu].some((n) => isNaN(n))) continue;

    const paragraphChars = tag === 'p:sp'
      ? allEls(child, 'a:p').map((p) => allEls(p, 'a:t').reduce((sum, t) => sum + (t.textContent || '').length, 0))
      : [];

    map.set(id, {
      xPct: (xEmu / slideSize.cx) * 100,
      yPct: (yEmu / slideSize.cy) * 100,
      wPct: (cxEmu / slideSize.cx) * 100,
      hPct: (cyEmu / slideSize.cy) * 100,
      paragraphChars,
    });
  }
  return map;
}

// Reads the actual animation primitive used, rather than PowerPoint's
// numeric presetID gallery (~200 entries, not something worth hardcoding
// from memory when it's easy to get subtly wrong) - <p:set> alone means an
// instant "Appear", an <p:animEffect filter="fade.../wipe..."> names itself
// directly, and <p:anim>/<p:animScale> driving position/size implies some
// kind of "Fly"-like motion. Falls back to 'fade' - a safe, common default -
// when nothing more specific is recognizable.
function detectEffect(par: Element): BuildEffect {
  const hasAnimEffect = allEls(par, 'p:animEffect').length > 0;
  const hasAnim = allEls(par, 'p:anim').length > 0 || allEls(par, 'p:animScale').length > 0;
  if (allEls(par, 'p:set').length && !hasAnimEffect && !hasAnim) return 'appear';
  const animEffect = firstEl(par, 'p:animEffect');
  const filter = animEffect?.getAttribute('filter') || '';
  if (filter.startsWith('fade')) return 'fade';
  if (filter.startsWith('wipe')) return 'wipe';
  if (hasAnim) return 'fly';
  return 'fade';
}

// Splits a shape's box proportionally by paragraph character count (with a
// floor so an empty bullet still gets a sliver of height) when a build step
// only targets some of its paragraphs - e.g. a single placeholder set to
// reveal one bullet per click. This is an approximation (real text layout
// depends on font/wrapping we don't have access to), but it tracks uneven
// bullet lengths far better than an equal-height split, and it's the
// documented trade-off, not a silent guess.
function paragraphRangeToRegion(shape: ShapeGeom, st: number, end: number, effect: BuildEffect): BuildRegion {
  const lengths = shape.paragraphChars.map((n) => Math.max(n, 8));
  const total = lengths.reduce((a, b) => a + b, 0) || 1;
  const clampedSt = Math.max(0, Math.min(st, lengths.length - 1));
  const clampedEnd = Math.max(clampedSt, Math.min(end, lengths.length - 1));
  let before = 0;
  for (let i = 0; i < clampedSt; i++) before += lengths[i];
  let within = 0;
  for (let i = clampedSt; i <= clampedEnd; i++) within += lengths[i];
  return {
    xPct: shape.xPct,
    yPct: shape.yPct + shape.hPct * (before / total),
    wPct: shape.wPct,
    hPct: Math.max(shape.hPct * (within / total), 1.5),
    effect,
  };
}

// One slide's worth of click-triggered build steps, read out of its own
// <p:timing> tree. Only looks at the main click sequence
// (p:tnLst > p:par > p:cTn[nodeType=mainSeq] > p:childTnLst) - each direct
// child <p:par> there is one click step in the common "on click, one build
// at a time" pattern PowerPoint generates for ordinary bullet/object
// animation. If anything in the tree looks like something this can't
// safely approximate (a motion path, or an exit/emphasis effect where an
// entrance was expected), the WHOLE slide is skipped - better to show it
// fully built, like today, than to reveal it incorrectly.
function extractBuildsForSlide(slideDoc: Document, slideSize: { cx: number; cy: number }): SlideBuilds | undefined {
  try {
    const timing = firstEl(slideDoc, 'p:timing');
    if (!timing) return undefined;

    const shapes = collectShapeGeometry(slideDoc, slideSize);
    if (!shapes.size) return undefined;

    const mainSeqCTn = allEls(timing, 'p:cTn').find((c) => c.getAttribute('nodeType') === 'mainSeq');
    if (!mainSeqCTn) return undefined;
    const mainChildLst = firstEl(mainSeqCTn, 'p:childTnLst');
    if (!mainChildLst) return undefined;

    const clickPars = Array.from(mainChildLst.children).filter((c) => c.tagName === 'p:par');
    const steps: SlideBuildStep[] = [];

    for (const par of clickPars) {
      if (allEls(par, 'p:animMotion').length) return undefined; // motion path - bail on the whole slide

      const presetClasses = allEls(par, 'p:cTn')
        .map((c) => c.getAttribute('presetClass'))
        .filter((v): v is string => !!v);
      if (presetClasses.some((c) => c !== 'entr')) return undefined; // exit/emphasis mixed in - bail
      if (!presetClasses.length) continue; // nothing recognizable here - skip this step, don't bail

      const effect = detectEffect(par);
      const regions: BuildRegion[] = [];
      for (const tgt of allEls(par, 'p:spTgt')) {
        const spid = tgt.getAttribute('spid');
        if (!spid) continue;
        const shape = shapes.get(spid);
        if (!shape) continue; // couldn't place this shape - leave it unmasked, never bail over one shape

        const txEl = firstEl(tgt, 'p:txEl');
        const pRg = txEl ? firstEl(txEl, 'p:pRg') : null;
        if (pRg && shape.paragraphChars.length > 0) {
          const st = Number(pRg.getAttribute('st'));
          const end = Number(pRg.getAttribute('end'));
          if (!isNaN(st) && !isNaN(end)) {
            regions.push(paragraphRangeToRegion(shape, st, end, effect));
            continue;
          }
        }
        regions.push({ xPct: shape.xPct, yPct: shape.yPct, wPct: shape.wPct, hPct: shape.hPct, effect });
      }
      if (regions.length) steps.push({ regions });
    }

    return steps.length ? { steps } : undefined;
  } catch {
    return undefined;
  }
}

// --- Entry point -----------------------------------------------------------
// Runs the whole extraction pass once per upload. Never throws - always
// resolves to *something* usable (possibly empty maps), so a parsing
// hiccup never blocks the actual upload/conversion flow in FileUpload.tsx.
export async function extractPptxMeta(file: File): Promise<PptxMeta> {
  const empty: PptxMeta = { notesByPage: {}, transitionsByPage: {}, buildsByPage: {}, slideCount: 0 };
  try {
    const zip = await JSZip.loadAsync(file);
    const slidePaths = await getOrderedSlidePaths(zip);
    if (!slidePaths.length) return empty;

    const slideSize = await getSlideSize(zip);

    const notesByPage: Record<number, string> = {};
    const transitionsByPage: Record<number, SlideTransition> = {};
    const buildsByPage: Record<number, SlideBuilds> = {};

    for (let i = 0; i < slidePaths.length; i++) {
      const page = i + 1;
      const [notes, transition] = await Promise.all([
        extractNotesForSlide(zip, slidePaths[i]),
        extractTransitionForSlide(zip, slidePaths[i]),
      ]);
      if (notes) notesByPage[page] = notes;
      if (transition) transitionsByPage[page] = transition;

      // Build steps need the slide's own XML directly (shape geometry +
      // <p:timing>), same best-effort/never-throws contract as the notes
      // and transition passes above - a slide this can't confidently parse
      // just gets no `builds` entry and renders fully built, as it does today.
      if (slideSize) {
        try {
          const slideDoc = await readXml(zip, slidePaths[i]);
          const builds = slideDoc ? extractBuildsForSlide(slideDoc, slideSize) : undefined;
          if (builds) buildsByPage[page] = builds;
        } catch {
          // skip - never blocks notes/transitions/the rest of the upload
        }
      }
    }

    return { notesByPage, transitionsByPage, buildsByPage, slideCount: slidePaths.length };
  } catch (err) {
    console.warn('⚠️ PPTX metadata extraction skipped:', err);
    return empty;
  }
}
