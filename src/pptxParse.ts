// --- Client-side PPTX metadata extraction ----------------------------------
//
// PPTX is just a ZIP of XML parts. This runs once, in the browser, right at
// upload time (see FileUpload.tsx), completely independent of the existing
// Code.gs pipeline that converts the file to a PDF for on-screen rendering.
// It never touches how slides are displayed by default - it only pulls out
// things Code.gs doesn't give us today:
//
//   1. Speaker notes per slide (ppt/notesSlides/notesSlideN.xml)
//   2. The slide's <p:transition> effect + duration, for a CSS-based replay
//      on "next slide" that approximates what PowerPoint itself would show
//   3. Click-triggered BUILD animations (bullet-by-bullet / object-by-object
//      reveals) - see the big comment above extractBuildsForSlide below for
//      exactly what's supported and why everything else safely falls back
//      to showing the slide fully built, instead of guessing wrong.
//
// All three are keyed by slide position (1-based, matching the PDF page
// numbers Code.gs produces), and all are best-effort: if anything about a
// given slide can't be parsed, that slide just ends up with no notes / no
// transition / no build info (falls back to a plain cut / fully-built slide)
// instead of blowing up the upload.

import JSZip from 'jszip';

export type TransitionKind = 'fade' | 'slide' | 'cut';
export interface SlideTransition {
  kind: TransitionKind;
  durationMs: number;
  direction?: 'l' | 'r' | 'u' | 'd'; // only meaningful for 'slide'
}

// One rectangle to reveal, in % of slide width/height (0-100) so it's
// resolution-independent - the viewer just multiplies by whatever pixel
// size the slide is actually being displayed at.
export interface BuildBox { xPct: number; yPct: number; wPct: number; hPct: number; }
// One "Next" press worth of reveal - almost always a single box (one bullet
// or one object), but PowerPoint lets several things enter on the same
// click, hence an array.
export interface BuildStep { boxes: BuildBox[]; }
// Presence of an entry for a given page means "this slide has a build we
// understood well enough to trust" - absence means "render it fully built,
// like every slide today". There's no separate on/off flag; not being here
// *is* the fallback.
export interface SlideBuildInfo { steps: BuildStep[]; }

export interface PptxMeta {
  notesByPage: Record<number, string>;
  transitionsByPage: Record<number, SlideTransition>;
  buildsByPage: Record<number, SlideBuildInfo>;
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

// The timing/build tree is deeply nested and re-uses the same tag names at
// different depths (a <p:cTn> can contain other <p:cTn>s many levels down),
// so the flat/deep lookups above are actively wrong there - we need the
// *immediate* child with a given local name, ignoring namespace prefixes.
function localName(el: Element): string {
  const t = el.tagName;
  const i = t.indexOf(':');
  return i === -1 ? t : t.slice(i + 1);
}
function directChildren(el: Element, tag: string): Element[] {
  return Array.from(el.children).filter((c) => localName(c) === tag);
}
function directChild(el: Element, tag: string): Element | null {
  return directChildren(el, tag)[0] || null;
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

// --- Build (bullet/object reveal) animations --------------------------------
//
// This does NOT attempt to reproduce every PowerPoint animation - that's an
// enormous, genuinely open-ended surface (motion paths, Morph, SmartArt,
// chart-by-series, emphasis effects on already-visible content...). Instead
// it recognizes exactly one well-defined, extremely common pattern - "each
// click reveals one more bullet/object, nothing else happens" - and bails
// out (returns undefined -> slide renders fully built, exactly like today)
// the moment anything doesn't match that pattern with confidence.
//
// What it needs, concretely, per click-triggered step in <p:timing>:
//   - a real shape (<p:sp> or <p:pic>, not grouped, not a chart/table/SmartArt
//     <p:graphicFrame>) that we can find a bounding box for - either the
//     shape's own <a:xfrm>, or (one level of fallback) the matching
//     placeholder's <a:xfrm> on the slide's layout
//   - either the whole shape appearing, or a paragraph-range (<p:pRg>) within
//     it - never finer than that (<p:charRg> = per-character = bail)
//   - an entrance effect (presetClass="entr") - anything else (exit/emphasis
//     mid-sequence, since our model assumes everything starts hidden and
//     only ever gets revealed) also bails
//
// We deliberately don't care *how* PowerPoint enters it (fade/wipe/fly-in/
// whatever) - the viewer always reveals it the same simple way, by removing
// a mask (see BuildRevealOverlay.tsx). So the animation *style* PowerPoint
// authored is not reproduced, only the *reveal order/grouping* is - which is
// exactly the behavior the feature request asked for.
//
// The geometry this produces is only ever a bounding box - never the shape's
// actual text/image content - because the viewer never re-draws the shape at
// all. It just samples the real, already-rendered PDF page's own pixels at
// that box and covers/uncovers them. That sidesteps needing to reproduce
// fonts, theme colors, bullet glyphs, or backgrounds here entirely, and is
// why this stays this short.

function slideHasPictureOrPatternBackground(slideDoc: Document): boolean {
  const cSld = firstEl(slideDoc, 'p:cSld');
  const bg = cSld ? directChild(cSld, 'bg') : null;
  const bgPr = bg ? directChild(bg, 'bgPr') : null;
  if (!bgPr) return false;
  return !!directChild(bgPr, 'blipFill') || !!directChild(bgPr, 'pattFill');
}

async function getSlideSize(zip: JSZip): Promise<{ cx: number; cy: number } | null> {
  const presDoc = await readXml(zip, 'ppt/presentation.xml');
  const sldSz = presDoc ? firstEl(presDoc, 'p:sldSz') : null;
  if (!sldSz) return null;
  const cx = Number(sldSz.getAttribute('cx'));
  const cy = Number(sldSz.getAttribute('cy'));
  if (!isFinite(cx) || !isFinite(cy) || cx <= 0 || cy <= 0) return null;
  return { cx, cy };
}

// Every <p:sp>/<p:pic>/<p:graphicFrame>/<p:grpSp> in the slide, indexed by
// its shape id (<p:cNvPr id="...">) so animation targets (<p:spTgt spid="...">)
// can be resolved back to an actual element.
function buildShapeIndex(slideDoc: Document): Map<string, Element> {
  const map = new Map<string, Element>();
  for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp']) {
    for (const el of allEls(slideDoc, tag)) {
      const cNvPr = firstEl(el, 'p:cNvPr');
      const id = cNvPr?.getAttribute('id');
      if (id && !map.has(id)) map.set(id, el);
    }
  }
  return map;
}

function isInsideGroup(shapeEl: Element): boolean {
  let node = shapeEl.parentElement;
  while (node) {
    const ln = localName(node);
    if (ln === 'grpSp') return true;
    if (ln === 'spTree') return false;
    node = node.parentElement;
  }
  return false;
}

function getOwnXfrm(shapeEl: Element): { x: number; y: number; cx: number; cy: number } | null {
  const spPr = directChild(shapeEl, 'spPr') || directChild(shapeEl, 'grpSpPr');
  const xfrm = spPr ? directChild(spPr, 'xfrm') : null;
  const off = xfrm ? directChild(xfrm, 'off') : null;
  const ext = xfrm ? directChild(xfrm, 'ext') : null;
  if (!off || !ext) return null;
  const x = Number(off.getAttribute('x'));
  const y = Number(off.getAttribute('y'));
  const cx = Number(ext.getAttribute('cx'));
  const cy = Number(ext.getAttribute('cy'));
  if ([x, y, cx, cy].some((n) => !isFinite(n))) return null;
  return { x, y, cx, cy };
}

// One-level fallback for placeholders (title/body/content boxes) that have
// never been individually resized, and so have no <a:xfrm> of their own on
// the slide itself - PowerPoint then draws them at the position defined on
// the slide's layout. We stop at the layout (don't also fall back further to
// the slide master) - if even the layout doesn't pin it down, that's rare
// enough, and unusual enough, to just bail instead of guessing further.
async function resolvePlaceholderXfrmFromLayout(
  zip: JSZip,
  slidePath: string,
  shapeEl: Element
): Promise<{ x: number; y: number; cx: number; cy: number } | null> {
  const ph = firstEl(shapeEl, 'p:ph');
  if (!ph) return null;
  const phType = ph.getAttribute('type') || 'body';
  const phIdx = ph.getAttribute('idx');

  const slideDir = slidePath.split('/').slice(0, -1).join('/');
  const slideFile = slidePath.split('/').pop()!;
  const relsPath = `${slideDir}/_rels/${slideFile}.rels`;
  const relsDoc = await readXml(zip, relsPath);
  if (!relsDoc) return null;

  let layoutPath: string | null = null;
  for (const rel of allEls(relsDoc, 'Relationship')) {
    if ((rel.getAttribute('Type') || '').endsWith('/slideLayout')) {
      const target = rel.getAttribute('Target');
      if (target) layoutPath = resolveRelPath(relsPath, target);
      break;
    }
  }
  if (!layoutPath) return null;
  const layoutDoc = await readXml(zip, layoutPath);
  if (!layoutDoc) return null;

  for (const layoutShape of allEls(layoutDoc, 'p:sp')) {
    const layoutPh = firstEl(layoutShape, 'p:ph');
    if (!layoutPh) continue;
    const lIdx = layoutPh.getAttribute('idx');
    const lType = layoutPh.getAttribute('type') || 'body';
    const idxMatches = phIdx != null && lIdx != null && phIdx === lIdx;
    const typeOnlyMatches = phIdx == null && lIdx == null && phType === lType;
    if (idxMatches || typeOnlyMatches) {
      const xfrm = getOwnXfrm(layoutShape);
      if (xfrm) return xfrm;
    }
  }
  return null;
}

async function resolveShapeGeometry(
  zip: JSZip,
  slidePath: string,
  shapeEl: Element
): Promise<{ x: number; y: number; cx: number; cy: number } | null> {
  return getOwnXfrm(shapeEl) || resolvePlaceholderXfrmFromLayout(zip, slidePath, shapeEl);
}

function paragraphCountFor(shapeEl: Element): number {
  const txBody = directChild(shapeEl, 'txBody');
  return txBody ? directChildren(txBody, 'p').length : 0;
}

async function extractBuildsForSlide(
  zip: JSZip,
  slidePath: string,
  slideCx: number,
  slideCy: number
): Promise<SlideBuildInfo | undefined> {
  try {
    const slideDoc = await readXml(zip, slidePath);
    if (!slideDoc) return undefined;
    // Cheap early bail - a photo/pattern background makes pixel-sampled
    // masking (see BuildRevealOverlay.tsx) unreliable, so there's no point
    // even parsing the timing tree. (Solid colors and gradients are fine,
    // and don't need special-casing here - the viewer verifies those itself
    // against the actual rendered pixels right before it trusts a mask.)
    if (slideHasPictureOrPatternBackground(slideDoc)) return undefined;

    const timing = firstEl(slideDoc, 'p:timing');
    if (!timing) return undefined;

    const mainSeq = allEls(timing, 'p:seq').find(
      (s) => directChild(s, 'cTn')?.getAttribute('nodeType') === 'mainSeq'
    );
    if (!mainSeq) return undefined;
    const mainCtn = directChild(mainSeq, 'cTn');
    const childTnLst = mainCtn ? directChild(mainCtn, 'childTnLst') : null;
    if (!childTnLst) return undefined;
    const topPars = directChildren(childTnLst, 'par');
    if (!topPars.length) return undefined;

    const shapeIndex = buildShapeIndex(slideDoc);
    const steps: BuildStep[] = [];
    let currentBoxes: BuildBox[] = [];
    let sawAnyClick = false;

    for (const par of topPars) {
      const ctn = directChild(par, 'cTn');
      const nodeType = ctn?.getAttribute('nodeType');
      const presetClass = ctn?.getAttribute('presetClass');

      // Anything other than a click-triggered step or one chained
      // automatically off it (same step, no extra click needed) is timing
      // machinery we don't model - bail rather than guess.
      if (nodeType !== 'clickEffect' && nodeType !== 'withEffect' && nodeType !== 'afterEffect') return undefined;
      // Our model assumes everything starts hidden and only ever gets
      // revealed - an exit or emphasis effect mid-sequence breaks that
      // assumption, so bail rather than show something wrong.
      if (presetClass && presetClass !== 'entr') return undefined;
      // Per-character reveals are finer-grained than a bounding-box mask
      // can represent.
      if (allEls(par, 'p:charRg').length) return undefined;

      if (nodeType === 'clickEffect') {
        if (currentBoxes.length) steps.push({ boxes: currentBoxes });
        currentBoxes = [];
        sawAnyClick = true;
      }

      const targets = allEls(par, 'p:spTgt');
      if (!targets.length) {
        // A clickEffect step with nothing we can locate a shape for (e.g.
        // it's actually a sound/media cue) can't be represented - bail. A
        // with/afterEffect with no shape target is harmless to just skip.
        if (nodeType === 'clickEffect') return undefined;
        continue;
      }

      for (const tgt of targets) {
        const spid = tgt.getAttribute('spid');
        const shapeEl = spid ? shapeIndex.get(spid) : undefined;
        if (!shapeEl) return undefined;
        if (localName(shapeEl) !== 'sp' && localName(shapeEl) !== 'pic') return undefined; // graphicFrame/group - out of scope
        if (isInsideGroup(shapeEl)) return undefined;

        const geometry = await resolveShapeGeometry(zip, slidePath, shapeEl);
        if (!geometry) return undefined;

        const xPct = (geometry.x / slideCx) * 100;
        const yPct = (geometry.y / slideCy) * 100;
        const wPct = (geometry.cx / slideCx) * 100;
        const hPct = (geometry.cy / slideCy) * 100;

        const pRg = firstEl(tgt, 'p:pRg');
        if (pRg) {
          const total = paragraphCountFor(shapeEl);
          const st = Number(pRg.getAttribute('st'));
          const end = Number(pRg.getAttribute('end'));
          if (!total || !isFinite(st) || !isFinite(end) || st < 0 || end >= total || end < st) return undefined;
          const bandH = hPct / total;
          currentBoxes.push({ xPct, yPct: yPct + bandH * st, wPct, hPct: bandH * (end - st + 1) });
        } else {
          currentBoxes.push({ xPct, yPct, wPct, hPct });
        }
      }
    }
    if (currentBoxes.length) steps.push({ boxes: currentBoxes });

    if (!sawAnyClick || !steps.length) return undefined;
    return { steps };
  } catch {
    return undefined; // best-effort, same convention as notes/transitions
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
    const buildsByPage: Record<number, SlideBuildInfo> = {};

    for (let i = 0; i < slidePaths.length; i++) {
      const page = i + 1;
      const [notes, transition, builds] = await Promise.all([
        extractNotesForSlide(zip, slidePaths[i]),
        extractTransitionForSlide(zip, slidePaths[i]),
        slideSize ? extractBuildsForSlide(zip, slidePaths[i], slideSize.cx, slideSize.cy) : Promise.resolve(undefined),
      ]);
      if (notes) notesByPage[page] = notes;
      if (transition) transitionsByPage[page] = transition;
      if (builds) buildsByPage[page] = builds;
    }

    return { notesByPage, transitionsByPage, buildsByPage, slideCount: slidePaths.length };
  } catch (err) {
    console.warn('⚠️ PPTX metadata extraction skipped:', err);
    return empty;
  }
}
