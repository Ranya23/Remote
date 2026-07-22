import { useEffect, useState } from 'react';
import type { SlideBuildInfo, BuildBox } from './pptxParse';

interface MaskRect extends BuildBox {
  fill: string; // a solid CSS color, or a top-to-bottom linear-gradient
}

interface Props {
  build: SlideBuildInfo | undefined;
  buildStep: number; // how many build steps are already revealed
  canvasEl: HTMLCanvasElement | null; // the actual rendered PDF page, straight from react-pdf's canvasRef
  renderTick: number; // bumped by the caller every time canvasEl has freshly finished rendering, to trigger re-sampling
}

// Reads a handful of points around a box's edges from the canvas that's
// already showing this slide fully built, and only trusts it as maskable if
// those samples agree closely enough to look right as a flat/gradient
// patch. Anything noisier - a photo, a pattern, another shape poking in -
// means we can't reliably hide it, so the caller abandons masking for the
// whole slide rather than risk a mask that visibly doesn't match.
function sampleBox(ctx: CanvasRenderingContext2D, canvasW: number, canvasH: number, box: BuildBox): string | null {
  const x = Math.round((box.xPct / 100) * canvasW);
  const y = Math.round((box.yPct / 100) * canvasH);
  const w = Math.max(1, Math.round((box.wPct / 100) * canvasW));
  const h = Math.max(1, Math.round((box.hPct / 100) * canvasH));
  const clampX = (v: number) => Math.min(Math.max(v, 0), canvasW - 1);
  const clampY = (v: number) => Math.min(Math.max(v, 0), canvasH - 1);

  const corners: Array<[number, number]> = [
    [x + 2, y + 2], [x + w - 2, y + 2],
    [x + 2, y + h - 2], [x + w - 2, y + h - 2],
  ];
  const topMid: [number, number] = [x + w / 2, y + 2];
  const bottomMid: [number, number] = [x + w / 2, y + h - 2];

  const readPixel = ([px, py]: [number, number]) => {
    const d = ctx.getImageData(clampX(px), clampY(py), 1, 1).data;
    return [d[0], d[1], d[2]] as [number, number, number];
  };

  const samples = [...corners, topMid, bottomMid].map(readPixel);
  const avg = samples.reduce((a, c) => [a[0] + c[0], a[1] + c[1], a[2] + c[2]], [0, 0, 0]).map((v) => v / samples.length);
  const maxDelta = samples.reduce((m, c) => {
    const d = Math.abs(c[0] - avg[0]) + Math.abs(c[1] - avg[1]) + Math.abs(c[2] - avg[2]);
    return Math.max(m, d);
  }, 0);
  // Too much variation anywhere in the box - don't trust a flat/gradient fill here.
  if (maxDelta > 40) return null;

  const rgb = (c: number[]) => `rgb(${Math.round(c[0])}, ${Math.round(c[1])}, ${Math.round(c[2])})`;
  const top = readPixel(topMid);
  const bottom = readPixel(bottomMid);
  const topBottomDelta = Math.abs(top[0] - bottom[0]) + Math.abs(top[1] - bottom[1]) + Math.abs(top[2] - bottom[2]);
  if (topBottomDelta > 12) return `linear-gradient(to bottom, ${rgb(top)}, ${rgb(bottom)})`;
  return rgb(avg);
}

export default function BuildRevealOverlay({ build, buildStep, canvasEl, renderTick }: Props) {
  const [masks, setMasks] = useState<MaskRect[] | null>(null);

  useEffect(() => {
    if (!build || !build.steps.length || !canvasEl) {
      setMasks(null);
      return;
    }
    try {
      const ctx = canvasEl.getContext('2d');
      if (!ctx) { setMasks(null); return; }
      const w = canvasEl.width;
      const h = canvasEl.height;
      if (!w || !h) { setMasks(null); return; }

      const rects: MaskRect[] = [];
      let allOk = true;
      outer: for (const step of build.steps) {
        for (const box of step.boxes) {
          const fill = sampleBox(ctx, w, h, box);
          if (!fill) { allOk = false; break outer; }
          rects.push({ ...box, fill });
        }
      }
      setMasks(allOk ? rects : null);
    } catch (err) {
      // getImageData can throw on a tainted canvas - shouldn't happen here
      // (the PDF is decoded from a same-origin blob: URL, never fetched
      // cross-origin) but fail safe if it ever does.
      console.warn('⚠️ Build-animation masking skipped for this slide:', err);
      setMasks(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [build, canvasEl, renderTick]);

  if (!masks || !build) return null;

  // masks[] is flattened across all steps, in the same order steps/boxes
  // were produced in - so "reveal everything through step N" is just "the
  // first `revealedCount` masks are hidden".
  const revealedCount = build.steps.slice(0, buildStep).reduce((n, s) => n + s.boxes.length, 0);

  return (
    <div className="absolute inset-0 pointer-events-none" style={{ zIndex: 5 }}>
      {masks.map((m, i) => (
        <div
          key={i}
          style={{
            position: 'absolute',
            left: `${m.xPct}%`,
            top: `${m.yPct}%`,
            width: `${m.wPct}%`,
            height: `${m.hPct}%`,
            background: m.fill,
            opacity: i < revealedCount ? 0 : 1,
            transition: 'opacity 200ms ease-out',
          }}
        />
      ))}
    </div>
  );
}
