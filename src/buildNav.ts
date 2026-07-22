import type { SlideBuildInfo } from './pptxParse';

// Present.tsx's and MobileRemote.tsx's own FlatSlide interfaces both satisfy
// this structurally - neither needs to import the other's type, they just
// both need a `build` field of this shape.
export interface HasBuild {
  build?: SlideBuildInfo;
}

export function stepCountFor(slide: HasBuild | undefined): number {
  return slide?.build?.steps.length || 0;
}

// How many steps should already be showing the *moment* you arrive at a
// slide by advancing forward into it - bullet 1 is visible immediately,
// not after an extra click (per the feature request), so this is 1 for any
// slide with builds, 0 for a slide with none.
function initialStepFor(slide: HasBuild | undefined): number {
  return stepCountFor(slide) > 0 ? 1 : 0;
}

// "Next" was pressed while sitting at flat slide `flatIdx` (0-based, into
// `flatSlides`), `buildStep` reveals in. Only moves to the next real slide
// once every build step on the current one has been used up - and lands the
// new slide with its own first bullet/object already showing.
export function computeNext(flatSlides: HasBuild[], flatIdx: number, buildStep: number): { flatIdx: number; buildStep: number } {
  const steps = stepCountFor(flatSlides[flatIdx]);
  if (buildStep < steps) return { flatIdx, buildStep: buildStep + 1 };
  const nextIdx = Math.min(flatIdx + 1, flatSlides.length - 1);
  return { flatIdx: nextIdx, buildStep: initialStepFor(flatSlides[nextIdx]) };
}

// "Previous": steps back one bullet/object if past the first one, otherwise
// moves to the previous slide - landing on it FULLY BUILT. That matches
// PowerPoint: going back never re-hides bullets on a slide you've already
// shown.
export function computePrev(flatSlides: HasBuild[], flatIdx: number, buildStep: number): { flatIdx: number; buildStep: number } {
  const initial = initialStepFor(flatSlides[flatIdx]);
  if (buildStep > initial) return { flatIdx, buildStep: buildStep - 1 };
  const prevIdx = Math.max(flatIdx - 1, 0);
  return { flatIdx: prevIdx, buildStep: stepCountFor(flatSlides[prevIdx]) };
}

// Jumping straight to a slide (thumbnail tap, first/last buttons, a remote
// reconnecting mid-session) always lands fully built too - same reasoning.
export function computeJump(flatSlides: HasBuild[], targetIdx: number): { flatIdx: number; buildStep: number } {
  const clamped = Math.max(0, Math.min(targetIdx, flatSlides.length - 1));
  return { flatIdx: clamped, buildStep: stepCountFor(flatSlides[clamped]) };
}
