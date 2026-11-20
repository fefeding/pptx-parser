// Ribbon / menu icon set: 16×16 stroke glyphs, drawn with `currentColor` so they inherit
// the surrounding text colour (and therefore also the disabled / active states).
//
// Each entry is the *inner* markup of an <svg viewBox="0 0 16 16">; filled parts opt out with
// fill="currentColor" stroke="none". Keep them geometric — hand-written curves are hard to
// keep consistent at 16px.
const GLYPHS: Record<string, string> = {
  save: '<path d="M2.5 2.5h8.4l2.6 2.6v8.4H2.5z"/><path d="M5 2.5v3.6h6V2.5"/><rect x="5" y="9.2" width="6" height="4.3"/>',
  undo: '<path d="M6.3 4.4L2.9 7.6l3.4 3.2"/><path d="M2.9 7.6h6.3a4.1 4.1 0 0 1 0 8.2H6"/>',
  redo: '<path d="M9.7 4.4l3.4 3.2-3.4 3.2"/><path d="M13.1 7.6H6.8a4.1 4.1 0 0 0 0 8.2H10"/>',
  present: '<path d="M4.6 3.1l8.4 4.9-8.4 4.9z" fill="currentColor" stroke="none"/>',
  presentHere: '<path d="M3 3.1l7.2 4.2L3 11.5z" fill="currentColor" stroke="none"/><path d="M12.5 3v10"/>',
  endShow: '<rect x="4" y="4" width="8" height="8" rx="1"/>',
  help: '<circle cx="8" cy="8" r="6.2"/><path d="M6.3 6.3a1.8 1.8 0 1 1 3.1 1.3c-.6.6-1.4 1-1.4 2"/><circle cx="8" cy="11.8" r=".7" fill="currentColor" stroke="none"/>',
  newSlide:
    '<rect x="2" y="2.5" width="8.5" height="11" rx="1"/><path d="M4 5.8h4.5M4 8.4h4.5M4 11h3"/><path d="M13 7.6v5.6M10.2 10.4h5.6"/>',
  duplicateSlide: '<rect x="1.8" y="1.8" width="8.6" height="8.6" rx="1"/><path d="M5.6 13.8h8.6V5.4"/>',
  deleteSlide:
    '<path d="M2.8 4.4h10.4M6 4.4V2.9h4v1.5M4.4 4.4l.7 9h5.8l.7-9"/><path d="M6.9 7v4M9.1 7v4"/>',
  layout: '<rect x="2" y="2.5" width="12" height="11" rx="1"/><path d="M2 6.2h12M9 6.2v7.3"/>',
  showSlide: '<path d="M1.8 8S4.3 4.2 8 4.2 14.2 8 14.2 8 11.7 11.8 8 11.8 1.8 8 1.8 8z"/><circle cx="8" cy="8" r="1.7"/>',
  hideSlide:
    '<path d="M1.8 8S4.3 4.2 8 4.2 14.2 8 14.2 8 11.7 11.8 8 11.8 1.8 8 1.8 8z"/><circle cx="8" cy="8" r="1.7"/><path d="M2.6 13.4l10.8-10.8"/>',
  text: '<path d="M4.4 12.2L8 3.4l3.6 8.8"/><path d="M5.8 9.1h4.4"/>',
  shape: '<rect x="2" y="2.5" width="12" height="11" rx="1.6"/>',
  picture:
    '<rect x="1.6" y="3" width="12.8" height="10" rx="1"/><circle cx="5.4" cy="6.4" r="1.1"/><path d="M2.2 11.6l3.4-3.1 2.9 2.5 2-1.8 3.3 3"/>',
  table: '<rect x="2" y="3" width="12" height="10" rx="1"/><path d="M2 6.3h12M2 9.7h12M6.7 3v10M10.3 3v10"/>',
  chart:
    '<path d="M2.2 13.4h11.6"/><rect x="3.4" y="8.4" width="2.4" height="4" fill="currentColor" stroke="none"/><rect x="6.9" y="5.2" width="2.4" height="7.2" fill="currentColor" stroke="none"/><rect x="10.4" y="3.2" width="2.4" height="9.2" fill="currentColor" stroke="none"/>',
  media: '<rect x="1.5" y="3.5" width="13" height="9" rx="1.4"/><path d="M6.6 6l4 2-4 2z" fill="currentColor" stroke="none"/>',
  paste: '<path d="M6 3.2H4.2a1 1 0 0 0-1 1v8.6a1 1 0 0 0 1 1h7.6a1 1 0 0 0 1-1V4.2a1 1 0 0 0-1-1H10"/><rect x="6" y="1.7" width="4" height="2.8" rx=".7"/>',
  cut: '<circle cx="4.2" cy="12" r="1.7"/><circle cx="11.8" cy="12" r="1.7"/><path d="M5.4 10.8L11.8 2.6M10.6 10.8L4.2 2.6"/>',
  copy: '<rect x="5.6" y="5.6" width="7.9" height="7.9" rx="1"/><path d="M10.4 5.6V3.5a1 1 0 0 0-1-1H3.6a1 1 0 0 0-1 1v5.8a1 1 0 0 0 1 1h2.1"/>',
  duplicate: '<rect x="2.4" y="2.4" width="6.8" height="6.8" rx="1"/><rect x="6.8" y="6.8" width="6.8" height="6.8" rx="1"/>',
  group:
    '<rect x="1.8" y="4.6" width="6.8" height="6.8" rx="1" stroke-dasharray="2 1.5"/><rect x="7.4" y="2.6" width="6.8" height="6.8" rx="1" stroke-dasharray="2 1.5"/>',
  ungroup: '<rect x="1.6" y="2.6" width="6" height="6" rx="1"/><rect x="8.4" y="7.4" width="6" height="6" rx="1"/>',
  zFront:
    '<rect x="5.8" y="5.8" width="8" height="8" rx="1" fill="currentColor" fill-opacity=".22"/><rect x="2.2" y="2.2" width="8" height="8" rx="1"/>',
  zBack: '<rect x="2.2" y="2.2" width="8" height="8" rx="1" fill="currentColor" fill-opacity=".22"/><rect x="5.8" y="5.8" width="8" height="8" rx="1"/>',
  alignLeft: '<path d="M2.2 2v12"/><path d="M4.8 5h8.4M4.8 10.6h5.4"/>',
  alignCenter: '<path d="M8 2v12"/><path d="M4.2 5h7.6M5.9 10.6h4.2"/>',
  alignRight: '<path d="M13.8 2v12"/><path d="M2.8 5h8.4M5.8 10.6h5.4"/>',
  alignJustify: '<path d="M2.2 2v12"/><path d="M4.4 5h9M4.4 8h9M4.4 11h9"/>',
  alignTop: '<path d="M2 2.2h12"/><path d="M5 4.8v8M10.6 4.8v5.4"/>',
  alignMiddle: '<path d="M2 8h12"/><path d="M5 4.6v8M10.6 5.9v5.4"/>',
  alignBottom: '<path d="M2 13.8h12"/><path d="M5 3.2v8M10.6 5.8v5.4"/>',
  distH: '<path d="M2.4 2v12M13.6 2v12"/><rect x="6.4" y="5" width="3.2" height="6" rx=".6"/>',
  distV: '<path d="M2 2.4h12M2 13.6h12"/><rect x="5" y="6.4" width="6" height="3.2" rx=".6"/>',
  lock: '<rect x="3.6" y="7" width="8.8" height="6.8" rx="1.2"/><path d="M5.7 7V5.2a2.3 2.3 0 0 1 4.6 0V7"/>',
  unlock: '<rect x="3.6" y="7" width="8.8" height="6.8" rx="1.2"/><path d="M5.7 7V5.2a2.3 2.3 0 0 1 4.4-.7"/>',
  flipH: '<path d="M8 1.6v12.8" stroke-dasharray="2 2"/><path d="M6.2 4.6L2.4 8l3.8 3.4z" fill="currentColor" stroke="none"/><path d="M9.8 4.6L13.6 8l-3.8 3.4z"/>',
  flipV: '<path d="M1.6 8h12.8" stroke-dasharray="2 2"/><path d="M4.6 6.2L8 2.4l3.4 3.8z" fill="currentColor" stroke="none"/><path d="M4.6 9.8L8 13.6l3.4-3.8z"/>',
  shadow: '<circle cx="9.2" cy="9.2" r="4.6" fill="currentColor" fill-opacity=".3" stroke="none"/><path d="M4.4 6.2A5 5 0 0 1 9.6 2.6a4.8 4.8 0 0 1 4.2 6.9"/>',
  theme: '<path d="M8 14.4A6.4 6.4 0 1 1 8 1.6a6.4 6.4 0 0 1 0 12.8z"/><circle cx="5.5" cy="6.6" r=".95" fill="currentColor" stroke="none"/><circle cx="8" cy="5.1" r=".95" fill="currentColor" stroke="none"/><circle cx="10.5" cy="6.6" r=".95" fill="currentColor" stroke="none"/>',
  slideSize: '<rect x="1.6" y="3.6" width="12.8" height="8.8" rx="1"/><path d="M5 1.8v2M11 1.8v2M5 12.2v2M11 12.2v2"/>',
  background:
    '<path d="M6.9 2.3l5.5 5.5-4.9 4.9-5.5-5.5z"/><path d="M4.6 4.6l3.2 3.2"/><path d="M14 9.4c0 1-.7 1.8-1.5 1.8s-1.5-.8-1.5-1.8 1.5-2.3 1.5-2.3 1.5 1.3 1.5 2.3z" fill="currentColor" stroke="none"/>',
  transition:
    '<rect x="1.8" y="4" width="7" height="8" rx="1"/><path d="M13.2 6.2H10M12 4.8l1.2 1.4L12 7.6"/><path d="M9.4 12H7M8 10.6L6.8 12 8 13.4"/>',
  animation:
    '<circle cx="8" cy="8" r="2.1"/><path d="M8 2.2a5.8 5.8 0 0 1 5.8 5.8" stroke-dasharray="2 1.8"/><path d="M8 13.8A5.8 5.8 0 0 1 2.2 8" stroke-dasharray="2 1.8"/>',
  zoomIn: '<circle cx="6.9" cy="6.9" r="4.6"/><path d="M10.3 10.3l4 4"/><path d="M4.9 6.9h4M6.9 4.9v4"/>',
  zoomOut: '<circle cx="6.9" cy="6.9" r="4.6"/><path d="M10.3 10.3l4 4"/><path d="M4.9 6.9h4"/>',
  fit: '<path d="M2.2 5.6V2.2h3.4M13.8 5.6V2.2h-3.4M2.2 10.4v3.4h3.4M13.8 10.4v3.4h-3.4"/>',
  grid: '<rect x="2" y="2" width="12" height="12" rx="1"/><path d="M6 2v12M10 2v12M2 6h12M2 10h12"/>',
  snap: '<path d="M3.6 2.6v5.2a4.4 4.4 0 0 0 8.8 0V2.6"/><path d="M3.6 2.6h2.6M9.8 2.6h2.6"/><path d="M1.8 13.4h12.4" stroke-dasharray="2 2"/>',
  notes: '<rect x="2.6" y="1.8" width="10.8" height="12.4" rx="1"/><path d="M5 5.2h6M5 7.8h6M5 10.4h4"/>',
  move: '<path d="M8 2.2v11.6M2.2 8h11.6"/><path d="M8 2.2L6.2 4M8 2.2L9.8 4M8 13.8L6.2 12M8 13.8l1.6-1.8M2.2 8L4 6.2M2.2 8L4 9.8M13.8 8L12 6.2M13.8 8L12 9.8"/>',
  close: '<path d="M4 4l8 8M12 4l-8 8"/>',
  chevron: '<path d="M4.2 6.4L8 10.2l3.8-3.8"/>',
  up: '<path d="M8 12.6V4M4.6 7.4L8 4l3.4 3.4"/>',
  down: '<path d="M8 3.4V12M4.6 8.6L8 12l3.4-3.4"/>',
  add: '<path d="M8 3.4v9.2M3.4 8h9.2"/>',
  bullets:
    '<circle cx="3.4" cy="4.2" r="1.1" fill="currentColor" stroke="none"/><circle cx="3.4" cy="8" r="1.1" fill="currentColor" stroke="none"/><circle cx="3.4" cy="11.8" r="1.1" fill="currentColor" stroke="none"/><path d="M6.4 4.2h7.2M6.4 8h7.2M6.4 11.8h7.2"/>',
  fontColor:
    '<path d="M3.8 10.4L7 3.6l3.2 6.8"/><path d="M5 8.4h4"/><rect x="3" y="12" width="10" height="2.6" rx=".7" fill="currentColor" stroke="none"/>',
  fillColor:
    '<path d="M7 2.4l5.6 5.6-5 5-5.6-5.6z"/><path d="M14 9.6c0 1-.7 1.8-1.5 1.8s-1.5-.8-1.5-1.8 1.5-2.4 1.5-2.4 1.5 1.4 1.5 2.4z" fill="currentColor" stroke="none"/>',
  lineColor:
    '<path d="M11.6 2.4l2 2-7.8 7.8-2.7.7.7-2.7z"/><rect x="2.4" y="12.6" width="11.2" height="2.2" rx=".7" fill="currentColor" stroke="none"/>',
  lineSpacing:
    '<path d="M2 3.4h12M2 8h12M2 12.6h7"/><path d="M12.2 10.6v4M10.4 13l1.8 1.6 1.8-1.6"/>',
  props: '<rect x="2" y="2.5" width="12" height="11" rx="1"/><path d="M4.6 5.6h6.8M4.6 8h6.8M4.6 10.4h4"/>',
  download: '<path d="M8 2.6v7.6M4.8 7.2L8 10.4l3.2-3.2"/><path d="M2.8 13.4h10.4"/>',
  swap: '<path d="M2.5 5.6h9.3M9.4 3.2l2.4 2.4-2.4 2.4"/><path d="M13.5 10.4H4.2M6.6 8L4.2 10.4l2.4 2.4"/>',
  edit: '<path d="M11.4 2.6l2 2-7.6 7.6-2.6.6.6-2.6z"/><path d="M2.4 13.6h11.2"/>',
  menu: '<path d="M2.5 4.4h11M2.5 8h11M2.5 11.6h11"/>'
};

/** SVG markup for an icon name (unknown names render nothing, never throw). */
export function iconSvg(name: string, size = 16): string {
  const body = GLYPHS[name];
  if (!body) return '';
  return (
    `<svg class="ico" viewBox="0 0 16 16" width="${size}" height="${size}" fill="none" ` +
    `stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" ` +
    `aria-hidden="true" focusable="false">${body}</svg>`
  );
}

/** Icon as a DOM node, wrapped in a span so flex layout can align it with the label. */
export function icon(name: string, size = 16): HTMLElement {
  const span = document.createElement('span');
  span.className = 'rb-ico';
  span.innerHTML = iconSvg(name, size);
  return span;
}