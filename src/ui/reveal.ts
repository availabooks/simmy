/**
 * Bring `el` into view inside the box that scrolls it, the way scrollIntoView({ block: "nearest" })
 * would, but without touching the host page. scrollIntoView scrolls every scrolling ancestor,
 * the document included, so a narration line below the fold dragged the whole page down to it.
 * The search stops at the shadow root (parentElement is null there), so it never leaves simmy.
 */
export function revealInBox(el: HTMLElement): void {
  let box = el.parentElement;
  while (box && !/(auto|scroll)/.test(getComputedStyle(box).overflowY)) {
    box = box.parentElement;
  }
  if (!box) {
    return;
  }
  const line = el.getBoundingClientRect();
  // Where the line sits within the box's visible area.
  const top = line.top - box.getBoundingClientRect().top - box.clientTop;
  if (top < 0) {
    box.scrollTop += top;
  } else if (top + line.height > box.clientHeight) {
    box.scrollTop += Math.min(top, top + line.height - box.clientHeight);
  }
}
