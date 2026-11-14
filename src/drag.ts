/**
 * Drag-to-reorder with a grip handle. Pointer events rather than HTML drag-and-drop, so it also works
 * with touch (for phones later) and textareas in the rows keep normal text selection.
 * The page scrolls when you drag near the top or bottom edge, for long lists.
 */
export interface DragOptions {
  /** The rows (one element per patient) that can be reordered, in order. */
  items: HTMLElement[];
  from: number;
  /** Shown next to the pointer while dragging, e.g. the patient's name. */
  label: string;
  onDrop: (from: number, to: number) => void;
}

const START_DISTANCE = 4;
const EDGE = 70;

export function startDrag(event: PointerEvent, options: DragOptions): void {
  if (event.button !== 0) return;
  const grip = event.currentTarget as HTMLElement;
  const { items, from } = options;
  const source = items[from];
  const startY = event.clientY;
  let pointerY = event.clientY;
  let pointerX = event.clientX;
  let dragging = false;
  let target = from;
  let frame = 0;
  const ghost = document.createElement('div');
  ghost.className = 'drag-ghost';
  ghost.textContent = options.label;

  grip.setPointerCapture(event.pointerId);
  event.preventDefault();

  // Where everything was when the drag started (in page coordinates), so the gap that opens up
  // doesn't move the insertion points while you drag.
  const rects = items.map((el) => {
    const box = el.getBoundingClientRect();
    return { top: box.top + scrollY, bottom: box.bottom + scrollY };
  });
  const spacing = items.length > 1 ? Math.max(0, rects[1].top - rects[0].bottom) : 0;
  const room = rects[from].bottom - rects[from].top + spacing;

  /** Insertion point: before item i (0..items.length). */
  const insertionAt = (y: number) => {
    const at = y + scrollY;
    for (let i = 0; i < items.length; i++) {
      if (at < (rects[i].top + rects[i].bottom) / 2) return i;
    }
    return items.length;
  };

  /**
   * The patients between the old and new place slide over to open a gap where the dragged one will
   * land, and the dragged one moves into it (shown as a dashed outline), as in a sortable list.
   */
  const mark = () => {
    const to = target > from ? target - 1 : target;
    items.forEach((el, i) => {
      let shift = 0;
      if (i === from) shift = to > from ? rects[to].bottom - rects[from].bottom : to < from ? rects[to].top - rects[from].top : 0;
      else if (to > from && i > from && i <= to) shift = -room;
      else if (to < from && i >= to && i < from) shift = room;
      el.style.transform = shift ? `translateY(${shift}px)` : '';
    });
  };

  const tick = () => {
    if (!dragging) return;
    const speed = pointerY < EDGE ? -(EDGE - pointerY) / 4 : pointerY > innerHeight - EDGE ? (pointerY - (innerHeight - EDGE)) / 4 : 0;
    if (speed) {
      window.scrollBy(0, speed);
      target = insertionAt(pointerY);
      mark();
    }
    frame = requestAnimationFrame(tick);
  };

  const move = (e: PointerEvent) => {
    pointerY = e.clientY;
    pointerX = e.clientX;
    if (!dragging) {
      if (Math.abs(pointerY - startY) < START_DISTANCE) return;
      dragging = true;
      for (const item of items) item.classList.add('drag-shift');
      source.classList.add('dragging');
      document.body.classList.add('is-dragging');
      document.body.append(ghost);
      frame = requestAnimationFrame(tick);
    }
    ghost.style.transform = `translate(${pointerX + 14}px, ${pointerY - 12}px)`;
    target = insertionAt(pointerY);
    mark();
  };

  const finish = (drop: boolean) => {
    grip.removeEventListener('pointermove', move);
    grip.removeEventListener('pointerup', up);
    grip.removeEventListener('pointercancel', cancel);
    cancelAnimationFrame(frame);
    ghost.remove();
    document.body.classList.remove('is-dragging');
    source.classList.remove('dragging');
    for (const item of items) {
      item.classList.remove('drag-shift');
      item.style.transform = '';
    }
    if (!drop || !dragging) return;
    const to = target > from ? target - 1 : target;
    if (to !== from) options.onDrop(from, to);
  };
  const up = () => finish(true);
  const cancel = () => finish(false);

  grip.addEventListener('pointermove', move);
  grip.addEventListener('pointerup', up);
  grip.addEventListener('pointercancel', cancel);
}
