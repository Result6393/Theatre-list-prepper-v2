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

  /** Insertion point: before item i (0..items.length). */
  const insertionAt = (y: number) => {
    for (let i = 0; i < items.length; i++) {
      const box = items[i].getBoundingClientRect();
      if (y < box.top + box.height / 2) return i;
    }
    return items.length;
  };

  const mark = () => {
    for (const item of items) item.classList.remove('drop-before', 'drop-after');
    if (target === from || target === from + 1) return;
    if (target < items.length) items[target].classList.add('drop-before');
    else items[items.length - 1].classList.add('drop-after');
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
    for (const item of items) item.classList.remove('drop-before', 'drop-after');
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
