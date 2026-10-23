import type { ComponentChildren } from "preact";
import { useEffect, useRef, useState } from "preact/hooks";

/** A dropdown: `trigger` opens it, and a press anywhere outside closes it. */
export function Menu({
  trigger,
  children,
}: {
  trigger: (open: boolean, toggle: () => void) => ComponentChildren;
  children: (close: () => void) => ComponentChildren;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLSpanElement>(null);
  useEffect(() => {
    if (!open) {
      return;
    }
    // composedPath sees through a shadow root, where event.target is retargeted to the host.
    const onDown = (event: PointerEvent) => {
      if (ref.current && !event.composedPath().includes(ref.current)) {
        setOpen(false);
      }
    };
    document.addEventListener("pointerdown", onDown);
    return () => document.removeEventListener("pointerdown", onDown);
  }, [open]);
  return (
    <span class="xd-menu-wrap" ref={ref}>
      {trigger(open, () => setOpen(!open))}
      {open && <div class="xd-menu">{children(() => setOpen(false))}</div>}
    </span>
  );
}
