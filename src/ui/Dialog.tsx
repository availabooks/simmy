import type { ComponentChildren, RefObject } from "preact";
import { useLayoutEffect, useRef } from "preact/hooks";

export type DialogProps = {
  title: string;
  /** The close button, Escape, and Cancel (or Close). */
  onCancel: () => void;
  /** OK, and Enter. Without it the dialog only has a Close button. */
  onSubmit?: () => void;
  submitLabel?: string;
  cancelLabel?: string;
  /** Buttons on the left of the button row, such as Go To's "Special…". */
  extra?: ComponentChildren;
  width?: number;
  /** What to focus when the dialog opens; the first field otherwise. */
  initialFocus?: RefObject<HTMLElement>;
  children: ComponentChildren;
};

const FOCUSABLE = "input:not([disabled]), select:not([disabled]), button:not([disabled]), [tabindex='0']";

/**
 * An Excel dialog, laid out like Excel for Windows: the title and a close button across the
 * top, the content, then the buttons along the bottom right. It covers the Excel window (not
 * the page), keeps focus inside, and keeps its keys from reaching the grid or the page.
 */
export function Dialog({
  title,
  onCancel,
  onSubmit,
  submitLabel = "OK",
  cancelLabel = onSubmit ? "Cancel" : "Close",
  extra,
  width = 380,
  initialFocus,
  children,
}: DialogProps) {
  const formRef = useRef<HTMLFormElement>(null);

  useLayoutEffect(() => {
    const first = formRef.current?.querySelector<HTMLElement>(`.xd-dialog-body :is(${FOCUSABLE})`);
    const target = initialFocus?.current ?? first;
    target?.focus({ preventScroll: true });
    if (target instanceof HTMLInputElement) {
      target.select();
    }
  }, []);

  const onKeyDown = (event: KeyboardEvent) => {
    event.stopPropagation();
    if (event.key === "Escape") {
      event.preventDefault();
      onCancel();
      return;
    }
    if (event.key !== "Tab" || !formRef.current) {
      return;
    }
    // Keep Tab inside the dialog, as a modal dialog does.
    const items = [...formRef.current.querySelectorAll<HTMLElement>(FOCUSABLE)];
    const root = formRef.current.getRootNode() as Document | ShadowRoot;
    const at = items.indexOf(root.activeElement as HTMLElement);
    const next = event.shiftKey ? (at <= 0 ? items.length - 1 : at - 1) : at === items.length - 1 ? 0 : at + 1;
    event.preventDefault();
    items[next]?.focus();
  };

  return (
    <div class="xd-dialog-veil">
      <form
        ref={formRef}
        class="xd-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        style={{ width }}
        onKeyDown={onKeyDown}
        onSubmit={(event) => {
          event.preventDefault();
          (onSubmit ?? onCancel)();
        }}
      >
        <div class="xd-dialog-title">
          <span>{title}</span>
          <button type="button" class="xd-dialog-x" title="Close" aria-label="Close" onClick={onCancel}>
            <svg viewBox="0 0 12 12" width="10" height="10" aria-hidden="true">
              <path d="M1 1l10 10M11 1L1 11" stroke="currentColor" stroke-width="1.2" />
            </svg>
          </button>
        </div>
        <div class="xd-dialog-body">{children}</div>
        <div class="xd-dialog-actions">
          <span class="xd-dialog-extra">{extra}</span>
          {onSubmit && (
            <button type="submit" class="xd-dialog-btn is-default">
              {submitLabel}
            </button>
          )}
          <button
            type={onSubmit ? "button" : "submit"}
            class={`xd-dialog-btn${onSubmit ? "" : " is-default"}`}
            onClick={onSubmit ? onCancel : undefined}
          >
            {cancelLabel}
          </button>
        </div>
      </form>
    </div>
  );
}
