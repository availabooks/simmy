/**
 * Excel's dialogs for going to cells and for defined names, labelled as in Excel for Windows.
 */
import { useRef, useState } from "preact/hooks";
import type { Names } from "../engine/sheet.ts";
import { Dialog } from "./Dialog.tsx";

/** The names, alphabetically, as Excel lists them. */
export function sortedNames(names: Names): string[] {
  return Object.keys(names).sort((a, b) => a.localeCompare(b));
}

/**
 * Go To (F5 or Ctrl+G): pick a name, or type a reference, a name or a spill (A1#). `onGo`
 * goes there and says whether it could.
 */
export function GoToDialog({
  names,
  onGo,
  onCancel,
}: {
  names: Names;
  onGo: (reference: string) => boolean;
  onCancel: () => void;
}) {
  const [reference, setReference] = useState("");
  const [problem, setProblem] = useState<string | null>(null);
  const referenceRef = useRef<HTMLInputElement>(null);
  const go = (text: string) => {
    if (!onGo(text)) {
      setProblem("Reference isn't valid.");
    }
  };
  return (
    <Dialog
      title="Go To"
      width={300}
      onCancel={onCancel}
      onSubmit={() => go(reference)}
      initialFocus={referenceRef}
      extra={
        <button type="button" class="xd-dialog-btn" disabled title="Special… (not available here)">
          Special…
        </button>
      }
    >
      <label class="xd-field-label" id="xd-goto-list">
        Go to:
      </label>
      <div class="xd-listbox" role="listbox" aria-labelledby="xd-goto-list">
        {sortedNames(names).map((name) => (
          <div
            key={name}
            role="option"
            tabIndex={0}
            aria-selected={reference === name}
            class={`xd-option${reference === name ? " is-on" : ""}`}
            onClick={() => {
              setReference(name);
              setProblem(null);
            }}
            onDblClick={() => go(name)}
          >
            {name}
          </div>
        ))}
      </div>
      <label class="xd-field-label">
        Reference:
        <input
          ref={referenceRef}
          class="xd-field"
          value={reference}
          spellcheck={false}
          onInput={(event) => {
            setReference(event.currentTarget.value);
            setProblem(null);
          }}
        />
      </label>
      {problem && (
        <p class="xd-dialog-error" role="alert">
          {problem}
        </p>
      )}
    </Dialog>
  );
}
