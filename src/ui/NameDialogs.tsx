/**
 * Excel's dialogs for going to cells and for defined names, labelled as in Excel for Windows.
 */
import { useRef, useState } from "preact/hooks";
import {
  absoluteRef,
  lookupName,
  nameProblem,
  nameRange,
  parseRefersTo,
  refersToText,
  type Sides,
} from "../engine/names.ts";
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

/** What New Name or Edit Name saves. */
export type NameEntry = { name: string; ref: string; comment: string };

/**
 * New Name (Formulas › Define Name), and Edit Name from the Name Manager: Name, Scope (always
 * Workbook here), Comment and Refers to.
 */
export function DefineNameDialog({
  title,
  initial,
  names,
  editing,
  sheet,
  onSave,
  onCancel,
}: {
  title: "New Name" | "Edit Name";
  initial: NameEntry;
  names: Names;
  /** The name being edited, which the new name may keep. */
  editing?: string;
  sheet: string;
  onSave: (entry: NameEntry) => void;
  onCancel: () => void;
}) {
  const [name, setName] = useState(initial.name);
  const [comment, setComment] = useState(initial.comment);
  const [refersTo, setRefersTo] = useState(initial.ref);
  const [problem, setProblem] = useState<string | null>(null);
  const save = () => {
    const trimmed = name.trim();
    if (nameProblem(trimmed) !== null) {
      setProblem(
        "The name that you entered is not valid. A name must begin with a letter or an underscore, " +
          "can't contain spaces or other invalid characters, and can't look like a cell reference.",
      );
      return;
    }
    const taken = lookupName(names, trimmed);
    if (taken && taken.name.toUpperCase() !== editing?.toUpperCase()) {
      setProblem("The name that you entered already exists. Enter a unique name.");
      return;
    }
    const range = parseRefersTo(refersTo, sheet);
    if (!range) {
      setProblem("There's a problem with this formula.");
      return;
    }
    onSave({ name: trimmed, ref: absoluteRef(range), comment: comment.slice(0, 255) });
  };
  return (
    <Dialog title={title} width={420} onCancel={onCancel} onSubmit={save}>
      <label class="xd-field-row">
        <span>Name:</span>
        <input
          class="xd-field"
          value={name}
          spellcheck={false}
          onInput={(event) => {
            setName(event.currentTarget.value);
            setProblem(null);
          }}
        />
      </label>
      <label class="xd-field-row">
        <span>Scope:</span>
        <select class="xd-field" disabled>
          <option>Workbook</option>
        </select>
      </label>
      <label class="xd-field-row is-top">
        <span>Comment:</span>
        <textarea
          class="xd-field"
          value={comment}
          maxLength={255}
          onInput={(event) => setComment(event.currentTarget.value)}
        />
      </label>
      <label class="xd-field-row">
        <span>Refers to:</span>
        <input
          class="xd-field"
          value={refersTo}
          spellcheck={false}
          onInput={(event) => {
            setRefersTo(event.currentTarget.value);
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

/** Create Names from Selection (Formulas › Create from Selection, or Ctrl+Shift+F3). */
export function CreateNamesDialog({
  initial,
  onCreate,
  onCancel,
}: {
  initial: Sides;
  onCreate: (sides: Sides) => void;
  onCancel: () => void;
}) {
  const [sides, setSides] = useState(initial);
  const boxes: [keyof Sides, string][] = [
    ["top", "Top row"],
    ["left", "Left column"],
    ["bottom", "Bottom row"],
    ["right", "Right column"],
  ];
  return (
    <Dialog
      title="Create Names from Selection"
      width={300}
      onCancel={onCancel}
      onSubmit={() => onCreate(sides)}
    >
      <span>Create names from values in the:</span>
      {boxes.map(([side, label]) => (
        <label key={side} class="xd-check">
          <input
            type="checkbox"
            checked={sides[side]}
            onChange={(event) => setSides({ ...sides, [side]: event.currentTarget.checked })}
          />
          {label}
        </label>
      ))}
    </Dialog>
  );
}

/**
 * The Name Manager (Formulas › Name Manager, or Ctrl+F3): every name with its value, what it
 * refers to, its scope and its comment. New… and Edit… open the name dialog over it.
 */
export function NameManagerDialog({
  names,
  comments,
  valueOf,
  sheet,
  onNew,
  onEdit,
  onDelete,
  onClose,
}: {
  names: Names;
  comments: Record<string, string>;
  /** What a name's Value column shows. */
  valueOf: (ref: string) => string;
  sheet: string;
  onNew: () => void;
  onEdit: (name: string) => void;
  onDelete: (name: string) => void;
  onClose: () => void;
}) {
  const list = sortedNames(names);
  const [picked, setPicked] = useState<string | null>(list[0] ?? null);
  const [confirming, setConfirming] = useState(false);
  const selected = picked !== null && picked in names ? picked : (list[0] ?? null);
  const refersTo = (name: string) => {
    const range = nameRange(names[name]);
    return range ? refersToText(range, sheet) : names[name];
  };
  return (
    <>
      <Dialog title="Name Manager" width={600} onCancel={onClose} cancelLabel="Close">
        <div class="xd-dialog-toolbar">
          <button type="button" class="xd-dialog-btn" onClick={onNew}>
            New…
          </button>
          <button
            type="button"
            class="xd-dialog-btn"
            disabled={!selected}
            onClick={() => selected && onEdit(selected)}
          >
            Edit…
          </button>
          <button type="button" class="xd-dialog-btn" disabled={!selected} onClick={() => setConfirming(true)}>
            Delete
          </button>
          <span class="xd-dialog-spacer" />
          <button type="button" class="xd-dialog-btn" disabled title="Filter (not available here)">
            Filter
          </button>
        </div>
        <div class="xd-table" role="grid" aria-label="Names">
          <div class="xd-table-row is-head" role="row">
            <span role="columnheader">Name</span>
            <span role="columnheader">Value</span>
            <span role="columnheader">Refers To</span>
            <span role="columnheader">Scope</span>
            <span role="columnheader">Comment</span>
          </div>
          {list.map((name) => (
            <div
              key={name}
              role="row"
              tabIndex={0}
              aria-selected={name === selected}
              class={`xd-table-row${name === selected ? " is-on" : ""}`}
              onClick={() => setPicked(name)}
              onDblClick={() => onEdit(name)}
            >
              <span role="gridcell">{name}</span>
              <span role="gridcell">{valueOf(names[name])}</span>
              <span role="gridcell">{refersTo(name)}</span>
              <span role="gridcell">Workbook</span>
              <span role="gridcell">{comments[name] ?? ""}</span>
            </div>
          ))}
        </div>
        <label class="xd-field-label">
          Refers to:
          <input class="xd-field" value={selected ? refersTo(selected) : ""} disabled />
        </label>
      </Dialog>
      {confirming && selected && (
        <div class="xd-dialog-layer">
          <Dialog
            title="Microsoft Excel"
            width={340}
            submitLabel="OK"
            onCancel={() => setConfirming(false)}
            onSubmit={() => {
              setConfirming(false);
              onDelete(selected);
            }}
          >
            <p class="xd-dialog-message">Are you sure you want to delete the name {selected}?</p>
          </Dialog>
        </div>
      )}
    </>
  );
}

/** Paste Name (F3): pick a name to put into the formula. */
export function PasteNameDialog({
  names,
  onPaste,
  onCancel,
}: {
  names: Names;
  onPaste: (name: string) => void;
  onCancel: () => void;
}) {
  const list = sortedNames(names);
  const [picked, setPicked] = useState(list[0] ?? "");
  return (
    <Dialog title="Paste Name" width={280} onCancel={onCancel} onSubmit={() => picked && onPaste(picked)}>
      <span>Paste name</span>
      <div class="xd-listbox" role="listbox" aria-label="Paste name">
        {list.map((name) => (
          <div
            key={name}
            role="option"
            tabIndex={0}
            aria-selected={picked === name}
            class={`xd-option${picked === name ? " is-on" : ""}`}
            onClick={() => setPicked(name)}
            onDblClick={() => onPaste(name)}
          >
            {name}
          </div>
        ))}
      </div>
    </Dialog>
  );
}
