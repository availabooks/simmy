import type { ComponentChildren } from "preact";
import { useEffect, useRef, useState } from "preact/hooks";

/**
 * Which Excel to look like. "mac" and "win" are the desktop apps on their desktops. "web" is
 * Excel on its own, filling the stage the way it fills a browser tab, with nothing from any OS.
 */
export type Platform = "mac" | "win" | "web";
export type RibbonTab = "Home" | "Formulas";

const TABS_WIN = [
  "File",
  "Home",
  "Insert",
  "Draw",
  "Page Layout",
  "Formulas",
  "Data",
  "Review",
  "View",
  "Help",
];
const TABS_MAC = ["Home", "Insert", "Draw", "Page Layout", "Formulas", "Data", "Review", "View"];
const TABS_WEB = ["Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View"];
const TABS: Record<Platform, string[]> = { mac: TABS_MAC, win: TABS_WIN, web: TABS_WEB };

export const AUTOSUM_ITEMS: { label: string; fn: string }[] = [
  { label: "Sum", fn: "SUM" },
  { label: "Average", fn: "AVERAGE" },
  { label: "Count Numbers", fn: "COUNT" },
  { label: "Max", fn: "MAX" },
  { label: "Min", fn: "MIN" },
];

const STATISTICAL = [
  "AVERAGE",
  "COUNT",
  "COUNTIF",
  "CORREL",
  "MAX",
  "MEDIAN",
  "MIN",
  "MODE.SNGL",
  "STDEV.P",
  "STDEV.S",
  "VAR.S",
];

export type RibbonProps = {
  platform: Platform;
  tab: RibbonTab;
  onTab: (tab: RibbonTab) => void;
  numberFormat: string;
  paneOpen: boolean;
  /** The ribbon is collapsed and showing over the sheet for now, so its corner button pins it open. */
  collapsed: boolean;
  /** Collapse the ribbon to its tabs, or pin it open again. */
  onToggleRibbon: () => void;
  onTogglePane: () => void;
  onInsertFunction: (fn: string) => void;
  onStartFormula: () => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onClear: () => void;
  onFillDown: () => void;
};

/** The tab row. `tab` is null while the ribbon is collapsed, so no tab shows as selected. */
export function RibbonTabs({
  platform,
  tab,
  onTab,
  onToggleRibbon,
}: Pick<RibbonProps, "platform" | "onTab" | "onToggleRibbon"> & { tab: RibbonTab | null }) {
  const tabs = TABS[platform];
  return (
    <div class="xd-tabs" role="tablist">
      {tabs.map((name) => {
        const live = name === "Home" || name === "Formulas";
        return (
          <button
            key={name}
            type="button"
            role="tab"
            aria-selected={name === tab}
            class={`xd-tab${name === tab ? " is-active" : ""}${live ? "" : " is-idle"}`}
            onClick={() => live && onTab(name as RibbonTab)}
            onDblClick={onToggleRibbon}
          >
            {name}
          </button>
        );
      })}
      <span class="xd-tabs-spacer" />
      {platform !== "web" && (
        <>
          <span class="xd-tab-side is-idle">
            <Svg w={16}>
              <path d="M3 4h10v7H7l-3 3v-3H3z" fill="none" stroke="currentColor" />
            </Svg>
            Comments
          </span>
          <span class="xd-share is-idle">
            <Svg w={16}>
              <path d="M8 2v8M5 5l3-3 3 3M3 9v4h10V9" fill="none" stroke="currentColor" stroke-width="1.2" />
            </Svg>
            Share
          </span>
        </>
      )}
    </div>
  );
}

export function Ribbon(props: RibbonProps) {
  const toggleTitle = props.collapsed ? "Pin the Ribbon (Ctrl+F1)" : "Collapse the Ribbon (Ctrl+F1)";
  return (
    <div class="xd-ribbon">
      <button
        type="button"
        class="xd-ribbon-toggle"
        title={toggleTitle}
        aria-label={toggleTitle}
        onMouseDown={(event) => event.preventDefault()}
        onClick={props.onToggleRibbon}
      >
        <Svg w={12}>
          {props.collapsed ? (
            <path
              d="M10 1.5l4.5 4.5-1.6.6-2.4 2.4.2 3-1 1-2.7-2.7-4 4-.7-.7 4-4L3.6 6.9l1-1 3 .2 2.4-2.4z"
              fill="currentColor"
            />
          ) : (
            <path d="M4 13l6-6 6 6" fill="none" stroke="currentColor" stroke-width="1.8" />
          )}
        </Svg>
      </button>
      <div class="xd-ribbon-groups">
        {props.tab === "Home" ? <HomeGroups {...props} /> : <FormulaGroups {...props} />}
      </div>
      <Group label="Add-ins" fixed>
        <BigButton
          label="Excel Educator"
          title="Show the Excel Educator task pane"
          active={props.paneOpen}
          onClick={props.onTogglePane}
          icon={
            <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
              <rect x="2" y="2" width="28" height="28" rx="3" fill="#217346" />
              <rect x="7" y="7" width="18" height="18" rx="1" fill="#fff" />
              <path d="M7 13h18M7 19h18M13 7v18M19 7v18" stroke="#217346" stroke-width="1.3" />
            </svg>
          }
        />
      </Group>
    </div>
  );
}

function HomeGroups(props: RibbonProps) {
  return (
    <>
      <Group label="Clipboard">
        <BigButton label="Paste" dropdown onClick={props.onPaste} icon={<PasteIcon />} />
        <div class="xd-stack">
          <SmallButton title="Cut" label="Cut" onClick={props.onCut}>
            <Svg>
              <circle cx="5" cy="14" r="2.5" fill="none" stroke="#616161" />
              <circle cx="15" cy="14" r="2.5" fill="none" stroke="#616161" />
              <path d="M6.5 12L14 3M13.5 12L6 3" stroke="#616161" stroke-width="1.2" />
            </Svg>
          </SmallButton>
          <SmallButton title="Copy" label="Copy" onClick={props.onCopy}>
            <Svg>
              <rect x="6.5" y="3.5" width="9" height="11" rx="1" fill="#fff" stroke="#616161" />
              <path d="M4.5 6.5v10h8" fill="none" stroke="#616161" />
            </Svg>
          </SmallButton>
          <SmallButton title="Format Painter" label="Format Painter">
            <Svg>
              <rect x="3.5" y="3.5" width="12" height="5" rx="1" fill="#fff" stroke="#616161" />
              <path d="M15.5 6H17v4H10v3" fill="none" stroke="#616161" />
              <rect x="8.5" y="13" width="3" height="5" fill="#d29a2b" />
            </Svg>
          </SmallButton>
        </div>
      </Group>
      <Group label="Font">
        <div class="xd-font">
          <div class="xd-row">
            <span class="xd-select xd-select-font is-idle">Calibri (Body)</span>
            <span class="xd-select xd-select-size is-idle">11</span>
            <SmallButton title="Increase Font Size">
              <span class="xd-glyph">
                A<sup>^</sup>
              </span>
            </SmallButton>
            <SmallButton title="Decrease Font Size">
              <span class="xd-glyph xd-glyph-sm">
                A<sup>ˇ</sup>
              </span>
            </SmallButton>
          </div>
          <div class="xd-row">
            <SmallButton title="Bold">
              <span class="xd-glyph">
                <b>B</b>
              </span>
            </SmallButton>
            <SmallButton title="Italic">
              <span class="xd-glyph">
                <i>I</i>
              </span>
            </SmallButton>
            <SmallButton title="Underline">
              <span class="xd-glyph">
                <u>U</u>
              </span>
            </SmallButton>
            <span class="xd-sep" />
            <SmallButton title="Borders">
              <Svg>
                <rect
                  x="3.5"
                  y="3.5"
                  width="13"
                  height="13"
                  fill="none"
                  stroke="#9b9b9b"
                  stroke-dasharray="1.5 1.5"
                />
                <path d="M3 16.5h14" stroke="#262626" stroke-width="1.5" />
              </Svg>
            </SmallButton>
            <SmallButton title="Fill Color">
              <Svg>
                <path d="M5 9l5-5 5 5-5 5z" fill="#fff" stroke="#616161" />
                <path d="M15 10c1 1.5 1.5 2.3 1.5 3a1.5 1.5 0 01-3 0c0-.7.5-1.5 1.5-3z" fill="#616161" />
                <rect x="2" y="16" width="16" height="3" fill="#ffff00" />
              </Svg>
            </SmallButton>
            <SmallButton title="Font Color">
              <Svg>
                <path d="M6 14l4-10 4 10M7.5 10.5h5" fill="none" stroke="#262626" stroke-width="1.3" />
                <rect x="2" y="16" width="16" height="3" fill="#e00000" />
              </Svg>
            </SmallButton>
          </div>
        </div>
      </Group>
      <Group label="Alignment" optional>
        <div class="xd-stack xd-stack-grid">
          <div class="xd-row">
            <SmallButton title="Top Align">
              <Lines kind="top" />
            </SmallButton>
            <SmallButton title="Middle Align" active>
              <Lines kind="middle" />
            </SmallButton>
            <SmallButton title="Bottom Align">
              <Lines kind="bottom" />
            </SmallButton>
            <SmallButton title="Wrap Text" label="Wrap Text">
              <Svg>
                <path d="M3 5h14M3 10h11a3 3 0 010 6h-3M3 15h5" fill="none" stroke="#616161" />
                <path d="M12.5 14l-2 2 2 2" fill="none" stroke="#2b7bd6" />
              </Svg>
            </SmallButton>
          </div>
          <div class="xd-row">
            <SmallButton title="Align Left">
              <Lines kind="left" />
            </SmallButton>
            <SmallButton title="Center">
              <Lines kind="center" />
            </SmallButton>
            <SmallButton title="Align Right">
              <Lines kind="right" />
            </SmallButton>
            <SmallButton title="Merge & Center" label="Merge & Center">
              <Svg>
                <rect x="2.5" y="4.5" width="15" height="11" fill="#fff" stroke="#616161" />
                <path d="M5 10h10M7 8l-2 2 2 2M13 8l2 2-2 2" fill="none" stroke="#2b7bd6" />
              </Svg>
            </SmallButton>
          </div>
        </div>
      </Group>
      <Group label="Number">
        <div class="xd-stack">
          <div class="xd-row">
            <span class="xd-select xd-select-number">{props.numberFormat}</span>
          </div>
          <div class="xd-row">
            <SmallButton title="Accounting Number Format">
              <span class="xd-glyph">$</span>
            </SmallButton>
            <SmallButton title="Percent Style">
              <span class="xd-glyph">%</span>
            </SmallButton>
            <SmallButton title="Comma Style">
              <span class="xd-glyph">9</span>
            </SmallButton>
            <SmallButton title="Increase Decimal">
              <span class="xd-glyph xd-glyph-xs">
                ←.0
                <br />
                .00
              </span>
            </SmallButton>
            <SmallButton title="Decrease Decimal">
              <span class="xd-glyph xd-glyph-xs">
                .00
                <br />
                →.0
              </span>
            </SmallButton>
          </div>
        </div>
      </Group>
      <Group label="Styles" collapsible>
        <BigButton label="Conditional Formatting" dropdown icon={<CondIcon />} />
        <BigButton label="Format as Table" dropdown icon={<TableIcon />} />
        <BigButton label="Cell Styles" dropdown icon={<StylesIcon />} />
      </Group>
      <Group label="Cells" collapsible>
        <BigButton label="Insert" dropdown icon={<CellsIcon mark="+" />} />
        <BigButton label="Delete" dropdown icon={<CellsIcon mark="×" />} />
        <BigButton label="Format" dropdown icon={<CellsIcon mark="≡" />} />
      </Group>
      <Group label="Editing">
        <div class="xd-stack">
          <AutoSumButton small onPick={props.onInsertFunction} />
          <SmallButton title="Fill Down" label="Fill" onClick={props.onFillDown}>
            <Svg>
              <rect x="4.5" y="2.5" width="11" height="15" fill="#fff" stroke="#616161" />
              <path d="M10 5v9M7 11l3 3 3-3" fill="none" stroke="#2b7bd6" stroke-width="1.3" />
            </Svg>
          </SmallButton>
          <SmallButton title="Clear" label="Clear" onClick={props.onClear}>
            <Svg>
              <path d="M3 13l7-8 6 5-6 7H6z" fill="#f4b9c3" stroke="#c4314b" />
              <path d="M6 17h11" stroke="#616161" />
            </Svg>
          </SmallButton>
        </div>
        <BigButton
          label="Sort & Filter"
          dropdown
          icon={
            <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
              <text x="3" y="13" font-size="10" font-family="Segoe UI, sans-serif" fill="#2b7bd6">
                A
              </text>
              <text x="3" y="27" font-size="10" font-family="Segoe UI, sans-serif" fill="#616161">
                Z
              </text>
              <path d="M12 6v18M9 21l3 3 3-3" fill="none" stroke="#616161" />
              <path d="M17 8h12l-4.5 6v8l-3 2v-10z" fill="#fff" stroke="#616161" />
            </svg>
          }
        />
        <BigButton
          label="Find & Select"
          dropdown
          icon={
            <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
              <circle cx="13" cy="13" r="7.5" fill="#fff" stroke="#616161" stroke-width="1.5" />
              <path d="M18.5 18.5L27 27" stroke="#616161" stroke-width="2.5" stroke-linecap="round" />
            </svg>
          }
        />
      </Group>
    </>
  );
}

function FormulaGroups(props: RibbonProps) {
  const book = (glyph: string, color: string) => (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <path d="M6 5h17a2 2 0 012 2v20H8a2 2 0 01-2-2z" fill="#fff" stroke="#616161" />
      <path d="M6 23a2 2 0 012-2h17" fill="none" stroke="#616161" />
      <text
        x="15.5"
        y="17"
        text-anchor="middle"
        font-size="11"
        font-weight="700"
        font-family="Segoe UI, sans-serif"
        fill={color}
      >
        {glyph}
      </text>
    </svg>
  );
  return (
    <>
      <Group label="Function Library">
        <BigButton
          label="Insert Function"
          onClick={props.onStartFormula}
          icon={
            <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
              <text
                x="16"
                y="23"
                text-anchor="middle"
                font-size="19"
                font-style="italic"
                font-family="Cambria, Georgia, serif"
                fill="#262626"
              >
                fx
              </text>
            </svg>
          }
        />
        <AutoSumButton onPick={props.onInsertFunction} />
        <BigButton label="Recently Used" dropdown icon={book("★", "#d29a2b")} />
        <BigButton label="Financial" dropdown icon={book("$", "#217346")} />
        <BigButton label="Logical" dropdown icon={book("?", "#2b7bd6")} />
        <BigButton label="Text" dropdown icon={book("A", "#616161")} />
        <BigButton label="Date & Time" dropdown icon={book("◷", "#2b7bd6")} />
        <BigButton label="Lookup & Reference" dropdown icon={book("⌕", "#616161")} />
        <BigButton label="Math & Trig" dropdown icon={book("θ", "#8a3ab9")} />
        <Menu
          trigger={(open, toggle) => (
            <BigButton
              label="More Functions"
              dropdown
              active={open}
              onClick={toggle}
              icon={book("…", "#616161")}
            />
          )}
        >
          {(close) => (
            <>
              <div class="xd-menu-head">Statistical</div>
              {STATISTICAL.map((fn) => (
                <button
                  key={fn}
                  type="button"
                  class="xd-menu-item"
                  onClick={() => {
                    close();
                    props.onInsertFunction(fn);
                  }}
                >
                  {fn}
                </button>
              ))}
            </>
          )}
        </Menu>
      </Group>
      <Group label="Defined Names" collapsible>
        <BigButton label="Name Manager" icon={book("N", "#616161")} />
      </Group>
      <Group label="Formula Auditing" collapsible>
        <div class="xd-stack">
          <SmallButton title="Trace Precedents" label="Trace Precedents">
            <Svg>
              <path d="M3 15h4M13 5h4M7 15l6-10" fill="none" stroke="#2b7bd6" />
            </Svg>
          </SmallButton>
          <SmallButton title="Trace Dependents" label="Trace Dependents">
            <Svg>
              <path d="M3 5h4M13 15h4M7 5l6 10" fill="none" stroke="#2b7bd6" />
            </Svg>
          </SmallButton>
          <SmallButton title="Show Formulas" label="Show Formulas">
            <Svg>
              <text x="3" y="15" font-size="10" font-style="italic" fill="#262626">
                fx
              </text>
            </Svg>
          </SmallButton>
        </div>
      </Group>
      <Group label="Calculation" collapsible>
        <BigButton label="Calculation Options" dropdown icon={book("=", "#217346")} />
      </Group>
    </>
  );
}

function AutoSumButton({ onPick, small = false }: { onPick: (fn: string) => void; small?: boolean }) {
  return (
    <Menu
      trigger={(open, toggle) =>
        small ? (
          <SmallButton title="AutoSum" label="AutoSum" dropdown active={open} onClick={toggle}>
            <span class="xd-glyph xd-sigma">Σ</span>
          </SmallButton>
        ) : (
          <BigButton
            label="AutoSum"
            dropdown
            active={open}
            onClick={toggle}
            icon={
              <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
                <text
                  x="16"
                  y="25"
                  text-anchor="middle"
                  font-size="24"
                  font-family="Segoe UI, sans-serif"
                  fill="#262626"
                >
                  Σ
                </text>
              </svg>
            }
          />
        )
      }
    >
      {(close) =>
        AUTOSUM_ITEMS.map((item) => (
          <button
            key={item.fn}
            type="button"
            class="xd-menu-item"
            onClick={() => {
              close();
              onPick(item.fn);
            }}
          >
            {item.label}
          </button>
        ))
      }
    </Menu>
  );
}

function Menu({
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

function Group({
  label,
  children,
  fixed = false,
  collapsible = false,
  optional = false,
}: {
  label: string;
  children: ComponentChildren;
  fixed?: boolean;
  /** Hidden first when the window narrows. */
  collapsible?: boolean;
  /** Hidden next. */
  optional?: boolean;
}) {
  return (
    <div
      class={`xd-group${fixed ? " is-fixed" : ""}${collapsible ? " is-collapsible" : ""}${optional ? " is-optional" : ""}`}
    >
      <div class="xd-group-body">{children}</div>
      <div class="xd-group-label">{label}</div>
    </div>
  );
}

function BigButton({
  label,
  icon,
  dropdown = false,
  active = false,
  title,
  onClick,
}: {
  label: string;
  icon: ComponentChildren;
  dropdown?: boolean;
  active?: boolean;
  title?: string;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      class={`xd-big${active ? " is-active" : ""}${onClick ? "" : " is-idle"}`}
      title={idleTitle(title ?? label, onClick)}
      aria-disabled={onClick ? undefined : "true"}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {icon}
      <span class="xd-big-label">
        {label}
        {dropdown && <Caret />}
      </span>
    </button>
  );
}

function SmallButton({
  title,
  label,
  children,
  dropdown = false,
  active = false,
  onClick,
}: {
  title: string;
  label?: string;
  children: ComponentChildren;
  dropdown?: boolean;
  active?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      class={`xd-small${active ? " is-active" : ""}${onClick ? "" : " is-idle"}`}
      title={idleTitle(title, onClick)}
      aria-disabled={onClick ? undefined : "true"}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {children}
      {label && <span class="xd-small-label">{label}</span>}
      {dropdown && <Caret />}
    </button>
  );
}

/** The simulator draws all of Excel's ribbon but runs only part of it. The rest says so when hovered. */
function idleTitle(title: string, onClick: (() => void) | undefined): string {
  return onClick ? title : `${title} (not available here)`;
}

function Caret() {
  return (
    <svg class="xd-caret" viewBox="0 0 8 8" width="8" height="8" aria-hidden="true">
      <path d="M1.5 3l2.5 2.5L6.5 3" fill="none" stroke="currentColor" />
    </svg>
  );
}

export function Svg({ children, w = 20 }: { children: ComponentChildren; w?: number }) {
  return (
    <svg viewBox="0 0 20 20" width={w} height={w} aria-hidden="true">
      {children}
    </svg>
  );
}

function Lines({ kind }: { kind: "top" | "middle" | "bottom" | "left" | "center" | "right" }) {
  if (kind === "top" || kind === "middle" || kind === "bottom") {
    const y = kind === "top" ? 5 : kind === "middle" ? 8 : 11;
    return (
      <Svg>
        <path d={`M3 ${kind === "top" ? 3 : kind === "middle" ? 10 : 17}h14`} stroke="#616161" />
        <path d={`M7 ${y}h6M7 ${y + 2}h6M7 ${y + 4}h6`} stroke="#2b7bd6" />
      </Svg>
    );
  }
  const rows = [14, 9, 14, 9];
  return (
    <Svg>
      {rows.map((w, i) => {
        const x = kind === "left" ? 3 : kind === "center" ? 10 - w / 2 : 17 - w;
        return <path key={i} d={`M${x} ${4.5 + i * 3.7}h${w}`} stroke="#616161" />;
      })}
    </Svg>
  );
}

function PasteIcon() {
  return (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect x="4" y="5" width="18" height="23" rx="2" fill="#c9a26b" stroke="#8a6a3d" />
      <rect x="9" y="3" width="8" height="5" rx="1" fill="#fff" stroke="#616161" />
      <path d="M14 13h13v16H14z" fill="#fff" stroke="#616161" />
      <path d="M17 18h7M17 21h7M17 24h5" stroke="#2b7bd6" />
    </svg>
  );
}

function CondIcon() {
  return (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect x="4.5" y="5.5" width="23" height="21" fill="#fff" stroke="#616161" />
      <rect x="6" y="7" width="9" height="5" fill="#f8696b" />
      <rect x="6" y="14" width="9" height="5" fill="#ffeb84" />
      <rect x="6" y="21" width="9" height="4" fill="#63be7b" />
      <path d="M17 9.5h9M17 16.5h9M17 23h9" stroke="#9b9b9b" />
    </svg>
  );
}

function TableIcon() {
  return (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect x="4.5" y="5.5" width="23" height="21" fill="#fff" stroke="#616161" />
      <rect x="5" y="6" width="22" height="5" fill="#2b7bd6" />
      <path d="M5 16h22M5 21h22M12.5 11v15M20 11v15" stroke="#9b9b9b" />
      <rect x="5" y="16" width="22" height="5" fill="#dce9f7" opacity=".8" />
    </svg>
  );
}

function StylesIcon() {
  return (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect x="4.5" y="7.5" width="11" height="8" fill="#ffc7ce" stroke="#9c0006" />
      <rect x="16.5" y="7.5" width="11" height="8" fill="#c6efce" stroke="#006100" />
      <rect x="4.5" y="17.5" width="11" height="8" fill="#ffeb9c" stroke="#9c5700" />
      <rect x="16.5" y="17.5" width="11" height="8" fill="#fff" stroke="#616161" />
    </svg>
  );
}

function CellsIcon({ mark }: { mark: string }) {
  return (
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect x="4.5" y="6.5" width="20" height="18" fill="#fff" stroke="#616161" />
      <path d="M4.5 12.5h20M4.5 18.5h20M11.5 6.5v18M18 6.5v18" stroke="#c8c8c8" />
      <circle cx="24" cy="23" r="6" fill={mark === "×" ? "#c4314b" : mark === "+" ? "#217346" : "#616161"} />
      <text
        x="24"
        y="26.5"
        text-anchor="middle"
        font-size="10"
        font-weight="700"
        fill="#fff"
        font-family="Segoe UI, sans-serif"
      >
        {mark}
      </text>
    </svg>
  );
}
