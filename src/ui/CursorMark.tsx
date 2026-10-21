/** The autopilot pointer, also used on the Take control button. */
export function CursorMark(props: { size: number }) {
  return (
    <svg class="xd-cursor-mark" viewBox="0 0 24 24" width={props.size} height={props.size} aria-hidden="true">
      <path
        d="M4 2.5v17l4.6-4.2 3 6.4 3.1-1.4-3-6.3h6.3z"
        fill="#fff"
        stroke="#111"
        stroke-width="1.4"
        stroke-linejoin="round"
      />
    </svg>
  );
}
