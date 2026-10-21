// The docs page's own behavior: the demo's settings, its HTML snippet, the event log,
// copy buttons, and the table of contents following the reader.

const demo = document.getElementById("demo");
const form = document.getElementById("playground");
const snippet = document.getElementById("demo-snippet");
const log = document.getElementById("event-log");

/** Each setting's default, which the element gets by leaving the attribute off. */
const DEFAULTS = { platform: "auto", mode: "movie", ribbon: "open", autoplay: "true" };

function settings() {
  const data = new FormData(form);
  return Object.fromEntries(Object.keys(DEFAULTS).map((name) => [name, String(data.get(name))]));
}

function apply() {
  const chosen = settings();
  const attrs = [`src="${demo.getAttribute("src")}"`];
  for (const [name, value] of Object.entries(chosen)) {
    if (value === DEFAULTS[name]) {
      demo.removeAttribute(name);
    } else {
      demo.setAttribute(name, value);
      attrs.push(`${name}="${value}"`);
    }
  }
  snippet.textContent = [
    `<script type="module" src="${document.querySelector('script[src*="simmy.js"]').src}"></script>`,
    "",
    `<simmy-app\n  ${attrs.join("\n  ")}\n></simmy-app>`,
  ].join("\n");
}

form.addEventListener("change", apply);
apply();

function logEvent(name, text) {
  if (log.querySelector(".muted")) {
    log.replaceChildren();
  }
  const item = document.createElement("li");
  const label = document.createElement("b");
  label.textContent = name;
  item.append(label, ` ${text}`);
  log.append(item);
  log.scrollTop = log.scrollHeight;
}

demo.addEventListener("simmy:ready", (event) => {
  const mod = event.detail;
  logEvent("simmy:ready", `${mod.workbook?.name ?? "module"}, ${mod.tasks?.length ?? 0} tasks`);
});
demo.addEventListener("simmy:graded", (event) => {
  const { earned, points } = event.detail;
  logEvent("simmy:graded", `${earned} of ${points} points`);
});
demo.addEventListener("simmy:error", (event) => {
  logEvent("simmy:error", String(event.detail?.message ?? event.detail));
});

// Every code block gets a copy button; the snippet's sits in its header.
for (const pre of document.querySelectorAll(".content pre")) {
  const text = pre.textContent;
  const button = document.createElement("button");
  button.type = "button";
  button.className = "copy copy-floating";
  button.textContent = "Copy";
  button.addEventListener("click", () => copy(button, text));
  pre.append(button);
}
for (const button of document.querySelectorAll("[data-copy]")) {
  button.addEventListener("click", () => copy(button, document.getElementById(button.dataset.copy).textContent));
}

async function copy(button, text) {
  try {
    await navigator.clipboard.writeText(text);
    button.textContent = "Copied";
  } catch {
    button.textContent = "Press Ctrl+C";
  }
  setTimeout(() => (button.textContent = "Copy"), 1500);
}

// Mark the section being read in the table of contents.
const links = new Map(
  [...document.querySelectorAll(".toc a")].map((link) => [link.getAttribute("href").slice(1), link]),
);
const watcher = new IntersectionObserver(
  (entries) => {
    for (const entry of entries) {
      if (entry.isIntersecting) {
        for (const link of links.values()) {
          link.classList.toggle("is-current", link === links.get(entry.target.id));
        }
      }
    }
  },
  { rootMargin: "-20% 0px -70% 0px" },
);
for (const id of links.keys()) {
  const section = document.getElementById(id);
  if (section) {
    watcher.observe(section);
  }
}
