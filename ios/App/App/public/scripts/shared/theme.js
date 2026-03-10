(function initThemeSwitcher(global) {
  const THEME_STORAGE_KEY = "app_theme_v1";
  const THEMES = ["classic", "coastal"];

  function readTheme() {
    try {
      const raw = localStorage.getItem(THEME_STORAGE_KEY);
      if (THEMES.includes(raw)) return raw;
    } catch (e) {
      console.warn("Theme read failed", e);
    }
    return "classic";
  }

  function writeTheme(theme) {
    try {
      localStorage.setItem(THEME_STORAGE_KEY, theme);
    } catch (e) {
      console.warn("Theme write failed", e);
    }
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    if (!root) return;
    if (theme === "classic") root.removeAttribute("data-theme");
    else root.setAttribute("data-theme", theme);
  }

  function nextTheme(current) {
    const idx = THEMES.indexOf(current);
    if (idx < 0) return THEMES[0];
    return THEMES[(idx + 1) % THEMES.length];
  }

  function ensureToggle() {
    const sidebars = Array.from(document.querySelectorAll(".sidebar-nav"));
    sidebars.forEach((sidebar) => {
      if (sidebar.querySelector(".theme-switch")) return;
      const wrap = document.createElement("div");
      wrap.className = "theme-switch";

      const button = document.createElement("button");
      button.type = "button";
      button.className = "theme-toggle-btn";
      wrap.appendChild(button);
      sidebar.appendChild(wrap);

      const renderLabel = () => {
        const active = readTheme();
        button.textContent = active === "classic" ? "Theme: Classic" : "Theme: Coastal";
      };

      button.addEventListener("click", () => {
        const current = readTheme();
        const updated = nextTheme(current);
        writeTheme(updated);
        applyTheme(updated);
        renderLabel();
      });

      renderLabel();
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    applyTheme(readTheme());
    ensureToggle();
  });

  applyTheme(readTheme());
})(window);
