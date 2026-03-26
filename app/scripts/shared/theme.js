(function initThemeSwitcher(global) {
  const THEME_STORAGE_KEY = "app_theme_v1";
  const DASHBOARD_THEME_STORAGE_KEY = "dashboard_theme_v1";
  const THEMES = ["classic", "coastal"];
  const DASHBOARD_THEMES = ["dark", "light"];

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

  function isDashboardPage() {
    const body = document.body;
    if (!body) return false;
    const pageId = String(body.getAttribute("data-page-id") || "").trim().toLowerCase();
    return pageId === "investor-dashboard" || pageId === "deal-ownership";
  }

  function readDashboardTheme() {
    try {
      const raw = localStorage.getItem(DASHBOARD_THEME_STORAGE_KEY);
      if (DASHBOARD_THEMES.includes(raw)) return raw;
    } catch (e) {
      console.warn("Dashboard theme read failed", e);
    }
    return "light";
  }

  function writeDashboardTheme(theme) {
    try {
      localStorage.setItem(DASHBOARD_THEME_STORAGE_KEY, theme);
    } catch (e) {
      console.warn("Dashboard theme write failed", e);
    }
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    if (!root) return;
    if (theme === "classic") root.removeAttribute("data-theme");
    else root.setAttribute("data-theme", theme);
  }

  function applyDashboardTheme(theme) {
    const body = document.body;
    if (!body || !isDashboardPage()) return;
    body.setAttribute("data-dashboard-theme", theme === "light" ? "light" : "dark");
  }

  function nextTheme(current) {
    const idx = THEMES.indexOf(current);
    if (idx < 0) return THEMES[0];
    return THEMES[(idx + 1) % THEMES.length];
  }

  function nextDashboardTheme(current) {
    return current === "light" ? "dark" : "light";
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
        if (isDashboardPage()) {
          const active = readDashboardTheme();
          button.textContent = active === "light" ? "Theme: Light Gray" : "Theme: Dark";
          return;
        }

        const active = readTheme();
        button.textContent = active === "classic" ? "Theme: Classic" : "Theme: Coastal";
      };

      button.addEventListener("click", () => {
        if (isDashboardPage()) {
          const current = readDashboardTheme();
          const updated = nextDashboardTheme(current);
          writeDashboardTheme(updated);
          applyDashboardTheme(updated);
          renderLabel();
          return;
        }

        const current = readTheme();
        const updated = nextTheme(current);
        writeTheme(updated);
        applyTheme(updated);
        renderLabel();
      });

      renderLabel();
    });
  }

  function applyStoredThemes() {
    if (isDashboardPage()) applyTheme("classic");
    else applyTheme(readTheme());
    applyDashboardTheme(readDashboardTheme());
  }

  document.addEventListener("DOMContentLoaded", () => {
    applyStoredThemes();
    ensureToggle();
    setTimeout(ensureToggle, 0);
  });

  global.addEventListener("load", ensureToggle, { once: true });

  applyStoredThemes();
})(window);
