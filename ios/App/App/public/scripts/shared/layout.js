(function initPlutusLayout(global) {
  const appConfig = global.PlutusAppConfig;
  if (!appConfig) return;

  function renderSidebarNav() {
    const activePageId = appConfig.getCurrentNavPageId();
    const navItems = appConfig.getNavPages().map((page) => {
      const activeClass = page.pageId === activePageId ? " active" : "";
      return `
        <li>
          <a class="nav-link${activeClass}" href="${appConfig.buildPageHref(page.pageId)}">
            <span class="dot"></span> ${page.label}
          </a>
        </li>
      `;
    }).join("");

    global.document.querySelectorAll("[data-sidebar-nav]").forEach((sidebar) => {
      sidebar.innerHTML = `
        <div class="sidebar-title">Navigation</div>
        <ul class="nav-list">${navItems}</ul>
      `;
    });
  }

  function resolveRouteLinks() {
    global.document.querySelectorAll("[data-route-id]").forEach((link) => {
      const routeId = String(link.getAttribute("data-route-id") || "").trim();
      if (!routeId) return;
      link.setAttribute("href", appConfig.buildPageHref(routeId));
    });
  }

  function initializeLayout() {
    renderSidebarNav();
    resolveRouteLinks();
  }

  if (global.document.readyState === "loading") {
    global.document.addEventListener("DOMContentLoaded", initializeLayout, { once: true });
  } else {
    initializeLayout();
  }
})(window);
