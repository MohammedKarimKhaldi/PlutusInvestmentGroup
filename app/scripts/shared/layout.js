(function initPlutusLayout(global) {
  const appConfig = global.PlutusAppConfig;
  if (!appConfig) return;

  async function renderSidebarNav() {
    const activePageId = appConfig.getCurrentNavPageId();
    const navPages = appConfig.getNavPages();
    const appCore = global.AppCore;
    const visiblePages = [];

    for (const page of navPages) {
      if (appCore && typeof appCore.getPageAccessStatus === "function") {
        try {
          const access = await appCore.getPageAccessStatus(page.pageId);
          if (access && access.restricted && !access.allowed) continue;
        } catch {
          // If access resolution fails, keep the route visible rather than breaking navigation.
        }
      }
      visiblePages.push(page);
    }

    const navItems = visiblePages.map((page) => {
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

  async function resolveRouteLinks() {
    const appCore = global.AppCore;
    const links = Array.from(global.document.querySelectorAll("[data-route-id]"));

    for (const link of links) {
      const routeId = String(link.getAttribute("data-route-id") || "").trim();
      if (!routeId) continue;

      if (appCore && typeof appCore.getPageAccessStatus === "function") {
        try {
          const access = await appCore.getPageAccessStatus(routeId);
          if (access && access.restricted && !access.allowed) {
            link.hidden = true;
            link.removeAttribute("href");
            continue;
          }
        } catch {
          // Fall through and keep the route link usable.
        }
      }

      link.hidden = false;
      link.setAttribute("href", appConfig.buildPageHref(routeId));
    }
  }

  async function initializeLayout() {
    await renderSidebarNav();
    await resolveRouteLinks();
  }

  if (global.document.readyState === "loading") {
    global.document.addEventListener("DOMContentLoaded", () => {
      initializeLayout();
    }, { once: true });
  } else {
    initializeLayout();
  }
})(window);
