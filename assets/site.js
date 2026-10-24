/* Theme toggle, mobile menu, and publication filter. */
(function () {
  var root = document.documentElement;
  var KEY = "theme";

  // Runs in <head> so the saved theme applies before the page paints.
  try {
    var saved = localStorage.getItem(KEY);
    if (saved === "light" || saved === "dark") root.setAttribute("data-theme", saved);
  } catch (e) {}

  function currentTheme() {
    var set = root.getAttribute("data-theme");
    if (set) return set;
    return window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }

  document.addEventListener("DOMContentLoaded", function () {
    var toggle = document.querySelector(".theme-toggle");
    if (toggle) {
      toggle.setAttribute("aria-pressed", String(currentTheme() === "dark"));
      toggle.addEventListener("click", function () {
        var next = currentTheme() === "dark" ? "light" : "dark";
        root.setAttribute("data-theme", next);
        toggle.setAttribute("aria-pressed", String(next === "dark"));
        try { localStorage.setItem(KEY, next); } catch (e) {}
      });
    }

    var menu = document.querySelector(".nav-toggle");
    var nav = document.getElementById("site-nav");
    if (menu && nav) {
      menu.addEventListener("click", function () {
        var open = menu.getAttribute("aria-expanded") !== "true";
        menu.setAttribute("aria-expanded", String(open));
        nav.classList.toggle("open", open);
      });
    }

    var filter = document.getElementById("pub-filter");
    var count = document.getElementById("pub-count");
    if (filter) {
      filter.addEventListener("input", function () {
        var q = filter.value.trim().toLowerCase();
        var shown = 0;
        document.querySelectorAll(".pub-group").forEach(function (group) {
          var any = false;
          group.querySelectorAll("li").forEach(function (li) {
            var hit = !q || li.textContent.toLowerCase().indexOf(q) !== -1;
            li.hidden = !hit;
            if (hit) { any = true; shown++; }
          });
          group.hidden = !any;
        });
        if (count) {
          var total = count.getAttribute("data-total");
          count.textContent = q ? shown + " of " + total + " publications" : total + " publications";
        }
      });
    }
  });
})();
