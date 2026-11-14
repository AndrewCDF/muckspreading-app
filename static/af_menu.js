// The Menu button in the bottom bar slides up a panel of options; tap outside, press Escape or pick an option to close it.
(function () {
  const button = document.getElementById("af-menu-button");
  const menu = document.getElementById("af-menu");
  const backdrop = document.getElementById("af-menu-backdrop");
  if (!button || !menu || !backdrop) { return; }
  const icon = button.querySelector("strong");
  function isOpen() { return document.body.classList.contains("af-menu-open"); }
  function setOpen(open) {
    button.setAttribute("aria-expanded", open ? "true" : "false");
    if (icon) { icon.textContent = open ? "✕" : "☰"; }
    if (open) {
      menu.hidden = false;
      backdrop.hidden = false;
      requestAnimationFrame(function () { document.body.classList.add("af-menu-open"); });
      const current = menu.querySelector("[aria-current]") || menu.querySelector(".af-menu-item");
      if (current) { current.focus({ preventScroll: true }); }
    } else {
      document.body.classList.remove("af-menu-open");
      setTimeout(function () { if (!isOpen()) { menu.hidden = true; backdrop.hidden = true; } }, 260);
    }
  }
  window.afCloseMenu = function () { setOpen(false); };
  button.addEventListener("click", function () { setOpen(!isOpen()); });
  backdrop.addEventListener("click", function () { setOpen(false); });
  menu.addEventListener("click", function (event) { if (event.target.closest(".af-menu-item")) { setOpen(false); } });
  document.addEventListener("keydown", function (event) { if (event.key === "Escape" && isOpen()) { setOpen(false); button.focus(); } });
  window.addEventListener("pageshow", function () { document.body.classList.remove("af-menu-open"); menu.hidden = true; backdrop.hidden = true; button.setAttribute("aria-expanded", "false"); if (icon) { icon.textContent = "☰"; } });
})();
