// An in-app viewer for PDFs and photos. On an iPhone home-screen app a PDF or photo opened on its own has
// no browser buttons, so there is no way back; this shows it over the page with Close and Save / Share instead.
//   <a href="/invoice/history/3/show" data-viewer="View invoice">           opens a link in the viewer
//   <button formaction="/invoice/preview.pdf" data-viewer-submit="Preview"> sends its form and shows the result
(function () {
  const PDFJS = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/";
  let overlay, body, titleEl, saveButton, currentFile = null, objectUrl = null, pushedState = false;

  function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text) node.textContent = text;
    return node;
  }

  function build() {
    if (overlay) return;
    overlay = element("div", "af-viewer");
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.hidden = true;
    const bar = element("div", "af-viewer-bar");
    const close = element("button", "af-viewer-close", "✕ Close");
    close.type = "button";
    close.addEventListener("click", () => closeViewer());
    titleEl = element("p", "af-viewer-title");
    saveButton = element("button", "af-viewer-save", "Save / Share");
    saveButton.type = "button";
    saveButton.addEventListener("click", saveCurrent);
    bar.append(close, titleEl, saveButton);
    body = element("div", "af-viewer-body");
    overlay.append(bar, body);
    document.body.append(overlay);
    document.addEventListener("keydown", (event) => { if (event.key === "Escape" && !overlay.hidden) closeViewer(); });
    window.addEventListener("popstate", () => { if (!overlay.hidden) closeViewer(true); });
  }

  function show(title) {
    build();
    titleEl.textContent = title || "";
    body.replaceChildren(element("p", "af-viewer-message", "Loading…"));
    saveButton.hidden = true;
    currentFile = null;
    overlay.hidden = false;
    document.documentElement.classList.add("af-viewer-open");
    requestAnimationFrame(() => overlay.classList.add("is-open"));
    if (!pushedState) {  // the phone's back gesture closes the viewer instead of leaving the page
      history.pushState({ afViewer: true }, "");
      pushedState = true;
    }
  }

  function closeViewer(fromHistory) {
    if (!overlay || overlay.hidden) return;
    overlay.classList.remove("is-open");
    overlay.hidden = true;
    document.documentElement.classList.remove("af-viewer-open");
    body.replaceChildren();
    if (objectUrl) { URL.revokeObjectURL(objectUrl); objectUrl = null; }
    if (pushedState) {
      pushedState = false;
      if (!fromHistory) history.back();
    }
  }

  function message(text) {
    body.replaceChildren(element("p", "af-viewer-message", text));
  }

  function filenameFrom(response, fallback) {
    const header = response.headers.get("Content-Disposition") || "";
    const match = header.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
    return match ? decodeURIComponent(match[1]) : fallback;
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.onload = resolve;
      script.onerror = reject;
      document.head.append(script);
    });
  }

  async function renderPdf(blob) {
    try {
      if (!window.pdfjsLib) {
        await loadScript(PDFJS + "pdf.min.js");
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + "pdf.worker.min.js";
      }
      const pdf = await window.pdfjsLib.getDocument({ data: await blob.arrayBuffer() }).promise;
      const pages = element("div", "af-viewer-pages");
      body.replaceChildren(pages);
      const width = Math.min(body.clientWidth - 24, 900);
      const ratio = window.devicePixelRatio || 1;
      for (let number = 1; number <= pdf.numPages; number += 1) {
        const page = await pdf.getPage(number);
        const scale = width / page.getViewport({ scale: 1 }).width;
        const viewport = page.getViewport({ scale: scale * ratio });
        const canvas = element("canvas", "af-viewer-page");
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.width = Math.floor(viewport.width / ratio) + "px";
        pages.append(canvas);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      }
    } catch (error) {
      // Without the PDF reader (no internet), fall back to the browser's own view of the file, still under the Close bar.
      const frame = element("iframe", "af-viewer-frame");
      frame.src = objectUrl;
      frame.title = titleEl.textContent || "PDF";
      body.replaceChildren(frame);
    }
  }

  async function showResponse(response, fallbackName) {
    if (!response.ok) {
      const text = (await response.text()).trim();
      message(text && text.length < 300 && !text.startsWith("<") ? text : "This could not be opened. Check the details and try again.");
      return;
    }
    const blob = await response.blob();
    const name = filenameFrom(response, fallbackName);
    currentFile = new File([blob], name, { type: blob.type });
    objectUrl = URL.createObjectURL(blob);
    saveButton.hidden = false;
    if (blob.type.startsWith("image/")) {
      const image = element("img", "af-viewer-image");
      image.src = objectUrl;
      image.alt = titleEl.textContent || "Photo";
      body.replaceChildren(image);
    } else if (blob.type === "application/pdf") {
      await renderPdf(blob);
    } else {
      message("This file cannot be shown here. Use Save / Share to open it.");
    }
  }

  async function saveCurrent() {
    if (!currentFile) return;
    try {
      if (navigator.canShare && navigator.canShare({ files: [currentFile] })) {
        await navigator.share({ files: [currentFile], title: currentFile.name });
        return;
      }
    } catch (error) {
      if (error && error.name === "AbortError") return;  // they closed the share sheet
    }
    const link = element("a");
    link.href = objectUrl;
    link.download = currentFile.name;
    document.body.append(link);
    link.click();
    link.remove();
  }

  async function openUrl(url, title, fallbackName) {
    show(title);
    try {
      await showResponse(await fetch(url, { credentials: "same-origin" }), fallbackName || "file");
    } catch (error) {
      message("This could not be loaded. Check the connection and try again.");
    }
  }

  async function openForm(form, submitter, title) {
    if (!form.reportValidity()) return;
    const data = new FormData(form);
    if (submitter && submitter.name) data.append(submitter.name, submitter.value);
    show(title);
    try {
      const response = await fetch(submitter.getAttribute("formaction") || form.action, { method: "POST", body: data, credentials: "same-origin" });
      await showResponse(response, "preview.pdf");
    } catch (error) {
      message("This could not be loaded. Check the connection and try again.");
    }
  }

  document.addEventListener("click", (event) => {
    const link = event.target.closest("a[data-viewer]");
    if (link) {
      event.preventDefault();
      openUrl(link.href, link.dataset.viewer || link.textContent.trim(), link.dataset.viewerName);
      return;
    }
    const button = event.target.closest("button[data-viewer-submit]");
    if (button && button.form) {
      event.preventDefault();
      openForm(button.form, button, button.dataset.viewerSubmit || button.textContent.trim());
    }
  });

  window.afViewer = { openUrl, openForm, close: () => closeViewer() };
})();
