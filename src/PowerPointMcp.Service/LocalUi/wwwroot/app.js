"use strict";
(() => {
  const token = new URLSearchParams(location.hash.slice(1)).get("token") || "";
  const $ = (selector) => document.querySelector(selector);
  const status = (text) => { $("#status").textContent = text; };

  async function api(path, params) {
    const url = new URL(path, location.origin);
    for (const [key, value] of Object.entries(params || {})) {
      if (value !== undefined && value !== null) url.searchParams.set(key, typeof value === "string" ? value : JSON.stringify(value));
    }
    const response = await fetch(url, { headers: { "X-PptMcp-Token": token }, cache: "no-store" });
    const body = await response.json();
    if (!response.ok || body.success === false) throw new Error(body.errorMessage || `HTTP ${response.status}`);
    return body;
  }
  const command = (name, args) => api("/api/command", { command: name, session: $("#session").value, args: args ? JSON.stringify(args) : undefined });

  function fill(tableSelector, rows, cells) {
    const body = $(`${tableSelector} tbody`);
    body.replaceChildren();
    for (const row of rows) {
      const tr = document.createElement("tr");
      for (const cell of cells(row)) {
        const td = document.createElement("td");
        if (cell && typeof cell === "object") { td.textContent = cell.text ?? ""; if (cell.className) td.className = cell.className; }
        else td.textContent = cell ?? "";
        tr.append(td);
      }
      body.append(tr);
    }
    if (rows.length === 0) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 6; td.textContent = "Nothing to show.";
      tr.append(td); body.append(tr);
    }
    return body;
  }

  async function loadSessions() {
    const { sessions } = await api("/api/sessions");
    const select = $("#session");
    const current = select.value;
    select.replaceChildren();
    if (!sessions.length) select.append(new Option("(no open presentation)", ""));
    for (const session of sessions) select.append(new Option(session.presentationPath, session.sessionId));
    if ([...select.options].some((option) => option.value === current)) select.value = current;
  }

  async function showSlides() {
    const result = await command("deck.summary", { includeTheme: false, maxSlides: 100 });
    const body = fill("#slide-table", result.slides || [], (slide) => [slide.slideIndex, slide.title || "", slide.layoutName || "", slide.shapeCount, slide.hidden ? "yes" : "", (slide.notes || "").slice(0, 80)]);
    [...body.rows].forEach((tr, index) => {
      const slide = (result.slides || [])[index];
      if (!slide) return;
      tr.tabIndex = 0;
      const open = () => showPreview(slide.slideIndex);
      tr.addEventListener("click", open);
      tr.addEventListener("keydown", (event) => { if (event.key === "Enter") open(); });
    });
  }

  async function showPreview(slideIndex) {
    status(`Rendering slide ${slideIndex}…`);
    const result = await command("preview.snapshot", { slideIndex, includeSketch: true });
    const figure = $("#preview");
    const image = figure.querySelector("img");
    if (result.imagePath) {
      const response = await fetch(new URL(`/api/file?path=${encodeURIComponent(result.imagePath)}`, location.origin), { headers: { "X-PptMcp-Token": token } });
      if (response.ok) {
        if (image.src.startsWith("blob:")) URL.revokeObjectURL(image.src);
        image.src = URL.createObjectURL(await response.blob());
        image.alt = `Slide ${slideIndex} as rendered by PowerPoint`;
        figure.hidden = false;
      }
    }
    figure.querySelector("figcaption").textContent = `Slide ${slideIndex}`;
    const sketch = $("#sketch");
    sketch.textContent = result.sketch || "";
    sketch.hidden = !result.sketch;
    status("");
  }

  async function showReview() {
    const result = await command("review.validate", { limit: 200 });
    fill("#finding-table", result.findings || [], (finding) => [finding.slideIndex, { text: finding.severity, className: `sev-${finding.severity}` }, finding.code, finding.message, finding.suggestion || ""]);
  }

  async function showPictures() {
    const result = await command("asset.inspect");
    fill("#picture-table", result.pictures || [], (picture) => [picture.slideIndex, picture.name, picture.effectivePpi ?? "?", picture.decorative ? "(decorative)" : picture.altText || "", (picture.issues || []).join(", ")]);
  }

  async function showDesign() {
    const profiles = await command("design.list-profiles");
    fill("#profile-table", profiles.profiles || [], (profile) => [profile.name, profile.source, profile.version || "", profile.description || ""]);
    const components = await command("design.list-components");
    fill("#component-table", components.components || [], (component) => [component.name, component.profileVersion || "", component.instances, component.outdated, (component.slides || []).join(", ")]);
  }

  async function showAbout() {
    const { report } = await api("/api/capabilities");
    const list = $("#about-list");
    list.replaceChildren();
    const add = (term, value) => { const dt = document.createElement("dt"); dt.textContent = term; const dd = document.createElement("dd"); dd.textContent = value; list.append(dt, dd); };
    add("Server", report.serverVersion);
    add("PowerPoint", report.powerPoint.version || report.powerPoint.note || "unknown");
    add("Open presentations", String(report.openSessions));
    add("Tools", (report.tools || []).map((tool) => tool.tool).join(", "));
    for (const guarantee of report.guarantees) add("•", guarantee);
    const workflows = $("#workflow-list");
    workflows.replaceChildren();
    for (const workflow of report.workflows || []) {
      const item = document.createElement("li");
      item.textContent = `${workflow.name} — ${workflow.when} ${workflow.steps.join(" → ")}`;
      workflows.append(item);
    }
  }

  const views = { slides: showSlides, review: showReview, pictures: showPictures, design: showDesign, about: showAbout };
  let currentView = "slides";

  async function show(view) {
    currentView = view;
    for (const tab of document.querySelectorAll("nav button")) tab.setAttribute("aria-selected", String(tab.dataset.view === view));
    for (const section of document.querySelectorAll(".view")) section.hidden = section.id !== view;
    if (view !== "about" && !$("#session").value) { status("Open a presentation through the agent or CLI, then Refresh."); return; }
    status("Loading…");
    try { await views[view](); status(""); }
    catch (error) { status(error.message); }
  }

  for (const tab of document.querySelectorAll("nav button")) tab.addEventListener("click", () => show(tab.dataset.view));
  $("#refresh").addEventListener("click", async () => { await loadSessions().catch((error) => status(error.message)); show(currentView); });
  $("#session").addEventListener("change", () => show(currentView));
  if (!token) status("Open the full URL printed by the server (it contains the access token).");
  loadSessions().then(() => show("slides")).catch((error) => status(error.message));
})();
