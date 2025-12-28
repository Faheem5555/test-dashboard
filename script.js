/* =========================
   Power BI Prototype Dashboard Website
   FULL IMPLEMENTATION (no skips)

   ✅ FEATURE INDEX (what this JS includes)
   1) Fixed Power BI-like canvas (dynamic size with presets) — no infinite scroll; visuals don't push each other
   2) Visual picker dropdown (category-wise)
   3) Supported visuals (distinct):
      - Pie, Donut (Chart.js)
      - Treemap (custom prototype)
      - Ribbon (custom SVG prototype)
      - Line (multi-trend), Area, Stacked Area
      - Clustered/Stacked/100% Bar
      - Clustered/Stacked/100% Column
      - Line & Clustered Column (combo)
      - Line & Stacked Column (combo)
      - Scatter
      - KPI (prototype)
      - Card (prototype)
      - Multi-row Card (prototype)
      - Text Box (editable + formatting)
      - Upload Image (acts like a visual)
   4) Interaction:
      - Select visual -> selection outline
      - Delete (✖) shows only when selected
      - Click another visual to switch selection
      - Click empty canvas -> deselect, Format pane stays visible (Canvas settings)
      - Bring-to-front on selection (Power BI-like)
   5) Drag + Resize:
      - Smooth dragging by header
      - Resize handles visible only when selected
      - Clamped within canvas on ALL sides (left/top/right/bottom)
   6) Format Pane (always visible):
      - Canvas settings when none selected:
         • preset selector (16:9: 1280×720, 1920×1080)
         • width/height editable + apply
         • background color, opacity, background image URL/data URL
         • upload canvas background theme JSON (Browse)
      - Visual settings when selected:
         • title
         • x/y/w/h
         • series name edits + per-series color edits
         • reset to theme palette
         • image replace for image visual
         • text box formatting
   7) Theme import (Power BI-like):
      - Import Theme JSON
      - Toast: “Theme import successful”
      - Applies to existing + new visuals immediately
      - Sample theme download
   8) Canvas background theme upload:
      - Supported file JSON:
        { "backgroundColor": "...", "backgroundImage": "...", "opacity": 0.9 }
      - Applies instantly + toast
      - Sample canvas background download
   9) Dashboard save/load:
      - Download dashboard JSON (includes canvas, theme, visuals, series names/colors, image base64, text box props)
      - Upload dashboard JSON:
          • If non-default state -> confirmation modal (Cancel / Discard & Load)
          • If empty default dashboard -> NO confirmation modal
      - ✅ Modal never shows on page load (forced hidden)
  10) Power BI-like “Fit to screen”:
      - Canvas scales down to always fully visible (no canvas clipped off screen)
========================= */

(() => {
  // -------------------------
  // Helpers first (avoid any init-order issues)
  // -------------------------
  function clamp(v, min, max) {
    return Math.max(min, Math.min(max, v));
  }

  function escapeHtml(str){
    return String(str ?? "")
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }
  function escapeAttr(str){ return escapeHtml(str).replaceAll("\n"," "); }

  // -------------------------
  // DOM
  // -------------------------
  const canvas = document.getElementById("canvas");
  const statusPill = document.getElementById("statusPill");

  const visualPickerToggle = document.getElementById("visualPickerToggle");
  const visualPickerMenu = document.getElementById("visualPickerMenu");

  const formatBody = document.getElementById("formatBody");
  const toastEl = document.getElementById("toast");

  const importThemeBtn = document.getElementById("importThemeBtn");
  const importThemeInput = document.getElementById("importThemeInput");

  const downloadDashboardBtn = document.getElementById("downloadDashboardBtn");
  const uploadDashboardBtn = document.getElementById("uploadDashboardBtn");
  const uploadDashboardInput = document.getElementById("uploadDashboardInput");

  const downloadSampleThemeBtn = document.getElementById("downloadSampleThemeBtn");
  const downloadSampleCanvasBgBtn = document.getElementById("downloadSampleCanvasBgBtn");

  const modalBackdrop = document.getElementById("modalBackdrop");
  const modalCancelBtn = document.getElementById("modalCancelBtn");
  const modalConfirmBtn = document.getElementById("modalConfirmBtn");

  const imageUploadInput = document.getElementById("imageUploadInput");

  // A viewport/container (used for fit-to-screen). Falls back safely.
  const canvasViewport =
    document.getElementById("canvasViewport") ||
    document.getElementById("canvasHost") ||
    canvas?.parentElement;

  // -------------------------
  // ✅ HARD FIX: ensure discard modal is NOT visible on page load
  // -------------------------
  if (modalBackdrop) {
    modalBackdrop.hidden = true;
    modalBackdrop.style.display = "none";
    modalBackdrop.setAttribute("aria-hidden", "true");
  }
  if (uploadDashboardInput) uploadDashboardInput.value = "";

  // -------------------------
  // Canvas size (dynamic + presets)
  // -------------------------
  const CANVAS_PRESETS = [
    { key: "p1280", label: "16:9 (1280 × 720)", w: 1280, h: 720 },
    { key: "p1920", label: "16:9 (1920 × 1080)", w: 1920, h: 1080 }
  ];

  let canvasW = 1280;
  let canvasH = 720;
  let canvasPresetKey = "p1280"; // default 16:9

  function updateStatusPill(){
    if (statusPill) statusPill.textContent = `Canvas: ${canvasW} × ${canvasH}`;
  }

  function applyCanvasSize(){
    if (!canvas) return;

    canvas.style.width = `${canvasW}px`;
    canvas.style.height = `${canvasH}px`;

    // keep right side clamp perfect after resizing
    clampAllVisualsInsideCanvas();

    // update UI + fit-to-screen
    updateStatusPill();
    fitCanvasToScreen();
    renderFormatPane();
  }

  // Fit canvas to visible area like Power BI "fit to page"
  function fitCanvasToScreen(){
    if (!canvas || !canvasViewport) return;

    const rect = canvasViewport.getBoundingClientRect();
    const padding = 18;

    const availW = Math.max(100, rect.width - padding * 2);
    const availH = Math.max(100, rect.height - padding * 2);

    const sx = availW / canvasW;
    const sy = availH / canvasH;
    const scale = Math.min(1, sx, sy);

    // scale the canvas itself (children scale too)
    canvas.style.transformOrigin = "top left";
    canvas.style.transform = `scale(${scale})`;

    // reserve space in layout so it doesn't clip
    // (we set the parent min size if possible)
    if (canvasViewport) {
      canvasViewport.style.overflow = "hidden";
    }
  }

  window.addEventListener("resize", () => fitCanvasToScreen());

  // -------------------------
  // Dummy retail data (realistic growth Jan→Dec)
  // -------------------------
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  const mkGrowth = (start, end) => {
    const arr = [];
    for (let i=0; i<12; i++){
      const t = i / 11;
      const v = start + (end - start) * t + (Math.sin(i*0.9)*0.03*(end-start));
      arr.push(Math.round(v));
    }
    return arr;
  };

  const retail = {
    Sales: mkGrowth(120, 220),
    Profit: mkGrowth(18, 32),
    Orders: mkGrowth(800, 1450),
    Regions: {
      "North": mkGrowth(40, 78),
      "South": mkGrowth(36, 70),
      "East":  mkGrowth(28, 58),
      "West":  mkGrowth(32, 66),
    },
    Categories: [
      { name: "Electronics", value: 38 },
      { name: "Fashion", value: 26 },
      { name: "Grocery", value: 18 },
      { name: "Home", value: 12 },
      { name: "Other", value: 6 },
    ]
  };

  // -------------------------
  // Theme (Power BI-like)
  // -------------------------
  const defaultTheme = {
    name: "Default Dark Prototype",
    dataColors: ["#3b82f6","#22c55e","#f59e0b","#a855f7","#ef4444","#06b6d4","#eab308","#f97316","#84cc16","#14b8a6"],
    background: "#0f1115",
    foreground: "#e8ecf2",
    textClasses: {
      title: { fontFace: "Segoe UI Semibold", color: "#e8ecf2", fontSize: 12 },
      label: { fontFace: "Segoe UI", color: "#cfd6e1", fontSize: 10 }
    },
    chart: {
      plotBg: "rgba(255,255,255,0.00)",
      grid: "rgba(255,255,255,0.08)"
    }
  };

  const sampleTheme = {
    name: "Green Theme (Sample)",
    dataColors: ["#22c55e","#16a34a","#84cc16","#10b981","#06b6d4","#f59e0b","#ef4444","#a855f7"],
    background: "#0f1115",
    foreground: "#e8ecf2",
    textClasses: {
      title: { fontFace: "Segoe UI Semibold", color: "#e8ecf2", fontSize: 12 },
      label: { fontFace: "Segoe UI", color: "#d5ffe3", fontSize: 10 }
    },
    chart: {
      plotBg: "rgba(255,255,255,0.00)",
      grid: "rgba(34,197,94,0.14)"
    }
  };

  const sampleCanvasBg = {
    backgroundColor: "#0b0d12",
    backgroundImage: "",
    opacity: 1
  };

  let theme = structuredClone(defaultTheme);
  let canvasBg = structuredClone(sampleCanvasBg);

  function showToast(msg){
    if (!toastEl) return;
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toastEl.classList.remove("show"), 2200);
  }

  function downloadObjectAsJson(obj, filename){
    const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function readJsonFile(file){
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        try { resolve(JSON.parse(reader.result)); }
        catch(e){ reject(e); }
      };
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }

  function readImageAsDataUrl(file){
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = reject;
      r.readAsDataURL(file);
    });
  }

  function isDefaultState(){
    return state.order.length === 0 &&
      JSON.stringify(theme) === JSON.stringify(defaultTheme) &&
      JSON.stringify(canvasBg) === JSON.stringify(sampleCanvasBg) &&
      canvasW === 1280 && canvasH === 720;
  }

  function applyCanvasBackground(){
    document.documentElement.style.setProperty("--canvasBgColor", canvasBg.backgroundColor || "#0b0d12");
    const img = (canvasBg.backgroundImage && String(canvasBg.backgroundImage).trim())
      ? `url("${canvasBg.backgroundImage}")`
      : "none";
    document.documentElement.style.setProperty("--canvasBgImage", img);
    document.documentElement.style.setProperty("--canvasBgOpacity", String(
      clamp(Number(canvasBg.opacity ?? 1), 0, 1)
    ));
  }

  // Apply initial canvas bg + size
  applyCanvasBackground();
  updateStatusPill();
  applyCanvasSize();

  // -------------------------
  // Chart defaults
  // ✅ (1) remove background gridlines + remove axis border lines
  // -------------------------
  function makeChartDefaults(){
    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";

    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 300 },
      plugins: {
        legend: {
          display: true,
          labels: { color: labelColor, boxWidth: 10, boxHeight: 10, usePointStyle: true }
        },
        title: { display: false },
        tooltip: { enabled: true }
      },
      scales: {
        x: {
          ticks: { color: labelColor },
          grid: { display: false },     // ✅ no grid lines
          border: { display: false }    // ✅ no axis line
        },
        y: {
          ticks: { color: labelColor },
          grid: { display: false },     // ✅ no grid lines
          border: { display: false }    // ✅ no axis line
        }
      }
    };
  }

  function paletteColor(i){
    const arr = theme?.dataColors?.length ? theme.dataColors : defaultTheme.dataColors;
    return arr[i % arr.length];
  }

  // -------------------------
  // Visual picker registry
  // -------------------------
  const VISUALS = [
    {
      category: "Charts",
      items: [
        { type: "pie", name: "Pie Chart", icon: "◔" },
        { type: "donut", name: "Donut Chart", icon: "◕" },
        { type: "treemap", name: "Treemap", icon: "▧" },
        { type: "ribbon", name: "Ribbon Chart", icon: "〰" },

        { type: "line", name: "Line Chart (multiple trends)", icon: "╱" },
        { type: "area", name: "Area Chart", icon: "▱" },
        { type: "stackedArea", name: "Stacked Area Chart", icon: "▰" },

        { type: "clusteredBar", name: "Clustered Bar Chart", icon: "▤" },
        { type: "stackedBar", name: "Stacked Bar Chart", icon: "▥" },
        { type: "stackedBar100", name: "100% Stacked Bar Chart", icon: "▦" },

        { type: "clusteredColumn", name: "Clustered Column Chart", icon: "▮" },
        { type: "stackedColumn", name: "Stacked Column Chart", icon: "▯" },
        { type: "stackedColumn100", name: "100% Stacked Column Chart", icon: "▰" },

        { type: "lineClusteredColumn", name: "Line & Clustered Column Chart", icon: "⟂" },
        { type: "lineStackedColumn", name: "Line & Stacked Column Chart", icon: "⟋" },

        { type: "scatter", name: "Scatter Chart", icon: "∘" },
      ]
    },
    {
      category: "Cards",
      items: [
        { type: "kpi", name: "KPI", icon: "📈" },
        { type: "card", name: "Card", icon: "🧾" },
        { type: "multirowCard", name: "Multi-row Card", icon: "📋" },
        { type: "textBox", name: "Text Box", icon: "🅣" }
      ]
    },
    {
      category: "Other",
      items: [
        { type: "image", name: "Upload Image", icon: "🖼" }
      ]
    }
  ];

  const DEFAULT_SIZES = {
    pie: { w: 360, h: 270 },
    donut: { w: 360, h: 270 },
    treemap: { w: 420, h: 280 },
    ribbon: { w: 460, h: 280 },

    line: { w: 520, h: 300 },
    area: { w: 520, h: 300 },
    stackedArea: { w: 520, h: 300 },

    clusteredBar: { w: 520, h: 300 },
    stackedBar: { w: 520, h: 300 },
    stackedBar100: { w: 520, h: 300 },

    clusteredColumn: { w: 520, h: 300 },
    stackedColumn: { w: 520, h: 300 },
    stackedColumn100: { w: 520, h: 300 },

    lineClusteredColumn: { w: 560, h: 320 },
    lineStackedColumn: { w: 560, h: 320 },

    scatter: { w: 520, h: 300 },

    kpi: { w: 320, h: 180 },
    card: { w: 300, h: 150 },
    multirowCard: { w: 380, h: 220 },
    textBox: { w: 360, h: 200 },

    image: { w: 420, h: 260 }
  };

  // -------------------------
  // App state
  // -------------------------
  const state = {
    visuals: new Map(),
    order: [],
    selectedId: null,
    nextId: 1
  };

  // -------------------------
  // Visual picker render + open/close
  // -------------------------
  function renderVisualPicker(){
    if (!visualPickerMenu) return;
    visualPickerMenu.innerHTML = "";

    VISUALS.forEach(section => {
      const sec = document.createElement("div");
      sec.className = "menuSection";

      const head = document.createElement("div");
      head.className = "menuHeader";
      head.textContent = section.category;
      sec.appendChild(head);

      const grid = document.createElement("div");
      grid.className = "menuGrid";

      section.items.forEach(item => {
        const btn = document.createElement("div");
        btn.className = "menuItem";
        btn.setAttribute("role", "menuitem");
        btn.tabIndex = 0;

        btn.innerHTML = `
          <div class="ico" aria-hidden="true">${escapeHtml(item.icon)}</div>
          <div class="lbl">${escapeHtml(item.name)}</div>
        `;

        btn.addEventListener("click", () => {
          closePicker();
          if (item.type === "image") addImageVisualFlow();
          else addVisual(item.type);
        });

        btn.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            btn.click();
          }
        });

        grid.appendChild(btn);
      });

      sec.appendChild(grid);
      visualPickerMenu.appendChild(sec);
    });
  }

  function openPicker(){
    if (!visualPickerMenu || !visualPickerToggle) return;
    visualPickerMenu.classList.add("open");
    visualPickerToggle.setAttribute("aria-expanded", "true");
  }
  function closePicker(){
    if (!visualPickerMenu || !visualPickerToggle) return;
    visualPickerMenu.classList.remove("open");
    visualPickerToggle.setAttribute("aria-expanded", "false");
  }

  if (visualPickerToggle) {
    visualPickerToggle.addEventListener("click", (e) => {
      e.stopPropagation();
      if (visualPickerMenu.classList.contains("open")) closePicker();
      else openPicker();
    });
  }

  document.addEventListener("click", () => closePicker());
  renderVisualPicker();

  // -------------------------
  // Selection behavior
  // -------------------------
  if (canvas) {
    canvas.addEventListener("mousedown", (e) => {
      if (e.target === canvas || e.target === canvasWatermarkOrBefore(e.target)) {
        setSelected(null);
      }
    });
  }

  function canvasWatermarkOrBefore(target){
    const wm = document.getElementById("canvasWatermark");
    return target === wm;
  }

  function setSelected(id){
    state.selectedId = id;

    state.order.forEach(vid => {
      const el = document.getElementById(`v_${vid}`);
      if (!el) return;
      if (vid === id) el.classList.add("selected");
      else el.classList.remove("selected");
    });

    if (id && state.order.includes(id)) {
      state.order = state.order.filter(x => x !== id);
      state.order.push(id);
      updateZOrder();
    }

    renderFormatPane();
  }

  function updateZOrder(){
    state.order.forEach((vid, idx) => {
      const el = document.getElementById(`v_${vid}`);
      if (el) el.style.zIndex = String(10 + idx);
    });
  }

  // -------------------------
  // Add visuals
  // -------------------------
  function addVisual(type){
    const id = `vis${state.nextId++}`;
    const size = DEFAULT_SIZES[type] || { w: 520, h: 300 };

    const offset = (state.order.length * 18) % 140;
    const x = clamp(40 + offset, 0, canvasW - size.w);
    const y = clamp(40 + offset, 0, canvasH - size.h);

    const v = {
      id,
      type,
      title: defaultTitleFor(type),
      x, y, w: size.w, h: size.h,
      series: buildDefaultSeries(type),
      chart: null,

      // Card/KPI/Text extras
      card: buildDefaultCard(type),
      text: buildDefaultText(type)
    };

    state.visuals.set(id, v);
    state.order.push(id);

    const el = createVisualElement(v);
    canvas.appendChild(el);
    updateZOrder();

    renderVisualContent(v);
    setSelected(id);
  }

  async function addImageVisualFlow(){
    if (!imageUploadInput) return;
    imageUploadInput.value = "";
    imageUploadInput.click();

    imageUploadInput.onchange = async () => {
      const f = imageUploadInput.files?.[0];
      if (!f) return;

      const dataUrl = await readImageAsDataUrl(f);

      const id = `vis${state.nextId++}`;
      const size = DEFAULT_SIZES.image;
      const offset = (state.order.length * 18) % 140;
      const x = clamp(40 + offset, 0, canvasW - size.w);
      const y = clamp(40 + offset, 0, canvasH - size.h);

      const v = {
        id,
        type: "image",
        title: "Image",
        x, y, w: size.w, h: size.h,
        series: [],
        imageDataUrl: dataUrl,
        chart: null,
        card: null,
        text: null
      };

      state.visuals.set(id, v);
      state.order.push(id);

      const el = createVisualElement(v);
      canvas.appendChild(el);
      updateZOrder();

      renderVisualContent(v);
      setSelected(id);
    };
  }

  function defaultTitleFor(type){
    const map = {
      pie: "Sales by Category",
      donut: "Profit Split (Category)",
      treemap: "Category Treemap",
      ribbon: "Ribbon (Share Trend)",
      line: "Monthly Sales (Trends)",
      area: "Monthly Profit (Area)",
      stackedArea: "Regional Sales (Stacked Area)",
      clusteredBar: "Sales by Region (Bar)",
      stackedBar: "Sales by Region (Stacked Bar)",
      stackedBar100: "Sales Share by Region (100% Bar)",
      clusteredColumn: "Orders by Region (Column)",
      stackedColumn: "Orders by Region (Stacked Column)",
      stackedColumn100: "Orders Share by Region (100% Column)",
      lineClusteredColumn: "Sales + Profit (Combo)",
      lineStackedColumn: "Sales + Regional Mix (Combo)",
      scatter: "Profit vs Sales (Scatter)",

      kpi: "KPI",
      card: "Card",
      multirowCard: "Multi-row card",
      textBox: "Text box"
    };
    return map[type] || "Visual";
  }

  function buildDefaultSeries(type){
    if (type === "pie" || type === "donut") {
      return retail.Categories.map((c, i) => ({
        key: c.name,
        name: c.name,
        color: paletteColor(i),
        overrideColor: false
      }));
    }

    const regionKeys = Object.keys(retail.Regions);
    const singleSeries = [{ key: "Value", name: "Value", color: paletteColor(0), overrideColor: false }];
    const multiSeries = regionKeys.map((k, i) => ({ key: k, name: k, color: paletteColor(i), overrideColor: false }));

    switch(type){
      case "line":
      case "stackedArea":
      case "stackedBar":
      case "stackedBar100":
      case "stackedColumn":
      case "stackedColumn100":
      case "lineStackedColumn":
        return multiSeries;

      default:
        return singleSeries;
    }
  }

  function buildDefaultCard(type){
    if (type === "kpi") {
      return {
        label: "Profit YTD",
        value: "32.4K",
        trend: "+18.6%",
        trendUp: true
      };
    }
    if (type === "card") {
      return {
        label: "Total Sales",
        value: "220K"
      };
    }
    if (type === "multirowCard") {
      return {
        rows: [
          { k: "Sales", v: "220K" },
          { k: "Profit", v: "32K" },
          { k: "Orders", v: "1,450" },
          { k: "YoY", v: "+80%" }
        ]
      };
    }
    return null;
  }

  function buildDefaultText(type){
    if (type !== "textBox") return null;
    return {
      html: "Type your text here…",
      fontSize: 16,
      color: "#e8ecf2",
      bold: false,
      italic: false,
      align: "left",
      bg: "rgba(255,255,255,0.02)"
    };
  }

  // -------------------------
  // Visual DOM
  // -------------------------
  function createVisualElement(v){
    const el = document.createElement("div");
    el.className = "visual";
    el.id = `v_${v.id}`;
    el.style.left = `${v.x}px`;
    el.style.top  = `${v.y}px`;
    el.style.width = `${v.w}px`;
    el.style.height = `${v.h}px`;

    el.innerHTML = `
      <div class="vHeader" data-drag="1">
        <div class="vTitle" id="title_${v.id}">${escapeHtml(v.title)}</div>
        <div class="vHeaderSpacer"></div>
        <button class="vDelete" title="Delete visual" aria-label="Delete visual">✖</button>
      </div>
      <div class="vBody" id="body_${v.id}"></div>

      <div class="handle nw" data-handle="nw"></div>
      <div class="handle ne" data-handle="ne"></div>
      <div class="handle sw" data-handle="sw"></div>
      <div class="handle se" data-handle="se"></div>
      <div class="handle n"  data-handle="n"></div>
      <div class="handle s"  data-handle="s"></div>
      <div class="handle w"  data-handle="w"></div>
      <div class="handle e"  data-handle="e"></div>
    `;

    el.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      setSelected(v.id);
    });

    el.querySelector(".vDelete").addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(v.id);
    });

    const header = el.querySelector(".vHeader");
    header.addEventListener("mousedown", (e) => startDrag(e, v.id));

    el.querySelectorAll(".handle").forEach(h => {
      h.addEventListener("mousedown", (e) => startResize(e, v.id, h.dataset.handle));
    });

    return el;
  }

  function removeVisual(id){
    const v = state.visuals.get(id);
    if (!v) return;

    if (v.chart) {
      try { v.chart.destroy(); } catch {}
      v.chart = null;
    }

    state.visuals.delete(id);
    state.order = state.order.filter(x => x !== id);

    const el = document.getElementById(`v_${id}`);
    if (el) el.remove();

    if (state.selectedId === id) setSelected(null);
    updateZOrder();
  }

  // -------------------------
  // Drag & Resize (clamped on ALL sides)
  // -------------------------
  let dragCtx = null;

  function startDrag(e, id){
    if (e.button !== 0) return;

    const v = state.visuals.get(id);
    if (!v) return;

    setSelected(id);

    dragCtx = {
      mode: "drag",
      id,
      startX: e.clientX,
      startY: e.clientY,
      origX: v.x,
      origY: v.y
    };

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp, { once:true });
  }

  function startResize(e, id, handle){
    if (e.button !== 0) return;

    const v = state.visuals.get(id);
    if (!v) return;

    setSelected(id);

    dragCtx = {
      mode: "resize",
      id,
      handle,
      startX: e.clientX,
      startY: e.clientY,
      origX: v.x, origY: v.y,
      origW: v.w, origH: v.h
    };

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp, { once:true });
  }

  function onMove(e){
    if (!dragCtx) return;

    const v = state.visuals.get(dragCtx.id);
    if (!v) return;

    const dx = e.clientX - dragCtx.startX;
    const dy = e.clientY - dragCtx.startY;

    if (dragCtx.mode === "drag") {
      v.x = clamp(dragCtx.origX + dx, 0, canvasW - v.w);
      v.y = clamp(dragCtx.origY + dy, 0, canvasH - v.h);
      applyVisualRect(v);
    } else {
      const minW = 220;
      const minH = 140;

      let x = dragCtx.origX;
      let y = dragCtx.origY;
      let w = dragCtx.origW;
      let h = dragCtx.origH;

      const hnd = dragCtx.handle;

      const applyW = (newW) => clamp(newW, minW, canvasW - x);
      const applyH = (newH) => clamp(newH, minH, canvasH - y);

      if (hnd.includes("e")) w = applyW(dragCtx.origW + dx);
      if (hnd.includes("s")) h = applyH(dragCtx.origH + dy);

      if (hnd.includes("w")) {
        const newX = clamp(dragCtx.origX + dx, 0, dragCtx.origX + dragCtx.origW - minW);
        const newW = dragCtx.origW + (dragCtx.origX - newX);
        x = newX;
        w = clamp(newW, minW, canvasW - x);
      }

      if (hnd.includes("n")) {
        const newY = clamp(dragCtx.origY + dy, 0, dragCtx.origY + dragCtx.origH - minH);
        const newH = dragCtx.origH + (dragCtx.origY - newY);
        y = newY;
        h = clamp(newH, minH, canvasH - y);
      }

      v.x = x; v.y = y; v.w = w; v.h = h;
      applyVisualRect(v);
      if (v.chart) v.chart.resize();
    }
  }

  function onUp(){
    window.removeEventListener("mousemove", onMove);
    dragCtx = null;
    renderFormatPane();
  }

  function applyVisualRect(v){
    const el = document.getElementById(`v_${v.id}`);
    if (!el) return;
    el.style.left = `${v.x}px`;
    el.style.top  = `${v.y}px`;
    el.style.width = `${v.w}px`;
    el.style.height = `${v.h}px`;
  }

  function clampAllVisualsInsideCanvas(){
    state.order.forEach(id => {
      const v = state.visuals.get(id);
      if (!v) return;
      v.x = clamp(v.x, 0, Math.max(0, canvasW - v.w));
      v.y = clamp(v.y, 0, Math.max(0, canvasH - v.h));

      // Also clamp size if canvas becomes smaller
      v.w = clamp(v.w, 220, Math.max(220, canvasW - v.x));
      v.h = clamp(v.h, 140, Math.max(140, canvasH - v.y));

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
    });
  }

  // -------------------------
  // Render visual content
  // -------------------------
  function renderVisualContent(v){
    const body = document.getElementById(`body_${v.id}`);
    if (!body) return;

    if (v.chart) {
      try { v.chart.destroy(); } catch {}
      v.chart = null;
    }
    body.innerHTML = "";

    // Image
    if (v.type === "image") {
      const host = document.createElement("div");
      host.className = "imageHost";
      const img = document.createElement("img");
      img.alt = "Uploaded visual image";
      img.src = v.imageDataUrl || "";
      host.appendChild(img);
      body.appendChild(host);
      return;
    }

    // KPI
    if (v.type === "kpi") {
      body.appendChild(makeKpiPrototype(v));
      return;
    }

    // Card
    if (v.type === "card") {
      body.appendChild(makeCardPrototype(v));
      return;
    }

    // Multi-row card
    if (v.type === "multirowCard") {
      body.appendChild(makeMultirowCardPrototype(v));
      return;
    }

    // Text box
    if (v.type === "textBox") {
      body.appendChild(makeTextBoxPrototype(v));
      return;
    }

    // Treemap
    if (v.type === "treemap") {
      body.appendChild(makeTreemapPrototype(v));
      return;
    }

    // Ribbon
    if (v.type === "ribbon") {
      body.appendChild(makeRibbonPrototype(v));
      return;
    }

    // Chart.js visuals
    const host = document.createElement("div");
    host.className = "chartHost";
    const c = document.createElement("canvas");
    c.id = `c_${v.id}`;
    host.appendChild(c);
    body.appendChild(host);

    const ctx = c.getContext("2d");
    const ds = buildChartJsData(v);

    v.chart = new Chart(ctx, ds.config);
    applyThemeToSingleVisual(v);
  }

  function makeKpiPrototype(v){
    const wrap = document.createElement("div");
    wrap.style.width = "100%";
    wrap.style.height = "100%";
    wrap.style.display = "flex";
    wrap.style.flexDirection = "column";
    wrap.style.justifyContent = "center";
    wrap.style.padding = "14px";

    const label = document.createElement("div");
    label.style.opacity = "0.85";
    label.style.fontSize = "12px";
    label.textContent = v.card?.label || "KPI";

    const value = document.createElement("div");
    value.style.fontSize = "34px";
    value.style.fontWeight = "800";
    value.style.marginTop = "6px";
    value.textContent = v.card?.value || "0";

    const trend = document.createElement("div");
    trend.style.marginTop = "8px";
    trend.style.fontSize = "13px";
    trend.style.display = "flex";
    trend.style.alignItems = "center";
    trend.style.gap = "8px";

    const up = !!v.card?.trendUp;
    const arrow = document.createElement("span");
    arrow.textContent = up ? "▲" : "▼";
    arrow.style.fontWeight = "900";
    arrow.style.opacity = "0.9";

    const tval = document.createElement("span");
    tval.textContent = v.card?.trend || "+0%";

    trend.appendChild(arrow);
    trend.appendChild(tval);

    wrap.appendChild(label);
    wrap.appendChild(value);
    wrap.appendChild(trend);
    return wrap;
  }

  function makeCardPrototype(v){
    const wrap = document.createElement("div");
    wrap.style.width = "100%";
    wrap.style.height = "100%";
    wrap.style.display = "flex";
    wrap.style.flexDirection = "column";
    wrap.style.justifyContent = "center";
    wrap.style.padding = "14px";

    const label = document.createElement("div");
    label.style.opacity = "0.85";
    label.style.fontSize = "12px";
    label.textContent = v.card?.label || "Card";

    const value = document.createElement("div");
    value.style.fontSize = "32px";
    value.style.fontWeight = "800";
    value.style.marginTop = "8px";
    value.textContent = v.card?.value || "0";

    wrap.appendChild(label);
    wrap.appendChild(value);
    return wrap;
  }

  function makeMultirowCardPrototype(v){
    const wrap = document.createElement("div");
    wrap.style.width = "100%";
    wrap.style.height = "100%";
    wrap.style.display = "flex";
    wrap.style.flexDirection = "column";
    wrap.style.padding = "14px";
    wrap.style.gap = "10px";

    const rows = Array.isArray(v.card?.rows) ? v.card.rows : [];

    rows.forEach(r => {
      const line = document.createElement("div");
      line.style.display = "flex";
      line.style.justifyContent = "space-between";
      line.style.gap = "12px";

      const k = document.createElement("div");
      k.style.opacity = "0.85";
      k.style.fontSize = "12px";
      k.textContent = r.k;

      const val = document.createElement("div");
      val.style.fontWeight = "800";
      val.style.fontSize = "13px";
      val.textContent = r.v;

      line.appendChild(k);
      line.appendChild(val);
      wrap.appendChild(line);
    });

    return wrap;
  }

  function makeTextBoxPrototype(v){
    const wrap = document.createElement("div");
    wrap.style.width = "100%";
    wrap.style.height = "100%";
    wrap.style.padding = "12px";
    wrap.style.overflow = "hidden";

    const t = v.text || buildDefaultText("textBox");

    const box = document.createElement("div");
    box.className = "textBoxHost";
    box.contentEditable = "true";
    box.spellcheck = false;

    box.style.width = "100%";
    box.style.height = "100%";
    box.style.outline = "none";
    box.style.whiteSpace = "pre-wrap";
    box.style.wordBreak = "break-word";
    box.style.fontSize = `${t.fontSize || 16}px`;
    box.style.color = t.color || "#e8ecf2";
    box.style.fontWeight = t.bold ? "800" : "400";
    box.style.fontStyle = t.italic ? "italic" : "normal";
    box.style.textAlign = t.align || "left";
    box.style.background = t.bg || "rgba(255,255,255,0.02)";
    box.style.borderRadius = "10px";
    box.style.padding = "10px";

    box.innerHTML = t.html || "Type your text here…";

    box.addEventListener("input", () => {
      v.text = v.text || buildDefaultText("textBox");
      v.text.html = box.innerHTML;
    });

    wrap.appendChild(box);
    return wrap;
  }

  function buildChartJsData(v){
    const cfgBase = {
      type: "bar",
      data: { labels: months, datasets: [] },
      options: makeChartDefaults()
    };

    const regionSeries = v.series.map((s) => retail.Regions[s.key] || mkGrowth(20, 60));
    const single = mkGrowth(30, 55);

    const makeLineDs = (s, values, fill=false) => ({
      label: s.name,
      data: values,
      borderColor: s.color,
      backgroundColor: withAlpha(s.color, fill ? 0.22 : 0.15),
      tension: 0.35,
      pointRadius: 2,
      pointHoverRadius: 3,
      fill: fill ? "origin" : false,
      borderWidth: 2
    });

    const makeBarDs = (s, values) => ({
      label: s.name,
      data: values,
      backgroundColor: withAlpha(s.color, 0.75),
      borderColor: withAlpha(s.color, 0.95),
      borderWidth: 1.2,
      borderRadius: 6,
      barPercentage: 0.75,
      categoryPercentage: 0.70
    });

    switch(v.type){
      case "pie":
      case "donut": {
        const labels = v.series.map(s => s.name);
        const values = retail.Categories.map(c => c.value);
        const colors = v.series.map(s => s.color);

        return {
          config: {
            type: "pie",
            data: {
              labels,
              datasets: [{
                label: "Value",
                data: values,
                backgroundColor: colors.map(c => withAlpha(c, 0.82)),
                borderColor: colors.map(c => withAlpha(c, 1)),
                borderWidth: 1
              }]
            },
            options: {
              ...makeChartDefaults(),
              cutout: v.type === "donut" ? "58%" : "0%"
            }
          }
        };
      }

      case "line": {
        cfgBase.type = "line";
        cfgBase.data.datasets = v.series.map((s, i) => makeLineDs(s, regionSeries[i], false));
        return { config: cfgBase };
      }

      case "area": {
        cfgBase.type = "line";
        const s = v.series[0] || { name: "Profit", color: paletteColor(0) };
        cfgBase.data.datasets = [ makeLineDs({ ...s, name: v.series[0]?.name || "Profit" }, retail.Profit, true) ];
        return { config: cfgBase };
      }

      case "stackedArea": {
        cfgBase.type = "line";
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;
        cfgBase.data.datasets = v.series.map((s, i) => ({ ...makeLineDs(s, regionSeries[i], true), fill: "origin" }));
        return { config: cfgBase };
      }

      case "clusteredBar": {
        cfgBase.type = "bar";
        cfgBase.options.indexAxis = "y";
        cfgBase.data.labels = ["North","South","East","West"];
        cfgBase.data.datasets = [{
          label: v.series[0]?.name || "Sales",
          data: [78, 70, 58, 66],
          backgroundColor: withAlpha(v.series[0]?.color || paletteColor(0), 0.75),
          borderRadius: 7
        }];
        return { config: cfgBase };
      }

      case "stackedBar":
      case "stackedBar100": {
        cfgBase.type = "bar";
        cfgBase.options.indexAxis = "y";
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;

        const datasets = v.series.map((s, i) => makeBarDs(s, regionSeries[i]));
        cfgBase.data.datasets = datasets;

        if (v.type === "stackedBar100") {
          cfgBase.data = toPercentStacked(months, datasets);
          cfgBase.options.plugins.tooltip.callbacks = {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.formattedValue}%`
          };
        }
        return { config: cfgBase };
      }

      case "clusteredColumn": {
        cfgBase.type = "bar";
        cfgBase.data.labels = ["North","South","East","West"];
        cfgBase.data.datasets = [{
          label: v.series[0]?.name || "Orders",
          data: [420, 405, 350, 390],
          backgroundColor: withAlpha(v.series[0]?.color || paletteColor(0), 0.75),
          borderRadius: 7
        }];
        return { config: cfgBase };
      }

      case "stackedColumn":
      case "stackedColumn100": {
        cfgBase.type = "bar";
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;

        const datasets = v.series.map((s, i) => makeBarDs(s, regionSeries[i]));
        cfgBase.data.datasets = datasets;

        if (v.type === "stackedColumn100") {
          cfgBase.data = toPercentStacked(months, datasets);
          cfgBase.options.plugins.tooltip.callbacks = {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.formattedValue}%`
          };
        }
        return { config: cfgBase };
      }

      case "lineClusteredColumn": {
        const barColor = v.series[0]?.color || paletteColor(0);
        const lineColor = v.series[1]?.color || paletteColor(1);

        return {
          config: {
            type: "bar",
            data: {
              labels: months,
              datasets: [
                {
                  type: "bar",
                  label: v.series[0]?.name || "Sales",
                  data: retail.Sales,
                  backgroundColor: withAlpha(barColor, 0.70),
                  borderColor: withAlpha(barColor, 0.95),
                  borderWidth: 1,
                  borderRadius: 6
                },
                {
                  type: "line",
                  label: v.series[1]?.name || "Profit",
                  data: retail.Profit,
                  borderColor: withAlpha(lineColor, 1),
                  backgroundColor: withAlpha(lineColor, 0.18),
                  pointRadius: 2,
                  tension: 0.35,
                  yAxisID: "y"
                }
              ]
            },
            options: makeChartDefaults()
          }
        };
      }

      case "lineStackedColumn": {
        const cfg = {
          type: "bar",
          data: {
            labels: months,
            datasets: [
              ...v.series.map((s, i) => ({
                type: "bar",
                label: s.name,
                data: regionSeries[i],
                backgroundColor: withAlpha(s.color, 0.68),
                borderColor: withAlpha(s.color, 0.92),
                borderWidth: 1,
                borderRadius: 5,
                stack: "mix"
              })),
              {
                type: "line",
                label: "Profit",
                data: retail.Profit,
                borderColor: withAlpha(paletteColor(6), 1),
                backgroundColor: withAlpha(paletteColor(6), 0.18),
                pointRadius: 2,
                tension: 0.35,
                yAxisID: "y"
              }
            ]
          },
          options: (() => {
            const o = makeChartDefaults();
            o.scales.x.stacked = true;
            o.scales.y.stacked = true;
            return o;
          })()
        };
        return { config: cfg };
      }

      case "scatter": {
        const s = v.series[0] || { name: "Points", color: paletteColor(0) };
        const points = months.map((m, i) => ({ x: retail.Sales[i], y: retail.Profit[i] }));

        return {
          config: {
            type: "scatter",
            data: {
              datasets: [{
                label: s.name,
                data: points,
                backgroundColor: withAlpha(s.color, 0.75),
                borderColor: withAlpha(s.color, 1),
                pointRadius: 4
              }]
            },
            options: makeChartDefaults()
          }
        };
      }

      default: {
        cfgBase.type = "line";
        cfgBase.data.datasets = [ makeLineDs({ name:"Value", color: paletteColor(0) }, single, true) ];
        return { config: cfgBase };
      }
    }
  }

  // -------------------------
  // Prototypes
  // -------------------------
  function makeTreemapPrototype(v){
    const wrap = document.createElement("div");
    wrap.style.width = "100%";
    wrap.style.height = "100%";
    wrap.style.display = "grid";
    wrap.style.gridTemplateColumns = "1.3fr 1fr";
    wrap.style.gridTemplateRows = "1fr 1fr";
    wrap.style.gap = "8px";

    const cats = v.series.length ? v.series : buildDefaultSeries("pie");
    const boxes = [
      { label: cats[0]?.name || "Electronics", color: cats[0]?.color || paletteColor(0) },
      { label: cats[1]?.name || "Fashion", color: cats[1]?.color || paletteColor(1) },
      { label: cats[2]?.name || "Grocery", color: cats[2]?.color || paletteColor(2) },
    ];

    const big = document.createElement("div");
    big.style.gridRow = "1 / span 2";
    big.style.borderRadius = "10px";
    big.style.border = "1px solid rgba(255,255,255,0.10)";
    big.style.background = withAlpha(boxes[0].color, 0.28);
    big.style.padding = "10px";
    big.innerHTML = `<div style="font-weight:700;font-size:12px;">${escapeHtml(boxes[0].label)}</div>
                     <div style="opacity:.75;font-size:11px;margin-top:4px;">Largest share</div>`;
    wrap.appendChild(big);

    for (let i=1; i<boxes.length; i++){
      const b = document.createElement("div");
      b.style.borderRadius = "10px";
      b.style.border = "1px solid rgba(255,255,255,0.10)";
      b.style.background = withAlpha(boxes[i].color, 0.24);
      b.style.padding = "10px";
      b.innerHTML = `<div style="font-weight:700;font-size:12px;">${escapeHtml(boxes[i].label)}</div>
                     <div style="opacity:.75;font-size:11px;margin-top:4px;">Category block</div>`;
      wrap.appendChild(b);
    }

    return wrap;
  }

  function makeRibbonPrototype(v){
    const host = document.createElement("div");
    host.style.width = "100%";
    host.style.height = "100%";
    host.style.border = "1px solid rgba(255,255,255,0.10)";
    host.style.borderRadius = "12px";
    host.style.overflow = "hidden";
    host.style.background = "rgba(255,255,255,0.01)";

    const series = (v.series && v.series.length > 1)
      ? v.series
      : Object.keys(retail.Regions).map((k,i)=>({ key:k, name:k, color: paletteColor(i), overrideColor:false }));

    const svg = document.createElementNS("http://www.w3.org/2000/svg","svg");
    svg.setAttribute("viewBox","0 0 600 260");
    svg.setAttribute("preserveAspectRatio","none");
    svg.style.width = "100%";
    svg.style.height = "100%";

    // ✅ (1) remove ribbon background grid lines entirely (no extra lines)
    const bands = [
      { y0: 60,  y1: 90,  wobble: 18 },
      { y0: 100, y1: 135, wobble: 22 },
      { y0: 145, y1: 175, wobble: 14 },
      { y0: 182, y1: 210, wobble: 20 },
    ];

    series.slice(0,4).forEach((s, i) => {
      const b = bands[i] || bands[bands.length-1];
      const c = withAlpha(s.color, 0.35);

      const path = document.createElementNS("http://www.w3.org/2000/svg","path");
      const w = b.wobble;

      const dTop = `M0 ${b.y0}
        C150 ${b.y0-w}, 300 ${b.y0+w}, 450 ${b.y0-w}
        S600 ${b.y0+w}, 600 ${b.y0}`;
      const dBot = `L600 ${b.y1}
        C450 ${b.y1+w}, 300 ${b.y1-w}, 150 ${b.y1+w}
        S0 ${b.y1-w}, 0 ${b.y1} Z`;

      path.setAttribute("d", dTop + " " + dBot);
      path.setAttribute("fill", c);
      path.setAttribute("stroke", withAlpha(s.color, 0.7));
      path.setAttribute("stroke-width", "1.2");
      svg.appendChild(path);

      const t = document.createElementNS("http://www.w3.org/2000/svg","text");
      t.setAttribute("x","14");
      t.setAttribute("y", String(b.y0 + 18));
      t.setAttribute("fill", theme?.textClasses?.label?.color || "rgba(232,236,242,0.85)");
      t.setAttribute("font-size","12");
      t.setAttribute("font-family","Segoe UI, Arial");
      t.textContent = s.name;
      svg.appendChild(t);
    });

    host.appendChild(svg);
    return host;
  }

  function withAlpha(hex, a){
    if (!hex) return `rgba(255,255,255,${a})`;
    if (hex.startsWith("rgba")) return hex;
    if (hex.startsWith("rgb(")) return hex.replace("rgb(", "rgba(").replace(")", `,${a})`);

    const h = hex.replace("#","").trim();
    const full = h.length === 3 ? h.split("").map(ch => ch+ch).join("") : h.padEnd(6, "0").slice(0,6);
    const r = parseInt(full.slice(0,2), 16);
    const g = parseInt(full.slice(2,4), 16);
    const b = parseInt(full.slice(4,6), 16);
    return `rgba(${r},${g},${b},${a})`;
  }

  function toPercentStacked(labels, datasets){
    const sums = labels.map((_, i) => datasets.reduce((acc, d) => acc + (Number(d.data[i]) || 0), 0));
    const newDatasets = datasets.map(d => ({
      ...d,
      data: d.data.map((v, i) => {
        const s = sums[i] || 1;
        return Math.round((Number(v) / s) * 100);
      })
    }));
    return { labels, datasets: newDatasets };
  }

  // -------------------------
  // Theme apply
  // -------------------------
  function applyThemeToSingleVisual(v){
    if (["treemap","ribbon","kpi","card","multirowCard","textBox","image"].includes(v.type)) {
      renderVisualContent(v);
      return;
    }
    if (!v.chart) return;

    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";

    if (v.chart.options?.plugins?.legend?.labels) {
      v.chart.options.plugins.legend.labels.color = labelColor;
    }

    // keep "no grid lines" on theme updates too
    if (v.chart.options?.scales?.x?.grid) v.chart.options.scales.x.grid.display = false;
    if (v.chart.options?.scales?.y?.grid) v.chart.options.scales.y.grid.display = false;
    if (v.chart.options?.scales?.x?.border) v.chart.options.scales.x.border.display = false;
    if (v.chart.options?.scales?.y?.border) v.chart.options.scales.y.border.display = false;

    syncSeriesToChart(v);
    v.chart.update();
  }

  function syncSeriesToChart(v){
    if (!v.chart) return;

    if (v.type === "pie" || v.type === "donut") {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.backgroundColor = v.series.map(s => withAlpha(s.color, 0.82));
      ds.borderColor = v.series.map(s => withAlpha(s.color, 1));
      v.chart.data.labels = v.series.map(s => s.name);
      return;
    }

    if (v.type === "scatter") {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.label = v.series[0]?.name || ds.label;
      ds.backgroundColor = withAlpha(v.series[0]?.color || paletteColor(0), 0.75);
      ds.borderColor = withAlpha(v.series[0]?.color || paletteColor(0), 1);
      return;
    }

    if (v.type === "line" || v.type === "area" || v.type === "stackedArea") {
      const dsets = v.chart.data.datasets || [];
      dsets.forEach((ds, i) => {
        const s = v.series[i] || v.series[0];
        if (!s) return;
        ds.label = s.name;
        ds.borderColor = withAlpha(s.color, 1);
        ds.backgroundColor = withAlpha(s.color, (v.type === "line" ? 0.12 : 0.22));
      });
      return;
    }

    if (["stackedBar","stackedBar100","stackedColumn","stackedColumn100"].includes(v.type)) {
      const dsets = v.chart.data.datasets || [];
      dsets.forEach((ds, i) => {
        const s = v.series[i];
        if (!s) return;
        ds.label = s.name;
        ds.backgroundColor = withAlpha(s.color, 0.70);
        ds.borderColor = withAlpha(s.color, 0.95);
      });
      return;
    }

    if (["clusteredBar","clusteredColumn"].includes(v.type)) {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.label = v.series[0]?.name || ds.label;
      ds.backgroundColor = withAlpha(v.series[0]?.color || paletteColor(0), 0.75);
      return;
    }

    if (v.type === "lineClusteredColumn") {
      const dsets = v.chart.data.datasets || [];
      if (dsets[0]) {
        const s0 = v.series[0] || { name:"Sales", color: paletteColor(0) };
        dsets[0].label = s0.name;
        dsets[0].backgroundColor = withAlpha(s0.color, 0.70);
        dsets[0].borderColor = withAlpha(s0.color, 0.95);
      }
      if (dsets[1]) {
        const s1 = v.series[1] || { name:"Profit", color: paletteColor(1) };
        dsets[1].label = s1.name;
        dsets[1].borderColor = withAlpha(s1.color, 1);
        dsets[1].backgroundColor = withAlpha(s1.color, 0.18);
      }
      return;
    }

    if (v.type === "lineStackedColumn") {
      const dsets = v.chart.data.datasets || [];
      const n = v.series.length;
      for (let i=0; i<n; i++){
        if (!dsets[i]) continue;
        const s = v.series[i];
        dsets[i].label = s.name;
        dsets[i].backgroundColor = withAlpha(s.color, 0.68);
        dsets[i].borderColor = withAlpha(s.color, 0.92);
      }
      return;
    }
  }

  function applyThemeEverywhere(){
    state.order.forEach((id) => {
      const v = state.visuals.get(id);
      if (!v) return;

      v.series?.forEach((s, i) => {
        if (!s.overrideColor) s.color = paletteColor(i);
      });

      applyThemeToSingleVisual(v);
    });

    renderFormatPane();
  }

  // -------------------------
  // Format pane (always visible)
  // -------------------------
  function renderFormatPane(){
    const selected = state.selectedId ? state.visuals.get(state.selectedId) : null;

    // Canvas settings
    if (!selected) {
      const presetOptions = CANVAS_PRESETS.map(p => `
        <option value="${escapeAttr(p.key)}"${p.key === canvasPresetKey ? " selected" : ""}>
          ${escapeHtml(p.label)}
        </option>
      `).join("");

      formatBody.innerHTML = `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Canvas settings</div>
            <div class="fSmall">No visual selected</div>
          </div>

          <div class="row one">
            <div class="field">
              <div class="label">Type</div>
              <select class="input" id="canvasPreset">
                ${presetOptions}
              </select>
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Width (px)</div>
              <input class="input" id="canvasW" value="${escapeAttr(String(canvasW))}" />
            </div>
            <div class="field">
              <div class="label">Height (px)</div>
              <input class="input" id="canvasH" value="${escapeAttr(String(canvasH))}" />
            </div>
          </div>

          <div class="smallBtnRow">
            <button class="btn" id="applyCanvasSizeBtn">Apply size</button>
            <button class="btn" id="resetCanvasSizeBtn">Reset</button>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Background color</div>
              <input class="input" id="canvasBgColor" value="${escapeAttr(canvasBg.backgroundColor || "#0b0d12")}" />
            </div>
            <div class="field">
              <div class="label">Opacity (0–1)</div>
              <input class="input" id="canvasBgOpacity" value="${escapeAttr(String(canvasBg.opacity ?? 1))}" />
            </div>
          </div>

          <div class="row one">
            <div class="field">
              <div class="label">Background image URL / data URL (optional)</div>
              <input class="input" id="canvasBgImage" value="${escapeAttr(canvasBg.backgroundImage || "")}" placeholder="https://... or data:image/..." />
            </div>
          </div>

          <div class="smallBtnRow">
            <button class="btn" id="applyCanvasBgBtn">Apply background</button>
            <button class="btn" id="browseCanvasBgBtn">Browse (Upload Canvas BG JSON)</button>
            <input id="canvasBgUploadInput" type="file" accept=".json,application/json" hidden />
          </div>

          <div class="fSmall" style="margin-top:8px;">
            Upload format supported: <code>{ "backgroundColor": "...", "backgroundImage": "...", "opacity": 0.9 }</code>
          </div>
        </div>

        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Theme</div>
            <div class="fSmall">${escapeHtml(theme?.name || "Theme")}</div>
          </div>

          <div class="smallBtnRow">
            <button class="btn" id="importThemeBtn2">Import Theme JSON</button>
            <button class="btn" id="resetThemeBtn">Reset Theme</button>
          </div>

          <div class="fSmall" style="margin-top:8px;">
            Theme applies instantly to existing + new visuals.
          </div>
        </div>
      `;

      // Preset
      const presetSel = document.getElementById("canvasPreset");
      presetSel.onchange = () => {
        const key = presetSel.value;
        const p = CANVAS_PRESETS.find(x => x.key === key);
        if (!p) return;
        canvasPresetKey = key;
        canvasW = p.w;
        canvasH = p.h;
        applyCanvasSize();
        showToast("Canvas size updated");
      };

      // Apply custom
      document.getElementById("applyCanvasSizeBtn").onclick = () => {
        const w = Number(document.getElementById("canvasW").value);
        const h = Number(document.getElementById("canvasH").value);

        if (isFinite(w) && w >= 600) canvasW = Math.round(w);
        if (isFinite(h) && h >= 400) canvasH = Math.round(h);

        // If user customizes, keep preset as closest (or unchanged)
        // (simple rule: if matches preset exactly, set it)
        const match = CANVAS_PRESETS.find(p => p.w === canvasW && p.h === canvasH);
        if (match) canvasPresetKey = match.key;

        applyCanvasSize();
        showToast("Canvas size applied");
      };

      document.getElementById("resetCanvasSizeBtn").onclick = () => {
        canvasPresetKey = "p1280";
        canvasW = 1280; canvasH = 720;
        applyCanvasSize();
        showToast("Canvas reset");
      };

      // Background apply
      document.getElementById("applyCanvasBgBtn").onclick = () => {
        const c = document.getElementById("canvasBgColor").value.trim();
        const o = Number(document.getElementById("canvasBgOpacity").value.trim());
        const img = document.getElementById("canvasBgImage").value.trim();

        canvasBg.backgroundColor = c || canvasBg.backgroundColor;
        canvasBg.opacity = isFinite(o) ? clamp(o, 0, 1) : canvasBg.opacity;
        canvasBg.backgroundImage = img || "";

        applyCanvasBackground();
        showToast("Canvas background updated");
      };

      // Browse canvas bg
      const browseBtn = document.getElementById("browseCanvasBgBtn");
      const input = document.getElementById("canvasBgUploadInput");
      browseBtn.onclick = () => { input.value=""; input.click(); };
      input.onchange = async () => {
        const f = input.files?.[0];
        if (!f) return;
        try{
          const cfg = await readJsonFile(f);
          canvasBg = normalizeCanvasBg(cfg);
          applyCanvasBackground();
          showToast("Canvas background theme applied");
          renderFormatPane();
        }catch{
          showToast("Invalid canvas background file");
        }
      };

      // Theme buttons
      document.getElementById("importThemeBtn2").onclick = () => importThemeBtn.click();
      document.getElementById("resetThemeBtn").onclick = () => {
        theme = structuredClone(defaultTheme);
        applyThemeEverywhere();
        showToast("Theme reset");
      };

      return;
    }

    // Visual settings
    const v = selected;

    const posHtml = `
      <div class="fSection">
        <div class="fHeader">
          <div class="fHeaderTitle">Visual</div>
          <div class="fSmall">${escapeHtml(humanType(v.type))}</div>
        </div>

        <div class="row one">
          <div class="field">
            <div class="label">Title</div>
            <input class="input" id="fmtTitle" value="${escapeAttr(v.title)}" />
          </div>
        </div>

        <div class="row">
          <div class="field">
            <div class="label">X</div>
            <input class="input" id="fmtX" value="${escapeAttr(String(v.x))}" />
          </div>
          <div class="field">
            <div class="label">Y</div>
            <input class="input" id="fmtY" value="${escapeAttr(String(v.y))}" />
          </div>
        </div>

        <div class="row">
          <div class="field">
            <div class="label">Width</div>
            <input class="input" id="fmtW" value="${escapeAttr(String(v.w))}" />
          </div>
          <div class="field">
            <div class="label">Height</div>
            <input class="input" id="fmtH" value="${escapeAttr(String(v.h))}" />
          </div>
        </div>

        <div class="smallBtnRow">
          <button class="btn" id="applyPosBtn">Apply</button>
          <button class="btn btnDanger" id="deleteBtn">Delete</button>
        </div>
      </div>
    `;

    const isDataColorVisual = !["image","kpi","card","multirowCard","textBox"].includes(v.type);

    const colorsHtml = (v.type === "image")
      ? `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Image</div>
            <div class="fSmall">Upload / replace</div>
          </div>
          <div class="smallBtnRow">
            <button class="btn" id="replaceImageBtn">Replace image</button>
          </div>
        </div>
      `
      : isDataColorVisual
      ? `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Data colors</div>
            <div class="fSmall">Edit names + colors</div>
          </div>

          ${renderSeriesEditor(v)}

          <div class="smallBtnRow">
            <button class="btn" id="resetVisualColorsBtn">Reset to theme palette</button>
          </div>

          <div class="fSmall" style="margin-top:8px;">
            Tip: Renaming here updates the legend labels too.
          </div>
        </div>
      `
      : "";

    const textHtml = (v.type === "textBox")
      ? `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Text</div>
            <div class="fSmall">Format</div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Font size</div>
              <input class="input" id="txtSize" value="${escapeAttr(String(v.text?.fontSize ?? 16))}" />
            </div>
            <div class="field">
              <div class="label">Text color</div>
              <input class="colorInput" id="txtColor" type="color" value="${escapeAttr(ensureHex(v.text?.color || "#e8ecf2"))}" />
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Align</div>
              <select class="input" id="txtAlign">
                <option value="left"${(v.text?.align==="left")?" selected":""}>Left</option>
                <option value="center"${(v.text?.align==="center")?" selected":""}>Center</option>
                <option value="right"${(v.text?.align==="right")?" selected":""}>Right</option>
              </select>
            </div>
            <div class="field">
              <div class="label">Background</div>
              <input class="input" id="txtBg" value="${escapeAttr(v.text?.bg || "rgba(255,255,255,0.02)")}" />
            </div>
          </div>

          <div class="smallBtnRow">
            <button class="btn" id="txtBoldBtn">${v.text?.bold ? "Unbold" : "Bold"}</button>
            <button class="btn" id="txtItalicBtn">${v.text?.italic ? "Unitalic" : "Italic"}</button>
            <button class="btn" id="txtApplyBtn">Apply</button>
          </div>

          <div class="fSmall" style="margin-top:8px;">
            Tip: Click inside the text box to edit content.
          </div>
        </div>
      `
      : "";

    formatBody.innerHTML = posHtml + colorsHtml + textHtml;

    // Title
    document.getElementById("fmtTitle").oninput = (e) => {
      v.title = e.target.value;
      const t = document.getElementById(`title_${v.id}`);
      if (t) t.textContent = v.title;
    };

    // Position apply (use current canvasW/H)
    document.getElementById("applyPosBtn").onclick = () => {
      const nx = Number(document.getElementById("fmtX").value);
      const ny = Number(document.getElementById("fmtY").value);
      const nw = Number(document.getElementById("fmtW").value);
      const nh = Number(document.getElementById("fmtH").value);

      if (isFinite(nx)) v.x = clamp(nx, 0, canvasW - v.w);
      if (isFinite(ny)) v.y = clamp(ny, 0, canvasH - v.h);
      if (isFinite(nw)) v.w = clamp(nw, 220, canvasW - v.x);
      if (isFinite(nh)) v.h = clamp(nh, 140, canvasH - v.y);

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
      renderFormatPane();
    };

    document.getElementById("deleteBtn").onclick = () => removeVisual(v.id);

    // Series editor
    if (isDataColorVisual) {
      v.series.forEach((s, idx) => {
        const nameEl = document.getElementById(`sname_${v.id}_${idx}`);
        const colEl  = document.getElementById(`scol_${v.id}_${idx}`);

        if (nameEl) {
          nameEl.oninput = (e) => {
            s.name = e.target.value;
            applyThemeToSingleVisual(v);
          };
        }
        if (colEl) {
          colEl.oninput = (e) => {
            s.color = e.target.value;
            s.overrideColor = true;
            applyThemeToSingleVisual(v);
          };
        }
      });

      const resetBtn = document.getElementById("resetVisualColorsBtn");
      if (resetBtn) {
        resetBtn.onclick = () => {
          v.series.forEach((s, i) => {
            s.color = paletteColor(i);
            s.overrideColor = false;
          });
          applyThemeToSingleVisual(v);
          renderFormatPane();
          showToast("Colors reset to theme palette");
        };
      }
    }

    // Replace image
    if (v.type === "image") {
      document.getElementById("replaceImageBtn").onclick = () => {
        imageUploadInput.value = "";
        imageUploadInput.click();
        imageUploadInput.onchange = async () => {
          const f = imageUploadInput.files?.[0];
          if (!f) return;
          v.imageDataUrl = await readImageAsDataUrl(f);
          renderVisualContent(v);
          showToast("Image replaced");
        };
      };
    }

    // Text formatting
    if (v.type === "textBox") {
      document.getElementById("txtBoldBtn").onclick = () => {
        v.text = v.text || buildDefaultText("textBox");
        v.text.bold = !v.text.bold;
        renderVisualContent(v);
        renderFormatPane();
      };
      document.getElementById("txtItalicBtn").onclick = () => {
        v.text = v.text || buildDefaultText("textBox");
        v.text.italic = !v.text.italic;
        renderVisualContent(v);
        renderFormatPane();
      };
      document.getElementById("txtApplyBtn").onclick = () => {
        v.text = v.text || buildDefaultText("textBox");
        const size = Number(document.getElementById("txtSize").value);
        const color = document.getElementById("txtColor").value;
        const align = document.getElementById("txtAlign").value;
        const bg = document.getElementById("txtBg").value;

        if (isFinite(size)) v.text.fontSize = clamp(size, 10, 72);
        v.text.color = color || v.text.color;
        v.text.align = align;
        v.text.bg = bg || v.text.bg;

        renderVisualContent(v);
        showToast("Text formatting applied");
      };
    }
  }

  function renderSeriesEditor(v){
    if (!v.series || v.series.length === 0) {
      return `<div class="fSmall">No series for this visual.</div>`;
    }

    return v.series.map((s, i) => `
      <div class="colorRow">
        <input class="input" id="sname_${v.id}_${i}" value="${escapeAttr(s.name)}" />
        <input class="colorInput" id="scol_${v.id}_${i}" type="color" value="${escapeAttr(ensureHex(s.color))}" title="Change series color" />
      </div>
    `).join("");
  }

  function ensureHex(c){
    if (!c) return "#3b82f6";
    if (c.startsWith("#")) {
      if (c.length === 4) return "#" + c.slice(1).split("").map(x=>x+x).join("");
      return c.slice(0,7);
    }
    return paletteColor(0);
  }

  function normalizeCanvasBg(obj){
    const out = structuredClone(sampleCanvasBg);
    if (obj && typeof obj === "object") {
      if (typeof obj.backgroundColor === "string") out.backgroundColor = obj.backgroundColor;
      if (typeof obj.backgroundImage === "string") out.backgroundImage = obj.backgroundImage;
      if (obj.opacity !== undefined) out.opacity = clamp(Number(obj.opacity), 0, 1);
    }
    return out;
  }

  function humanType(type){
    const found = VISUALS.flatMap(s => s.items).find(x => x.type === type);
    return found ? found.name : type;
  }

  // -------------------------
  // Theme import
  // -------------------------
  if (importThemeBtn) {
    importThemeBtn.addEventListener("click", () => {
      importThemeInput.value = "";
      importThemeInput.click();
    });
  }

  if (importThemeInput) {
    importThemeInput.addEventListener("change", async () => {
      const f = importThemeInput.files?.[0];
      if (!f) return;
      try{
        const t = await readJsonFile(f);
        theme = normalizeTheme(t);
        applyThemeEverywhere();
        showToast("Theme import successful");
      }catch{
        showToast("Invalid theme JSON");
      }
    });
  }

  function normalizeTheme(t){
    const out = structuredClone(defaultTheme);
    if (!t || typeof t !== "object") return out;

    if (typeof t.name === "string") out.name = t.name;
    if (Array.isArray(t.dataColors) && t.dataColors.length) out.dataColors = t.dataColors.slice();
    if (typeof t.background === "string") out.background = t.background;
    if (typeof t.foreground === "string") out.foreground = t.foreground;

    out.textClasses = out.textClasses || {};
    if (t.textClasses && typeof t.textClasses === "object") {
      out.textClasses.title = { ...out.textClasses.title, ...(t.textClasses.title || {}) };
      out.textClasses.label = { ...out.textClasses.label, ...(t.textClasses.label || {}) };
    }

    out.chart = out.chart || {};
    if (t.chart && typeof t.chart === "object") {
      out.chart = { ...out.chart, ...t.chart };
    }

    return out;
  }

  // -------------------------
  // Download/Upload dashboard
  // -------------------------
  if (downloadDashboardBtn) {
    downloadDashboardBtn.addEventListener("click", () => {
      const payload = serializeDashboard();
      downloadObjectAsJson(payload, "dashboard.json");
      showToast("Dashboard downloaded");
    });
  }

  if (uploadDashboardBtn) {
    uploadDashboardBtn.addEventListener("click", () => {
      uploadDashboardInput.value = "";
      uploadDashboardInput.click();
    });
  }

  if (uploadDashboardInput) {
    uploadDashboardInput.addEventListener("change", async () => {
      const f = uploadDashboardInput.files?.[0];
      if (!f) return;

      const loadIt = async () => {
        try{
          const obj = await readJsonFile(f);
          await loadDashboard(obj);
          showToast("Dashboard loaded");
        }catch{
          showToast("Invalid dashboard file");
        } finally {
          uploadDashboardInput.value = "";
        }
      };

      if (!isDefaultState()) openDiscardModal(loadIt);
      else loadIt();
    });
  }

  function openDiscardModal(onConfirm){
    if (!modalBackdrop) return;

    if (isDefaultState()) {
      onConfirm?.();
      return;
    }

    modalBackdrop.hidden = false;
    modalBackdrop.style.display = "flex";
    modalBackdrop.setAttribute("aria-hidden", "false");

    const close = () => {
      modalBackdrop.hidden = true;
      modalBackdrop.style.display = "none";
      modalBackdrop.setAttribute("aria-hidden", "true");
    };

    modalCancelBtn.onclick = () => close();
    modalConfirmBtn.onclick = async () => {
      close();
      await onConfirm();
    };
  }

  function serializeDashboard(){
    const visuals = state.order.map(id => {
      const v = state.visuals.get(id);
      if (!v) return null;

      return {
        id: v.id,
        type: v.type,
        title: v.title,
        x: v.x, y: v.y, w: v.w, h: v.h,
        series: (v.series || []).map(s => ({
          key: s.key,
          name: s.name,
          color: s.color,
          overrideColor: !!s.overrideColor
        })),
        imageDataUrl: v.type === "image" ? (v.imageDataUrl || "") : undefined,
        card: (["kpi","card","multirowCard"].includes(v.type)) ? (v.card || null) : undefined,
        text: (v.type === "textBox") ? (v.text || null) : undefined
      };
    }).filter(Boolean);

    return {
      version: 2,
      canvas: {
        width: canvasW,
        height: canvasH,
        presetKey: canvasPresetKey,
        background: canvasBg
      },
      theme,
      visuals
    };
  }

  async function loadDashboard(obj){
    clearAllVisuals();

    // canvas
    const cw = Number(obj?.canvas?.width);
    const ch = Number(obj?.canvas?.height);
    if (isFinite(cw) && cw >= 600) canvasW = Math.round(cw);
    if (isFinite(ch) && ch >= 400) canvasH = Math.round(ch);

    const pk = String(obj?.canvas?.presetKey || "");
    if (CANVAS_PRESETS.some(p => p.key === pk)) canvasPresetKey = pk;

    // background
    if (obj?.canvas?.background) canvasBg = normalizeCanvasBg(obj.canvas.background);
    else canvasBg = structuredClone(sampleCanvasBg);
    applyCanvasBackground();

    // theme
    theme = normalizeTheme(obj?.theme || defaultTheme);

    // apply canvas size (also clamps)
    applyCanvasSize();

    // visuals
    const visuals = Array.isArray(obj?.visuals) ? obj.visuals : [];
    visuals.forEach(vs => {
      const id = String(vs.id || `vis${state.nextId++}`);
      const type = vs.type;

      const v = {
        id,
        type,
        title: vs.title || defaultTitleFor(type),

        x: clamp(Number(vs.x)||40, 0, canvasW-220),
        y: clamp(Number(vs.y)||40, 0, canvasH-140),
        w: clamp(Number(vs.w)||DEFAULT_SIZES[type]?.w||520, 220, canvasW),
        h: clamp(Number(vs.h)||DEFAULT_SIZES[type]?.h||300, 140, canvasH),

        series: Array.isArray(vs.series)
          ? vs.series.map((s,i)=>({
              key: s.key || s.name || `S${i+1}`,
              name: s.name || s.key || `Series ${i+1}`,
              color: s.color || paletteColor(i),
              overrideColor: !!s.overrideColor
            }))
          : buildDefaultSeries(type),

        imageDataUrl: type === "image" ? (vs.imageDataUrl || "") : undefined,
        card: (["kpi","card","multirowCard"].includes(type)) ? (vs.card || buildDefaultCard(type)) : buildDefaultCard(type),
        text: (type === "textBox") ? (vs.text || buildDefaultText("textBox")) : null,
        chart: null
      };

      v.series?.forEach((s,i)=>{
        if (!s.overrideColor) s.color = paletteColor(i);
      });

      // clamp strictly in bounds
      v.x = clamp(v.x, 0, Math.max(0, canvasW - v.w));
      v.y = clamp(v.y, 0, Math.max(0, canvasH - v.h));

      state.visuals.set(id, v);
      state.order.push(id);

      const el = createVisualElement(v);
      canvas.appendChild(el);
      renderVisualContent(v);
    });

    updateZOrder();
    setSelected(null);
    state.nextId = Math.max(state.nextId, state.order.length + 1);

    applyThemeEverywhere();
    fitCanvasToScreen();
  }

  function clearAllVisuals(){
    state.order.forEach(id => {
      const v = state.visuals.get(id);
      if (v?.chart) {
        try { v.chart.destroy(); } catch {}
      }
    });
    state.visuals.clear();
    state.order = [];
    state.selectedId = null;
    Array.from(canvas.querySelectorAll(".visual")).forEach(el => el.remove());
  }

  // -------------------------
  // Sample downloads
  // -------------------------
  if (downloadSampleThemeBtn) {
    downloadSampleThemeBtn.addEventListener("click", () => {
      downloadObjectAsJson(sampleTheme, "sample-theme.json");
      showToast("Sample theme downloaded");
    });
  }

  if (downloadSampleCanvasBgBtn) {
    downloadSampleCanvasBgBtn.addEventListener("click", () => {
      downloadObjectAsJson(sampleCanvasBg, "sample-canvas-background.json");
      showToast("Sample canvas background downloaded");
    });
  }

  // -------------------------
  // Start: always show canvas settings
  // -------------------------
  renderFormatPane();
  fitCanvasToScreen();

})();
