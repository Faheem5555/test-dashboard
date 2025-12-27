/* =========================
   Power BI Prototype Dashboard Website
   FULL IMPLEMENTATION (no skips)
   - Fixed canvas 1280×720, no infinite scroll
   - Visual picker dropdown (category-wise) with icons
   - Many visuals with distinct prototypes (Chart.js + custom prototype blocks)
   - Absolute-position visuals (overlap allowed, no pushing down)
   - Drag + resize (smooth) + selection outline + delete only on selection
   - Always-visible Format pane:
       • Canvas settings when none selected
       • Visual formatting when selected
       • Per-series/slice colors + editable legend/series names
   - Upload Image visual (behaves like a visual)
   - Theme import (Power BI-like):
       • Theme JSON import
       • Toast “Theme import successful”
       • Applies to existing + new visuals immediately
   - Canvas background upload (file-based JSON format)
   - Download/Upload dashboard (JSON), with discard confirmation logic
========================= */

(() => {
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

  // -------------------------
  // Fixed canvas size (Hard rule)
  // -------------------------
  const CANVAS_W = 1280;
  const CANVAS_H = 720;
  statusPill.textContent = `Canvas: ${CANVAS_W} × ${CANVAS_H}`;

  // -------------------------
  // Dummy retail data (realistic growth Jan→Dec)
  // -------------------------
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // base gradual growth, approx +80% YoY by end (visual only)
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
    // Regions for multi-series
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
  // Theme (Power BI-like) - applied globally
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

  // Sample theme JSON (downloadable)
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

  // Canvas background format (uploadable JSON)
  // Supports:
  //  - backgroundColor: string
  //  - backgroundImage: string (URL or data URL)
  //  - opacity: 0..1
  const sampleCanvasBg = {
    backgroundColor: "#0b0d12",
    backgroundImage: "",
    opacity: 1
  };

  let theme = structuredClone(defaultTheme);

  // Canvas background state
  let canvasBg = structuredClone(sampleCanvasBg);

  // Apply initial canvas bg
  applyCanvasBackground();

  // -------------------------
  // Visual registry (picker)
  // -------------------------
  // NOTE: prototypes must be distinct and correct-looking; some are Chart.js, some are custom visual prototypes.
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
      category: "Other",
      items: [
        { type: "image", name: "Upload Image", icon: "🖼" }
      ]
    }
  ];

  // Default visual sizes (medium, readable, not full width)
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

    image: { w: 420, h: 260 }
  };

  // -------------------------
  // App state
  // -------------------------
  const state = {
    visuals: new Map(), // id -> visual object
    order: [],          // z-order
    selectedId: null,
    nextId: 1
  };

  // Visual object schema (saved/loaded):
  // {
  //   id, type, title,
  //   x,y,w,h,
  //   series: [{ key, name, color, overrideColor:boolean }],
  //   // per visual:
  //   imageDataUrl?: string,
  //   // internal:
  //   chart?: Chart instance
  // }

  // -------------------------
  // Helpers
  // -------------------------
  const clamp = (v, min, max) => Math.max(min, Math.min(max, v));

  function showToast(msg){
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
      JSON.stringify(canvasBg) === JSON.stringify(sampleCanvasBg);
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

  // Theme → Chart.js common options
  function makeChartDefaults(){
    const titleColor = theme?.textClasses?.title?.color || theme.foreground || "#e8ecf2";
    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";
    const gridColor  = theme?.chart?.grid || "rgba(255,255,255,0.08)";

    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 300 },
      plugins: {
        legend: {
          display: true,
          labels: { color: labelColor, boxWidth: 10, boxHeight: 10, usePointStyle: true }
        },
        title: {
          display: false
        },
        tooltip: {
          enabled: true
        }
      },
      scales: {
        x: {
          ticks: { color: labelColor },
          grid: { color: gridColor }
        },
        y: {
          ticks: { color: labelColor },
          grid: { color: gridColor }
        }
      }
    };
  }

  // Generate palette color for series index
  function paletteColor(i){
    const arr = theme?.dataColors?.length ? theme.dataColors : defaultTheme.dataColors;
    return arr[i % arr.length];
  }

  // -------------------------
  // Visual picker render + open/close
  // -------------------------
  function renderVisualPicker(){
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
          if (item.type === "image") {
            addImageVisualFlow();
          } else {
            addVisual(item.type);
          }
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
    visualPickerMenu.classList.add("open");
    visualPickerToggle.setAttribute("aria-expanded", "true");
  }
  function closePicker(){
    visualPickerMenu.classList.remove("open");
    visualPickerToggle.setAttribute("aria-expanded", "false");
  }

  visualPickerToggle.addEventListener("click", (e) => {
    e.stopPropagation();
    if (visualPickerMenu.classList.contains("open")) closePicker();
    else openPicker();
  });

  document.addEventListener("click", () => closePicker());

  renderVisualPicker();

  // -------------------------
  // Selection behavior (Hard rules)
  // -------------------------
  canvas.addEventListener("mousedown", (e) => {
    // Clicking empty canvas deselects all
    if (e.target === canvas || e.target === canvasWatermarkOrBefore(e.target)) {
      setSelected(null);
    }
  });

  function canvasWatermarkOrBefore(target){
    // canvas ::before isn't a node. watermark is.
    const wm = document.getElementById("canvasWatermark");
    return target === wm;
  }

  function setSelected(id){
    state.selectedId = id;

    // Update visual DOM selection classes
    state.order.forEach(vid => {
      const el = document.getElementById(`v_${vid}`);
      if (!el) return;
      if (vid === id) el.classList.add("selected");
      else el.classList.remove("selected");
    });

    // Bring selected to front (like clicking in Power BI)
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

    // place near top-left with slight offset to avoid full overlap on add
    const offset = (state.order.length * 18) % 140;
    const x = clamp(40 + offset, 0, CANVAS_W - size.w);
    const y = clamp(40 + offset, 0, CANVAS_H - size.h);

    const v = {
      id,
      type,
      title: defaultTitleFor(type),
      x, y, w: size.w, h: size.h,
      series: buildDefaultSeries(type),
      chart: null
    };

    state.visuals.set(id, v);
    state.order.push(id);

    const el = createVisualElement(v);
    canvas.appendChild(el);
    updateZOrder();

    // create chart/prototype
    renderVisualContent(v);

    // auto select new
    setSelected(id);
  }

  async function addImageVisualFlow(){
    imageUploadInput.value = "";
    imageUploadInput.click();

    imageUploadInput.onchange = async () => {
      const f = imageUploadInput.files?.[0];
      if (!f) return;

      const dataUrl = await readImageAsDataUrl(f);

      const id = `vis${state.nextId++}`;
      const size = DEFAULT_SIZES.image;
      const offset = (state.order.length * 18) % 140;
      const x = clamp(40 + offset, 0, CANVAS_W - size.w);
      const y = clamp(40 + offset, 0, CANVAS_H - size.h);

      const v = {
        id,
        type: "image",
        title: "Image",
        x, y, w: size.w, h: size.h,
        series: [],
        imageDataUrl: dataUrl,
        chart: null
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
      scatter: "Profit vs Sales (Scatter)"
    };
    return map[type] || "Visual";
  }

  function buildDefaultSeries(type){
    // By default, series colors follow theme palette.
    // If user changes a color, we set overrideColor=true for that series.
    if (type === "pie" || type === "donut") {
      return retail.Categories.map((c, i) => ({
        key: c.name,
        name: c.name,
        color: paletteColor(i),
        overrideColor: false
      }));
    }

    // Multi-series visuals: Regions
    const regionKeys = Object.keys(retail.Regions);
    // For single-series visuals that still show legend, we keep 1 dataset
    const singleSeries = [{ key: "Value", name: "Value", color: paletteColor(0), overrideColor: false }];

    const multiSeries = regionKeys.map((k, i) => ({
      key: k, name: k, color: paletteColor(i), overrideColor: false
    }));

    switch(type){
      case "line":
      case "stackedArea":
      case "clusteredBar":
      case "stackedBar":
      case "stackedBar100":
      case "clusteredColumn":
      case "stackedColumn":
      case "stackedColumn100":
      case "lineStackedColumn":
        return multiSeries;

      case "area":
      case "scatter":
      case "lineClusteredColumn":
      case "treemap":
      case "ribbon":
      default:
        return singleSeries;
    }
  }

  // -------------------------
  // Visual DOM creation
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

      <!-- resize handles (visible only when selected) -->
      <div class="handle nw" data-handle="nw"></div>
      <div class="handle ne" data-handle="ne"></div>
      <div class="handle sw" data-handle="sw"></div>
      <div class="handle se" data-handle="se"></div>
      <div class="handle n"  data-handle="n"></div>
      <div class="handle s"  data-handle="s"></div>
      <div class="handle w"  data-handle="w"></div>
      <div class="handle e"  data-handle="e"></div>
    `;

    // Selection
    el.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      setSelected(v.id);
    });

    // Delete only when selected (button exists always but hidden via CSS)
    el.querySelector(".vDelete").addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(v.id);
    });

    // Dragging by header
    const header = el.querySelector(".vHeader");
    header.addEventListener("mousedown", (e) => startDrag(e, v.id));

    // Resizing by handles
    el.querySelectorAll(".handle").forEach(h => {
      h.addEventListener("mousedown", (e) => startResize(e, v.id, h.dataset.handle));
    });

    return el;
  }

  function removeVisual(id){
    const v = state.visuals.get(id);
    if (!v) return;

    // Destroy chart instance if exists
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
  // Drag & Resize (smooth, overlap allowed)
  // -------------------------
  let dragCtx = null;

  function startDrag(e, id){
    // Only left button
    if (e.button !== 0) return;

    const v = state.visuals.get(id);
    if (!v) return;

    setSelected(id);

    const el = document.getElementById(`v_${id}`);
    const startX = e.clientX;
    const startY = e.clientY;

    dragCtx = {
      mode: "drag",
      id,
      startX, startY,
      origX: v.x,
      origY: v.y,
      el
    };

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp, { once:true });
  }

  function startResize(e, id, handle){
    if (e.button !== 0) return;

    const v = state.visuals.get(id);
    if (!v) return;

    setSelected(id);

    const startX = e.clientX;
    const startY = e.clientY;

    dragCtx = {
      mode: "resize",
      id,
      handle,
      startX, startY,
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
      v.x = clamp(dragCtx.origX + dx, 0, CANVAS_W - v.w);
      v.y = clamp(dragCtx.origY + dy, 0, CANVAS_H - v.h);
      applyVisualRect(v);
    } else {
      // resize
      const minW = 220;
      const minH = 170;

      let x = dragCtx.origX;
      let y = dragCtx.origY;
      let w = dragCtx.origW;
      let h = dragCtx.origH;

      const hnd = dragCtx.handle;

      const applyW = (newW) => clamp(newW, minW, CANVAS_W - x);
      const applyH = (newH) => clamp(newH, minH, CANVAS_H - y);

      if (hnd.includes("e")) w = applyW(dragCtx.origW + dx);
      if (hnd.includes("s")) h = applyH(dragCtx.origH + dy);

      if (hnd.includes("w")) {
        const newX = clamp(dragCtx.origX + dx, 0, dragCtx.origX + dragCtx.origW - minW);
        const newW = dragCtx.origW + (dragCtx.origX - newX);
        x = newX;
        w = clamp(newW, minW, CANVAS_W - x);
      }

      if (hnd.includes("n")) {
        const newY = clamp(dragCtx.origY + dy, 0, dragCtx.origY + dragCtx.origH - minH);
        const newH = dragCtx.origH + (dragCtx.origY - newY);
        y = newY;
        h = clamp(newH, minH, CANVAS_H - y);
      }

      v.x = x; v.y = y; v.w = w; v.h = h;
      applyVisualRect(v);

      // Chart.js needs resize
      if (v.chart) v.chart.resize();
    }
  }

  function onUp(){
    window.removeEventListener("mousemove", onMove);
    dragCtx = null;

    // Keep format pane values accurate
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

  // -------------------------
  // Render visual content (Chart.js + prototypes)
  // -------------------------
  function renderVisualContent(v){
    const body = document.getElementById(`body_${v.id}`);
    if (!body) return;

    // Clean existing content + destroy chart
    if (v.chart) {
      try { v.chart.destroy(); } catch {}
      v.chart = null;
    }
    body.innerHTML = "";

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

    if (v.type === "treemap") {
      body.appendChild(makeTreemapPrototype(v));
      return;
    }

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
    const opts = makeChartDefaults();
    const ds = buildChartJsData(v);

    v.chart = new Chart(ctx, ds.config);
    // Apply theme styling (legend labels, grid, etc.)
    applyThemeToSingleVisual(v);
  }

  function buildChartJsData(v){
    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";
    const gridColor  = theme?.chart?.grid || "rgba(255,255,255,0.08)";

    const cfgBase = {
      type: "bar",
      data: { labels: months, datasets: [] },
      options: {
        ...makeChartDefaults(),
        plugins: {
          ...makeChartDefaults().plugins,
          legend: {
            ...makeChartDefaults().plugins.legend,
            labels: { ...makeChartDefaults().plugins.legend.labels, color: labelColor }
          }
        },
        scales: {
          x: { ...makeChartDefaults().scales.x, grid: { color: gridColor }, ticks: { color: labelColor } },
          y: { ...makeChartDefaults().scales.y, grid: { color: gridColor }, ticks: { color: labelColor } }
        }
      }
    };

    // helper to pull region series values
    const regionSeries = v.series.map((s) => retail.Regions[s.key] || mkGrowth(20, 60));

    const single = mkGrowth(30, 55);

    const makeLineDs = (s, values, fill=false, stacked=false) => ({
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

    const makeBarDs = (s, values, stacked=false) => ({
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

        const cfg = {
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
            plugins: {
              ...makeChartDefaults().plugins,
              legend: { ...makeChartDefaults().plugins.legend }
            },
            cutout: v.type === "donut" ? "58%" : "0%"
          }
        };
        return { config: cfg };
      }

      case "line": {
        cfgBase.type = "line";
        cfgBase.data.labels = months;
        cfgBase.data.datasets = v.series.map((s, i) => makeLineDs(s, regionSeries[i], false));
        return { config: cfgBase };
      }

      case "area": {
        cfgBase.type = "line";
        cfgBase.data.labels = months;
        // single series (Profit)
        const s = v.series[0] || { name: "Profit", color: paletteColor(0) };
        cfgBase.data.datasets = [ makeLineDs({ ...s, name: v.series[0]?.name || "Profit" }, retail.Profit, true) ];
        return { config: cfgBase };
      }

      case "stackedArea": {
        cfgBase.type = "line";
        cfgBase.data.labels = months;
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;

        cfgBase.data.datasets = v.series.map((s, i) => ({
          ...makeLineDs(s, regionSeries[i], true),
          fill: "origin"
        }));
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
        cfgBase.options.plugins.legend.display = true;
        return { config: cfgBase };
      }

      case "stackedBar":
      case "stackedBar100": {
        cfgBase.type = "bar";
        cfgBase.options.indexAxis = "y";
        cfgBase.data.labels = months;
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;

        const datasets = v.series.map((s, i) => makeBarDs(s, regionSeries[i], true));
        cfgBase.data.datasets = datasets;

        if (v.type === "stackedBar100") {
          cfgBase.options.plugins.tooltip.callbacks = {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.formattedValue}%`
          };
          cfgBase.data = toPercentStacked(months, datasets);
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
        cfgBase.data.labels = months;
        cfgBase.options.scales.x.stacked = true;
        cfgBase.options.scales.y.stacked = true;

        const datasets = v.series.map((s, i) => makeBarDs(s, regionSeries[i], true));
        cfgBase.data.datasets = datasets;

        if (v.type === "stackedColumn100") {
          cfgBase.options.plugins.tooltip.callbacks = {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.formattedValue}%`
          };
          cfgBase.data = toPercentStacked(months, datasets);
        }
        return { config: cfgBase };
      }

      case "lineClusteredColumn": {
        // Mixed: clustered columns (Sales) + line (Profit)
        const barColor = v.series[0]?.color || paletteColor(0);
        const lineColor = v.series[1]?.color || paletteColor(1);

        const cfg = {
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
        };
        return { config: cfg };
      }

      case "lineStackedColumn": {
        // Mixed: stacked columns (regions) + line (Profit)
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
        // Profit vs Sales scatter, plus trend-ish look
        const s = v.series[0] || { name: "Points", color: paletteColor(0) };
        const points = months.map((m, i) => ({ x: retail.Sales[i], y: retail.Profit[i] }));
        const cfg = {
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
          options: (() => {
            const o = makeChartDefaults();
            o.scales.x.title = { display:true, text:"Sales", color: (theme?.textClasses?.label?.color || "#cfd6e1") };
            o.scales.y.title = { display:true, text:"Profit", color: (theme?.textClasses?.label?.color || "#cfd6e1") };
            return o;
          })()
        };
        return { config: cfg };
      }

      default: {
        // fallback (shouldn't happen)
        cfgBase.type = "line";
        cfgBase.data.datasets = [ makeLineDs({ name:"Value", color: paletteColor(0) }, single, true) ];
        return { config: cfgBase };
      }
    }
  }

  // Treemap prototype: distinct non-Chart.js layout blocks
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
      { label: cats[0]?.name || "Electronics", color: cats[0]?.color || paletteColor(0), span:"r2" },
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

  // Ribbon prototype: distinct SVG “ribbons”
  function makeRibbonPrototype(v){
    const host = document.createElement("div");
    host.style.width = "100%";
    host.style.height = "100%";
    host.style.border = "1px solid rgba(255,255,255,0.10)";
    host.style.borderRadius = "12px";
    host.style.overflow = "hidden";
    host.style.background = "rgba(255,255,255,0.01)";

    const series = (v.series && v.series.length > 1) ? v.series : Object.keys(retail.Regions).map((k,i)=>({
      key:k, name:k, color: paletteColor(i), overrideColor:false
    }));

    const svg = document.createElementNS("http://www.w3.org/2000/svg","svg");
    svg.setAttribute("viewBox","0 0 600 260");
    svg.setAttribute("preserveAspectRatio","none");
    svg.style.width = "100%";
    svg.style.height = "100%";

    // background grid
    const grid = document.createElementNS("http://www.w3.org/2000/svg","path");
    grid.setAttribute("d","M0 210 H600 M0 160 H600 M0 110 H600 M0 60 H600");
    grid.setAttribute("stroke", theme?.chart?.grid || "rgba(255,255,255,0.10)");
    grid.setAttribute("stroke-width","1");
    grid.setAttribute("fill","none");
    svg.appendChild(grid);

    // Ribbons (simple bezier bands)
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

      // top curve
      const dTop = `M0 ${b.y0}
        C150 ${b.y0-w}, 300 ${b.y0+w}, 450 ${b.y0-w}
        S600 ${b.y0+w}, 600 ${b.y0}`;
      // bottom curve
      const dBot = `L600 ${b.y1}
        C450 ${b.y1+w}, 300 ${b.y1-w}, 150 ${b.y1+w}
        S0 ${b.y1-w}, 0 ${b.y1} Z`;

      path.setAttribute("d", dTop + " " + dBot);
      path.setAttribute("fill", c);
      path.setAttribute("stroke", withAlpha(s.color, 0.7));
      path.setAttribute("stroke-width", "1.2");
      svg.appendChild(path);

      // label
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
    // Accept rgb/rgba too
    if (!hex) return `rgba(255,255,255,${a})`;
    if (hex.startsWith("rgba")) return hex;
    if (hex.startsWith("rgb(")) return hex.replace("rgb(", "rgba(").replace(")", `,${a})`);

    const h = hex.replace("#","").trim();
    const full = h.length === 3
      ? h.split("").map(ch => ch+ch).join("")
      : h.padEnd(6, "0").slice(0,6);

    const r = parseInt(full.slice(0,2), 16);
    const g = parseInt(full.slice(2,4), 16);
    const b = parseInt(full.slice(4,6), 16);
    return `rgba(${r},${g},${b},${a})`;
  }

  function toPercentStacked(labels, datasets){
    // Convert stacked values into 100% per category index
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
    // For non-chart prototypes, just re-render them to update colors
    if (v.type === "treemap" || v.type === "ribbon") {
      renderVisualContent(v);
      return;
    }
    if (!v.chart) return;

    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";
    const gridColor  = theme?.chart?.grid || "rgba(255,255,255,0.08)";

    // legend label color
    if (v.chart.options?.plugins?.legend?.labels) {
      v.chart.options.plugins.legend.labels.color = labelColor;
    }

    // scales
    if (v.chart.options?.scales?.x?.ticks) v.chart.options.scales.x.ticks.color = labelColor;
    if (v.chart.options?.scales?.y?.ticks) v.chart.options.scales.y.ticks.color = labelColor;
    if (v.chart.options?.scales?.x?.grid) v.chart.options.scales.x.grid.color = gridColor;
    if (v.chart.options?.scales?.y?.grid) v.chart.options.scales.y.grid.color = gridColor;

    // apply dataset colors from v.series if relevant
    syncSeriesToChart(v);

    v.chart.update();
  }

  function syncSeriesToChart(v){
    if (!v.chart) return;

    // pie/donut: dataset[0].backgroundColor = series colors
    if (v.type === "pie" || v.type === "donut") {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.backgroundColor = v.series.map(s => withAlpha(s.color, 0.82));
      ds.borderColor = v.series.map(s => withAlpha(s.color, 1));
      v.chart.data.labels = v.series.map(s => s.name);
      return;
    }

    // treemap/ribbon handled elsewhere (not chart)
    // scatter: dataset label + point color
    if (v.type === "scatter") {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.label = v.series[0]?.name || ds.label;
      ds.backgroundColor = withAlpha(v.series[0]?.color || paletteColor(0), 0.75);
      ds.borderColor = withAlpha(v.series[0]?.color || paletteColor(0), 1);
      return;
    }

    // line/area/stacked area: dataset count matches series
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

    // bar/column + stacked variants: multiple datasets often map series
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

    // clusteredBar/clusteredColumn are single dataset
    if (["clusteredBar","clusteredColumn"].includes(v.type)) {
      const ds = v.chart.data.datasets?.[0];
      if (!ds) return;
      ds.label = v.series[0]?.name || ds.label;
      ds.backgroundColor = withAlpha(v.series[0]?.color || paletteColor(0), 0.75);
      return;
    }

    // combos:
    if (v.type === "lineClusteredColumn") {
      const dsets = v.chart.data.datasets || [];
      // dataset 0: bar, dataset 1: line
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
      // first N bars map to v.series; last is line "Profit" (kept themed)
      const n = v.series.length;
      for (let i=0; i<n; i++){
        if (!dsets[i]) continue;
        const s = v.series[i];
        dsets[i].label = s.name;
        dsets[i].backgroundColor = withAlpha(s.color, 0.68);
        dsets[i].borderColor = withAlpha(s.color, 0.92);
      }
      // keep last line visible; user colors series bars only (by requirement)
      return;
    }
  }

  function applyThemeEverywhere(){
    // Update app background/foreground if provided (Power BI theme-ish)
    // Keep the UI stable: just use theme colors lightly.
    // We still apply chart/series palette strongly (required).
    // Update non-overridden series colors from palette:
    state.order.forEach((id) => {
      const v = state.visuals.get(id);
      if (!v) return;

      // if series exists, update those not overridden
      v.series?.forEach((s, i) => {
        if (!s.overrideColor) s.color = paletteColor(i);
      });

      // update title label color (format pane shows it; header remains clean)
      applyThemeToSingleVisual(v);
    });

    renderFormatPane();
  }

  // -------------------------
  // Format pane (always visible; never disappears)
  // -------------------------
  function renderFormatPane(){
    const selected = state.selectedId ? state.visuals.get(state.selectedId) : null;

    if (!selected) {
      // Canvas settings view
      formatBody.innerHTML = `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Canvas settings</div>
            <div class="fSmall">No visual selected</div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Canvas width</div>
              <input class="input" value="${CANVAS_W}" disabled />
            </div>
            <div class="field">
              <div class="label">Canvas height</div>
              <input class="input" value="${CANVAS_H}" disabled />
            </div>
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
            <button class="btn" id="applyCanvasBgBtn">Apply</button>
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
            Theme applies to chart palette + labels + grid instantly (existing + new visuals).
          </div>
        </div>
      `;

      // Hook canvas apply
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

    // Visual settings view
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
      : `
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
      `;

    formatBody.innerHTML = posHtml + colorsHtml;

    // Hook: title + position
    document.getElementById("fmtTitle").oninput = (e) => {
      v.title = e.target.value;
      const t = document.getElementById(`title_${v.id}`);
      if (t) t.textContent = v.title;
    };

    document.getElementById("applyPosBtn").onclick = () => {
      const nx = Number(document.getElementById("fmtX").value);
      const ny = Number(document.getElementById("fmtY").value);
      const nw = Number(document.getElementById("fmtW").value);
      const nh = Number(document.getElementById("fmtH").value);

      if (isFinite(nx)) v.x = clamp(nx, 0, CANVAS_W - v.w);
      if (isFinite(ny)) v.y = clamp(ny, 0, CANVAS_H - v.h);
      if (isFinite(nw)) v.w = clamp(nw, 220, CANVAS_W - v.x);
      if (isFinite(nh)) v.h = clamp(nh, 170, CANVAS_H - v.y);

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
      renderFormatPane();
    };

    document.getElementById("deleteBtn").onclick = () => removeVisual(v.id);

    // Hook: series editor
    if (v.type !== "image") {
      v.series.forEach((s, idx) => {
        const nameEl = document.getElementById(`sname_${v.id}_${idx}`);
        const colEl  = document.getElementById(`scol_${v.id}_${idx}`);

        if (nameEl) {
          nameEl.oninput = (e) => {
            s.name = e.target.value;
            applyThemeToSingleVisual(v); // updates legend labels
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

    // Hook: replace image
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
    // If rgba/rgb, return palette as fallback (color input needs hex)
    if (!c) return "#3b82f6";
    if (c.startsWith("#")) {
      if (c.length === 4) return "#" + c.slice(1).split("").map(x=>x+x).join("");
      return c.slice(0,7);
    }
    // fallback: use palette color
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
  // Theme import (Power BI-like)
  // -------------------------
  importThemeBtn.addEventListener("click", () => {
    importThemeInput.value = "";
    importThemeInput.click();
  });

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

  function normalizeTheme(t){
    const out = structuredClone(defaultTheme);
    if (!t || typeof t !== "object") return out;

    if (typeof t.name === "string") out.name = t.name;

    // Support Power BI theme generator-like schema:
    // - dataColors: array
    // - foreground/background
    // - textClasses.title/label
    // (Your uploaded sample.json matches this)
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
  // Download/Upload dashboard (with discard confirm)
  // -------------------------
  downloadDashboardBtn.addEventListener("click", () => {
    const payload = serializeDashboard();
    downloadObjectAsJson(payload, "dashboard.json");
    showToast("Dashboard downloaded");
  });

  uploadDashboardBtn.addEventListener("click", () => {
    uploadDashboardInput.value = "";
    uploadDashboardInput.click();
  });

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
      }
    };

    // Discard confirmation only if NOT empty/default
    if (!isDefaultState()) {
      openDiscardModal(loadIt);
    } else {
      loadIt();
    }
  });

  function openDiscardModal(onConfirm){
    modalBackdrop.hidden = false;

    const close = () => { modalBackdrop.hidden = true; };

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
        imageDataUrl: v.type === "image" ? (v.imageDataUrl || "") : undefined
      };
    }).filter(Boolean);

    return {
      version: 1,
      canvas: {
        width: CANVAS_W,
        height: CANVAS_H,
        background: canvasBg
      },
      theme,
      visuals
    };
  }

  async function loadDashboard(obj){
    // Clear current
    clearAllVisuals();

    // Apply canvas + theme
    if (obj?.canvas?.background) {
      canvasBg = normalizeCanvasBg(obj.canvas.background);
      applyCanvasBackground();
    } else {
      canvasBg = structuredClone(sampleCanvasBg);
      applyCanvasBackground();
    }

    theme = normalizeTheme(obj?.theme || defaultTheme);
    applyThemeEverywhere(); // sets palette + styling

    // Restore visuals
    const visuals = Array.isArray(obj?.visuals) ? obj.visuals : [];
    visuals.forEach(vs => {
      const id = String(vs.id || `vis${state.nextId++}`);
      const v = {
        id,
        type: vs.type,
        title: vs.title || defaultTitleFor(vs.type),
        x: clamp(Number(vs.x)||40, 0, CANVAS_W-220),
        y: clamp(Number(vs.y)||40, 0, CANVAS_H-170),
        w: clamp(Number(vs.w)||DEFAULT_SIZES[vs.type]?.w||520, 220, CANVAS_W),
        h: clamp(Number(vs.h)||DEFAULT_SIZES[vs.type]?.h||300, 170, CANVAS_H),
        series: Array.isArray(vs.series) ? vs.series.map((s,i)=>({
          key: s.key || s.name || `S${i+1}`,
          name: s.name || s.key || `Series ${i+1}`,
          color: s.color || paletteColor(i),
          overrideColor: !!s.overrideColor
        })) : buildDefaultSeries(vs.type),
        imageDataUrl: vs.type === "image" ? (vs.imageDataUrl || "") : undefined,
        chart: null
      };

      // ensure series colors for non-overrides align with current theme palette
      v.series?.forEach((s,i)=>{
        if (!s.overrideColor) s.color = paletteColor(i);
      });

      state.visuals.set(id, v);
      state.order.push(id);

      const el = createVisualElement(v);
      canvas.appendChild(el);
      renderVisualContent(v);
    });

    updateZOrder();
    setSelected(null);

    // Ensure nextId doesn't collide
    state.nextId = Math.max(state.nextId, state.order.length + 1);
  }

  function clearAllVisuals(){
    // destroy charts + clear DOM
    state.order.forEach(id => {
      const v = state.visuals.get(id);
      if (v?.chart) {
        try { v.chart.destroy(); } catch {}
      }
    });
    state.visuals.clear();
    state.order = [];
    state.selectedId = null;

    // remove visual elements
    Array.from(canvas.querySelectorAll(".visual")).forEach(el => el.remove());
  }

  // -------------------------
  // Sample downloads
  // -------------------------
  downloadSampleThemeBtn.addEventListener("click", () => {
    downloadObjectAsJson(sampleTheme, "sample-theme.json");
    showToast("Sample theme downloaded");
  });

  downloadSampleCanvasBgBtn.addEventListener("click", () => {
    downloadObjectAsJson(sampleCanvasBg, "sample-canvas-background.json");
    showToast("Sample canvas background downloaded");
  });

  // -------------------------
  // Safety: Escape helpers
  // -------------------------
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
  // Start: format pane should always show canvas settings
  // -------------------------
  renderFormatPane();

  // -------------------------
  // Theme import successful popup matches requirement
  // (handled via showToast)
  // -------------------------

  // -------------------------
  // OPTIONAL: Load a couple of starter visuals for demo (comment out if you want empty start)
  // -------------------------
  // addVisual("line");
  // addVisual("donut");

})();