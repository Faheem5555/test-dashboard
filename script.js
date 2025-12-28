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
  // ✅ HARD FIX: ensure discard modal is NOT visible on page load
  // -------------------------
  if (modalBackdrop) {
    modalBackdrop.hidden = true;
    modalBackdrop.style.display = "none";
    modalBackdrop.setAttribute("aria-hidden", "true");
  }
  if (uploadDashboardInput) uploadDashboardInput.value = "";

  // -------------------------
  // App state
  // -------------------------
  const state = {
    visuals: new Map(), // id -> visual object
    order: [],          // z-order
    selectedId: null,
    nextId: 1
  };

  // -------------------------
  // Helpers
  // -------------------------
  function clamp(v, min, max) {
    return Math.max(min, Math.min(max, v));
  }

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

  // -------------------------
  // Canvas size (Dynamic) + 16:9 default
  // -------------------------
  let canvasW = 1280;
  let canvasH = 720;
  let canvasScale = 1;

  // Wrap canvas in a scaler element so we can fit-to-screen without layout overflow.
  // (No HTML rewrite required; we do it safely at runtime.)
  let canvasScaler = null;
  if (canvas && canvas.parentElement) {
    canvasScaler = document.getElementById("canvasScaler");
    if (!canvasScaler) {
      canvasScaler = document.createElement("div");
      canvasScaler.id = "canvasScaler";
      canvasScaler.className = "canvasScaler";
      canvas.parentElement.insertBefore(canvasScaler, canvas);
      canvasScaler.appendChild(canvas);
    }
  }

  function setCanvasCssVars(){
    document.documentElement.style.setProperty("--canvasW", `${canvasW}px`);
    document.documentElement.style.setProperty("--canvasH", `${canvasH}px`);
  }

  function updateStatusPill(){
    if (statusPill) statusPill.textContent = `Canvas: ${canvasW} × ${canvasH}`;
  }

  function computeFitScale(){
    // Fit only inside the center canvas frame viewport (Power BI-like "fit to screen").
    const frame = document.querySelector(".canvasFrame");
    if (!frame) return 1;

    const r = frame.getBoundingClientRect();
    const pad = 6; // small breathing room
    const availW = Math.max(1, r.width - pad);
    const availH = Math.max(1, r.height - pad);

    return Math.min(1, availW / canvasW, availH / canvasH);
  }

  function applyCanvasScale(){
    canvasScale = computeFitScale();

    // Layout box must also shrink: use the scaler wrapper dimensions.
    if (canvasScaler) {
      canvasScaler.style.width = `${Math.round(canvasW * canvasScale)}px`;
      canvasScaler.style.height = `${Math.round(canvasH * canvasScale)}px`;
    }

    if (canvas) {
      canvas.style.transformOrigin = "top left";
      canvas.style.transform = `scale(${canvasScale})`;
      canvas.style.width = `${canvasW}px`;
      canvas.style.height = `${canvasH}px`;
    }
  }

  function clampAllVisualsToCanvas(){
    state.order.forEach((id) => {
      const v = state.visuals.get(id);
      if (!v) return;

      const minW = 220;
      const minH = 170;

      v.w = clamp(v.w, minW, canvasW);
      v.h = clamp(v.h, minH, canvasH);

      v.x = clamp(v.x, 0, canvasW - v.w);
      v.y = clamp(v.y, 0, canvasH - v.h);

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
    });
  }

  function setCanvasSize(w, h, { fromPreset = false } = {}){
    const nw = clamp(Number(w) || 1280, 320, 4000);
    const nh = clamp(Number(h) || 720, 180, 3000);

    canvasW = Math.round(nw);
    canvasH = Math.round(nh);

    setCanvasCssVars();
    updateStatusPill();
    clampAllVisualsToCanvas();
    applyCanvasScale();

    // Keep format pane aligned and not floating outside
    renderFormatPane();

    if (fromPreset) showToast(`Canvas set to ${canvasW} × ${canvasH}`);
  }

  // Initialize default 16:9 canvas
  setCanvasSize(1280, 720);

  window.addEventListener("resize", () => {
    applyCanvasScale();
  });

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

  // ✅ True default state: no visuals AND default theme AND default canvas background AND default canvas size
  function isDefaultState(){
    return state.order.length === 0 &&
      JSON.stringify(theme) === JSON.stringify(defaultTheme) &&
      JSON.stringify(canvasBg) === JSON.stringify(sampleCanvasBg) &&
      canvasW === 1280 &&
      canvasH === 720;
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

  applyCanvasBackground();

  // Theme → Chart.js common options
  function makeChartDefaults(){
    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";

    // ✅ Power BI clean style request:
    // - No background grid lines
    // - No X/Y axis lines
    // - Keep only ticks/labels + data
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 300 },
      plugins: {
        legend: {
          display: true,
          labels: { color: labelColor, boxWidth: 10, usePointStyle: true }
        },
        title: { display: false },
        tooltip: { enabled: true }
      },
      scales: {
        x: {
          ticks: { color: labelColor, display: true },
          grid: { display: false },
          border: { display: false }
        },
        y: {
          ticks: { color: labelColor, display: true },
          grid: { display: false },
          border: { display: false }
        }
      }
    };
  }

  // -------------------------
  // Chart data labels plugin (values only)
  // -------------------------
  const PBI_VALUE_LABELS_PLUGIN = {
    id: "pbiValueLabels",
    afterDatasetsDraw(chart, args, pluginOptions) {
      const ctx = chart.ctx;
      const type = chart.config.type;
      const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.92)";
      const fontSize = 10;

      ctx.save();
      ctx.font = `${fontSize}px Segoe UI, Arial`;
      ctx.fillStyle = labelColor;
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";

      const format = (v) => {
        if (v === null || v === undefined) return "";
        if (typeof v === "number") return v.toLocaleString();
        if (typeof v === "string") return v;
        return String(v);
      };

      const drawText = (text, x, y) => {
        if (!text) return;
        ctx.fillText(text, x, y);
      };

      // Pie/Donut
      if (type === "pie" || type === "doughnut") {
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          if (!meta || meta.hidden) return;
          meta.data.forEach((arc, i) => {
            const v = ds.data?.[i];
            const p = arc.getCenterPoint ? arc.getCenterPoint() : arc.tooltipPosition();
            drawText(format(v), p.x, p.y - 4);
          });
        });
        ctx.restore();
        return;
      }

      // Other cartesian charts
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (!meta || meta.hidden) return;
        meta.data.forEach((pt, i) => {
          const raw = ds.data?.[i];
          const v = (raw && typeof raw === "object" && ("y" in raw)) ? raw.y : raw;
          const p = pt.tooltipPosition ? pt.tooltipPosition() : (pt.getCenterPoint ? pt.getCenterPoint() : null);
          if (!p) return;
          const offset = (type === "bar" || (ds.type === "bar")) ? 6 : 4;
          drawText(format(v), p.x, p.y - offset);
        });
      });

      ctx.restore();
    }
  };

  if (typeof Chart !== "undefined" && Chart?.register) {
    try { Chart.register(PBI_VALUE_LABELS_PLUGIN); } catch {}
  }

  function paletteColor(i){
    const arr = theme?.dataColors?.length ? theme.dataColors : defaultTheme.dataColors;
    return arr[i % arr.length];
  }

  // -------------------------
  // Visual registry (picker)
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
      category: "Other",
      items: [
        { type: "kpi", name: "KPI", icon: "⦿" },
        { type: "card", name: "Card", icon: "▭" },
        { type: "multirowCard", name: "Multi-row Card", icon: "≡" },
        { type: "textBox", name: "Text Box", icon: "T" },
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

    kpi: { w: 320, h: 170 },
    card: { w: 320, h: 170 },
    multirowCard: { w: 360, h: 220 },
    textBox: { w: 360, h: 200 },

    image: { w: 420, h: 260 }
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
  // Selection behavior (Hard rules)
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

    // Bring selected to front
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
      chart: null
    };

    // Defaults for non-chart visuals
    if (type === "kpi") {
      v.kpi = { label: "Sales", value: "" };
    }
    if (type === "card") {
      v.card = { label: "Orders", value: "" };
    }
    if (type === "multirowCard") {
      v.multirow = { rows: null };
    }
    if (type === "textBox") {
      v.textBox = {
        text: "Double-click to edit",
        fontSize: 18,
        color: theme?.textClasses?.title?.color || "#e8ecf2",
        bg: "rgba(0,0,0,0.15)",
        bold: true,
        align: "left"
      };
    }

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
      scatter: "Profit vs Sales (Scatter)",
      kpi: "KPI",
      card: "Card",
      multirowCard: "Multi-row Card",
      textBox: "Text Box"
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

    // Cards / KPI / Text Box don't have series colors
    if (["kpi","card","multirowCard","textBox"].includes(type)) {
      return [];
    }

    const regionKeys = Object.keys(retail.Regions);
    const singleSeries = [{ key: "Value", name: "Value", color: paletteColor(0), overrideColor: false }];
    const multiSeries = regionKeys.map((k, i) => ({ key: k, name: k, color: paletteColor(i), overrideColor: false }));

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

    el.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      setSelected(v.id);
    });

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
    if (e.button !== 0) return;

    const v = state.visuals.get(id);
    if (!v) return;

    setSelected(id);

    // allow text editing without accidental drag if clicked inside textbox editor
    if (v.type === "textBox" && e.target && e.target.classList && e.target.classList.contains("textBoxEditor")) {
      return;
    }

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

    // ✅ correct mouse delta when canvas is scaled to fit screen
    const dx = (e.clientX - dragCtx.startX) / (canvasScale || 1);
    const dy = (e.clientY - dragCtx.startY) / (canvasScale || 1);

    if (dragCtx.mode === "drag") {
      v.x = clamp(dragCtx.origX + dx, 0, canvasW - v.w);
      v.y = clamp(dragCtx.origY + dy, 0, canvasH - v.h);
      applyVisualRect(v);
    } else {
      const minW = 220;
      const minH = 170;

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

  // -------------------------
  // Render visual content (Chart.js + prototypes)
  // -------------------------
  function renderVisualContent(v){
    const body = document.getElementById(`body_${v.id}`);
    if (!body) return;

    if (v.chart) {
      try { v.chart.destroy(); } catch {}
      v.chart = null;
    }
    body.innerHTML = "";

    if (v.type === "kpi") {
      body.appendChild(makeKpiVisual(v));
      return;
    }

    if (v.type === "card") {
      body.appendChild(makeCardVisual(v));
      return;
    }

    if (v.type === "multirowCard") {
      body.appendChild(makeMultirowCardVisual(v));
      return;
    }

    if (v.type === "textBox") {
      body.appendChild(makeTextBoxVisual(v));
      return;
    }

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
            options: (() => {
              const o = makeChartDefaults();
              // titles allowed (labels), but no grid/axis borders
              o.scales.x.title = { display:true, text:"Sales", color: (theme?.textClasses?.label?.color || "#cfd6e1") };
              o.scales.y.title = { display:true, text:"Profit", color: (theme?.textClasses?.label?.color || "#cfd6e1") };
              o.scales.x.grid.display = false;
              o.scales.y.grid.display = false;
              o.scales.x.border.display = false;
              o.scales.y.border.display = false;
              return o;
            })()
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

  // Treemap prototype
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

  // Ribbon prototype (SVG) — grid removed for clean style
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

    // (Grid removed for clean Power BI style)

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

  // -------------------------
  // KPI / Card / Multi-row Card (prototype visuals)
  // -------------------------
  function makeKpiVisual(v){
    const wrap = document.createElement("div");
    wrap.className = "cardLike";

    const value = Math.round(retail.Sales[retail.Sales.length-1] || 0);
    const prev  = Math.round(retail.Sales[retail.Sales.length-2] || value);
    const delta = prev ? (((value - prev) / prev) * 100) : 0;

    wrap.innerHTML = `
      <div class="kpiLabel">${escapeHtml(v.kpi?.label || "Sales")}</div>
      <div class="kpiValue">${escapeHtml(v.kpi?.value || value.toLocaleString())}</div>
      <div class="kpiTrend">
        <span class="kpiArrow">${delta >= 0 ? "▲" : "▼"}</span>
        <span>${Math.abs(delta).toFixed(1)}%</span>
        <span class="kpiHint">vs last period</span>
      </div>
    `;
    return wrap;
  }

  function makeCardVisual(v){
    const wrap = document.createElement("div");
    wrap.className = "cardLike";
    const value = Math.round(retail.Orders[retail.Orders.length-1] || 0);
    wrap.innerHTML = `
      <div class="kpiLabel">${escapeHtml(v.card?.label || "Orders")}</div>
      <div class="kpiValue">${escapeHtml(v.card?.value || value.toLocaleString())}</div>
      <div class="kpiHint">Single value</div>
    `;
    return wrap;
  }

  function makeMultirowCardVisual(v){
    const wrap = document.createElement("div");
    wrap.className = "multirowCard";

    const rows = v.multirow?.rows || [
      { label: "Sales", value: (retail.Sales[11] || 0).toLocaleString() },
      { label: "Profit", value: (retail.Profit[11] || 0).toLocaleString() },
      { label: "Orders", value: (retail.Orders[11] || 0).toLocaleString() }
    ];

    wrap.innerHTML = rows.map(r => `
      <div class="mRow">
        <div class="mLabel">${escapeHtml(r.label)}</div>
        <div class="mValue">${escapeHtml(String(r.value))}</div>
      </div>
    `).join("");

    return wrap;
  }

  // -------------------------
  // Text Box visual (editable)
  // -------------------------
  function ensureTextBoxState(v){
    if (!v.textBox) {
      v.textBox = {
        text: "Double-click to edit",
        fontSize: 18,
        color: theme?.textClasses?.title?.color || "#e8ecf2",
        bg: "rgba(0,0,0,0.15)",
        bold: true,
        align: "left"
      };
    }
  }

  function applyTextBoxStyles(el, v){
    ensureTextBoxState(v);
    const t = v.textBox;

    el.style.fontSize = `${clamp(Number(t.fontSize)||18, 8, 120)}px`;
    el.style.color = t.color || "#e8ecf2";
    el.style.background = t.bg || "transparent";
    el.style.fontWeight = t.bold ? "800" : "500";
    el.style.textAlign = t.align || "left";
  }

  function makeTextBoxVisual(v){
    ensureTextBoxState(v);

    const host = document.createElement("div");
    host.className = "textBoxHost";

    const editor = document.createElement("div");
    editor.className = "textBoxEditor";
    editor.setAttribute("contenteditable", "true");
    editor.setAttribute("spellcheck", "false");
    editor.innerText = v.textBox.text || "";

    applyTextBoxStyles(editor, v);

    editor.addEventListener("input", () => {
      v.textBox.text = editor.innerText;
    });

    editor.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      setSelected(v.id);
    });

    host.appendChild(editor);
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
    if (v.type === "treemap" || v.type === "ribbon") {
      renderVisualContent(v);
      return;
    }
    if (!v.chart) return;

    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";

    if (v.chart.options?.plugins?.legend?.labels) {
      v.chart.options.plugins.legend.labels.color = labelColor;
    }

    // ✅ keep labels/ticks, but remove grid + axis borders
    if (v.chart.options?.scales?.x) {
      if (v.chart.options.scales.x.ticks) v.chart.options.scales.x.ticks.color = labelColor;
      v.chart.options.scales.x.grid = { ...(v.chart.options.scales.x.grid || {}), display: false };
      v.chart.options.scales.x.border = { ...(v.chart.options.scales.x.border || {}), display: false };
    }
    if (v.chart.options?.scales?.y) {
      if (v.chart.options.scales.y.ticks) v.chart.options.scales.y.ticks.color = labelColor;
      v.chart.options.scales.y.grid = { ...(v.chart.options.scales.y.grid || {}), display: false };
      v.chart.options.scales.y.border = { ...(v.chart.options.scales.y.border || {}), display: false };
    }

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

      if (v.type === "textBox") {
        // keep textbox readable with theme
        if (v.textBox && !v.textBox.color) v.textBox.color = theme?.textClasses?.title?.color || "#e8ecf2";
        renderVisualContent(v);
      }
    });

    renderFormatPane();
  }

  // -------------------------
  // Format pane (always visible; never disappears)
  // -------------------------
  function renderFormatPane(){
    const selected = state.selectedId ? state.visuals.get(state.selectedId) : null;

    if (!selected) {
      formatBody.innerHTML = `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Canvas settings</div>
            <div class="fSmall">No visual selected</div>
          </div>

          <div class="row one">
            <div class="field">
              <div class="label">Canvas size preset</div>
              <select class="select" id="canvasPreset">
                <option value="1280x720">1280 × 720 (16:9)</option>
                <option value="1920x1080">1920 × 1080 (16:9)</option>
                <option value="custom">Custom</option>
              </select>
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Canvas width</div>
              <input class="input" id="canvasWInput" value="${escapeAttr(String(canvasW))}" />
            </div>
            <div class="field">
              <div class="label">Canvas height</div>
              <input class="input" id="canvasHInput" value="${escapeAttr(String(canvasH))}" />
            </div>
          </div>

          <div class="fSmall" style="margin-top:4px;">
            Applies immediately. Visuals are clamped inside the page.
          </div>

          <div class="row" style="margin-top:10px;">
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
            Theme applies to chart palette + labels instantly (existing + new visuals).
          </div>
        </div>
      `;

      // preset selection (set correctly based on current size)
      const preset = document.getElementById("canvasPreset");
      if (preset) {
        const cur = `${canvasW}x${canvasH}`;
        if (cur === "1280x720") preset.value = "1280x720";
        else if (cur === "1920x1080") preset.value = "1920x1080";
        else preset.value = "custom";
      }

      const wInput = document.getElementById("canvasWInput");
      const hInput = document.getElementById("canvasHInput");

      const applyFromInputs = () => {
        const w = Number(wInput.value);
        const h = Number(hInput.value);
        setCanvasSize(w, h);
      };

      // ✅ applies immediately
      wInput.oninput = () => {
        if (preset) preset.value = "custom";
        applyFromInputs();
      };
      hInput.oninput = () => {
        if (preset) preset.value = "custom";
        applyFromInputs();
      };

      if (preset) {
        preset.onchange = () => {
          if (preset.value === "1280x720") {
            wInput.value = "1280";
            hInput.value = "720";
            setCanvasSize(1280, 720, { fromPreset: true });
          } else if (preset.value === "1920x1080") {
            wInput.value = "1920";
            hInput.value = "1080";
            setCanvasSize(1920, 1080, { fromPreset: true });
          }
        };
      }

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

    const textBoxHtml = (v.type === "textBox")
      ? `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Text Box</div>
            <div class="fSmall">Formatting</div>
          </div>

          <div class="row one">
            <div class="field">
              <div class="label">Text</div>
              <textarea class="input" id="tbText" rows="4" style="resize:vertical;">${escapeHtml(v.textBox?.text || "")}</textarea>
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Font size</div>
              <input class="input" id="tbSize" value="${escapeAttr(String(v.textBox?.fontSize ?? 18))}" />
            </div>
            <div class="field">
              <div class="label">Font color</div>
              <input class="colorInput" id="tbColor" type="color" value="${escapeAttr(ensureHex(v.textBox?.color || "#e8ecf2"))}" />
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Background color</div>
              <input class="input" id="tbBg" value="${escapeAttr(String(v.textBox?.bg || "rgba(0,0,0,0.15)"))}" />
            </div>
            <div class="field">
              <div class="label">Alignment</div>
              <div class="miniBtns">
                <button class="miniBtn" id="tbAlignL">Left</button>
                <button class="miniBtn" id="tbAlignC">Center</button>
                <button class="miniBtn" id="tbAlignR">Right</button>
              </div>
            </div>
          </div>

          <div class="miniBtns">
            <button class="miniBtn" id="tbBold">Bold</button>
          </div>

          <div class="fSmall" style="margin-top:8px;">
            Tip: You can also edit directly in the visual.
          </div>
        </div>
      `
      : "";

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
      : (v.type === "kpi" || v.type === "card" || v.type === "multirowCard" || v.type === "textBox")
        ? ``
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

    formatBody.innerHTML = posHtml + textBoxHtml + colorsHtml;

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

      if (isFinite(nx)) v.x = clamp(nx, 0, canvasW - v.w);
      if (isFinite(ny)) v.y = clamp(ny, 0, canvasH - v.h);
      if (isFinite(nw)) v.w = clamp(nw, 220, canvasW - v.x);
      if (isFinite(nh)) v.h = clamp(nh, 170, canvasH - v.y);

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
      renderFormatPane();
    };

    document.getElementById("deleteBtn").onclick = () => removeVisual(v.id);

    // Text Box controls
    if (v.type === "textBox") {
      ensureTextBoxState(v);

      const tbText = document.getElementById("tbText");
      const tbSize = document.getElementById("tbSize");
      const tbColor = document.getElementById("tbColor");
      const tbBg = document.getElementById("tbBg");

      const refreshTextBox = () => {
        const el = document.querySelector(`#body_${v.id} .textBoxEditor`);
        if (el) applyTextBoxStyles(el, v);
      };

      tbText.oninput = () => {
        v.textBox.text = tbText.value;
        const el = document.querySelector(`#body_${v.id} .textBoxEditor`);
        if (el) el.innerText = v.textBox.text;
      };

      tbSize.oninput = () => {
        const n = Number(tbSize.value);
        if (isFinite(n)) v.textBox.fontSize = clamp(n, 8, 120);
        refreshTextBox();
      };

      tbColor.oninput = () => {
        v.textBox.color = tbColor.value;
        refreshTextBox();
      };

      tbBg.oninput = () => {
        v.textBox.bg = tbBg.value;
        refreshTextBox();
      };

      const boldBtn = document.getElementById("tbBold");
      const setBoldActive = () => {
        if (!boldBtn) return;
        boldBtn.classList.toggle("active", !!v.textBox.bold);
      };
      boldBtn.onclick = () => {
        v.textBox.bold = !v.textBox.bold;
        setBoldActive();
        refreshTextBox();
      };
      setBoldActive();

      const alL = document.getElementById("tbAlignL");
      const alC = document.getElementById("tbAlignC");
      const alR = document.getElementById("tbAlignR");

      const setAlignActive = () => {
        alL.classList.toggle("active", v.textBox.align === "left");
        alC.classList.toggle("active", v.textBox.align === "center");
        alR.classList.toggle("active", v.textBox.align === "right");
      };

      alL.onclick = () => { v.textBox.align = "left"; setAlignActive(); refreshTextBox(); };
      alC.onclick = () => { v.textBox.align = "center"; setAlignActive(); refreshTextBox(); };
      alR.onclick = () => { v.textBox.align = "right"; setAlignActive(); refreshTextBox(); };

      setAlignActive();
    }

    if (v.type !== "image" && v.type !== "kpi" && v.type !== "card" && v.type !== "multirowCard" && v.type !== "textBox") {
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
  // Theme import (Power BI-like)
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
  // Download/Upload dashboard (with discard confirm)
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
        kpi: v.type === "kpi" ? (v.kpi || null) : undefined,
        card: v.type === "card" ? (v.card || null) : undefined,
        multirow: v.type === "multirowCard" ? (v.multirow || null) : undefined,
        textBox: v.type === "textBox" ? (v.textBox || null) : undefined
      };
    }).filter(Boolean);

    return {
      version: 2,
      canvas: {
        width: canvasW,
        height: canvasH,
        background: canvasBg
      },
      theme,
      visuals
    };
  }

  async function loadDashboard(obj){
    clearAllVisuals();

    // canvas size (must load first so clamping uses correct page bounds)
    const w = Number(obj?.canvas?.width || 1280);
    const h = Number(obj?.canvas?.height || 720);
    setCanvasSize(w, h);

    if (obj?.canvas?.background) {
      canvasBg = normalizeCanvasBg(obj.canvas.background);
    } else {
      canvasBg = structuredClone(sampleCanvasBg);
    }
    applyCanvasBackground();

    theme = normalizeTheme(obj?.theme || defaultTheme);
    applyThemeEverywhere();

    const visuals = Array.isArray(obj?.visuals) ? obj.visuals : [];
    visuals.forEach(vs => {
      const id = String(vs.id || `vis${state.nextId++}`);
      const v = {
        id,
        type: vs.type,
        title: vs.title || defaultTitleFor(vs.type),
        x: clamp(Number(vs.x)||40, 0, canvasW-220),
        y: clamp(Number(vs.y)||40, 0, canvasH-170),
        w: clamp(Number(vs.w)||DEFAULT_SIZES[vs.type]?.w||520, 220, canvasW),
        h: clamp(Number(vs.h)||DEFAULT_SIZES[vs.type]?.h||300, 170, canvasH),
        series: Array.isArray(vs.series)
          ? vs.series.map((s,i)=>({
              key: s.key || s.name || `S${i+1}`,
              name: s.name || s.key || `Series ${i+1}`,
              color: s.color || paletteColor(i),
              overrideColor: !!s.overrideColor
            }))
          : buildDefaultSeries(vs.type),
        imageDataUrl: vs.type === "image" ? (vs.imageDataUrl || "") : undefined,
        kpi: vs.type === "kpi" ? (vs.kpi || null) : undefined,
        card: vs.type === "card" ? (vs.card || null) : undefined,
        multirow: vs.type === "multirowCard" ? (vs.multirow || null) : undefined,
        textBox: vs.type === "textBox" ? (vs.textBox || null) : undefined,
        chart: null
      };

      if (v.type === "textBox") ensureTextBoxState(v);

      v.series?.forEach((s,i)=> {
        if (!s.overrideColor) s.color = paletteColor(i);
      });

      if (v.type === "kpi" && !v.kpi) v.kpi = { label: "Sales", value: "" };
      if (v.type === "card" && !v.card) v.card = { label: "Orders", value: "" };
      if (v.type === "multirowCard" && !v.multirow) v.multirow = { rows: null };

      state.visuals.set(id, v);
      state.order.push(id);

      const el = createVisualElement(v);
      canvas.appendChild(el);
      renderVisualContent(v);
    });

    updateZOrder();
    setSelected(null);
    state.nextId = Math.max(state.nextId, state.order.length + 1);
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
  // IMPORTANT: Empty dashboard on load (no starter visuals)
  // -------------------------
  // addVisual("line");
  // addVisual("donut");

})();
