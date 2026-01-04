/* =========================
   Power BI Prototype Dashboard Website
   FULL IMPLEMENTATION (no skips)
   Incremental updates added:
   - Chart grid/axis lines removed; show data labels values
   - Canvas size dropdown and editable width/height (16:9 default)
   - KPI / Card / Multi-row Card visuals added
   - Text Box visual added with editor controls
   - Boundary enforcement fixed so visuals never cross canvas edges
   - Pane alignment with canvas and screen-fit improvements
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

  const leftPane = document.querySelector(".leftPane");
  const rightPane = document.querySelector(".rightPane");
  const canvasFrame = document.querySelector(".canvasFrame");

  // -------------------------
  // ✅ HARD FIX: ensure discard modal is NOT visible on page load
  // -------------------------
  if (modalBackdrop) {
    modalBackdrop.hidden = true;
    modalBackdrop.style.display = "none";
    modalBackdrop.setAttribute("aria-hidden", "true");
  }

  // Also ensure upload input doesn't carry a cached value causing change events in some browsers
  if (uploadDashboardInput) uploadDashboardInput.value = "";

  // -------------------------
  // Fixed canvas size (Hard rule) -> now dynamic and editable
  // Default is 1280 x 720 (16:9)
  // -------------------------
  let CANVAS_W = 1280;
  let CANVAS_H = 720;

  function setCanvasCssSize(w, h) {
    document.documentElement.style.setProperty("--canvasW", `${w}px`);
    document.documentElement.style.setProperty("--canvasH", `${h}px`);
    // Update status pill text
    if (statusPill) statusPill.textContent = `Canvas: ${w} × ${h}`;
    // After canvas size changes, ensure visuals remain inside and panes align
    enforceAllVisualsInsideCanvas();
    alignPanesWithCanvas();
  }

  setCanvasCssSize(CANVAS_W, CANVAS_H);

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
  const sampleCanvasBg = {
    backgroundColor: "#0b0d12",
    backgroundImage: "",
    opacity: 1
  };

  let theme = structuredClone(defaultTheme);
  let canvasBg = structuredClone(sampleCanvasBg);

  // Apply initial canvas bg + CSS vars
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

  // -------------------------
  // Visual registry (picker)
  // Add KPI, Card, Multi-row Card, Textbox
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
        { type: "kpi", name: "KPI (single metric)", icon: "📈" },
        { type: "card", name: "Card visual", icon: "🃏" },
        { type: "multirowcard", name: "Multi-row Card", icon: "📋" },
        { type: "textbox", name: "Text Box", icon: "✎" }
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

    image: { w: 420, h: 260 },

    // new visuals
    kpi: { w: 300, h: 120 },
    card: { w: 300, h: 140 },
    multirowcard: { w: 340, h: 220 },
    textbox: { w: 520, h: 120 }
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

  // ✅ True default state: no visuals AND default theme AND default canvas background
  function isDefaultState(){
    return state.order.length === 0 &&
      JSON.stringify(theme) === JSON.stringify(defaultTheme) &&
      JSON.stringify(canvasBg) === JSON.stringify(sampleCanvasBg);
  }

  // -------------------------
  // Theme → Chart.js common options
  // Remove grid & axis lines; show only data labels (via plugin)
  // -------------------------
  function makeChartDefaults(){
    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";
    // gridColor not used (we hide grid)
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
        // datalabels handled via global plugin registered below
      },
      scales: {
        x: {
          ticks: { display: false, color: labelColor },
          grid: { display: false, drawBorder: false }
        },
        y: {
          ticks: { display: false, color: labelColor },
          grid: { display: false, drawBorder: false }
        }
      },
      elements: {
        point: { radius: 3 }
      }
    };
  }

  // Generate palette color for series index
  function paletteColor(i){
    const arr = theme?.dataColors?.length ? theme.dataColors : defaultTheme.dataColors;
    return arr[i % arr.length];
  }

  // -------------------------
  // Chart.js plugin: draw data labels (values)
  // Works for bars, lines (points), arcs (pie/donut), scatter
  // -------------------------
  Chart.register({
    id: 'powerbi_datalabels',
    afterDatasetsDraw(chart) {
      const ctx = chart.ctx;
      const { datasets } = chart.data;
      const fontColor = theme?.textClasses?.label?.color || "#cfd6e1";
      const fontSize = (theme?.textClasses?.label?.fontSize) ? Number(theme.textClasses.label.fontSize) + 1 : 11;
      ctx.save();
      ctx.font = `600 ${fontSize}px "Segoe UI", Arial, sans-serif`;
      ctx.fillStyle = fontColor;
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";

      datasets.forEach((dataset, dsIndex) => {
        const meta = chart.getDatasetMeta(dsIndex);
        if (!meta || !meta.data) return;

        meta.data.forEach((el, i) => {
          try {
            const value = dataset.data?.[i];
            if (value === undefined || value === null) return;
            const disp = String(value);

            const type = meta.type || chart.config.type || dataset.type;

            if (type === 'bar' || type === 'bar') {
              // For bar stacks the element has x, y
              const x = el.x ?? ((el.getCenterPoint && el.getCenterPoint())?.x);
              const y = el.y ?? (el.y ?? (el.getCenterPoint && el.getCenterPoint())?.y);
              if (x === undefined || y === undefined) return;
              // draw slightly above the top of the bar
              ctx.fillText(disp, x, (y - 6));
            } else if (type === 'line' || type === 'scatter') {
              const x = el.x ?? (el.getCenterPoint && el.getCenterPoint()?.x);
              const y = el.y ?? (el.getCenterPoint && el.getCenterPoint()?.y);
              if (x === undefined || y === undefined) return;
              ctx.fillText(disp, x, (y - 8));
            } else if (type === 'pie' || type === 'doughnut' || el?.startAngle !== undefined) {
              // arc element
              const cx = el.x ?? chart.chartArea?.left + (chart.chartArea?.width / 2);
              const cy = el.y ?? chart.chartArea?.top + (chart.chartArea?.height / 2);
              const start = el.startAngle ?? 0;
              const end = el.endAngle ?? 0;
              const angle = (start + end) / 2;
              const r = (el.outerRadius || Math.min(chart.chartArea.width, chart.chartArea.height) / 3) * 0.6;
              const tx = cx + Math.cos(angle) * r;
              const ty = cy + Math.sin(angle) * r;
              ctx.fillText(disp, tx, ty);
            } else {
              // fallback
              const pos = el.tooltipPosition ? el.tooltipPosition() : (el.getCenterPoint ? el.getCenterPoint() : null);
              if (pos) ctx.fillText(disp, pos.x, pos.y - 6);
            }
          } catch (e) {
            // ignore plugin drawing error
          }
        });
      });

      ctx.restore();
    }
  });

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
          else if (item.type === "textbox") addTextBoxVisual();
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

    // Additional default props for new visuals
    if (type === "kpi") {
      v.kpiValue = 12456;
      v.kpiLabel = "Total Sales";
    }
    if (type === "card") {
      v.cardValue = 832;
      v.cardLabel = "Active Users";
    }
    if (type === "multirowcard") {
      v.rows = [
        {label:"Sales", value: 12456},
        {label:"Profit", value: 2345},
        {label:"Orders", value: 395}
      ];
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

  function addTextBoxVisual(){
    const id = `vis${state.nextId++}`;
    const size = DEFAULT_SIZES.textbox;
    const offset = (state.order.length * 18) % 140;
    const x = clamp(40 + offset, 0, CANVAS_W - size.w);
    const y = clamp(40 + offset, 0, CANVAS_H - size.h);

    const v = {
      id,
      type: "textbox",
      title: "Text Box",
      x, y, w: size.w, h: size.h,
      series: [],
      text: "Editable text",
      fontSize: 16,
      color: "#e8ecf2",
      bgColor: "transparent",
      bold: false,
      align: "left",
      chart: null
    };

    state.visuals.set(id, v);
    state.order.push(id);

    const el = createVisualElement(v);
    canvas.appendChild(el);
    updateZOrder();

    renderVisualContent(v);
    setSelected(id);
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
      multirowcard: "Multi-row Card",
      textbox: "Text Box"
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
  // Enforce canvas boundaries using dynamic CANVAS_W/CANVAS_H
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
      v.x = clamp(dragCtx.origX + dx, 0, Math.max(0, CANVAS_W - v.w));
      v.y = clamp(dragCtx.origY + dy, 0, Math.max(0, CANVAS_H - v.h));
      applyVisualRect(v);
    } else {
      const minW = 220;
      const minH = 170;

      let x = dragCtx.origX;
      let y = dragCtx.origY;
      let w = dragCtx.origW;
      let h = dragCtx.origH;

      const hnd = dragCtx.handle;

      const applyW = (newW) => clamp(newW, minW, Math.max(minW, CANVAS_W - x));
      const applyH = (newH) => clamp(newH, minH, Math.max(minH, CANVAS_H - y));

      if (hnd.includes("e")) w = applyW(dragCtx.origW + dx);
      if (hnd.includes("s")) h = applyH(dragCtx.origH + dy);

      if (hnd.includes("w")) {
        const newX = clamp(dragCtx.origX + dx, 0, dragCtx.origX + dragCtx.origW - minW);
        const newW = dragCtx.origW + (dragCtx.origX - newX);
        x = newX;
        w = clamp(newW, minW, Math.max(minW, CANVAS_W - x));
      }

      if (hnd.includes("n")) {
        const newY = clamp(dragCtx.origY + dy, 0, dragCtx.origY + dragCtx.origH - minH);
        const newH = dragCtx.origH + (dragCtx.origY - newY);
        y = newY;
        h = clamp(newH, minH, Math.max(minH, CANVAS_H - y));
      }

      // final clamp ensuring object fits into canvas
      w = clamp(w, minW, Math.max(minW, CANVAS_W - x));
      h = clamp(h, minH, Math.max(minH, CANVAS_H - y));

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

    // enforce visual inside canvas (safety)
    v.w = clamp(v.w, 40, Math.max(40, CANVAS_W));
    v.h = clamp(v.h, 24, Math.max(24, CANVAS_H));
    v.x = clamp(v.x, 0, Math.max(0, CANVAS_W - v.w));
    v.y = clamp(v.y, 0, Math.max(0, CANVAS_H - v.h));

    el.style.left = `${v.x}px`;
    el.style.top  = `${v.y}px`;
    el.style.width = `${v.w}px`;
    el.style.height = `${v.h}px`;
  }

  // Ensures all visuals inside canvas (used when canvas size changes or load)
  function enforceAllVisualsInsideCanvas(){
    state.order.forEach(id => {
      const v = state.visuals.get(id);
      if (!v) return;
      v.w = clamp(v.w, 40, Math.max(40, CANVAS_W));
      v.h = clamp(v.h, 24, Math.max(24, CANVAS_H));
      v.x = clamp(v.x, 0, Math.max(0, CANVAS_W - v.w));
      v.y = clamp(v.y, 0, Math.max(0, CANVAS_H - v.h));
      applyVisualRect(v);
      if (v.chart) {
        try { v.chart.resize(); } catch {}
      }
    });
  }

  // -------------------------
  // Render visual content (Chart.js + prototypes)
  // - respects removed grid/axis and draws data labels via plugin
  // - also renders KPI, Card, Multi-row Card, Textbox
  // -------------------------
  function renderVisualContent(v){
    const body = document.getElementById(`body_${v.id}`);
    if (!body) return;

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

    // KPI / Card / Multi-row Card prototypes
    if (v.type === "kpi") {
      const host = document.createElement("div");
      host.className = "kpiHost";
      host.innerHTML = `
        <div class="kpiValue">${escapeHtml(String(v.kpiValue ?? 0))}</div>
        <div class="kpiLabel">${escapeHtml(v.kpiLabel || "Metric")}</div>
      `;
      body.appendChild(host);
      return;
    }

    if (v.type === "card") {
      const host = document.createElement("div");
      host.className = "kpiHost";
      host.innerHTML = `
        <div class="kpiValue">${escapeHtml(String(v.cardValue ?? 0))}</div>
        <div class="kpiLabel">${escapeHtml(v.cardLabel || "Card")}</div>
      `;
      body.appendChild(host);
      return;
    }

    if (v.type === "multirowcard") {
      const host = document.createElement("div");
      host.className = "kpiHost";
      host.style.gap = "8px";
      const rowsHtml = (v.rows || []).map(r => `<div style="display:flex;justify-content:space-between;width:100%"><div style="opacity:.85">${escapeHtml(String(r.label))}</div><div style="font-weight:700">${escapeHtml(String(r.value))}</div></div>`).join("");
      host.innerHTML = rowsHtml;
      body.appendChild(host);
      return;
    }

    if (v.type === "textbox") {
      const host = document.createElement("div");
      host.className = "textboxHost";
      const txt = document.createElement("div");
      txt.className = "textboxContent";
      txt.contentEditable = true;
      txt.innerHTML = escapeHtml(v.text || "");
      txt.style.fontSize = `${v.fontSize || 14}px`;
      txt.style.color = v.color || "#e8ecf2";
      txt.style.background = v.bgColor || "transparent";
      txt.style.fontWeight = v.bold ? "700" : "400";
      txt.style.textAlign = v.align || "left";
      txt.style.padding = "6px";
      txt.style.width = "100%";
      txt.style.height = "100%";
      txt.style.overflow = "auto";
      txt.addEventListener("input", (e) => {
        v.text = e.target.innerText;
      });
      host.appendChild(txt);
      body.appendChild(host);
      return;
    }

    // Otherwise, use Chart.js
    const host = document.createElement("div");
    host.className = "chartHost";
    const c = document.createElement("canvas");
    c.id = `c_${v.id}`;
    host.appendChild(c);
    body.appendChild(host);

    const ctx = c.getContext("2d");
    const ds = buildChartJsData(v);

    // Merge a shared plugin options if needed (plugin registered globally)
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
              cutout: v.type === "donut" ? "58%" : "0%",
              plugins: {
                ...makeChartDefaults().plugins,
                legend: { display: true, labels: { color: theme?.textClasses?.label?.color } }
              }
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
          cfgBase.options.plugins.tooltip = cfgBase.options.plugins.tooltip || {};
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
          cfgBase.options.plugins.tooltip = cfgBase.options.plugins.tooltip || {};
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
              // axis titles hidden per requirement; keep minimal axis visuals
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

  // Ribbon prototype (SVG)
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

    const grid = document.createElementNS("http://www.w3.org/2000/svg","path");
    grid.setAttribute("d","M0 210 H600 M0 160 H600 M0 110 H600 M0 60 H600");
    grid.setAttribute("stroke", theme?.chart?.grid || "rgba(255,255,255,0.10)");
    grid.setAttribute("stroke-width","1");
    grid.setAttribute("fill","none");
    svg.appendChild(grid);

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
    if (v.type === "treemap" || v.type === "ribbon") {
      renderVisualContent(v);
      return;
    }
    if (!v.chart) return;

    const labelColor = theme?.textClasses?.label?.color || "rgba(232,236,242,0.80)";
    const gridColor  = theme?.chart?.grid || "rgba(255,255,255,0.08)";

    if (v.chart.options?.plugins?.legend?.labels) {
      v.chart.options.plugins.legend.labels.color = labelColor;
    }
    if (v.chart.options?.scales?.x?.ticks) v.chart.options.scales.x.ticks.color = labelColor;
    if (v.chart.options?.scales?.y?.ticks) v.chart.options.scales.y.ticks.color = labelColor;
    if (v.chart.options?.scales?.x?.grid) v.chart.options.scales.x.grid.color = gridColor;
    if (v.chart.options?.scales?.y?.grid) v.chart.options.scales.y.grid.color = gridColor;

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
  // Format pane (always visible; never disappears)
  // Now includes canvas size dropdown and editable width/height that apply immediately
  // -------------------------
  function renderFormatPane(){
    const selected = state.selectedId ? state.visuals.get(state.selectedId) : null;

    // Canvas settings always shown at top of format pane (per requirement)
    const canvasSectionHtml = `
      <div class="fSection">
        <div class="fHeader">
          <div class="fHeaderTitle">Canvas settings</div>
          <div class="fSmall">No visual selected</div>
        </div>

        <div class="row">
          <div class="field">
            <div class="label">Canvas preset</div>
            <select class="select" id="canvasPresetSelect">
              <option value="1280x720">1280 × 720</option>
              <option value="1920x1080">1920 × 1080</option>
            </select>
          </div>
          <div class="field">
            <div class="label">Aspect ratio</div>
            <div class="fSmall">Default 16:9</div>
          </div>
        </div>

        <div class="row">
          <div class="field">
            <div class="label">Canvas width</div>
            <input class="input" id="canvasBgWidth" value="${escapeAttr(String(CANVAS_W))}" />
          </div>
          <div class="field">
            <div class="label">Canvas height</div>
            <input class="input" id="canvasBgHeight" value="${escapeAttr(String(CANVAS_H))}" />
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

        <div class="smallBtnRow" style="margin-top:8px;">
          <button class="btn" id="applyCanvasBgBtn">Apply</button>
          <button class="btn" id="browseCanvasBgBtn">Browse (Upload Canvas BG JSON)</button>
          <input id="canvasBgUploadInput" type="file" accept=".json,application/json" hidden />
        </div>

        <div class="fSmall" style="margin-top:8px;">
          Upload format supported: <code>{ "backgroundColor": "...", "backgroundImage": "...", "opacity": 0.9 }</code>
        </div>
      </div>
    `;

    // If a visual is selected, show its settings below the canvas section
    if (!selected) {
      const themeHtml = `
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

      formatBody.innerHTML = canvasSectionHtml + themeHtml;

      // Wire up canvas controls
      const preset = document.getElementById("canvasPresetSelect");
      const widthInput = document.getElementById("canvasBgWidth");
      const heightInput = document.getElementById("canvasBgHeight");
      const applyBtn = document.getElementById("applyCanvasBgBtn");
      const browseBtn = document.getElementById("browseCanvasBgBtn");
      const uploadInput = document.getElementById("canvasBgUploadInput");

      // set preset to current
      preset.value = `${CANVAS_W}x${CANVAS_H}`;

      preset.onchange = () => {
        const val = preset.value.split("x").map(Number);
        if (val.length === 2 && isFinite(val[0]) && isFinite(val[1])) {
          CANVAS_W = clamp(val[0], 200, 3840);
          CANVAS_H = clamp(val[1], 100, 2160);
          setCanvasCssSize(CANVAS_W, CANVAS_H);
          widthInput.value = String(CANVAS_W);
          heightInput.value = String(CANVAS_H);
        }
      };

      widthInput.oninput = () => {
        const nw = Number(widthInput.value);
        if (isFinite(nw) && nw > 0) {
          CANVAS_W = clamp(Math.round(nw), 200, 3840);
          setCanvasCssSize(CANVAS_W, CANVAS_H);
          preset.value = `${CANVAS_W}x${CANVAS_H}`;
        }
      };

      heightInput.oninput = () => {
        const nh = Number(heightInput.value);
        if (isFinite(nh) && nh > 0) {
          CANVAS_H = clamp(Math.round(nh), 100, 2160);
          setCanvasCssSize(CANVAS_W, CANVAS_H);
          preset.value = `${CANVAS_W}x${CANVAS_H}`;
        }
      };

      document.getElementById("canvasBgColor").oninput = (e) => {
        canvasBg.backgroundColor = e.target.value || canvasBg.backgroundColor;
      };
      document.getElementById("canvasBgOpacity").oninput = (e) => {
        const v = Number(e.target.value);
        canvasBg.opacity = isFinite(v) ? clamp(v, 0, 1) : canvasBg.opacity;
      };
      document.getElementById("canvasBgImage").oninput = (e) => {
        canvasBg.backgroundImage = e.target.value || "";
      };

      applyBtn.onclick = () => {
        applyCanvasBackground();
        showToast("Canvas background updated");
        // ensure css var size still matches inputs
        setCanvasCssSize(CANVAS_W, CANVAS_H);
      };

      browseBtn.onclick = () => { uploadInput.value=""; uploadInput.click(); };
      uploadInput.onchange = async () => {
        const f = uploadInput.files?.[0];
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

    // Visual settings view (when a visual is selected)
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

    // Editor sections depending on type
    let visualSpecificHtml = "";

    if (v.type === "textbox") {
      visualSpecificHtml = `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Text Box</div>
            <div class="fSmall">Edit content & style</div>
          </div>

          <div class="row one">
            <div class="field">
              <div class="label">Text content</div>
              <textarea id="txtContent" class="input" style="height:100px">${escapeAttr(v.text || "")}</textarea>
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Font size (px)</div>
              <input class="input" id="txtFontSize" value="${escapeAttr(String(v.fontSize || 16))}" />
            </div>
            <div class="field">
              <div class="label">Font color</div>
              <input class="input" id="txtColor" value="${escapeAttr(v.color || "#e8ecf2")}" />
            </div>
          </div>

          <div class="row">
            <div class="field">
              <div class="label">Background</div>
              <input class="input" id="txtBg" value="${escapeAttr(v.bgColor || "transparent")}" />
            </div>
            <div class="field">
              <div class="label">Format</div>
              <select class="select" id="txtAlign">
                <option value="left" ${v.align==="left"?"selected":""}>Left</option>
                <option value="center" ${v.align==="center"?"selected":""}>Center</option>
                <option value="right" ${v.align==="right"?"selected":""}>Right</option>
              </select>
            </div>
          </div>

          <div class="smallBtnRow" style="margin-top:8px;">
            <label style="display:inline-flex;align-items:center;gap:8px"><input type="checkbox" id="txtBold" ${v.bold?"checked":""}/> Bold</label>
            <button class="btn" id="applyTxtBtn">Apply</button>
          </div>
        </div>
      `;
    } else if (v.type === "image") {
      visualSpecificHtml = `
        <div class="fSection">
          <div class="fHeader">
            <div class="fHeaderTitle">Image</div>
            <div class="fSmall">Upload / replace</div>
          </div>
          <div class="smallBtnRow">
            <button class="btn" id="replaceImageBtn">Replace image</button>
          </div>
        </div>
      `;
    } else {
      visualSpecificHtml = `
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
    }

    formatBody.innerHTML = canvasSectionHtml + posHtml + visualSpecificHtml;

    // Wire canvas fields at top of pane
    document.getElementById("canvasBgColor").oninput = (e) => { canvasBg.backgroundColor = e.target.value || canvasBg.backgroundColor; applyCanvasBackground(); };
    document.getElementById("canvasBgOpacity").oninput = (e) => { const v = Number(e.target.value); canvasBg.opacity = isFinite(v) ? clamp(v,0,1) : canvasBg.opacity; applyCanvasBackground(); };
    document.getElementById("canvasBgImage").oninput = (e) => { canvasBg.backgroundImage = e.target.value || ""; applyCanvasBackground(); };

    // Visual fields
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

      if (isFinite(nx)) v.x = clamp(nx, 0, Math.max(0, CANVAS_W - v.w));
      if (isFinite(ny)) v.y = clamp(ny, 0, Math.max(0, CANVAS_H - v.h));
      if (isFinite(nw)) v.w = clamp(nw, 40, Math.max(40, CANVAS_W - v.x));
      if (isFinite(nh)) v.h = clamp(nh, 24, Math.max(24, CANVAS_H - v.y));

      applyVisualRect(v);
      if (v.chart) v.chart.resize();
      renderFormatPane();
    };

    document.getElementById("deleteBtn").onclick = () => removeVisual(v.id);

    if (v.type !== "image" && v.type !== "textbox") {
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

    if (v.type === "textbox") {
      document.getElementById("applyTxtBtn").onclick = () => {
        const content = document.getElementById("txtContent").value;
        const fs = Number(document.getElementById("txtFontSize").value) || 14;
        const color = document.getElementById("txtColor").value || "#e8ecf2";
        const bg = document.getElementById("txtBg").value || "transparent";
        const align = document.getElementById("txtAlign").value || "left";
        const bold = !!document.getElementById("txtBold").checked;

        v.text = content;
        v.fontSize = clamp(fs, 8, 96);
        v.color = color;
        v.bgColor = bg;
        v.align = align;
        v.bold = bold;

        renderVisualContent(v);
        showToast("Text box updated");
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
      // user explicitly initiated upload
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
        // include visual-specific fields
        kpiValue: v.kpiValue,
        kpiLabel: v.kpiLabel,
        cardValue: v.cardValue,
        cardLabel: v.cardLabel,
        rows: v.rows,
        text: v.text,
        fontSize: v.fontSize,
        colorText: v.color,
        bgColorText: v.bgColor,
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
    clearAllVisuals();

    if (obj?.canvas?.background) {
      canvasBg = normalizeCanvasBg(obj.canvas.background);
    } else {
      canvasBg = structuredClone(sampleCanvasBg);
    }
    applyCanvasBackground();

    // load canvas size if provided
    if (obj?.canvas?.width && obj?.canvas?.height) {
      CANVAS_W = clamp(Number(obj.canvas.width) || CANVAS_W, 200, 3840);
      CANVAS_H = clamp(Number(obj.canvas.height) || CANVAS_H, 100, 2160);
      setCanvasCssSize(CANVAS_W, CANVAS_H);
    } else {
      setCanvasCssSize(CANVAS_W, CANVAS_H);
    }

    theme = normalizeTheme(obj?.theme || defaultTheme);
    applyThemeEverywhere();

    const visuals = Array.isArray(obj?.visuals) ? obj.visuals : [];
    visuals.forEach(vs => {
      const id = String(vs.id || `vis${state.nextId++}`);
      const v = {
        id,
        type: vs.type,
        title: vs.title || defaultTitleFor(vs.type),
        x: clamp(Number(vs.x)||40, 0, CANVAS_W-220),
        y: clamp(Number(vs.y)||40, 0, CANVAS_H-170),
        w: clamp(Number(vs.w)||DEFAULT_SIZES[vs.type]?.w||520, 40, CANVAS_W),
        h: clamp(Number(vs.h)||DEFAULT_SIZES[vs.type]?.h||300, 24, CANVAS_H),
        series: Array.isArray(vs.series)
          ? vs.series.map((s,i)=>({
              key: s.key || s.name || `S${i+1}`,
              name: s.name || s.key || `Series ${i+1}`,
              color: s.color || paletteColor(i),
              overrideColor: !!s.overrideColor
            }))
          : buildDefaultSeries(vs.type),
        imageDataUrl: vs.type === "image" ? (vs.imageDataUrl || "") : undefined,
        chart: null
      };

      // visual specific fields
      if (vs.type === "kpi") { v.kpiValue = vs.kpiValue ?? 0; v.kpiLabel = vs.kpiLabel ?? "Metric"; }
      if (vs.type === "card") { v.cardValue = vs.cardValue ?? 0; v.cardLabel = vs.cardLabel ?? "Card"; }
      if (vs.type === "multirowcard") { v.rows = Array.isArray(vs.rows) ? vs.rows : []; }
      if (vs.type === "textbox") {
        v.text = vs.text || "Editable text";
        v.fontSize = vs.fontSize || 16;
        v.color = vs.colorText || "#e8ecf2";
        v.bgColor = vs.bgColorText || "transparent";
        v.bold = !!vs.bold;
        v.align = vs.align || "left";
      }

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
  // Pane alignment with canvas
  // Align left/right panes top padding to the canvas top position
  // -------------------------
  function alignPanesWithCanvas(){
    if (!canvas || !leftPane || !rightPane || !canvasFrame) return;
    // compute canvas top relative to the pane containers
    const canvasRect = canvas.getBoundingClientRect();
    const leftRect = leftPane.getBoundingClientRect();
    const rightRect = rightPane.getBoundingClientRect();

    // how much offset within left/right panes to align their content with the canvas top?
    const leftOffset = Math.max(8, canvasRect.top - leftRect.top);
    const rightOffset = Math.max(8, canvasRect.top - rightRect.top);

    leftPane.style.paddingTop = `${leftOffset}px`;
    rightPane.style.paddingTop = `${rightOffset}px`;
    document.documentElement.style.setProperty("--canvasOffsetTop", `${Math.min(leftOffset, rightOffset)}px`);
  }

  window.addEventListener("resize", () => {
    alignPanesWithCanvas();
  });

  // -------------------------
  // Utilities: ensure alignment after canvas size changes
  // -------------------------
  function setCanvasSizeImmediate(w, h){
    CANVAS_W = clamp(Math.round(w), 200, 3840);
    CANVAS_H = clamp(Math.round(h), 100, 2160);
    setCanvasCssSize(CANVAS_W, CANVAS_H);
  }

  // -------------------------
  // Start: format pane should always show canvas settings
  // -------------------------
  renderFormatPane();

  // -------------------------
  // IMPORTANT: Empty dashboard on load (no starter visuals)
  // -------------------------
  // addVisual("line");
  // addVisual("donut");

  // align panes initially
  setTimeout(() => alignPanesWithCanvas(), 120);

})();
