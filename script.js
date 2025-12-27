// script.js
/* =========================================================
   Power BI-Style Prototype Dashboard (no real analytics)
   - Fixed 1280x720 canvas
   - Gridstack drag/resize with overlap allowed
   - Chart.js visuals + custom prototypes (Treemap, Ribbon)
   - Always-visible format pane (dashboard vs visual context)
   - Per-series name + color editing
   - Visual-level background + theme styling
   - Dashboard-level background + theme
   - Upload image as draggable/resizable visual
========================================================= */

(() => {
  // ---------- Visual catalog (category-wise, like Power BI) ----------
  const VISUAL_CATALOG = [
    {
      category: "Common",
      items: [
        { key: "line", name: "Line Chart" },
        { key: "area", name: "Area Chart" },
        { key: "stackedArea", name: "Stacked Area Chart" },
        { key: "clusteredColumn", name: "Clustered Column Chart" },
        { key: "stackedColumn", name: "Stacked Column Chart" },
        { key: "pctStackedColumn", name: "100% Stacked Column Chart" },
        { key: "clusteredBar", name: "Clustered Bar Chart" },
        { key: "stackedBar", name: "Stacked Bar Chart" },
        { key: "pctStackedBar", name: "100% Stacked Bar Chart" },
      ]
    },
    {
      category: "Parts-to-Whole",
      items: [
        { key: "pie", name: "Pie Chart" },
        { key: "donut", name: "Donut Chart" },
        { key: "treemap", name: "Treemap (prototype)" },
      ]
    },
    {
      category: "Advanced",
      items: [
        { key: "ribbon", name: "Ribbon Chart (prototype)" },
        { key: "scatter", name: "Scatter Chart" },
        { key: "lineClusteredColumn", name: "Line & Clustered Column Chart" },
        { key: "lineStackedColumn", name: "Line & Stacked Column Chart" },
      ]
    }
  ];

  // ---------- DOM ----------
  const el = {
    // left
    visualCategory: document.getElementById("visualCategory"),
    visualType: document.getElementById("visualType"),
    addVisualBtn: document.getElementById("addVisualBtn"),
    deselectBtn: document.getElementById("deselectBtn"),
    imageUpload: document.getElementById("imageUpload"),

    // canvas
    canvas: document.getElementById("canvas"),
    grid: document.getElementById("grid"),
    canvasClickCatcher: document.getElementById("canvasClickCatcher"),

    // format pane
    formatContext: document.getElementById("formatContext"),
    dashboardFormat: document.getElementById("dashboardFormat"),
    visualFormat: document.getElementById("visualFormat"),

    // dashboard format fields
    dashboardBg: document.getElementById("dashboardBg"),
    dashboardBgReset: document.getElementById("dashboardBgReset"),
    dashboardTheme: document.getElementById("dashboardTheme"),
    defaultVisualBg: document.getElementById("defaultVisualBg"),
    defaultVisualBgReset: document.getElementById("defaultVisualBgReset"),
    defaultTitleColor: document.getElementById("defaultTitleColor"),
    defaultTitleReset: document.getElementById("defaultTitleReset"),

    // visual format fields
    visualTitle: document.getElementById("visualTitle"),
    visualBg: document.getElementById("visualBg"),
    visualBgReset: document.getElementById("visualBgReset"),
    visualTheme: document.getElementById("visualTheme"),
    applyVisualTheme: document.getElementById("applyVisualTheme"),
    seriesEditor: document.getElementById("seriesEditor"),
  };

  // ---------- State ----------
  const state = {
    grid: null,
    selectedId: null,
    visuals: new Map(), // id -> model
    nextId: 1,

    // dashboard formatting
    dashboard: {
      theme: "dark",
      bg: "#0b1220",
      defaultVisualBg: "rgba(255,255,255,0.04)",
      defaultTitleColor: "#e5e7eb",
    }
  };

  // ---------- Themes ----------
  const THEMES = {
    light: {
      dashboardBg: "#f5f7fb",
      defaultVisualBg: "rgba(0,0,0,0.03)",
      defaultTitleColor: "#111827",
      chartText: "#111827",
      gridBorder: "rgba(0,0,0,0.10)"
    },
    dark: {
      dashboardBg: "#0b1220",
      defaultVisualBg: "rgba(255,255,255,0.04)",
      defaultTitleColor: "#e5e7eb",
      chartText: "#e5e7eb",
      gridBorder: "rgba(255,255,255,0.12)"
    },
    brand: {
      dashboardBg: "#071428",
      defaultVisualBg: "rgba(59,130,246,0.10)",
      defaultTitleColor: "#e6f0ff",
      chartText: "#e6f0ff",
      gridBorder: "rgba(59,130,246,0.30)"
    }
  };

  // ---------- Helpers ----------
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  function cssVarSet(name, value){
    document.documentElement.style.setProperty(name, value);
  }

  function rgbaFromHex(hex, alpha){
    const h = hex.replace("#","");
    const r = parseInt(h.slice(0,2),16);
    const g = parseInt(h.slice(2,4),16);
    const b = parseInt(h.slice(4,6),16);
    return `rgba(${r},${g},${b},${alpha})`;
  }

  function normalize100Percent(seriesArrays){
    // seriesArrays: [ [v1,v2..], [v1,v2..], ... ] per series across categories
    const len = seriesArrays[0].length;
    const out = seriesArrays.map(arr => arr.slice());
    for(let i=0;i<len;i++){
      let sum = 0;
      for(let s=0;s<out.length;s++) sum += out[s][i];
      sum = sum || 1;
      for(let s=0;s<out.length;s++){
        out[s][i] = Math.round((out[s][i] / sum) * 1000) / 10; // 1 decimal
      }
    }
    return out;
  }

  function makeNicePalette(n, themeKey){
    // Stable, readable palette. Users can override per series.
    // We avoid hardcoding theme colors into charts; this palette is just starting defaults.
    const base = [
      "#3b82f6", "#22c55e", "#f59e0b", "#a855f7", "#ef4444",
      "#06b6d4", "#84cc16", "#f97316", "#14b8a6", "#eab308"
    ];
    const out = [];
    for(let i=0;i<n;i++) out.push(base[i % base.length]);
    return out;
  }

  function toCanvasSizePx(){
    // from CSS variables (fixed)
    return { w: 1280, h: 720 };
  }

  function makeDefaultGridPlacement(){
    // small offset cascade so they don't all appear exactly on top
    const idx = state.nextId - 1;
    const x = (idx * 2) % 12;      // in columns (grid has 24)
    const y = (idx * 2) % 10;
    return { x, y };
  }

  function ensureFormatContext(){
    if(state.selectedId){
      el.formatContext.textContent = "Visual";
      el.dashboardFormat.classList.add("hidden");
      el.visualFormat.classList.remove("hidden");
    } else {
      el.formatContext.textContent = "Dashboard";
      el.visualFormat.classList.add("hidden");
      el.dashboardFormat.classList.remove("hidden");
    }
  }

  function deselectVisual(){
    if(state.selectedId){
      const m = state.visuals.get(state.selectedId);
      if(m?.rootEl) m.rootEl.classList.remove("selected");
    }
    state.selectedId = null;
    ensureFormatContext();
    refreshDashboardFormatUI();
    clearSeriesEditor();
  }

  function selectVisual(id){
    if(state.selectedId === id) return;

    // deselect previous
    if(state.selectedId){
      const prev = state.visuals.get(state.selectedId);
      if(prev?.rootEl) prev.rootEl.classList.remove("selected");
    }

    state.selectedId = id;

    const m = state.visuals.get(id);
    if(m?.rootEl) m.rootEl.classList.add("selected");

    ensureFormatContext();
    refreshVisualFormatUI();
  }

  function clearSeriesEditor(){
    el.seriesEditor.innerHTML = "";
  }

  function setDashboardBackground(color){
    state.dashboard.bg = color;
    el.canvas.style.background = color;
    el.dashboardBg.value = color;
  }

  function applyDashboardTheme(themeKey){
    state.dashboard.theme = themeKey;
    const t = THEMES[themeKey] || THEMES.dark;

    setDashboardBackground(t.dashboardBg);

    // Default visual styling (CSS variables)
    state.dashboard.defaultVisualBg = t.defaultVisualBg;
    state.dashboard.defaultTitleColor = t.defaultTitleColor;
    cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
    cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);

    // Update UI pickers
    el.dashboardTheme.value = themeKey;

    // Default visual bg picker needs hex; we keep a "best effort"
    el.defaultVisualBg.value = themeKey === "light" ? "#ffffff" : (themeKey === "brand" ? "#3b82f6" : "#111827");
    el.defaultTitleColor.value = t.defaultTitleColor;

    // Update existing visuals borders to fit theme a bit
    for(const m of state.visuals.values()){
      if(m?.rootEl){
        m.rootEl.style.borderColor = t.gridBorder;
      }
      // Update chart tick/label colors to match theme for readability
      if(m?.chart){
        applyChartThemeToInstance(m.chart, themeKey);
        m.chart.update();
      }
    }
  }

  function applyChartThemeToInstance(chart, themeKey){
    const t = THEMES[themeKey] || THEMES.dark;
    const opts = chart.options || {};
    opts.plugins = opts.plugins || {};
    opts.plugins.legend = opts.plugins.legend || {};
    opts.plugins.legend.labels = opts.plugins.legend.labels || {};
    opts.plugins.legend.labels.color = t.chartText;

    opts.scales = opts.scales || {};
    for(const k of Object.keys(opts.scales)){
      const sc = opts.scales[k];
      sc.ticks = sc.ticks || {};
      sc.grid = sc.grid || {};
      sc.ticks.color = t.chartText;
      sc.grid.color = themeKey === "light" ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.10)";
    }

    // Title color handled in DOM header, not chart plugin title.
    chart.options = opts;
  }

  // ---------- Dummy retail-ish data (prototype only) ----------
  function makeRetailTimeLabels(){
    return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  }

  function makeGrowthSeries(base, growthPct){
    // gradual growth with small noise, ends higher
    const labels = makeRetailTimeLabels();
    const data = [];
    let v = base;
    for(let i=0;i<labels.length;i++){
      const step = (base * growthPct / 100) / (labels.length - 1);
      const noise = (Math.sin(i * 1.3) + Math.cos(i * 0.7)) * (base * 0.015);
      v = v + step + noise;
      data.push(Math.max(0, Math.round(v)));
    }
    return { labels, data };
  }

  function makeCategoryLabels(){
    return ["North","South","East","West"];
  }

  function makeChannelLabels(){
    return ["Online","Store","Wholesale","Marketplace"];
  }

  // ---------- Build Chart.js configs ----------
  function baseChartOptions(themeKey){
    const t = THEMES[themeKey] || THEMES.dark;
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 220 },
      plugins: {
        legend: {
          position: "bottom",
          labels: { color: t.chartText, boxWidth: 12, boxHeight: 12 }
        },
        tooltip: { enabled: true }
      },
      scales: {
        x: { ticks: { color: t.chartText }, grid: { color: themeKey === "light" ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.10)" } },
        y: { ticks: { color: t.chartText }, grid: { color: themeKey === "light" ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.10)" } }
      }
    };
  }

  function buildChartConfig(visualKey, model){
    const themeKey = state.dashboard.theme;
    const opts = baseChartOptions(themeKey);
    const palette = makeNicePalette(6, themeKey);

    // Default labels (can be edited in format pane)
    const months = makeRetailTimeLabels();
    const regions = makeCategoryLabels();
    const channels = makeChannelLabels();

    // Common series defaults
    const sNames = ["A","B","C"];
    const sColors = palette.slice(0, 3);

    // Model storage: labels + series
    if(!model.series) model.series = [];
    if(!model.labels) model.labels = [];

    switch(visualKey){
      case "pie":
      case "donut": {
        model.labels = channels.slice();
        const values = [42, 28, 18, 12]; // nice distribution
        const colors = makeNicePalette(values.length, themeKey);
        model.series = model.labels.map((name, i) => ({ name, color: colors[i], values: [values[i]] }));

        return {
          type: "doughnut",
          data: {
            labels: model.labels,
            datasets: [{
              label: "Sales Share",
              data: values,
              backgroundColor: colors.map(c => rgbaFromHex(c, 0.85)),
              borderColor: colors.map(c => rgbaFromHex(c, 1)),
              borderWidth: 1
            }]
          },
          options: {
            ...opts,
            cutout: visualKey === "donut" ? "62%" : "0%",
            scales: {} // no axes
          }
        };
      }

      case "line":
      case "area":
      case "stackedArea": {
        const s1 = makeGrowthSeries(120, 35);
        const s2 = makeGrowthSeries(90, 55);
        const s3 = makeGrowthSeries(60, 80);

        model.labels = months.slice();
        model.series = [
          { name: "A", color: sColors[0], values: s1.data },
          { name: "B", color: sColors[1], values: s2.data },
          { name: "C", color: sColors[2], values: s3.data },
        ];

        const isArea = visualKey === "area" || visualKey === "stackedArea";
        const stacked = visualKey === "stackedArea";

        return {
          type: "line",
          data: {
            labels: months,
            datasets: model.series.map((s, idx) => ({
              label: s.name,
              data: s.values,
              borderColor: rgbaFromHex(s.color, 1),
              backgroundColor: rgbaFromHex(s.color, isArea ? 0.25 : 0.10),
              fill: isArea ? true : false,
              tension: 0.35,
              pointRadius: 2,
              borderWidth: 2
            }))
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, stacked },
              y: { ...opts.scales.y, stacked }
            }
          }
        };
      }

      case "clusteredColumn":
      case "stackedColumn":
      case "pctStackedColumn": {
        model.labels = regions.slice();
        const sA = [220, 180, 210, 240];
        const sB = [140, 160, 120, 150];
        const sC = [80, 95, 75, 100];

        let arrays = [sA, sB, sC];
        if(visualKey === "pctStackedColumn"){
          arrays = normalize100Percent(arrays);
        }

        model.series = [
          { name: "A", color: sColors[0], values: arrays[0] },
          { name: "B", color: sColors[1], values: arrays[1] },
          { name: "C", color: sColors[2], values: arrays[2] },
        ];

        const stacked = (visualKey !== "clusteredColumn");
        return {
          type: "bar",
          data: {
            labels: model.labels,
            datasets: model.series.map(s => ({
              label: s.name,
              data: s.values,
              backgroundColor: rgbaFromHex(s.color, 0.75),
              borderColor: rgbaFromHex(s.color, 1),
              borderWidth: 1
            }))
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, stacked },
              y: { ...opts.scales.y, stacked, suggestedMax: visualKey === "pctStackedColumn" ? 100 : undefined }
            }
          }
        };
      }

      case "clusteredBar":
      case "stackedBar":
      case "pctStackedBar": {
        model.labels = regions.slice();
        const sA = [220, 180, 210, 240];
        const sB = [140, 160, 120, 150];
        const sC = [80, 95, 75, 100];

        let arrays = [sA, sB, sC];
        if(visualKey === "pctStackedBar"){
          arrays = normalize100Percent(arrays);
        }

        model.series = [
          { name: "A", color: sColors[0], values: arrays[0] },
          { name: "B", color: sColors[1], values: arrays[1] },
          { name: "C", color: sColors[2], values: arrays[2] },
        ];

        const stacked = (visualKey !== "clusteredBar");
        return {
          type: "bar",
          data: {
            labels: model.labels,
            datasets: model.series.map(s => ({
              label: s.name,
              data: s.values,
              backgroundColor: rgbaFromHex(s.color, 0.75),
              borderColor: rgbaFromHex(s.color, 1),
              borderWidth: 1
            }))
          },
          options: {
            ...opts,
            indexAxis: "y",
            scales: {
              x: { ...opts.scales.x, stacked, suggestedMax: visualKey === "pctStackedBar" ? 100 : undefined },
              y: { ...opts.scales.y, stacked }
            }
          }
        };
      }

      case "lineClusteredColumn":
      case "lineStackedColumn": {
        // Combo: columns by region + line trend
        model.labels = regions.slice();

        const colSeriesA = [220, 180, 210, 240];
        const colSeriesB = [140, 160, 120, 150];
        const lineSeries = [38, 44, 57, 68]; // e.g., profit %

        const stacked = (visualKey === "lineStackedColumn");

        model.series = [
          { name: "A", color: sColors[0], values: colSeriesA, kind: "bar" },
          { name: "B", color: sColors[1], values: colSeriesB, kind: "bar" },
          { name: "C", color: palette[3], values: lineSeries, kind: "line" }, // use a 4th color
        ];

        return {
          data: {
            labels: model.labels,
            datasets: [
              {
                type: "bar",
                label: model.series[0].name,
                data: model.series[0].values,
                backgroundColor: rgbaFromHex(model.series[0].color, 0.75),
                borderColor: rgbaFromHex(model.series[0].color, 1),
                borderWidth: 1
              },
              {
                type: "bar",
                label: model.series[1].name,
                data: model.series[1].values,
                backgroundColor: rgbaFromHex(model.series[1].color, 0.75),
                borderColor: rgbaFromHex(model.series[1].color, 1),
                borderWidth: 1
              },
              {
                type: "line",
                label: model.series[2].name,
                data: model.series[2].values,
                borderColor: rgbaFromHex(model.series[2].color, 1),
                backgroundColor: rgbaFromHex(model.series[2].color, 0.15),
                pointRadius: 3,
                borderWidth: 2,
                tension: 0.35,
                yAxisID: "y1"
              }
            ]
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, stacked },
              y: { ...opts.scales.y, stacked },
              y1: {
                position: "right",
                ticks: { color: (THEMES[themeKey] || THEMES.dark).chartText },
                grid: { drawOnChartArea: false },
                suggestedMin: 0,
                suggestedMax: 100
              }
            }
          }
        };
      }

      case "scatter": {
        // Scatter: Sales vs Profit with 3 clusters
        model.labels = ["Cluster A", "Cluster B", "Cluster C"];
        model.series = [
          { name: "A", color: sColors[0], values: [] },
          { name: "B", color: sColors[1], values: [] },
          { name: "C", color: sColors[2], values: [] },
        ];

        function pts(cx, cy, n){
          const out = [];
          for(let i=0;i<n;i++){
            out.push({
              x: Math.round(cx + (Math.random()-0.5)*18),
              y: Math.round(cy + (Math.random()-0.5)*14),
            });
          }
          return out;
        }

        model.series[0].values = pts(60, 28, 14);
        model.series[1].values = pts(90, 44, 14);
        model.series[2].values = pts(120, 60, 14);

        return {
          type: "scatter",
          data: {
            datasets: model.series.map(s => ({
              label: s.name,
              data: s.values,
              backgroundColor: rgbaFromHex(s.color, 0.8),
              borderColor: rgbaFromHex(s.color, 1),
              pointRadius: 4
            }))
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, title: { display: true, text: "Sales (Index)", color: (THEMES[themeKey] || THEMES.dark).chartText } },
              y: { ...opts.scales.y, title: { display: true, text: "Profit (Index)", color: (THEMES[themeKey] || THEMES.dark).chartText } }
            }
          }
        };
      }

      default:
        return null;
    }
  }

  // ---------- Custom prototype visuals (Treemap, Ribbon) ----------
  function renderTreemapPrototype(canvas, model){
    const ctx = canvas.getContext("2d");
    const w = canvas.width, h = canvas.height;

    ctx.clearRect(0,0,w,h);

    // Ensure model structure
    if(!model.labels || model.labels.length === 0){
      model.labels = ["Apparel","Footwear","Accessories","Beauty","Home"];
    }
    if(!model.series || model.series.length === 0){
      const colors = makeNicePalette(model.labels.length, state.dashboard.theme);
      const sizes = [32, 22, 18, 15, 13]; // %
      model.series = model.labels.map((name,i)=>({ name, color: colors[i], values: [sizes[i]] }));
    }

    // Simple slice-and-dice treemap layout
    const values = model.series.map(s => s.values[0]);
    const total = values.reduce((a,b)=>a+b,0) || 1;

    let x = 0, y = 0;
    let horizontal = true;
    let remainingW = w, remainingH = h;

    const items = model.series
      .map((s,i)=>({ s, v: values[i], frac: values[i]/total }))
      .sort((a,b)=>b.v-a.v);

    // soft grid background
    ctx.fillStyle = "rgba(255,255,255,0.03)";
    ctx.fillRect(0,0,w,h);

    for(let i=0;i<items.length;i++){
      const frac = items[i].frac;
      let rw, rh;
      if(horizontal){
        rw = remainingW;
        rh = Math.max(18, Math.round(remainingH * frac));
      } else {
        rw = Math.max(18, Math.round(remainingW * frac));
        rh = remainingH;
      }

      const rect = { x, y, w: rw, h: rh };

      // draw
      ctx.fillStyle = rgbaFromHex(items[i].s.color, 0.75);
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);

      ctx.strokeStyle = "rgba(255,255,255,0.22)";
      ctx.lineWidth = 1;
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);

      // label
      const themeKey = state.dashboard.theme;
      ctx.fillStyle = (THEMES[themeKey] || THEMES.dark).chartText;
      ctx.font = "12px Segoe UI, Arial";
      const label = `${items[i].s.name} • ${Math.round(items[i].v)}%`;
      if(rect.w > 70 && rect.h > 18){
        ctx.fillText(label, rect.x + 8, rect.y + 16);
      }

      // update remaining
      if(horizontal){
        y += rh;
        remainingH -= rh;
        if(remainingH < 60){
          horizontal = false;
          x = 0; y = 0;
          remainingW = w; remainingH = h;
          // stop early for simplicity
          break;
        }
      } else {
        x += rw;
        remainingW -= rw;
      }
    }

    // If we broke early, do a simple block layout for all items:
    if(items.length && (x === 0 && y === 0)){
      const cols = 2;
      const rows = Math.ceil(items.length / cols);
      const gap = 6;
      const cellW = Math.floor((w - gap*(cols+1)) / cols);
      const cellH = Math.floor((h - gap*(rows+1)) / rows);
      ctx.clearRect(0,0,w,h);
      ctx.fillStyle = "rgba(255,255,255,0.03)";
      ctx.fillRect(0,0,w,h);

      items.forEach((it, idx) => {
        const c = idx % cols;
        const r = Math.floor(idx / cols);
        const rx = gap + c*(cellW+gap);
        const ry = gap + r*(cellH+gap);
        ctx.fillStyle = rgbaFromHex(it.s.color, 0.75);
        ctx.fillRect(rx, ry, cellW, cellH);
        ctx.strokeStyle = "rgba(255,255,255,0.22)";
        ctx.strokeRect(rx + 0.5, ry + 0.5, cellW - 1, cellH - 1);
        ctx.fillStyle = (THEMES[state.dashboard.theme] || THEMES.dark).chartText;
        ctx.font = "12px Segoe UI, Arial";
        ctx.fillText(`${it.s.name}`, rx + 10, ry + 18);
      });
    }
  }

  function renderRibbonPrototype(canvas, model){
    const ctx = canvas.getContext("2d");
    const w = canvas.width, h = canvas.height;
    ctx.clearRect(0,0,w,h);

    // Data: ranks shift across periods (prototype)
    if(!model.labels || model.labels.length === 0){
      model.labels = ["Q1","Q2","Q3","Q4"];
    }
    if(!model.series || model.series.length === 0){
      const cats = ["Online","Store","Wholesale","Marketplace"];
      const colors = makeNicePalette(cats.length, state.dashboard.theme);
      model.series = cats.map((name,i)=>({ name, color: colors[i], values: [] }));
      // Each quarter: share values (just to draw ribbon thickness)
      // Online grows, Store stable, etc.
      const shares = [
        [26, 24, 18, 14],
        [24, 25, 16, 14],
        [28, 23, 15, 16],
        [32, 22, 14, 18],
      ];
      // transpose into series.values per quarter
      for(let s=0;s<model.series.length;s++){
        model.series[s].values = shares.map(q => q[s]);
      }
    }

    // background
    ctx.fillStyle = "rgba(255,255,255,0.03)";
    ctx.fillRect(0,0,w,h);

    const pad = 14;
    const innerW = w - pad*2;
    const innerH = h - pad*2;

    const steps = model.labels.length;
    const xStep = innerW / (steps - 1);

    // Compute stacked bands per step
    const sums = [];
    for(let i=0;i<steps;i++){
      sums[i] = model.series.reduce((a,s)=>a + (s.values[i] || 0), 0) || 1;
    }

    // For each step, compute y positions for each series band (top & bottom)
    const bands = []; // bands[i][s] = {y0,y1}
    for(let i=0;i<steps;i++){
      let y = pad;
      bands[i] = [];
      for(let s=0;s<model.series.length;s++){
        const v = (model.series[s].values[i] || 0);
        const bandH = innerH * (v / sums[i]);
        bands[i][s] = { y0: y, y1: y + bandH };
        y += bandH;
      }
    }

    // Draw ribbons for each series across steps (quadratic curves)
    for(let s=0;s<model.series.length;s++){
      ctx.beginPath();
      for(let i=0;i<steps;i++){
        const x = pad + i*xStep;
        const midY = (bands[i][s].y0 + bands[i][s].y1) / 2;
        if(i === 0) ctx.moveTo(x, midY);
        else {
          const prevX = pad + (i-1)*xStep;
          const prevMidY = (bands[i-1][s].y0 + bands[i-1][s].y1) / 2;
          const cx = (prevX + x) / 2;
          ctx.quadraticCurveTo(cx, prevMidY, x, midY);
        }
      }
      // stroke
      ctx.strokeStyle = rgbaFromHex(model.series[s].color, 1);
      ctx.lineWidth = 6;
      ctx.lineCap = "round";
      ctx.stroke();
    }

    // Step labels + separators
    const themeKey = state.dashboard.theme;
    ctx.fillStyle = (THEMES[themeKey] || THEMES.dark).chartText;
    ctx.font = "12px Segoe UI, Arial";
    ctx.textAlign = "center";
    for(let i=0;i<steps;i++){
      const x = pad + i*xStep;
      ctx.fillText(model.labels[i], x, h - 10);
      ctx.strokeStyle = themeKey === "light" ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.10)";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x, pad);
      ctx.lineTo(x, h - 26);
      ctx.stroke();
    }
  }

  function setupCustomCanvasHighDPI(canvas){
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.floor(rect.width * dpr));
    canvas.height = Math.max(1, Math.floor(rect.height * dpr));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  // ---------- Visual creation / rendering ----------
  function createGridItemShell(id, title){
    const wrapper = document.createElement("div");
    wrapper.className = "visual";
    wrapper.setAttribute("data-vid", String(id));
    wrapper.setAttribute("role", "group");
    wrapper.setAttribute("aria-label", `Visual: ${title}`);

    const header = document.createElement("div");
    header.className = "visual-header";

    const t = document.createElement("div");
    t.className = "visual-title";
    t.textContent = title;

    const actions = document.createElement("div");
    actions.className = "visual-actions";

    const del = document.createElement("button");
    del.className = "icon-btn danger";
    del.type = "button";
    del.title = "Delete visual";
    del.setAttribute("aria-label", "Delete visual");
    del.textContent = "✖";

    actions.appendChild(del);
    header.appendChild(t);
    header.appendChild(actions);

    const body = document.createElement("div");
    body.className = "visual-body";

    wrapper.appendChild(header);
    wrapper.appendChild(body);

    return { wrapper, header, titleEl: t, body, deleteBtn: del };
  }

  function makeDefaultVisualSize(){
    // Medium Power BI-like default (not full width)
    // grid columns: 24; cellHeight: 30; canvas is 1280px wide inside fixed frame
    // This yields a readable chart with title + legend.
    return { w: 10, h: 8 }; // ~ medium tile
  }

  function addVisual(visualKey, displayName){
    const id = String(state.nextId++);
    const size = makeDefaultVisualSize();
    const place = makeDefaultGridPlacement();

    const node = document.createElement("div");
    node.className = "grid-stack-item";
    node.innerHTML = `<div class="grid-stack-item-content"></div>`;

    state.grid.addWidget(node, { w: size.w, h: size.h, x: place.x, y: place.y, id });

    const content = node.querySelector(".grid-stack-item-content");
    const shell = createGridItemShell(id, displayName);
    content.appendChild(shell.wrapper);

    // Model
    const model = {
      id,
      type: visualKey,
      name: displayName,
      title: displayName,
      background: state.dashboard.defaultVisualBg,
      titleColor: state.dashboard.defaultTitleColor,
      themePreset: "inherit",
      rootEl: shell.wrapper,
      titleEl: shell.titleEl,
      bodyEl: shell.body,
      deleteBtn: shell.deleteBtn,
      chart: null,
      customCanvas: null,
      imageEl: null,
      labels: [],
      series: []
    };

    // Apply defaults
    shell.wrapper.style.background = model.background;
    shell.titleEl.style.color = model.titleColor;

    // Render chart/custom
    renderVisual(model);

    // Events
    shell.wrapper.addEventListener("mousedown", (e) => {
      // Prevent gridstack drag when interacting with controls
      if(e.target.closest("button") || e.target.closest("input") || e.target.closest("select")) return;
      selectVisual(id);
      e.stopPropagation();
    });

    shell.deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(id);
    });

    state.visuals.set(id, model);
    selectVisual(id);
  }

  function addImageVisual(file){
    const id = String(state.nextId++);
    const size = makeDefaultVisualSize();
    const place = makeDefaultGridPlacement();

    const node = document.createElement("div");
    node.className = "grid-stack-item";
    node.innerHTML = `<div class="grid-stack-item-content"></div>`;
    state.grid.addWidget(node, { w: size.w, h: size.h, x: place.x, y: place.y, id });

    const content = node.querySelector(".grid-stack-item-content");
    const shell = createGridItemShell(id, "Image");
    content.appendChild(shell.wrapper);

    const img = document.createElement("img");
    img.alt = "Uploaded visual";
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = "contain";
    img.style.borderRadius = "8px";

    shell.body.appendChild(img);

    const model = {
      id,
      type: "image",
      name: "Image",
      title: "Image",
      background: state.dashboard.defaultVisualBg,
      titleColor: state.dashboard.defaultTitleColor,
      themePreset: "inherit",
      rootEl: shell.wrapper,
      titleEl: shell.titleEl,
      bodyEl: shell.body,
      deleteBtn: shell.deleteBtn,
      chart: null,
      customCanvas: null,
      imageEl: img,
      labels: [],
      series: []
    };

    // background + title
    shell.wrapper.style.background = model.background;
    shell.titleEl.style.color = model.titleColor;

    // Load image
    const reader = new FileReader();
    reader.onload = () => { img.src = String(reader.result); };
    reader.readAsDataURL(file);

    // Events
    shell.wrapper.addEventListener("mousedown", (e) => {
      if(e.target.closest("button")) return;
      selectVisual(id);
      e.stopPropagation();
    });

    shell.deleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(id);
    });

    state.visuals.set(id, model);
    selectVisual(id);
  }

  function removeVisual(id){
    const m = state.visuals.get(id);
    if(!m) return;

    // Destroy chart if present
    if(m.chart){
      try { m.chart.destroy(); } catch {}
    }

    // Remove from grid
    const item = el.grid.querySelector(`.grid-stack-item[gs-id="${CSS.escape(id)}"]`)
      || el.grid.querySelector(`.grid-stack-item[data-gs-id="${CSS.escape(id)}"]`)
      || el.grid.querySelector(`.grid-stack-item[gs-id="${CSS.escape(id)}"]`);
    // Gridstack stores node ref; safest: find by dataset on content
    const node = [...el.grid.querySelectorAll(".grid-stack-item")].find(n => {
      const wid = n.getAttribute("gs-id") || n.getAttribute("data-gs-id") || n.gridstackNode?.id;
      return String(wid) === String(id);
    });
    if(node){
      state.grid.removeWidget(node);
    }

    state.visuals.delete(id);
    if(state.selectedId === id){
      deselectVisual();
    }
  }

  function renderVisual(model){
    // Clear body
    model.bodyEl.innerHTML = "";

    // Apply title
    model.titleEl.textContent = model.title;
    model.titleEl.style.color = model.titleColor;

    // Apply background
    model.rootEl.style.background = model.background;

    // Image visual already handled separately
    if(model.type === "image"){
      // (not used here)
      return;
    }

    // Custom prototypes
    if(model.type === "treemap" || model.type === "ribbon"){
      const c = document.createElement("canvas");
      c.setAttribute("aria-label", `${model.name} canvas`);
      model.bodyEl.appendChild(c);
      model.customCanvas = c;
      model.chart = null;

      // Render now and on resize
      const redraw = () => {
        setupCustomCanvasHighDPI(c);
        if(model.type === "treemap") renderTreemapPrototype(c, model);
        if(model.type === "ribbon") renderRibbonPrototype(c, model);
      };

      // Observe size changes
      const ro = new ResizeObserver(() => redraw());
      ro.observe(model.bodyEl);

      // store observer so it doesn't get GC'd unexpectedly
      model._ro = ro;

      redraw();
      return;
    }

    // Chart.js visuals
    const canvas = document.createElement("canvas");
    canvas.setAttribute("aria-label", `${model.name} chart`);
    model.bodyEl.appendChild(canvas);

    // Destroy existing chart
    if(model.chart){
      try { model.chart.destroy(); } catch {}
      model.chart = null;
    }

    const cfg = buildChartConfig(model.type, model);
    if(!cfg){
      // fallback placeholder
      const ctx = canvas.getContext("2d");
      ctx.fillText("Unsupported visual type", 10, 20);
      return;
    }

    // Chart.js config: handle mixed chart object
    let chart;
    if(cfg.type){
      chart = new Chart(canvas, cfg);
    } else {
      chart = new Chart(canvas, { type: "bar", ...cfg }); // cfg already includes data/options
    }

    model.chart = chart;
    applyChartThemeToInstance(chart, state.dashboard.theme);
    chart.update();
  }

  // ---------- Format UI binding ----------
  function refreshDashboardFormatUI(){
    el.dashboardTheme.value = state.dashboard.theme;
    el.dashboardBg.value = state.dashboard.bg;

    // best-effort: show defaults as hex-like choices
    const t = THEMES[state.dashboard.theme] || THEMES.dark;
    el.defaultTitleColor.value = t.defaultTitleColor;
    // choose a representative color for rgba defaults
    el.defaultVisualBg.value = (state.dashboard.theme === "light") ? "#ffffff" : (state.dashboard.theme === "brand" ? "#3b82f6" : "#111827");
  }

  function refreshVisualFormatUI(){
    const m = state.visuals.get(state.selectedId);
    if(!m) return;

    el.visualTitle.value = m.title;

    // Visual background can be rgba or hex. Show a sensible value:
    // If rgba, convert to a representative hex (best effort).
    const bg = m.background;
    if(bg.startsWith("#")){
      el.visualBg.value = bg;
    } else {
      // if it's rgba(...) choose white-ish control for user; does not change current bg until user picks
      el.visualBg.value = "#111827";
    }

    el.visualTheme.value = m.themePreset || "inherit";

    buildSeriesEditor(m);
  }

  function buildSeriesEditor(model){
    clearSeriesEditor();

    // Image visual: no series editing
    if(model.type === "image"){
      const div = document.createElement("div");
      div.className = "hint";
      div.textContent = "Image visual: no series/legend editing.";
      el.seriesEditor.appendChild(div);
      return;
    }

    // For pie/donut: series are slices (model.series)
    // For charts: series correspond to datasets
    // For treemap/ribbon: series are categories
    if(!model.series || model.series.length === 0){
      // Ensure it has something
      if(model.chart) {
        // pull from chart
        const ds = model.chart.data.datasets || [];
        model.series = ds.map((d, i) => ({
          name: d.label || `Series ${i+1}`,
          color: (d.borderColor && String(d.borderColor).startsWith("#")) ? d.borderColor : makeNicePalette(ds.length, state.dashboard.theme)[i],
          values: d.data || []
        }));
      }
    }

    model.series.forEach((s, idx) => {
      const row = document.createElement("div");
      row.className = "series-row";

      const name = document.createElement("input");
      name.type = "text";
      name.value = s.name;
      name.setAttribute("aria-label", `Series name ${idx+1}`);

      const color = document.createElement("input");
      color.type = "color";
      color.value = (s.color && s.color.startsWith("#")) ? s.color : "#3b82f6";
      color.setAttribute("aria-label", `Series color ${idx+1}`);

      // Name change
      name.addEventListener("input", () => {
        s.name = name.value || `Series ${idx+1}`;
        applySeriesEdits(model);
      });

      // Color change
      color.addEventListener("input", () => {
        s.color = color.value;
        applySeriesEdits(model);
      });

      row.appendChild(name);
      row.appendChild(color);
      el.seriesEditor.appendChild(row);
    });
  }

  function applySeriesEdits(model){
    // Update title text
    model.titleEl.textContent = model.title;

    // Custom prototypes
    if(model.type === "treemap" || model.type === "ribbon"){
      if(model.customCanvas){
        setupCustomCanvasHighDPI(model.customCanvas);
        if(model.type === "treemap") renderTreemapPrototype(model.customCanvas, model);
        if(model.type === "ribbon") renderRibbonPrototype(model.customCanvas, model);
      }
      return;
    }

    if(!model.chart) return;

    // Pie/Donut
    if(model.type === "pie" || model.type === "donut"){
      const labels = model.series.map(s => s.name);
      const colors = model.series.map(s => s.color);
      const values = model.series.map(s => (s.values?.[0] ?? 10));

      model.chart.data.labels = labels;
      model.chart.data.datasets[0].data = values;
      model.chart.data.datasets[0].backgroundColor = colors.map(c => rgbaFromHex(c, 0.85));
      model.chart.data.datasets[0].borderColor = colors.map(c => rgbaFromHex(c, 1));
      model.chart.update();
      return;
    }

    // Scatter
    if(model.type === "scatter"){
      model.chart.data.datasets.forEach((d, i) => {
        d.label = model.series[i]?.name ?? d.label;
        const c = model.series[i]?.color ?? "#3b82f6";
        d.backgroundColor = rgbaFromHex(c, 0.8);
        d.borderColor = rgbaFromHex(c, 1);
      });
      model.chart.update();
      return;
    }

    // Combo charts
    if(model.type === "lineClusteredColumn" || model.type === "lineStackedColumn"){
      model.chart.data.datasets.forEach((d, i) => {
        const s = model.series[i];
        if(!s) return;
        d.label = s.name;
        if(d.type === "line"){
          d.borderColor = rgbaFromHex(s.color, 1);
          d.backgroundColor = rgbaFromHex(s.color, 0.15);
        } else {
          d.backgroundColor = rgbaFromHex(s.color, 0.75);
          d.borderColor = rgbaFromHex(s.color, 1);
        }
      });
      model.chart.update();
      return;
    }

    // General charts (line/area/bar/column/stacked)
    model.chart.data.datasets.forEach((d, i) => {
      const s = model.series[i];
      if(!s) return;
      d.label = s.name;
      if(model.chart.config.type === "line"){
        d.borderColor = rgbaFromHex(s.color, 1);
        d.backgroundColor = rgbaFromHex(s.color, (model.type === "line") ? 0.10 : 0.25);
      } else {
        d.backgroundColor = rgbaFromHex(s.color, 0.75);
        d.borderColor = rgbaFromHex(s.color, 1);
      }
    });
    model.chart.update();
  }

  // ---------- Visual background + theme application ----------
  function applyVisualThemePreset(model, preset){
    model.themePreset = preset;

    const themeKey = preset === "inherit" ? state.dashboard.theme : preset;
    const t = THEMES[themeKey] || THEMES.dark;

    // Visual-specific background and title color
    // (User can still override background manually after applying.)
    model.background = t.defaultVisualBg;
    model.titleColor = t.defaultTitleColor;

    // Optional: themed shadow only for brand (keeps "no forced card shadow" rule)
    if(preset === "brand") model.rootEl.classList.add("themed-shadow");
    else model.rootEl.classList.remove("themed-shadow");

    // Apply to DOM
    model.rootEl.style.background = model.background;
    model.titleEl.style.color = model.titleColor;

    // Apply chart readability
    if(model.chart){
      applyChartThemeToInstance(model.chart, themeKey);
      model.chart.update();
    }
    if(model.customCanvas){
      setupCustomCanvasHighDPI(model.customCanvas);
      if(model.type === "treemap") renderTreemapPrototype(model.customCanvas, model);
      if(model.type === "ribbon") renderRibbonPrototype(model.customCanvas, model);
    }

    // Refresh UI controls
    refreshVisualFormatUI();
  }

  // ---------- Init visual dropdowns ----------
  function initVisualDropdowns(){
    // categories
    el.visualCategory.innerHTML = "";
    VISUAL_CATALOG.forEach((c, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = c.category;
      el.visualCategory.appendChild(opt);
    });

    const rebuildTypes = () => {
      const cat = VISUAL_CATALOG[Number(el.visualCategory.value)] || VISUAL_CATALOG[0];
      el.visualType.innerHTML = "";
      cat.items.forEach(it => {
        const opt = document.createElement("option");
        opt.value = it.key;
        opt.textContent = it.name;
        el.visualType.appendChild(opt);
      });
    };

    el.visualCategory.addEventListener("change", rebuildTypes);
    rebuildTypes();
  }

  // ---------- Gridstack init (fixed canvas, no infinite scroll) ----------
  function initGrid(){
    // fixed canvas; gridstack should not expand beyond it
    // Use 24 columns to feel like Power BI granularity.
    state.grid = GridStack.init({
      column: 24,
      cellHeight: 30,
      margin: 6,
      float: true,              // overlap-friendly
      disableOneColumnMode: true,
      animate: true,
      resizable: { handles: "se, sw, ne, nw, e, w, n, s" },
      draggable: { handle: ".visual-header, .visual-body" },
      alwaysShowResizeHandle: true,
      // Keep within fixed canvas
      // (Gridstack v11 uses 'containment' via draggable/resizable underlying; this works well in practice.)
      dragIn: false
    }, el.grid);

    // Prevent canvas click-catcher from blocking drag events on items:
    el.canvasClickCatcher.addEventListener("mousedown", () => {
      deselectVisual();
    });

    // Resize charts on grid resize
    state.grid.on("resizestop", (_e, item) => {
      const id = String(item?.id ?? item?.el?.getAttribute("gs-id") ?? item?.el?.gridstackNode?.id);
      const m = state.visuals.get(id);
      if(!m) return;
      if(m.chart) m.chart.resize();
      if(m.customCanvas){
        setupCustomCanvasHighDPI(m.customCanvas);
        if(m.type === "treemap") renderTreemapPrototype(m.customCanvas, m);
        if(m.type === "ribbon") renderRibbonPrototype(m.customCanvas, m);
      }
    });

    // If you want strictly no item leaving canvas, clamp on dragstop
    state.grid.on("dragstop", (_e, item) => {
      const node = item?.el?.gridstackNode;
      if(!node) return;
      node.x = clamp(node.x, 0, 24 - node.w);
      node.y = clamp(node.y, 0, 100); // y doesn't matter much; canvas is fixed but gridstack doesn't create infinite scroll
      state.grid.update(item.el, { x: node.x, y: node.y });
    });
  }

  // ---------- Bind UI ----------
  function bindUI(){
    // Add visual
    el.addVisualBtn.addEventListener("click", () => {
      const key = el.visualType.value;
      const cat = VISUAL_CATALOG[Number(el.visualCategory.value)] || VISUAL_CATALOG[0];
      const it = cat.items.find(x => x.key === key);
      addVisual(key, it ? it.name : key);
    });

    // Deselect
    el.deselectBtn.addEventListener("click", () => deselectVisual());

    // Upload image
    el.imageUpload.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if(!file) return;
      addImageVisual(file);
      e.target.value = ""; // allow re-upload same file
    });

    // Dashboard background
    el.dashboardBg.addEventListener("input", () => {
      setDashboardBackground(el.dashboardBg.value);
    });
    el.dashboardBgReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      setDashboardBackground(t.dashboardBg);
    });

    // Dashboard theme
    el.dashboardTheme.addEventListener("change", () => {
      applyDashboardTheme(el.dashboardTheme.value);
    });

    // Default visual background (new visuals only + optionally apply to selected)
    el.defaultVisualBg.addEventListener("input", () => {
      // store as rgba for nicer overlays
      state.dashboard.defaultVisualBg = rgbaFromHex(el.defaultVisualBg.value, 0.10);
      cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
    });
    el.defaultVisualBgReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      state.dashboard.defaultVisualBg = t.defaultVisualBg;
      cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
      el.defaultVisualBg.value = (state.dashboard.theme === "light") ? "#ffffff" : (state.dashboard.theme === "brand" ? "#3b82f6" : "#111827");
    });

    // Default title color (new visuals)
    el.defaultTitleColor.addEventListener("input", () => {
      state.dashboard.defaultTitleColor = el.defaultTitleColor.value;
      cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);
    });
    el.defaultTitleReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      state.dashboard.defaultTitleColor = t.defaultTitleColor;
      cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);
      el.defaultTitleColor.value = t.defaultTitleColor;
    });

    // Visual title
    el.visualTitle.addEventListener("input", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      m.title = el.visualTitle.value || m.name;
      m.titleEl.textContent = m.title;
    });

    // Visual bg
    el.visualBg.addEventListener("input", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      // store as rgba overlay to mimic Power BI tile backgrounds
      m.background = rgbaFromHex(el.visualBg.value, 0.10);
      m.rootEl.style.background = m.background;
    });
    el.visualBgReset.addEventListener("click", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      const themeKey = (m.themePreset === "inherit") ? state.dashboard.theme : m.themePreset;
      const t = THEMES[themeKey] || THEMES.dark;
      m.background = t.defaultVisualBg;
      m.rootEl.style.background = m.background;
    });

    // Visual theme preset apply
    el.applyVisualTheme.addEventListener("click", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      applyVisualThemePreset(m, el.visualTheme.value);
    });

    // Click outside any visual should show dashboard formatting (already via click-catcher)
    el.canvas.addEventListener("mousedown", (e) => {
      // If click hits a visual, selection logic will handle
      if(e.target.closest(".visual")) return;
      deselectVisual();
    });
  }

  // ---------- Boot ----------
  function init(){
    initVisualDropdowns();
    initGrid();
    bindUI();

    // Set initial theme & dashboard background
    applyDashboardTheme(state.dashboard.theme);

    // Initial UI shows dashboard formatting by default
    ensureFormatContext();
    refreshDashboardFormatUI();

    // Seed a couple visuals to feel real (optional but useful for clients)
    addVisual("line", "Line Chart");
    addVisual("donut", "Donut Chart");
    deselectVisual(); // ensure dashboard options are visible by default
  }

  // Start
  init();
})();
