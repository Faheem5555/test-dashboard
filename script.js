// script.js
/* =========================================================
   Power BI-Style Prototype Dashboard (FINAL)
   - Fixed 1280x720 canvas, no infinite scroll
   - GridStack v11 drag/resize + overlap allowed (disableCollision)
   - Correct-looking prototypes for all requested visuals:
     Pie, Donut, Treemap, Ribbon, Line (multi-series), Area, Stacked Area,
     Clustered/Stacked/100% Bar & Column, Combo charts, Scatter
   - Always-visible format pane (Dashboard vs Visual context)
   - Per-series color + name editing (mandatory)
   - Visual background + theme + Dashboard background + theme
   - Delete icon only on selected
   - Upload image as draggable/resizable visual
   - Dummy retail data (realistic trends)
========================================================= */

(() => {
  // ---------- Visual catalog (category-wise like Power BI) ----------
  const VISUAL_CATALOG = [
    {
      category: "Common",
      items: [
        { key: "pie", name: "Pie Chart" },
        { key: "donut", name: "Donut Chart" },
        { key: "line", name: "Line Chart (multiple trends)" },
        { key: "area", name: "Area Chart" },
        { key: "stackedArea", name: "Stacked Area Chart" },
        { key: "clusteredBar", name: "Clustered Bar Chart" },
        { key: "stackedBar", name: "Stacked Bar Chart" },
        { key: "pctStackedBar", name: "100% Stacked Bar Chart" },
        { key: "clusteredColumn", name: "Clustered Column Chart" },
        { key: "stackedColumn", name: "Stacked Column Chart" },
        { key: "pctStackedColumn", name: "100% Stacked Column Chart" },
      ]
    },
    {
      category: "Advanced (Prototype)",
      items: [
        { key: "treemap", name: "Treemap (visual prototype)" },
        { key: "ribbon", name: "Ribbon Chart (visual prototype)" },
        { key: "lineClusteredColumn", name: "Line & Clustered Column Chart" },
        { key: "lineStackedColumn", name: "Line & Stacked Column Chart" },
        { key: "scatter", name: "Scatter Chart" },
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

    // format
    formatContext: document.getElementById("formatContext"),
    dashboardFormat: document.getElementById("dashboardFormat"),
    visualFormat: document.getElementById("visualFormat"),

    // dashboard format
    dashboardBg: document.getElementById("dashboardBg"),
    dashboardBgReset: document.getElementById("dashboardBgReset"),
    dashboardTheme: document.getElementById("dashboardTheme"),
    defaultVisualBg: document.getElementById("defaultVisualBg"),
    defaultVisualBgReset: document.getElementById("defaultVisualBgReset"),
    defaultTitleColor: document.getElementById("defaultTitleColor"),
    defaultTitleReset: document.getElementById("defaultTitleReset"),

    // visual format
    visualTitle: document.getElementById("visualTitle"),
    visualBg: document.getElementById("visualBg"),
    visualBgReset: document.getElementById("visualBgReset"),
    visualTheme: document.getElementById("visualTheme"),
    applyVisualTheme: document.getElementById("applyVisualTheme"),
    seriesEditor: document.getElementById("seriesEditor"),
  };

  // ---------- Themes ----------
  const THEMES = {
    light: {
      dashboardBg: "#f5f7fb",
      defaultVisualBg: "rgba(0,0,0,0.03)",
      defaultTitleColor: "#111827",
      chartText: "#111827",
      gridLine: "rgba(0,0,0,0.08)",
      border: "rgba(0,0,0,0.12)"
    },
    dark: {
      dashboardBg: "#0b1220",
      defaultVisualBg: "rgba(255,255,255,0.04)",
      defaultTitleColor: "#e5e7eb",
      chartText: "#e5e7eb",
      gridLine: "rgba(255,255,255,0.10)",
      border: "rgba(255,255,255,0.12)"
    },
    brand: {
      dashboardBg: "#071428",
      defaultVisualBg: "rgba(59,130,246,0.10)",
      defaultTitleColor: "#e6f0ff",
      chartText: "#e6f0ff",
      gridLine: "rgba(59,130,246,0.22)",
      border: "rgba(59,130,246,0.30)"
    }
  };

  // ---------- State ----------
  const state = {
    grid: null,
    nextId: 1,
    selectedId: null,
    visuals: new Map(), // id -> model
    dashboard: {
      theme: "dark",
      bg: THEMES.dark.dashboardBg,
      defaultVisualBg: THEMES.dark.defaultVisualBg,
      defaultTitleColor: THEMES.dark.defaultTitleColor,
    }
  };

  // ---------- Helpers ----------
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  function cssVarSet(name, value){
    document.documentElement.style.setProperty(name, value);
  }

  function hexToRgb(hex){
    const h = hex.replace("#", "");
    const r = parseInt(h.slice(0,2),16);
    const g = parseInt(h.slice(2,4),16);
    const b = parseInt(h.slice(4,6),16);
    return { r, g, b };
  }
  function rgba(hex, a){
    const {r,g,b} = hexToRgb(hex);
    return `rgba(${r},${g},${b},${a})`;
  }

  function nicePalette(n){
    const base = [
      "#3b82f6", "#22c55e", "#f59e0b", "#a855f7", "#ef4444",
      "#06b6d4", "#84cc16", "#f97316", "#14b8a6", "#eab308"
    ];
    return Array.from({length:n}, (_,i)=>base[i % base.length]);
  }

  function normalize100(seriesArrays){
    // seriesArrays: [ [v...], [v...], ...]
    const len = seriesArrays[0].length;
    const out = seriesArrays.map(a => a.slice());
    for(let i=0;i<len;i++){
      let sum = 0;
      for(let s=0;s<out.length;s++) sum += out[s][i];
      sum = sum || 1;
      for(let s=0;s<out.length;s++){
        out[s][i] = Math.round((out[s][i]/sum)*1000)/10; // 1 decimal
      }
    }
    return out;
  }

  function setDashboardBackground(color){
    state.dashboard.bg = color;
    el.canvas.style.background = color;
    el.dashboardBg.value = color;
  }

  function applyChartThemeToChart(chart, themeKey){
    const t = THEMES[themeKey] || THEMES.dark;
    chart.options.plugins = chart.options.plugins || {};
    chart.options.plugins.legend = chart.options.plugins.legend || {};
    chart.options.plugins.legend.labels = chart.options.plugins.legend.labels || {};
    chart.options.plugins.legend.labels.color = t.chartText;

    chart.options.scales = chart.options.scales || {};
    for(const k of Object.keys(chart.options.scales)){
      const sc = chart.options.scales[k];
      sc.ticks = sc.ticks || {};
      sc.grid = sc.grid || {};
      sc.ticks.color = t.chartText;
      sc.grid.color = t.gridLine;
      if(sc.title){
        sc.title.color = t.chartText;
      }
    }
  }

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
        x: { ticks: { color: t.chartText }, grid: { color: t.gridLine } },
        y: { ticks: { color: t.chartText }, grid: { color: t.gridLine } }
      }
    };
  }

  function months(){ return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]; }
  function regions(){ return ["North","South","East","West"]; }
  function channels(){ return ["Online","Store","Wholesale","Marketplace"]; }

  function growthSeries(base, endBoostPct){
    const m = months();
    const data = [];
    let v = base;
    const totalSteps = m.length-1;
    const totalInc = base * (endBoostPct/100);
    for(let i=0;i<m.length;i++){
      const step = totalInc/totalSteps;
      const noise = (Math.sin(i*1.2)+Math.cos(i*0.7))*(base*0.015);
      v = v + step + noise;
      data.push(Math.max(0, Math.round(v)));
    }
    return { labels: m, data };
  }

  function makeDefaultPlacement(){
    // slight cascade so new visuals don't fully overlap by default
    const idx = state.nextId - 1;
    const x = (idx * 2) % 24;
    const y = (idx * 2) % 8;
    return { x, y };
  }

  function makeDefaultVisualSize(){
    // Power BI-like "medium" size (not full width)
    return { w: 10, h: 8 }; // 24 cols grid
  }

  // ---------- Selection + format context ----------
  function ensureFormatContext(){
    if(state.selectedId){
      el.formatContext.textContent = "Visual";
      el.dashboardFormat.classList.add("hidden");
      el.visualFormat.classList.remove("hidden");
    } else {
      el.formatContext.textContent = "Dashboard";
      el.visualFormat.classList.add("hidden");
      el.dashboardFormat.classList.remove("hidden");
      clearSeriesEditor();
    }
  }

  function deselect(){
    if(state.selectedId){
      const m = state.visuals.get(state.selectedId);
      if(m?.rootEl) m.rootEl.classList.remove("selected");
    }
    state.selectedId = null;
    ensureFormatContext();
    refreshDashboardUI();
  }

  function select(id){
    if(state.selectedId === id) return;

    if(state.selectedId){
      const prev = state.visuals.get(state.selectedId);
      if(prev?.rootEl) prev.rootEl.classList.remove("selected");
    }
    state.selectedId = id;

    const m = state.visuals.get(id);
    if(m?.rootEl) m.rootEl.classList.add("selected");

    ensureFormatContext();
    refreshVisualUI();
  }

  // ---------- Theme application ----------
  function applyDashboardTheme(themeKey){
    state.dashboard.theme = themeKey;
    const t = THEMES[themeKey] || THEMES.dark;

    // canvas background
    setDashboardBackground(t.dashboardBg);

    // defaults
    state.dashboard.defaultVisualBg = t.defaultVisualBg;
    state.dashboard.defaultTitleColor = t.defaultTitleColor;

    cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
    cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);

    // update existing visuals where they inherit
    for(const m of state.visuals.values()){
      // border tune
      m.rootEl.style.borderColor = t.border;

      const appliedTheme = (m.themePreset === "inherit") ? themeKey : m.themePreset;
      const tt = THEMES[appliedTheme] || t;

      // keep user-chosen background if it was explicitly set (we store flag)
      if(!m.userSetBg){
        m.background = tt.defaultVisualBg;
        m.rootEl.style.background = m.background;
      }
      if(!m.userSetTitleColor){
        m.titleColor = tt.defaultTitleColor;
        m.titleEl.style.color = m.titleColor;
      }

      if(m.chart){
        applyChartThemeToChart(m.chart, appliedTheme);
        m.chart.update();
      }
      if(m.customCanvas){
        scheduleCustomRedraw(m);
      }
    }

    refreshDashboardUI();
  }

  function applyVisualThemePreset(model, preset){
    model.themePreset = preset;
    const themeKey = (preset === "inherit") ? state.dashboard.theme : preset;
    const t = THEMES[themeKey] || THEMES.dark;

    // apply only to this visual (and mark as not user-set)
    model.background = t.defaultVisualBg;
    model.titleColor = t.defaultTitleColor;
    model.userSetBg = false;
    model.userSetTitleColor = false;

    model.rootEl.style.background = model.background;
    model.titleEl.style.color = model.titleColor;

    if(model.chart){
      applyChartThemeToChart(model.chart, themeKey);
      model.chart.update();
    }
    if(model.customCanvas){
      scheduleCustomRedraw(model);
    }
    refreshVisualUI();
  }

  // ---------- UI refresh ----------
  function refreshDashboardUI(){
    el.dashboardTheme.value = state.dashboard.theme;
    el.dashboardBg.value = state.dashboard.bg;

    // show representative picks for defaults
    // defaultVisualBg is rgba, so pick a base hex for control
    el.defaultVisualBg.value =
      state.dashboard.theme === "light" ? "#ffffff" :
      state.dashboard.theme === "brand" ? "#3b82f6" : "#111827";

    el.defaultTitleColor.value = (THEMES[state.dashboard.theme] || THEMES.dark).defaultTitleColor;
  }

  function refreshVisualUI(){
    const m = state.visuals.get(state.selectedId);
    if(!m) return;

    el.visualTitle.value = m.title;

    // visualBg picker expects hex; if user set it via picker we'll store lastHex
    el.visualBg.value = m.lastBgHex || (
      m.themePreset === "brand" ? "#3b82f6" :
      m.themePreset === "light" ? "#ffffff" : "#111827"
    );

    el.visualTheme.value = m.themePreset || "inherit";
    buildSeriesEditor(m);
  }

  function clearSeriesEditor(){
    el.seriesEditor.innerHTML = "";
  }

  // ---------- Create visual shell ----------
  function createShell(id, title){
    const wrapper = document.createElement("div");
    wrapper.className = "visual";
    wrapper.setAttribute("data-vid", id);

    const header = document.createElement("div");
    header.className = "visual-header";

    const titleEl = document.createElement("div");
    titleEl.className = "visual-title";
    titleEl.textContent = title;

    const actions = document.createElement("div");
    actions.className = "visual-actions";

    const delBtn = document.createElement("button");
    delBtn.className = "icon-btn danger";
    delBtn.type = "button";
    delBtn.title = "Delete visual";
    delBtn.textContent = "✖";

    actions.appendChild(delBtn);
    header.appendChild(titleEl);
    header.appendChild(actions);

    const body = document.createElement("div");
    body.className = "visual-body";

    wrapper.appendChild(header);
    wrapper.appendChild(body);

    return { wrapper, titleEl, body, delBtn };
  }

  // ---------- Rendering: Chart.js configs ----------
  function buildChartConfig(visualKey, model){
    const themeKey = (model.themePreset === "inherit") ? state.dashboard.theme : model.themePreset;
    const opts = baseChartOptions(themeKey);
    const pal = nicePalette(6);

    // init series/labels if empty
    model.labels = model.labels || [];
    model.series = model.series || [];

    switch(visualKey){
      case "pie":
      case "donut": {
        const lbl = channels();
        const values = [42, 28, 18, 12];
        const colors = nicePalette(lbl.length);

        model.labels = lbl.slice();
        model.series = lbl.map((name,i)=>({ name, color: colors[i], values:[values[i]] }));

        return {
          type: "doughnut",
          data: {
            labels: model.labels,
            datasets: [{
              label: "Sales Share",
              data: values,
              backgroundColor: colors.map(c => rgba(c, 0.85)),
              borderColor: colors.map(c => rgba(c, 1)),
              borderWidth: 1
            }]
          },
          options: {
            ...opts,
            cutout: (visualKey === "donut") ? "62%" : "0%",
            scales: {}
          }
        };
      }

      case "line":
      case "area":
      case "stackedArea": {
        const s1 = growthSeries(120, 35);
        const s2 = growthSeries(90, 55);
        const s3 = growthSeries(60, 80); // ~80% growth
        model.labels = months().slice();
        model.series = [
          { name:"A", color: pal[0], values: s1.data },
          { name:"B", color: pal[1], values: s2.data },
          { name:"C", color: pal[2], values: s3.data },
        ];

        const isArea = (visualKey === "area" || visualKey === "stackedArea");
        const stacked = (visualKey === "stackedArea");

        return {
          type: "line",
          data: {
            labels: model.labels,
            datasets: model.series.map((s)=>({
              label: s.name,
              data: s.values,
              borderColor: rgba(s.color, 1),
              backgroundColor: rgba(s.color, isArea ? 0.25 : 0.10),
              fill: isArea,
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
        model.labels = regions().slice();
        const A = [220, 180, 210, 240];
        const B = [140, 160, 120, 150];
        const C = [80, 95, 75, 100];
        let arrays = [A,B,C];
        if(visualKey === "pctStackedColumn") arrays = normalize100(arrays);

        model.series = [
          { name:"A", color: pal[0], values: arrays[0] },
          { name:"B", color: pal[1], values: arrays[1] },
          { name:"C", color: pal[2], values: arrays[2] },
        ];

        const stacked = (visualKey !== "clusteredColumn");

        return {
          type: "bar",
          data: {
            labels: model.labels,
            datasets: model.series.map(s=>({
              label: s.name,
              data: s.values,
              backgroundColor: rgba(s.color, 0.75),
              borderColor: rgba(s.color, 1),
              borderWidth: 1
            }))
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, stacked },
              y: { ...opts.scales.y, stacked, suggestedMax: (visualKey==="pctStackedColumn") ? 100 : undefined }
            }
          }
        };
      }

      case "clusteredBar":
      case "stackedBar":
      case "pctStackedBar": {
        model.labels = regions().slice();
        const A = [220, 180, 210, 240];
        const B = [140, 160, 120, 150];
        const C = [80, 95, 75, 100];
        let arrays = [A,B,C];
        if(visualKey === "pctStackedBar") arrays = normalize100(arrays);

        model.series = [
          { name:"A", color: pal[0], values: arrays[0] },
          { name:"B", color: pal[1], values: arrays[1] },
          { name:"C", color: pal[2], values: arrays[2] },
        ];

        const stacked = (visualKey !== "clusteredBar");

        return {
          type: "bar",
          data: {
            labels: model.labels,
            datasets: model.series.map(s=>({
              label: s.name,
              data: s.values,
              backgroundColor: rgba(s.color, 0.75),
              borderColor: rgba(s.color, 1),
              borderWidth: 1
            }))
          },
          options: {
            ...opts,
            indexAxis: "y",
            scales: {
              x: { ...opts.scales.x, stacked, suggestedMax: (visualKey==="pctStackedBar") ? 100 : undefined },
              y: { ...opts.scales.y, stacked }
            }
          }
        };
      }

      case "lineClusteredColumn":
      case "lineStackedColumn": {
        model.labels = regions().slice();
        const colA = [220, 180, 210, 240];
        const colB = [140, 160, 120, 150];
        const profitPct = [38, 44, 57, 68]; // profit trend

        const stacked = (visualKey === "lineStackedColumn");

        model.series = [
          { name:"A", color: pal[0], values: colA, kind:"bar" },
          { name:"B", color: pal[1], values: colB, kind:"bar" },
          { name:"C", color: pal[3], values: profitPct, kind:"line" },
        ];

        return {
          data: {
            labels: model.labels,
            datasets: [
              {
                type:"bar",
                label: model.series[0].name,
                data: model.series[0].values,
                backgroundColor: rgba(model.series[0].color,0.75),
                borderColor: rgba(model.series[0].color,1),
                borderWidth:1
              },
              {
                type:"bar",
                label: model.series[1].name,
                data: model.series[1].values,
                backgroundColor: rgba(model.series[1].color,0.75),
                borderColor: rgba(model.series[1].color,1),
                borderWidth:1
              },
              {
                type:"line",
                label: model.series[2].name,
                data: model.series[2].values,
                borderColor: rgba(model.series[2].color,1),
                backgroundColor: rgba(model.series[2].color,0.15),
                borderWidth:2,
                tension:0.35,
                pointRadius:3,
                yAxisID:"y1"
              }
            ]
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, stacked },
              y: { ...opts.scales.y, stacked },
              y1: {
                position:"right",
                grid:{ drawOnChartArea:false },
                ticks:{ color: (THEMES[themeKey]||THEMES.dark).chartText },
                suggestedMin:0,
                suggestedMax:100
              }
            }
          }
        };
      }

      case "scatter": {
        model.series = [
          { name:"A", color: pal[0], values: [] },
          { name:"B", color: pal[1], values: [] },
          { name:"C", color: pal[2], values: [] },
        ];

        function pts(cx, cy, n){
          const out=[];
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
            datasets: model.series.map(s=>({
              label: s.name,
              data: s.values,
              backgroundColor: rgba(s.color,0.8),
              borderColor: rgba(s.color,1),
              pointRadius:4
            }))
          },
          options: {
            ...opts,
            scales: {
              x: { ...opts.scales.x, title: { display:true, text:"Sales (Index)", color:(THEMES[themeKey]||THEMES.dark).chartText } },
              y: { ...opts.scales.y, title: { display:true, text:"Profit (Index)", color:(THEMES[themeKey]||THEMES.dark).chartText } }
            }
          }
        };
      }

      default:
        return null;
    }
  }

  // ---------- Custom prototypes: Treemap & Ribbon ----------
  function setupCanvasDPI(canvas){
    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.floor(rect.width * dpr));
    canvas.height = Math.max(1, Math.floor(rect.height * dpr));
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr,0,0,dpr,0,0);
    return ctx;
  }

  function renderTreemap(canvas, model){
    const ctx = setupCanvasDPI(canvas);
    const w = canvas.getBoundingClientRect().width;
    const h = canvas.getBoundingClientRect().height;

    ctx.clearRect(0,0,w,h);
    ctx.fillStyle = "rgba(255,255,255,0.03)";
    ctx.fillRect(0,0,w,h);

    if(!model.series || model.series.length === 0){
      const labels = ["Apparel","Footwear","Accessories","Beauty","Home"];
      const sizes  = [32, 22, 18, 15, 13];
      const colors = nicePalette(labels.length);
      model.series = labels.map((n,i)=>({ name:n, color:colors[i], values:[sizes[i]] }));
    }

    const items = model.series.map(s=>({ s, v:s.values[0] })).sort((a,b)=>b.v-a.v);
    const total = items.reduce((a,b)=>a+b.v,0) || 1;

    const gap = 6;
    const cols = 2;
    const rows = Math.ceil(items.length/cols);
    const cellW = Math.floor((w - gap*(cols+1)) / cols);
    const cellH = Math.floor((h - gap*(rows+1)) / rows);

    const themeKey = (model.themePreset==="inherit") ? state.dashboard.theme : model.themePreset;
    const txt = (THEMES[themeKey]||THEMES.dark).chartText;

    ctx.font = "12px Segoe UI, Arial";
    ctx.textBaseline = "top";

    items.forEach((it, idx)=>{
      const c = idx % cols;
      const r = Math.floor(idx/cols);
      const x = gap + c*(cellW+gap);
      const y = gap + r*(cellH+gap);

      ctx.fillStyle = rgba(it.s.color, 0.75);
      ctx.fillRect(x,y,cellW,cellH);

      ctx.strokeStyle = "rgba(255,255,255,0.22)";
      ctx.strokeRect(x+0.5,y+0.5,cellW-1,cellH-1);

      ctx.fillStyle = txt;
      ctx.fillText(`${it.s.name} • ${it.v}%`, x+10, y+10);
    });
  }

  function renderRibbon(canvas, model){
    const ctx = setupCanvasDPI(canvas);
    const w = canvas.getBoundingClientRect().width;
    const h = canvas.getBoundingClientRect().height;

    ctx.clearRect(0,0,w,h);
    ctx.fillStyle = "rgba(255,255,255,0.03)";
    ctx.fillRect(0,0,w,h);

    // periods
    if(!model.labels || model.labels.length===0) model.labels = ["Q1","Q2","Q3","Q4"];

    if(!model.series || model.series.length===0){
      const cats = ["Online","Store","Wholesale","Marketplace"];
      const colors = nicePalette(cats.length);
      model.series = cats.map((n,i)=>({ name:n, color:colors[i], values:[] }));
      // values per quarter (thickness)
      const shares = [
        [26,24,18,14],
        [24,25,16,14],
        [28,23,15,16],
        [32,22,14,18],
      ];
      for(let s=0;s<model.series.length;s++){
        model.series[s].values = shares.map(q => q[s]);
      }
    }

    const pad = 14;
    const innerW = w - pad*2;
    const innerH = h - pad*2;
    const steps = model.labels.length;
    const xStep = innerW/(steps-1);

    // compute stacked bands per step
    const sums = [];
    for(let i=0;i<steps;i++){
      sums[i] = model.series.reduce((a,s)=>a+(s.values[i]||0),0) || 1;
    }

    const bands = [];
    for(let i=0;i<steps;i++){
      let y = pad;
      bands[i] = [];
      for(let s=0;s<model.series.length;s++){
        const v = model.series[s].values[i]||0;
        const bh = innerH*(v/sums[i]);
        bands[i][s] = { y0:y, y1:y+bh };
        y += bh;
      }
    }

    // draw thick curved lines representing ribbons (prototype look)
    for(let s=0;s<model.series.length;s++){
      ctx.beginPath();
      for(let i=0;i<steps;i++){
        const x = pad + i*xStep;
        const midY = (bands[i][s].y0 + bands[i][s].y1)/2;
        if(i===0) ctx.moveTo(x, midY);
        else{
          const px = pad + (i-1)*xStep;
          const pY = (bands[i-1][s].y0 + bands[i-1][s].y1)/2;
          const cx = (px + x)/2;
          ctx.quadraticCurveTo(cx, pY, x, midY);
        }
      }
      ctx.strokeStyle = rgba(model.series[s].color,1);
      ctx.lineWidth = 7;
      ctx.lineCap = "round";
      ctx.stroke();
    }

    const themeKey = (model.themePreset==="inherit") ? state.dashboard.theme : model.themePreset;
    const txt = (THEMES[themeKey]||THEMES.dark).chartText;
    const grid = (THEMES[themeKey]||THEMES.dark).gridLine;

    ctx.font = "12px Segoe UI, Arial";
    ctx.fillStyle = txt;
    ctx.textAlign = "center";
    for(let i=0;i<steps;i++){
      const x = pad + i*xStep;
      ctx.fillText(model.labels[i], x, h-20);

      ctx.strokeStyle = grid;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x, pad);
      ctx.lineTo(x, h-34);
      ctx.stroke();
    }
  }

  // Throttled redraw to avoid ResizeObserver loops
  function scheduleCustomRedraw(model){
    if(!model.customCanvas) return;
    if(model._raf) cancelAnimationFrame(model._raf);
    model._raf = requestAnimationFrame(() => {
      if(model.type==="treemap") renderTreemap(model.customCanvas, model);
      if(model.type==="ribbon") renderRibbon(model.customCanvas, model);
    });
  }

  // ---------- Render visual ----------
  function renderVisual(model){
    // clear body
    model.bodyEl.innerHTML = "";

    // title & style
    model.titleEl.textContent = model.title;
    model.titleEl.style.color = model.titleColor;
    model.rootEl.style.background = model.background;

    // cleanup old chart
    if(model.chart){
      try{ model.chart.destroy(); }catch{}
      model.chart = null;
    }
    // disconnect old resize observer
    if(model._ro){
      try{ model._ro.disconnect(); }catch{}
      model._ro = null;
    }

    // image visual
    if(model.type === "image"){
      const img = document.createElement("img");
      img.alt = "Uploaded image";
      img.src = model.imageSrc;
      model.bodyEl.appendChild(img);
      model.imageEl = img;
      return;
    }

    // custom prototypes
    if(model.type === "treemap" || model.type === "ribbon"){
      const c = document.createElement("canvas");
      model.bodyEl.appendChild(c);
      model.customCanvas = c;

      // observe resize safely
      let raf = null;
      const ro = new ResizeObserver(() => {
        if(raf) cancelAnimationFrame(raf);
        raf = requestAnimationFrame(() => scheduleCustomRedraw(model));
      });
      ro.observe(model.bodyEl);
      model._ro = ro;

      scheduleCustomRedraw(model);
      return;
    }

    // Chart.js
    const canvas = document.createElement("canvas");
    model.bodyEl.appendChild(canvas);

    const cfg = buildChartConfig(model.type, model);
    if(!cfg){
      const ctx = canvas.getContext("2d");
      ctx.fillText("Unsupported visual", 10, 20);
      return;
    }

    const chart = cfg.type
      ? new Chart(canvas, cfg)
      : new Chart(canvas, { type:"bar", ...cfg });

    model.chart = chart;

    const appliedTheme = (model.themePreset==="inherit") ? state.dashboard.theme : model.themePreset;
    applyChartThemeToChart(model.chart, appliedTheme);
    model.chart.update();
  }

  // ---------- Series editor ----------
  function buildSeriesEditor(model){
    clearSeriesEditor();

    // Image visual has no series
    if(model.type === "image"){
      const msg = document.createElement("div");
      msg.className = "hint";
      msg.textContent = "Image visual: no series/legend editing.";
      el.seriesEditor.appendChild(msg);
      return;
    }

    // Ensure model.series exists for chart types
    if(!model.series || model.series.length===0){
      // derive from chart datasets if possible
      if(model.chart){
        model.series = model.chart.data.datasets.map((d,i)=>({
          name: d.label || `Series ${i+1}`,
          color: "#3b82f6",
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
      color.value = s.color.startsWith("#") ? s.color : "#3b82f6";
      color.setAttribute("aria-label", `Series color ${idx+1}`);

      name.addEventListener("input", () => {
        s.name = name.value || `Series ${idx+1}`;
        applySeriesEdits(model);
      });
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
    // Custom prototype redraw
    if(model.type==="treemap" || model.type==="ribbon"){
      scheduleCustomRedraw(model);
      return;
    }

    if(!model.chart) return;

    // Pie/Donut: series are slices
    if(model.type==="pie" || model.type==="donut"){
      const labels = model.series.map(s=>s.name);
      const colors = model.series.map(s=>s.color);
      const values = model.series.map(s=>(s.values?.[0] ?? 10));
      model.chart.data.labels = labels;
      model.chart.data.datasets[0].data = values;
      model.chart.data.datasets[0].backgroundColor = colors.map(c=>rgba(c,0.85));
      model.chart.data.datasets[0].borderColor = colors.map(c=>rgba(c,1));
      model.chart.update();
      return;
    }

    // Scatter: each dataset
    if(model.type==="scatter"){
      model.chart.data.datasets.forEach((d,i)=>{
        const s = model.series[i];
        if(!s) return;
        d.label = s.name;
        d.backgroundColor = rgba(s.color,0.8);
        d.borderColor = rgba(s.color,1);
      });
      model.chart.update();
      return;
    }

    // Combo charts
    if(model.type==="lineClusteredColumn" || model.type==="lineStackedColumn"){
      model.chart.data.datasets.forEach((d,i)=>{
        const s = model.series[i];
        if(!s) return;
        d.label = s.name;
        if(d.type==="line"){
          d.borderColor = rgba(s.color,1);
          d.backgroundColor = rgba(s.color,0.15);
        }else{
          d.backgroundColor = rgba(s.color,0.75);
          d.borderColor = rgba(s.color,1);
        }
      });
      model.chart.update();
      return;
    }

    // Line/Area/Stacked Area
    if(model.type==="line" || model.type==="area" || model.type==="stackedArea"){
      const isArea = (model.type!=="line");
      model.chart.data.datasets.forEach((d,i)=>{
        const s = model.series[i];
        if(!s) return;
        d.label = s.name;
        d.borderColor = rgba(s.color,1);
        d.backgroundColor = rgba(s.color, isArea ? 0.25 : 0.10);
      });
      model.chart.update();
      return;
    }

    // Bars/Columns
    model.chart.data.datasets.forEach((d,i)=>{
      const s = model.series[i];
      if(!s) return;
      d.label = s.name;
      d.backgroundColor = rgba(s.color,0.75);
      d.borderColor = rgba(s.color,1);
    });
    model.chart.update();
  }

  // ---------- Add / remove visuals ----------
  function addVisual(visualKey, displayName){
    const id = String(state.nextId++);
    const size = makeDefaultVisualSize();
    const pos = makeDefaultPlacement();

    // GridStack v11: use addWidget(options) and then inject content
    const widgetEl = state.grid.addWidget({
      x: pos.x, y: pos.y,
      w: size.w, h: size.h,
      id
    });

    const content = widgetEl.querySelector(".grid-stack-item-content");
    const shell = createShell(id, displayName);
    content.appendChild(shell.wrapper);

    const appliedTheme = state.dashboard.theme;

    const model = {
      id,
      type: visualKey,
      name: displayName,
      title: displayName,
      themePreset: "inherit",

      background: state.dashboard.defaultVisualBg,
      titleColor: state.dashboard.defaultTitleColor,
      userSetBg: false,
      userSetTitleColor: false,
      lastBgHex: null,

      rootEl: shell.wrapper,
      titleEl: shell.titleEl,
      bodyEl: shell.body,
      deleteBtn: shell.delBtn,

      chart: null,
      customCanvas: null,
      imageEl: null,
      imageSrc: null,

      labels: [],
      series: []
    };

    // apply defaults
    model.rootEl.style.background = model.background;
    model.titleEl.style.color = model.titleColor;
    model.rootEl.style.borderColor = (THEMES[appliedTheme]||THEMES.dark).border;

    // selection
    shell.wrapper.addEventListener("mousedown", (e) => {
      // allow chart clicks but keep selection
      if(e.target.closest("button") || e.target.closest("input") || e.target.closest("select")) return;
      select(id);
      e.stopPropagation();
    });

    // delete
    shell.delBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(id);
    });

    // render
    state.visuals.set(id, model);
    renderVisual(model);
    select(id);
  }

  function addImageVisual(file){
    const id = String(state.nextId++);
    const size = makeDefaultVisualSize();
    const pos = makeDefaultPlacement();

    const widgetEl = state.grid.addWidget({
      x: pos.x, y: pos.y,
      w: size.w, h: size.h,
      id
    });

    const content = widgetEl.querySelector(".grid-stack-item-content");
    const shell = createShell(id, "Image");
    content.appendChild(shell.wrapper);

    const model = {
      id,
      type: "image",
      name: "Image",
      title: "Image",
      themePreset: "inherit",

      background: state.dashboard.defaultVisualBg,
      titleColor: state.dashboard.defaultTitleColor,
      userSetBg: false,
      userSetTitleColor: false,
      lastBgHex: null,

      rootEl: shell.wrapper,
      titleEl: shell.titleEl,
      bodyEl: shell.body,
      deleteBtn: shell.delBtn,

      chart: null,
      customCanvas: null,
      imageEl: null,
      imageSrc: null,

      labels: [],
      series: []
    };

    model.rootEl.style.background = model.background;
    model.titleEl.style.color = model.titleColor;

    // selection
    shell.wrapper.addEventListener("mousedown", (e) => {
      if(e.target.closest("button")) return;
      select(id);
      e.stopPropagation();
    });

    // delete
    shell.delBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeVisual(id);
    });

    // read file
    const reader = new FileReader();
    reader.onload = () => {
      model.imageSrc = String(reader.result);
      renderVisual(model);
    };
    reader.readAsDataURL(file);

    state.visuals.set(id, model);
    select(id);
  }

  function removeVisual(id){
    const m = state.visuals.get(id);
    if(!m) return;

    // destroy chart
    if(m.chart){
      try{ m.chart.destroy(); }catch{}
      m.chart = null;
    }
    if(m._ro){
      try{ m._ro.disconnect(); }catch{}
      m._ro = null;
    }

    // remove widget element (find by gs-id)
    const node = [...el.grid.querySelectorAll(".grid-stack-item")].find(n => {
      const nid = n.getAttribute("gs-id") || n.getAttribute("data-gs-id") || n.gridstackNode?.id;
      return String(nid) === String(id);
    });
    if(node) state.grid.removeWidget(node);

    state.visuals.delete(id);
    if(state.selectedId === id) deselect();
  }

  // ---------- Grid init ----------
  function initGrid(){
    // Fixed 1280x720 => 720/30 = 24 rows. We cap rows to prevent infinite scroll.
    state.grid = GridStack.init({
      column: 24,
      cellHeight: 30,
      margin: 6,
      float: true,
      disableOneColumnMode: true,
      animate: true,
      alwaysShowResizeHandle: true,

      // IMPORTANT: allow overlap
      disableCollision: true,

      // keep inside canvas (best-effort)
      draggable: { handle: ".visual-header, .visual-body" },
      resizable: { handles: "se, sw, ne, nw, e, w, n, s" },

      minRow: 24,
      maxRow: 24
    }, el.grid);

    // click empty canvas => dashboard context
    el.canvasClickCatcher.addEventListener("mousedown", () => deselect());

    // keep charts responsive on resize
    state.grid.on("resizestop", (_e, item) => {
      const id = String(item?.id ?? item?.el?.gridstackNode?.id);
      const m = state.visuals.get(id);
      if(!m) return;
      if(m.chart) m.chart.resize();
      if(m.customCanvas) scheduleCustomRedraw(m);
    });

    // clamp within fixed grid bounds (x within 0..24-w, y within 0..24-h)
    state.grid.on("dragstop", (_e, item) => {
      const node = item?.el?.gridstackNode;
      if(!node) return;
      const nx = clamp(node.x, 0, 24 - node.w);
      const ny = clamp(node.y, 0, 24 - node.h);
      if(nx !== node.x || ny !== node.y){
        state.grid.update(item.el, { x: nx, y: ny });
      }
    });
  }

  // ---------- Dropdown init ----------
  function initDropdowns(){
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

  // ---------- Bind UI ----------
  function bindUI(){
    // add visual
    el.addVisualBtn.addEventListener("click", () => {
      const key = el.visualType.value;
      const cat = VISUAL_CATALOG[Number(el.visualCategory.value)] || VISUAL_CATALOG[0];
      const it = cat.items.find(x => x.key === key);
      addVisual(key, it ? it.name : key);
    });

    // deselect
    el.deselectBtn.addEventListener("click", () => deselect());

    // upload image
    el.imageUpload.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if(!file) return;
      addImageVisual(file);
      e.target.value = "";
    });

    // dashboard background
    el.dashboardBg.addEventListener("input", () => setDashboardBackground(el.dashboardBg.value));
    el.dashboardBgReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      setDashboardBackground(t.dashboardBg);
    });

    // dashboard theme
    el.dashboardTheme.addEventListener("change", () => applyDashboardTheme(el.dashboardTheme.value));

    // default visual background (new visuals only)
    el.defaultVisualBg.addEventListener("input", () => {
      // store as rgba overlay
      state.dashboard.defaultVisualBg = rgba(el.defaultVisualBg.value, 0.10);
      cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
    });
    el.defaultVisualBgReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      state.dashboard.defaultVisualBg = t.defaultVisualBg;
      cssVarSet("--default-visual-bg", state.dashboard.defaultVisualBg);
      refreshDashboardUI();
    });

    // default title color (new visuals only)
    el.defaultTitleColor.addEventListener("input", () => {
      state.dashboard.defaultTitleColor = el.defaultTitleColor.value;
      cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);
    });
    el.defaultTitleReset.addEventListener("click", () => {
      const t = THEMES[state.dashboard.theme] || THEMES.dark;
      state.dashboard.defaultTitleColor = t.defaultTitleColor;
      cssVarSet("--default-title-color", state.dashboard.defaultTitleColor);
      refreshDashboardUI();
    });

    // visual title
    el.visualTitle.addEventListener("input", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      m.title = el.visualTitle.value || m.name;
      m.titleEl.textContent = m.title;
    });

    // visual background
    el.visualBg.addEventListener("input", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      m.lastBgHex = el.visualBg.value;
      m.background = rgba(el.visualBg.value, 0.10);
      m.userSetBg = true;
      m.rootEl.style.background = m.background;
    });
    el.visualBgReset.addEventListener("click", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      const themeKey = (m.themePreset==="inherit") ? state.dashboard.theme : m.themePreset;
      const t = THEMES[themeKey] || THEMES.dark;
      m.background = t.defaultVisualBg;
      m.userSetBg = false;
      m.lastBgHex = null;
      m.rootEl.style.background = m.background;
      refreshVisualUI();
    });

    // visual theme preset apply
    el.applyVisualTheme.addEventListener("click", () => {
      const m = state.visuals.get(state.selectedId);
      if(!m) return;
      applyVisualThemePreset(m, el.visualTheme.value);
    });
  }

  // ---------- Boot ----------
  function init(){
    initDropdowns();
    initGrid();
    bindUI();

    // initial theme
    applyDashboardTheme(state.dashboard.theme);

    // show dashboard formatting by default
    deselect();

    // Optional seed visuals (comment out if you want blank start)
    addVisual("line", "Line Chart (multiple trends)");
    addVisual("donut", "Donut Chart");
    deselect();
  }

  // Override to ensure proper UI refresh order
  function refreshVisualUI(){
    const m = state.visuals.get(state.selectedId);
    if(!m) return;

    el.visualTitle.value = m.title;
    el.visualTheme.value = m.themePreset || "inherit";
    el.visualBg.value = m.lastBgHex || (
      (m.themePreset==="brand") ? "#3b82f6" :
      (m.themePreset==="light") ? "#ffffff" : "#111827"
    );
    buildSeriesEditor(m);
  }

  init();
})();
