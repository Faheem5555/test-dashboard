/* =========================
   Power BI Prototype Dashboard
========================= */

(() => {

  // -------------------------
  // DOM
  // -------------------------
  const canvas = document.getElementById("canvas");
  const visualType = document.getElementById("visualType");
  const addVisualBtn = document.getElementById("addVisualBtn");

  const formatStatus = document.getElementById("formatStatus");
  const fmtTitle = document.getElementById("fmtTitle");
  const fmtX = document.getElementById("fmtX");
  const fmtY = document.getElementById("fmtY");
  const fmtW = document.getElementById("fmtW");
  const fmtH = document.getElementById("fmtH");
  const bringFrontBtn = document.getElementById("bringFrontBtn");
  const seriesColors = document.getElementById("seriesColors");

  // 🔹 THEME CONTROLS (added)
  const fmtDashboardBg = document.getElementById("fmtDashboardBg");
  const fmtVisualBg = document.getElementById("fmtVisualBg");
  const fmtVisualBorder = document.getElementById("fmtVisualBorder");

  const imageUpload = document.getElementById("imageUpload");

  // -------------------------
  // Canvas config
  // -------------------------
  const CANVAS_W = 1280;
  const CANVAS_H = 720;
  const DEFAULT_W = 380;
  const DEFAULT_H = 260;

  const clamp = (n, min, max) => Math.max(min, Math.min(max, n));

  // -------------------------
  // Dummy data
  // -------------------------
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const profitThisYear = [42,45,48,52,56,60,65,70,76,83,91,100];
  const profitLastYear = profitThisYear.map(v => Math.round(v / 1.8));
  const categories = ["Apparel","Footwear","Accessories","Beauty","Home"];
  const categoryShare = [32,24,18,14,12];

  // -------------------------
  // State
  // -------------------------
  let visualCounter = 0;
  let selectedId = null;
  let zCounter = 10;
  const visuals = {};

  const palette = [
    "#2F5597","#ED7D31","#A5A5A5","#FFC000",
    "#5B9BD5","#70AD47","#264478"
  ];

  // -------------------------
  // Helpers
  // -------------------------
  function nextId() {
    visualCounter++;
    return `v_${visualCounter}`;
  }

  function deselectAll() {
    selectedId = null;
    Object.values(visuals).forEach(v => v.el.classList.remove("selected"));
    renderFormatPane();
  }

  function selectVisual(id) {
    selectedId = id;
    Object.values(visuals).forEach(v => v.el.classList.remove("selected"));
    visuals[id].el.classList.add("selected");
    renderFormatPane();
  }

  canvas.addEventListener("mousedown", e => {
    if (e.target === canvas) deselectAll();
  });

  // -------------------------
  // Visual shell
  // -------------------------
  function createVisualShell(type) {
    const id = nextId();

    const visual = document.createElement("div");
    visual.className = "visual";
    visual.style.width = `${DEFAULT_W}px`;
    visual.style.height = `${DEFAULT_H}px`;
    visual.style.left = "40px";
    visual.style.top = "40px";
    visual.style.zIndex = ++zCounter;

    const header = document.createElement("div");
    header.className = "visualHeader";

    const title = document.createElement("div");
    title.className = "visualTitle";
    title.textContent = type.toUpperCase();

    const del = document.createElement("button");
    del.className = "visualDelete";
    del.textContent = "✖";
    del.onclick = () => deleteVisual(id);

    header.appendChild(title);
    header.appendChild(del);

    const body = document.createElement("div");
    body.className = "visualBody";

    visual.appendChild(header);
    visual.appendChild(body);
    canvas.appendChild(visual);

    visual.addEventListener("mousedown", e => {
      e.stopPropagation();
      selectVisual(id);
    });

    visuals[id] = { id, type, el: visual, titleEl: title, bodyEl: body, chart: null, seriesMeta: [] };
    selectVisual(id);
    return visuals[id];
  }

  function deleteVisual(id) {
    if (!visuals[id]) return;
    visuals[id].el.remove();
    delete visuals[id];
    deselectAll();
  }

  // -------------------------
  // Chart helpers
  // -------------------------
  function chartCanvas(parent) {
    const c = document.createElement("canvas");
    parent.appendChild(c);
    return c.getContext("2d");
  }

  function withAlpha(hex, a) {
    const h = hex.replace("#","");
    const r = parseInt(h.substr(0,2),16);
    const g = parseInt(h.substr(2,2),16);
    const b = parseInt(h.substr(4,2),16);
    return `rgba(${r},${g},${b},${a})`;
  }

  // -------------------------
  // SERIES META (UPDATED)
  // -------------------------
  function setSeriesMetaForChart(v, type) {
    const chart = v.chart;
    const meta = [];

    // PIE / DONUT
    if (type === "pie" || type === "donut") {
      chart.data.labels.forEach((lab, i) => {
        meta.push({
          label: lab,
          getColor: () => chart.data.datasets[0].backgroundColor[i],
          setColor: c => {
            chart.data.datasets[0].backgroundColor[i] = c;
            chart.update();
          },
          setLabel: newLabel => {
            chart.data.labels[i] = newLabel;
            meta[i].label = newLabel;
            chart.update();
          }
        });
      });
    }

    // MULTI SERIES
    chart.data.datasets.forEach((ds, i) => {
      meta.push({
        label: ds.label,
        getColor: () => ds.borderColor,
        setColor: c => {
          ds.borderColor = c;
          ds.backgroundColor = withAlpha(c, 0.25);
          chart.update();
        },
        setLabel: newLabel => {
          ds.label = newLabel;
          meta[i].label = newLabel;
          chart.update();
        }
      });
    });

    v.seriesMeta = meta;
  }

  // -------------------------
  // Visual builders
  // -------------------------
  function buildPie(v) {
    const ctx = chartCanvas(v.bodyEl);
    v.chart = new Chart(ctx, {
      type: "pie",
      data: {
        labels: categories,
        datasets: [{
          data: categoryShare,
          backgroundColor: categories.map((_,i)=>palette[i])
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
    setSeriesMetaForChart(v, "pie");
  }

  // -------------------------
  // Add visual
  // -------------------------
  function addVisual(type) {
    const v = createVisualShell(type);
    if (type === "pie") buildPie(v);
    renderFormatPane();
  }

  addVisualBtn.onclick = () => addVisual(visualType.value);

  // -------------------------
  // FORMAT PANE (UPDATED)
  // -------------------------
  function renderFormatPane() {
    const v = selectedId ? visuals[selectedId] : null;

    if (!v) {
      seriesColors.innerHTML = `<div class="emptyState">No visual selected.</div>`;
      return;
    }

    seriesColors.innerHTML = "";
    v.seriesMeta.forEach(s => {
      const row = document.createElement("div");
      row.className = "seriesRow";

      // 🔹 Editable series name
      const nameInput = document.createElement("input");
      nameInput.className = "seriesNameInput";
      nameInput.value = s.label;
      nameInput.oninput = () => s.setLabel && s.setLabel(nameInput.value);

      const color = document.createElement("input");
      color.type = "color";
      color.value = normalizeToHex(s.getColor());
      color.oninput = () => s.setColor(color.value);

      row.appendChild(nameInput);
      row.appendChild(color);
      seriesColors.appendChild(row);
    });
  }

  function normalizeToHex(c) {
    if (c.startsWith("#")) return c;
    const m = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!m) return "#000000";
    return `#${(+m[1]).toString(16).padStart(2,"0")}${(+m[2]).toString(16).padStart(2,"0")}${(+m[3]).toString(16).padStart(2,"0")}`;
  }

  // -------------------------
  // THEME CONTROLS (ADDED)
  // -------------------------
  fmtDashboardBg.oninput = () => {
    canvas.style.background = fmtDashboardBg.value;
  };

  fmtVisualBg.oninput = () => {
    if (!selectedId) return;
    visuals[selectedId].el.style.background = fmtVisualBg.value;
  };

  fmtVisualBorder.oninput = () => {
    if (!selectedId) return;
    visuals[selectedId].el.style.borderColor = fmtVisualBorder.value;
  };

  // -------------------------
  // Start demo visuals
  // -------------------------
  addVisual("pie");

})();
