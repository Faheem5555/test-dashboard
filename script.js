/* =========================
   Power BI Prototype Dashboard
   (FULL UPDATED VERSION)
========================= */

(() => {

  /* ---------- DOM ---------- */
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
  const imageUpload = document.getElementById("imageUpload");

  /* ---------- CONSTANTS ---------- */
  const CANVAS_W = 1280;
  const CANVAS_H = 720;
  const DEFAULT_W = 380;
  const DEFAULT_H = 260;

  const clamp = (n, min, max) => Math.max(min, Math.min(max, n));

  /* ---------- DUMMY DATA ---------- */
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const profitThisYear = [42,45,48,52,56,60,65,70,76,83,91,100];
  const profitLastYear = profitThisYear.map(v => Math.round(v / 1.8));
  const salesThisYear = [220,230,245,265,290,315,350,380,410,450,490,540];
  const salesLastYear = salesThisYear.map(v => Math.round(v / 1.18));
  const categories = ["Apparel","Footwear","Accessories","Beauty","Home"];
  const categoryShare = [32,24,18,14,12];

  const scatterPoints = Array.from({ length: 18 }, (_, i) => ({
    x: Math.round(50 + i * 18 + (Math.random() * 20 - 10)),
    y: Math.round((12 + Math.sin(i/2)*6 + Math.random()*4) * 10) / 10
  }));

  /* ---------- STATE ---------- */
  let visualCounter = 0;
  let selectedId = null;
  let zCounter = 10;
  const visuals = {};

  const palette = [
    "#2F5597","#ED7D31","#A5A5A5","#FFC000",
    "#5B9BD5","#70AD47","#264478","#9E480E"
  ];

  /* ---------- HELPERS ---------- */
  const nextId = () => `v_${++visualCounter}`;

  const withAlpha = (hex, a) => {
    const h = hex.replace("#","");
    const r = parseInt(h.substr(0,2),16);
    const g = parseInt(h.substr(2,2),16);
    const b = parseInt(h.substr(4,2),16);
    return `rgba(${r},${g},${b},${a})`;
  };

  const normalizeToHex = c => {
    if (!c) return "#000000";
    if (c.startsWith("#")) return c;
    const m = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    return m ? `#${(+m[1]).toString(16).padStart(2,"0")}${(+m[2]).toString(16).padStart(2,"0")}${(+m[3]).toString(16).padStart(2,"0")}` : "#000000";
  };

  /* ---------- SELECTION ---------- */
  const deselectAll = () => {
    selectedId = null;
    Object.values(visuals).forEach(v => v.el.classList.remove("selected"));
    renderFormatPane();
  };

  const selectVisual = id => {
    selectedId = id;
    Object.values(visuals).forEach(v => v.el.classList.remove("selected"));
    visuals[id]?.el.classList.add("selected");
    renderFormatPane();
  };

  canvas.addEventListener("mousedown", e => {
    if (e.target === canvas) deselectAll();
  });

  /* ---------- DRAG + RESIZE ---------- */
  function enableDragResize(el, header, id) {
    let drag = null, resize = null;

    header.onmousedown = e => {
      if (e.target.closest(".visualDelete")) return;
      selectVisual(id);
      const r = el.getBoundingClientRect();
      const c = canvas.getBoundingClientRect();
      drag = { x:e.clientX, y:e.clientY, l:r.left-c.left, t:r.top-c.top };
      e.preventDefault();
    };

    el.querySelectorAll(".resizeHandle").forEach(h => {
      h.onmousedown = e => {
        selectVisual(id);
        const r = el.getBoundingClientRect();
        const c = canvas.getBoundingClientRect();
        resize = {
          h:h.dataset.handle,
          x:e.clientX, y:e.clientY,
          l:r.left-c.left, t:r.top-c.top,
          w:r.width, h0:r.height
        };
        e.preventDefault(); e.stopPropagation();
      };
    });

    window.onmousemove = e => {
      if (drag) {
        el.style.left = clamp(drag.l + e.clientX-drag.x,0,CANVAS_W-el.offsetWidth)+"px";
        el.style.top  = clamp(drag.t + e.clientY-drag.y,0,CANVAS_H-el.offsetHeight)+"px";
        syncFormat();
      }
      if (resize) {
        let w = resize.w + (e.clientX-resize.x);
        let h = resize.h0 + (e.clientY-resize.y);
        w = Math.max(240,w); h = Math.max(180,h);
        el.style.width = w+"px";
        el.style.height = h+"px";
        visuals[id].chart?.resize();
        visuals[id].ribbon?.render();
        syncFormat();
      }
    };
    window.onmouseup = () => drag = resize = null;
  }

  function syncFormat(){
    const v = visuals[selectedId];
    if(!v) return;
    fmtX.value = parseInt(v.el.style.left);
    fmtY.value = parseInt(v.el.style.top);
    fmtW.value = parseInt(v.el.style.width);
    fmtH.value = parseInt(v.el.style.height);
  }

  /* ---------- VISUAL CREATION ---------- */
  function createVisual(type){
    const id = nextId();
    const el = document.createElement("div");
    el.className = "visual";
    el.style.left = "40px";
    el.style.top = "40px";
    el.style.width = DEFAULT_W+"px";
    el.style.height = DEFAULT_H+"px";
    el.style.zIndex = ++zCounter;

    const header = document.createElement("div");
    header.className = "visualHeader";

    const title = document.createElement("div");
    title.className = "visualTitle";
    title.textContent = "Visual";

    const del = document.createElement("button");
    del.className = "visualDelete";
    del.textContent = "✖";
    del.onclick = e => { e.stopPropagation(); el.remove(); delete visuals[id]; deselectAll(); };

    header.append(title,del);

    const body = document.createElement("div");
    body.className = "visualBody";

    ["se","sw","ne","nw"].forEach(h=>{
      const d=document.createElement("div");
      d.className=`resizeHandle h-${h}`;
      d.dataset.handle=h;
      el.appendChild(d);
    });

    el.append(header,body);
    canvas.appendChild(el);
    enableDragResize(el,header,id);

    visuals[id]={ id, type, el, titleEl:title, bodyEl:body, chart:null, ribbon:null, seriesMeta:[] };
    el.onmousedown=e=>{e.stopPropagation();selectVisual(id);};
    selectVisual(id);
    return visuals[id];
  }

  /* ---------- CHART HELPERS ---------- */
  const chartCanvas = p => {
    const c=document.createElement("canvas");
    p.appendChild(c);
    return c.getContext("2d");
  };

  const baseOpts = () => ({
    responsive:true,
    maintainAspectRatio:false,
    plugins:{ legend:{display:true}, tooltip:{enabled:true} }
  });

  function setSeriesMeta(v){
    const m=[];
    const ch=v.chart;
    if(!ch) return;

    if(ch.data.labels){
      ch.data.labels.forEach((l,i)=>m.push({
        label:l,
        getColor:()=>ch.data.datasets[0].backgroundColor[i],
        setColor:c=>{ch.data.datasets[0].backgroundColor[i]=c;ch.update();}
      }));
    } else {
      ch.data.datasets.forEach(ds=>m.push({
        label:ds.label,
        getColor:()=>ds.borderColor||ds.backgroundColor,
        setColor:c=>{ds.borderColor=c;ds.backgroundColor=withAlpha(c,.25);ch.update();}
      }));
    }
    v.seriesMeta=m;
  }

  /* ---------- BUILDERS ---------- */
  function buildPie(v){
    const ctx=chartCanvas(v.bodyEl);
    v.chart=new Chart(ctx,{
      type:"pie",
      data:{labels:categories,datasets:[{data:categoryShare,backgroundColor:palette}]},
      options:baseOpts()
    });
    setSeriesMeta(v);
  }

  /* ---------- ADD VISUAL ---------- */
  async function addVisual(type){
    const v=createVisual(type);
    if(type==="pie") buildPie(v);
    renderFormatPane();
  }

  addVisualBtn.onclick=()=>addVisual(visualType.value);

  /* ---------- FORMAT PANE ---------- */
  function renderFormatPane(){
    const v=visuals[selectedId];
    if(!v){
      seriesColors.innerHTML=`<div class="emptyState">No visual selected.</div>`;
      fmtTitle.disabled=true;
      return;
    }

    fmtTitle.disabled=false;
    fmtTitle.value=v.titleEl.textContent;
    seriesColors.innerHTML="";

    if(!v.seriesMeta.length){
      seriesColors.innerHTML=`<div class="emptyState">No series</div>`;
      return;
    }

    v.seriesMeta.forEach((s,i)=>{
      const row=document.createElement("div");
      row.className="seriesRow";

      const name=document.createElement("input");
      name.className="seriesNameInput";
      name.value=s.label;
      name.oninput=()=>{
        s.label=name.value;
        if(v.chart?.data?.labels) v.chart.data.labels[i]=name.value;
        if(v.chart?.data?.datasets?.[i]) v.chart.data.datasets[i].label=name.value;
        v.chart?.update();
        v.ribbon?.render();
      };

      const col=document.createElement("input");
      col.type="color";
      col.className="colorInput";
      col.value=normalizeToHex(s.getColor());
      col.oninput=()=>s.setColor(col.value);

      row.append(name,col);
      seriesColors.appendChild(row);
    });
  }

  fmtTitle.oninput=()=>{
    const v=visuals[selectedId];
    if(v) v.titleEl.textContent=fmtTitle.value;
  };

  /* ---------- INIT ---------- */
  addVisual("pie");

})();
