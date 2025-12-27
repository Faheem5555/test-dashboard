const canvas = document.getElementById("canvas");
const addBtn = document.getElementById("addVisualBtn");
const visualType = document.getElementById("visualType");

const seriesColors = document.getElementById("seriesColors");
const fmtTitle = document.getElementById("fmtTitle");

const fmtDashboardBg = document.getElementById("fmtDashboardBg");
const fmtVisualBg = document.getElementById("fmtVisualBg");
const fmtVisualBorder = document.getElementById("fmtVisualBorder");

const imageUpload = document.getElementById("imageUpload");

let visuals = {};
let selectedId = null;
let counter = 0;

const palette = ["#ff4d4d","#4deeea","#5da5da","#60bd68","#264478"];

const categories = ["Apparel","Footwear","Accessories","Beauty","Home"];
const values = [32,24,18,14,12];

function addVisual(type){
  const id = `v${++counter}`;
  const el = document.createElement("div");
  el.className = "visual";
  el.style.width = "380px";
  el.style.height = "260px";
  el.style.left = `${40 + counter*10}px`;
  el.style.top = `${40 + counter*10}px`;

  const header = document.createElement("div");
  header.className = "visualHeader";

  const title = document.createElement("div");
  title.className = "visualTitle";
  title.textContent = type.toUpperCase();

  const del = document.createElement("span");
  del.className = "visualDelete";
  del.textContent = "✖";
  del.onclick = () => el.remove();

  header.append(title, del);

  const body = document.createElement("div");
  body.className = "visualBody";

  el.append(header, body);
  canvas.appendChild(el);

  let chart = null;
  let seriesMeta = [];

  if(type === "pie" || type === "donut"){
    const ctx = document.createElement("canvas");
    body.appendChild(ctx);

    chart = new Chart(ctx,{
      type: type === "pie" ? "pie" : "doughnut",
      data:{
        labels:[...categories],
        datasets:[{
          data: values,
          backgroundColor: [...palette]
        }]
      }
    });

    seriesMeta = categories.map((c,i)=>({
      label:c,
      getColor:()=>chart.data.datasets[0].backgroundColor[i],
      setColor:(col)=>{
        chart.data.datasets[0].backgroundColor[i]=col;
        chart.update();
      },
      setLabel:(txt)=>{
        chart.data.labels[i]=txt;
        seriesMeta[i].label=txt;
        chart.update();
      }
    }));
  }

  visuals[id]={el,title,chart,seriesMeta};
  el.onclick = ()=>selectVisual(id);
  selectVisual(id);
}

function selectVisual(id){
  selectedId=id;
  document.querySelectorAll(".visual").forEach(v=>v.classList.remove("selected"));
  visuals[id].el.classList.add("selected");
  fmtTitle.value = visuals[id].title.textContent;
  renderColors();
}

function renderColors(){
  seriesColors.innerHTML="";
  const v = visuals[selectedId];
  v.seriesMeta.forEach(s=>{
    const row = document.createElement("div");
    row.className="seriesRow";

    const name = document.createElement("input");
    name.className="seriesNameInput";
    name.value=s.label;
    name.oninput=()=>s.setLabel(name.value);

    const col = document.createElement("input");
    col.type="color";
    col.className="colorInput";
    col.value=s.getColor();
    col.oninput=()=>s.setColor(col.value);

    row.append(name,col);
    seriesColors.appendChild(row);
  });
}

fmtDashboardBg.oninput=()=>canvas.style.background=fmtDashboardBg.value;
fmtVisualBg.oninput=()=>visuals[selectedId].el.style.background=fmtVisualBg.value;
fmtVisualBorder.oninput=()=>visuals[selectedId].el.style.borderColor=fmtVisualBorder.value;

addBtn.onclick=()=>addVisual(visualType.value);
