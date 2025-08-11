// ---------- helpers ----------
const uniq = arr => [...new Set(arr.filter(Boolean))];
const currentYear = new Date().getFullYear();
const years = Array.from({ length: 31 }, (_, i) => currentYear - 20 + i);
const industries = [
  "Agriculture","Automotive","Aviation","Cement","Chemicals","Construction","Consumer Goods","Education",
  "Energy & Utilities","Financial Services","Food & Beverage","Government","Healthcare","Hospitality",
  "IT & Telecom","Logistics","Manufacturing","Metals & Mining","Oil & Gas","Pharmaceuticals","Real Estate",
  "Retail & E-commerce","Textiles","Transportation","Waste Management","Other"
];

// ---------- Scope 1/2 ----------
function unitsForActivity12(title){ return uniq(factors12.filter(f=>f.Title===title).map(f=>f.Unit)); }
function factorFor12(title, unit){ const h=factors12.find(f=>f.Title===title && f.Unit===unit); return h?Number(h.Factor):0; }

// ---------- Scope 3 (normalized workbook) ----------
const s3Categories = uniq(scope3.map(r=>r.Category)).sort();
function s3Types(category){ return uniq(scope3.filter(r=>r.Category===category).map(r=>r.Type)).sort(); }
function s3Groups(category, type){ return uniq(scope3.filter(r=>r.Category===category && r.Type===type).map(r=>r.Group)).sort(); }
function s3Subtypes(category, type, group){ return uniq(scope3.filter(r=>r.Category===category && r.Type===type && r.Group===group).map(r=>r.Subtype)).sort(); }
function s3Units(category, type, group, subtype){ return uniq(scope3.filter(r=>r.Category===category && r.Type===type && r.Group===group && r.Subtype===subtype).map(r=>r.Unit)); }
function s3Factor(category, type, group, subtype, unit){
  const r = scope3.find(r=>r.Category===category && r.Type===type && r.Group===group && r.Subtype===subtype && r.Unit===unit);
  return r?Number(r.Factor):0;
}

function optionize(list, selected=""){
  return list.map(v=>`<option value="${String(v)}" ${String(v)===String(selected)?"selected":""}>${v}</option>`).join("");
}

// ---------- Row builders ----------
function createRowS1S2(scope="Scope 1"){
  const tr = document.createElement("tr");
  const activity = activities[0] || "";
  const units = unitsForActivity12(activity);
  const unit = units[0] || "";
  const factor = factorFor12(activity, unit);

  tr.innerHTML = `
    <td><select class="country">${optionize(countries,"India")}</select></td>
    <td><select class="year">${optionize(years,currentYear)}</select></td>
    <td><select class="industry">${optionize(industries,"Manufacturing")}</select></td>
    <td><select class="activity">${optionize(activities, activity)}</select></td>
    <td><input type="number" class="quantity" min="0" step="any" placeholder="0.00"></td>
    <td><select class="unit">${optionize(units, unit)}</select></td>
    <td class="factor-cell">${factor || "—"}</td>
    <td><textarea class="description" maxlength="1024" placeholder="Add details (max 1024 chars)"></textarea></td>
    <td><button class="delete-row">Remove</button></td>
  `;
  tr.querySelector(".delete-row").onclick = () => tr.remove();

  const act = tr.querySelector(".activity");
  const uni = tr.querySelector(".unit");
  const cell = tr.querySelector(".factor-cell");
  act.onchange = () => { const u=unitsForActivity12(act.value); uni.innerHTML=optionize(u,u[0]); cell.textContent=factorFor12(act.value,uni.value)||"—"; };
  uni.onchange = () => { cell.textContent = factorFor12(act.value, uni.value) || "—"; };

  tr.dataset.scope = scope;
  return tr;
}

function createRowS3(){
  const tr = document.createElement("tr");

  const cat = s3Categories[0] || "";
  const flowType = s3Types(cat)[0] || "";           // 'Type' from workbook
  const group = s3Groups(cat, flowType)[0] || "";
  const sub = s3Subtypes(cat, flowType, group)[0] || "";
  const units = s3Units(cat, flowType, group, sub);
  const unit = units[0] || "";
  const factor = s3Factor(cat, flowType, group, sub, unit);

  tr.innerHTML = `
    <td><select class="country">${optionize(countries,"India")}</select></td>
    <td><select class="year">${optionize(years,currentYear)}</select></td>
    <td><select class="industry">${optionize(industries,"Manufacturing")}</select></td>

    <td>
      <select class="flowDirection">
        <option value="Upstream">Upstream</option>
        <option value="Downstream">Downstream</option>
      </select>
    </td>

    <td><select class="category">${optionize(s3Categories, cat)}</select></td>
    <td><select class="flowType">${optionize(s3Types(cat), flowType)}</select></td>
    <td><select class="group">${optionize(s3Groups(cat, flowType), group)}</select></td>
    <td><select class="subtype">${optionize(s3Subtypes(cat, flowType, group), sub)}</select></td>
    <td><select class="unit">${optionize(units, unit)}</select></td>

    <td class="factor-cell">${factor || "—"}</td>
    <td><input type="number" class="quantity" min="0" step="any" placeholder="0.00"></td>
    <td><textarea class="description" maxlength="1024" placeholder="Add details (max 1024 chars)"></textarea></td>
    <td><button class="delete-row">Remove</button></td>
  `;
  tr.querySelector(".delete-row").onclick = () => tr.remove();

  // Cascading selects (flowDirection is informational; not filtering workbook yet)
  const selCat = tr.querySelector(".category");
  const selType = tr.querySelector(".flowType");
  const selGroup = tr.querySelector(".group");
  const selSub = tr.querySelector(".subtype");
  const selUnit = tr.querySelector(".unit");
  const cell = tr.querySelector(".factor-cell");

  selCat.onchange = () => {
    const t = s3Types(selCat.value);
    selType.innerHTML = optionize(t, t[0]);
    const g = s3Groups(selCat.value, selType.value);
    selGroup.innerHTML = optionize(g, g[0]);
    const s = s3Subtypes(selCat.value, selType.value, selGroup.value);
    selSub.innerHTML = optionize(s, s[0]);
    const u = s3Units(selCat.value, selType.value, selGroup.value, selSub.value);
    selUnit.innerHTML = optionize(u, u[0]);
    cell.textContent = s3Factor(selCat.value, selType.value, selGroup.value, selSub.value, selUnit.value) || "—";
  };
  selType.onchange = () => {
    const g = s3Groups(selCat.value, selType.value);
    selGroup.innerHTML = optionize(g, g[0]);
    const s = s3Subtypes(selCat.value, selType.value, selGroup.value);
    selSub.innerHTML = optionize(s, s[0]);
    const u = s3Units(selCat.value, selType.value, selGroup.value, selSub.value);
    selUnit.innerHTML = optionize(u, u[0]);
    cell.textContent = s3Factor(selCat.value, selType.value, selGroup.value, selSub.value, selUnit.value) || "—";
  };
  selGroup.onchange = () => {
    const s = s3Subtypes(selCat.value, selType.value, selGroup.value);
    selSub.innerHTML = optionize(s, s[0]);
    const u = s3Units(selCat.value, selType.value, selGroup.value, selSub.value);
    selUnit.innerHTML = optionize(u, u[0]);
    cell.textContent = s3Factor(selCat.value, selType.value, selGroup.value, selSub.value, selUnit.value) || "—";
  };
  selSub.onchange = () => {
    const u = s3Units(selCat.value, selType.value, selGroup.value, selSub.value);
    selUnit.innerHTML = optionize(u, u[0]);
    cell.textContent = s3Factor(selCat.value, selType.value, selGroup.value, selSub.value, selUnit.value) || "—";
  };
  selUnit.onchange = () => {
    cell.textContent = s3Factor(selCat.value, selType.value, selGroup.value, selSub.value, selUnit.value) || "—";
  };

  tr.dataset.scope = "Scope 3";
  return tr;
}

// ---------- wiring ----------
function tbodyForScope(scope){ return document.getElementById(`rows-${scope.toLowerCase().replace(/\s+/g,"")}`); }
function addRow(scope){ const tb=tbodyForScope(scope); if(!tb) return; tb.appendChild(scope==="Scope 3" ? createRowS3() : createRowS1S2(scope)); }
document.querySelectorAll(".add-row").forEach(b=>b.addEventListener("click",()=>addRow(b.dataset.scope)));
["Scope 1","Scope 2","Scope 3"].forEach(addRow);

// ---------- collect + calculate ----------
function collectRows(){
  const rows = [];
  ["Scope 1","Scope 2"].forEach(scope=>{
    const tb = tbodyForScope(scope); if(!tb) return;
    [...tb.querySelectorAll("tr")].forEach(tr=>{
      rows.push({
        scope,
        year: tr.querySelector(".year")?.value || "",
        activity: tr.querySelector(".activity")?.value || "",
        unit: tr.querySelector(".unit")?.value || "",
        quantity: Number(tr.querySelector(".quantity")?.value || 0),
      });
    });
  });
  {
    const scope="Scope 3", tb=tbodyForScope(scope); if(tb){
      [...tb.querySelectorAll("tr")].forEach(tr=>{
        rows.push({
          scope,
          year: tr.querySelector(".year")?.value || "",
          flowDirection: tr.querySelector(".flowDirection")?.value || "",   // informational
          category: tr.querySelector(".category")?.value || "",
          flowType: tr.querySelector(".flowType")?.value || "",             // used in factor key
          group: tr.querySelector(".group")?.value || "",
          subtype: tr.querySelector(".subtype")?.value || "",
          unit: tr.querySelector(".unit")?.value || "",
          quantity: Number(tr.querySelector(".quantity")?.value || 0),
          factor: Number(tr.querySelector(".factor-cell")?.textContent || 0)
        });
      });
    }
  }
  return rows;
}

document.getElementById("calculate").onclick = () => {
  const rows = collectRows();
  fetch("/calculate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({rows})})
    .then(r=>r.json())
    .then(res=>{
      document.getElementById("total").textContent  = res.total.toFixed(4);
      document.getElementById("scope1").textContent = res["Scope 1"].toFixed(4);
      document.getElementById("scope2").textContent = res["Scope 2"].toFixed(4);
      document.getElementById("scope3").textContent = res["Scope 3"].toFixed(4);
      renderCharts(rows, res);
    })
    .catch(()=>{
      ["total","scope1","scope2","scope3"].forEach(id=>document.getElementById(id).textContent="0.00");
      renderCharts([], {total:0,"Scope 1":0,"Scope 2":0,"Scope 3":0});
    });
};

document.getElementById("clear-all").onclick = () => {
  ["Scope 1","Scope 2","Scope 3"].forEach(s=>{ const tb=tbodyForScope(s); if(tb) tb.innerHTML=""; addRow(s); });
  ["total","scope1","scope2","scope3"].forEach(id=>document.getElementById(id).textContent="0.00");
  renderCharts([], {total:0,"Scope 1":0,"Scope 2":0,"Scope 3":0});
};

// ---------- charts ----------
const plotTheme = { paper_bgcolor:'#12261f', plot_bgcolor:'#12261f', font:{color:'#e2e8f0'} };
function compute(rows){ return rows.map(r=>({...r, emission:(Number(r.quantity)||0)*(Number(r.factor)||0)})); }
function sumBy(rows, key){ const m={}; rows.forEach(r=>{const k=r[key]||"Unknown"; m[k]=(m[k]||0)+(r.emission||0);}); return m; }

function renderCharts(rows, totals){
  const data = compute(rows);

  Plotly.newPlot("chart-scope-mix", [{
    type:"pie", hole:.55,
    values:[totals["Scope 1"]||0,totals["Scope 2"]||0,totals["Scope 3"]||0],
    labels:["Scope 1","Scope 2","Scope 3"],
    marker:{colors:["#34d399","#22c55e","#15803d"]}
  }], {...plotTheme, margin:{t:10,b:10,l:10,r:10}}, {displayModeBar:false});

  Plotly.newPlot("chart-scope-bars", [{
    type:"bar",
    x:["Scope 1","Scope 2","Scope 3"],
    y:[totals["Scope 1"]||0,totals["Scope 2"]||0,totals["Scope 3"]||0]
  }], {...plotTheme, yaxis:{title:"kgCO₂e"}, margin:{t:10,b:40,l:40,r:10}}, {displayModeBar:false});

  const s3rows = data.filter(r=>r.scope==="Scope 3");
  const byCat = sumBy(s3rows,"category");
  Plotly.newPlot("chart-s3-category", [{
    type:"treemap",
    labels:Object.keys(byCat),
    parents:Object.keys(byCat).map(()=> ""),
    values:Object.values(byCat),
    marker:{colors:Object.values(byCat).map(()=>"#16a34a")}
  }], {...plotTheme, margin:{t:10,b:10,l:10,r:10}}, {displayModeBar:false});

  const topSubtype = Object.entries(sumBy(s3rows,"subtype")).sort((a,b)=>b[1]-a[1]).slice(0,8);
  Plotly.newPlot("chart-s3-top", [{
    type:"bar", orientation:"h",
    x: topSubtype.map(x=>x[1]),
    y: topSubtype.map(x=>x[0])
  }], {...plotTheme, xaxis:{title:"kgCO₂e"}, margin:{t:10,b:20,l:140,r:10}}, {displayModeBar:false});

  const byYear = {};
  data.forEach(r=>{ const y=String(r.year||""); if(!byYear[y]) byYear[y]={"Scope 1":0,"Scope 2":0,"Scope 3":0}; byYear[y][r.scope]+=r.emission||0; });
  const ys = Object.keys(byYear).filter(Boolean).sort();
  Plotly.newPlot("chart-year", [
    {name:"Scope 1",type:"bar",x:ys,y:ys.map(y=>byYear[y]["Scope 1"]||0)},
    {name:"Scope 2",type:"bar",x:ys,y:ys.map(y=>byYear[y]["Scope 2"]||0)},
    {name:"Scope 3",type:"bar",x:ys,y:ys.map(y=>byYear[y]["Scope 3"]||0)}
  ], {...plotTheme, barmode:"group", yaxis:{title:"kgCO₂e"}, margin:{t:10,b:40,l:40,r:10}}, {displayModeBar:false});
}
