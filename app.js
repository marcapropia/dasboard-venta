// ============================================================
// Sales Analytics by MP - app.js  (PARTE 1/2)
// ============================================================
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
let globalData = [];
let filters = { year: "all", month: "all", category: "all", brand: "all", mp: "all" };
let charts = {};
if (typeof lucide !== 'undefined') { lucide.createIcons(); }

const uploadView      = document.getElementById("upload-view");
const dashboardLayout = document.getElementById("dashboard-layout");
const fileInput       = document.getElementById("file-input");
const dropZone        = document.getElementById("drop-zone");
const uploadStatus    = document.getElementById("upload-status");
const statusText      = document.getElementById("status-text");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover",(e)=>{e.preventDefault();dropZone.classList.add("drag-active");});
dropZone.addEventListener("dragleave",(e)=>{e.preventDefault();dropZone.classList.remove("drag-active");});
dropZone.addEventListener("drop",(e)=>{e.preventDefault();dropZone.classList.remove("drag-active");if(e.dataTransfer.files.length)handleFile(e.dataTransfer.files[0]);});
fileInput.addEventListener("change",(e)=>{if(e.target.files.length)handleFile(e.target.files[0]);});

function handleFile(file){
    if(!file.name.match(/\.(xlsx|xls)$/i)){alert("Por favor carga un archivo Excel (.xlsx o .xls)");return;}
    uploadStatus.classList.remove("hidden");statusText.innerText="Leyendo archivo Excel...";
    const reader=new FileReader();
    reader.onload=(e)=>{
        try{
            const data=new Uint8Array(e.target.result);
            const wb=XLSX.read(data,{type:"array"});
            const ws=wb.Sheets[wb.SheetNames[0]];
            const jd=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
            statusText.innerText="Transformando y normalizando datos...";
            setTimeout(()=>processData(jd),100);
        }catch(err){console.error(err);alert("Error al leer el archivo. Verifica el formato.");uploadStatus.classList.add("hidden");}
    };
    reader.readAsArrayBuffer(file);
}

function normalizeNumber(val){
    if(typeof val==="number")return val;
    if(!val||typeof val!=="string")return 0;
    let s=val.trim();
    if(s.includes(".")&&!s.includes(",")){const p=s.split(".");if(p.length>1&&p[p.length-1].length===3)s=s.replace(/\./g,"");}
    const n=parseFloat(s);return isNaN(n)?0:n;
}
function isMonthColumn(c){if(!c||typeof c!=="string")return false;return MESES.some(m=>c.toLowerCase().includes(m.toLowerCase()));}
function extractMonthYear(c,fy){
    let fm=-1;for(let i=0;i<MESES.length;i++){if(c.toLowerCase().includes(MESES[i].toLowerCase())){fm=i;break;}}
    let y=fy||new Date().getFullYear();const m=c.match(/\b(20\d{2})\b/);if(m)y=parseInt(m[1]);return{month:fm,year:y};
}
function guessBrand(a){
    if(!a) return "Desconocida";
    const p = String(a).toUpperCase().split(/\s+/);
    let i = 0;
    const generics = ["AGUA", "PAN", "CAFE", "CAFÉ", "CERVEZA", "LECHE", "JUGO", "GALLETAS", "CHOCOLATE", "BEBIDA", "GASEOSA", "REFRESCO", "VINO", "LICOR", "BOTELLA", "LATA", "PACK", "CAJA", "BOLSA", "SIX", "DOCENA", "FRANCES", "MARRAQUETA", "DE", "MESA", "HORNEADO", "S/GAS", "C/GAS", "SABORIZADA", "UHT", "NATURAL"];
    while(i < p.length && generics.includes(p[i])) { i++; }
    if(i >= p.length) return "SIN MARCA";
    let brand = p[i];
    const prefixes = ["LA", "EL", "LOS", "LAS", "DEL", "SAN", "SANTA", "DON", "DOÑA"];
    if(prefixes.includes(brand) && i + 1 < p.length) {
        brand = brand + " " + p[i+1];
        if (brand === "LA CASCADA" || brand === "LA ESTRELLA" || brand === "LA FRANCESA") { /* known brands */ }
    }
    if(brand === "COCA" && p[i+1] === "COLA") brand = "COCA COLA";
    if(brand === "BON" && p[i+1] === "PAIN") brand = "BON PAIN";
    if(brand === "RED" && p[i+1] === "BULL") brand = "RED BULL";
    if(brand === "CHIQUI" && p[i+1] === "CHOC") brand = "CHIQUI CHOC";
    if(brand === "VILLA" && p[i+1] === "SANTA") brand = "VILLA SANTA";
    return brand;
}
function isMarcaPropia(articleName) {
    if(!articleName) return false;
    const lower = String(articleName).toLowerCase();
    return lower.includes('basix') || lower.includes('amarket') || lower.includes('bon pain');
}

function processData(matrix){
    if(matrix.length<2){alert("El archivo parece estar vac\u00edo.");return;}
    let hIdx=0;
    for(let i=0;i<Math.min(15,matrix.length);i++){
        if(matrix[i].some(c=>typeof c==="string"&&(c.toLowerCase().includes("articulo")||c.toLowerCase().includes("art\u00edculo")))){hIdx=i;break;}
    }
    const headers=matrix[hIdx].map(h=>String(h).trim());
    const colYears=[];let cy=new Date().getFullYear();
    for(let col=0;col<headers.length;col++){
        let yf=null;
        for(let r=0;r<hIdx;r++){const v=String(matrix[r][col]||"").trim();const m=v.match(/\b(20\d{2})\b/);if(m){yf=parseInt(m[1]);break;}}
        if(yf)cy=yf;colYears[col]=cy;
    }
    let ai=headers.findIndex(h=>h.toLowerCase().includes("articulo")||h.toLowerCase().includes("art\u00edculo"));
    let ci=headers.findIndex(h=>h.toLowerCase().includes("cat 3")||h.toLowerCase().includes("categoria")||h.toLowerCase().includes("categor\u00eda"));
    let bi=headers.findIndex(h=>h.toLowerCase() === "marca" || h.toLowerCase() === "brand" || h.toLowerCase().includes("marca"));
    let ni=headers.findIndex(h=>h.toLowerCase().includes("tipo de negocio") || h.toLowerCase().includes("negocio"));
    if(ai===-1)ai=0;if(ci===-1)ci=1;
    const mc=[];headers.forEach((h,idx)=>{if(isMonthColumn(h))mc.push({idx,...extractMonthYear(h,colYears[idx])});});
    if(mc.length===0){alert("No se detectaron columnas de meses en la fila "+(hIdx+1));uploadStatus.classList.add("hidden");return;}
    const nd=[];
    for(let i=hIdx+1;i<matrix.length;i++){
        const row=matrix[i];const art=row[ai];
        if(!art||String(art).toLowerCase().includes("total general")||String(art).toLowerCase()==="totales")continue;
        const cat=row[ci]||"Sin Categor\u00eda";
        const brand = (bi !== -1 && row[bi] && String(row[bi]).trim() !== "") ? String(row[bi]).trim().toUpperCase() : guessBrand(art);
        let rowIsMP = false;
        if (ni !== -1 && row[ni] !== undefined && row[ni] !== "") {
            const val = String(row[ni]).toLowerCase();
            rowIsMP = val.includes("propia") || val === "mp";
        } else {
            rowIsMP = isMarcaPropia(art) || isMarcaPropia(brand);
        }
        mc.forEach(col=>{
            const us=row[col.idx];const units=(us===""||us===undefined)?0:normalizeNumber(us);
            const isMP = rowIsMP;
            nd.push({article:String(art).trim(),category:String(cat).trim(),brand,year:col.year,month:col.month,monthName:MESES[col.month],units,isMP});
        });
    }
    globalData=nd;
    const artCount = new Set(nd.map(d=>d.article)).size;
    const catCount = new Set(nd.map(d=>d.category)).size;
    statusText.innerText=`${artCount} artículos · ${catCount} categorías · ${nd.length} registros`;
    console.log("Datos normalizados:",globalData.length,"registros");
    setTimeout(()=>initDashboard(), 500);
}

function initDashboard(){
    uploadView.classList.add("hidden");
    dashboardLayout.classList.remove("hidden");
    dashboardLayout.classList.add("flex");
    populateFilters();
    updateDashboard();
}

function getFilteredData(includePrevYear = false){
    return globalData.filter(d=>{
        if(filters.year!=="all") {
            if(includePrevYear) {
                if(d.year != filters.year && d.year != (filters.year - 1)) return false;
            } else {
                if(d.year != filters.year) return false;
            }
        }
        if(filters.month!=="all"&&d.monthName!==filters.month)return false;
        if(filters.category!=="all"&&d.category!==filters.category)return false;
        if(filters.brand!=="all"&&d.brand!==filters.brand)return false;
        if(filters.mp==="only"&&!d.isMP)return false;
        return true;
    });
}

// ✅ OPT: obtiene {current, withPrev} en una sola pasada
function getFilteredDataPair() {
    const hasYear = filters.year !== "all";
    const yearNum = hasYear ? parseInt(filters.year, 10) : null;
    const prevYearNum = hasYear ? yearNum - 1 : null;
    const monthFilter = filters.month;
    const catFilter = filters.category;
    const brandFilter = filters.brand;
    const mpOnly = filters.mp === "only";

    const current = [];
    const withPrev = [];
    const n = globalData.length;
    for (let i = 0; i < n; i++) {
        const d = globalData[i];
        if (monthFilter !== "all" && d.monthName !== monthFilter) continue;
        if (catFilter !== "all" && d.category !== catFilter) continue;
        if (brandFilter !== "all" && d.brand !== brandFilter) continue;
        if (mpOnly && !d.isMP) continue;
        if (!hasYear) {
            current.push(d);
            withPrev.push(d);
        } else if (d.year === yearNum) {
            current.push(d);
            withPrev.push(d);
        } else if (d.year === prevYearNum) {
            withPrev.push(d);
        }
    }
    return { current, withPrev };
}

let tomCategory, tomBrand, tomMonth;
function populateFilters(){
    const yrs=[...new Set(globalData.map(d=>d.year))].sort((a,b)=>a-b);
    const cats=[...new Set(globalData.map(d=>d.category))].sort();
    const brands=[...new Set(globalData.map(d=>d.brand))].sort();
    const activeMonths = MESES.filter(m => globalData.some(d => d.monthName === m));

    const ys=document.getElementById("filter-year");
    ys.innerHTML = '<option value="all">Todos</option>';
    yrs.forEach(y=>ys.innerHTML+=`<option value="${y}">${y}</option>`);

    const ms=document.getElementById("filter-month");
    ms.innerHTML = '<option value=""></option>';
    activeMonths.forEach(m=>ms.innerHTML+=`<option value="${m}">${m}</option>`);

    const cs=document.getElementById("filter-category");
    cs.innerHTML = '<option value=""></option>';
    cats.forEach(c=>cs.innerHTML+=`<option value="${c}">${c}</option>`);

    const bs=document.getElementById("filter-brand");
    bs.innerHTML = '<option value=""></option>';
    brands.forEach(b=>bs.innerHTML+=`<option value="${b}">${b}</option>`);

    // ✅ FIX: usar .onchange en vez de addEventListener (evita acumulación)
    ys.onchange = (e) => { filters.year = e.target.value; updateDashboard(); };

    const mpEl = document.getElementById("filter-mp");
    if(mpEl) {
        mpEl.value = filters.mp;
        mpEl.onchange = (e) => { filters.mp = e.target.value; updateDashboard(); };
    }

    const chartModeEl = document.getElementById("chart-mode");
    if (chartModeEl) {
        chartModeEl.onchange = () => { updateDashboard(); };
    }

    if(tomCategory) { tomCategory.destroy(); }
    if(tomBrand) { tomBrand.destroy(); }
    if(tomMonth) { tomMonth.destroy(); }

    tomMonth = new TomSelect("#filter-month", {
        create: false,
        plugins: ['clear_button'],
        placeholder: 'Mes...',
        onChange: function(val) { filters.month = val || 'all'; updateDashboard(); }
    });

    tomCategory = new TomSelect("#filter-category", {
        create: false,
        plugins: ['clear_button'],
        placeholder: 'Buscar categoría...',
        sortField: { field: "text", direction: "asc" },
        onChange: function(val) { filters.category = val || 'all'; updateDashboard(); }
    });

    tomBrand = new TomSelect("#filter-brand", {
        create: false,
        plugins: ['clear_button'],
        placeholder: 'Buscar marca...',
        sortField: { field: "text", direction: "asc" },
        onChange: function(val) { filters.brand = val || 'all'; updateDashboard(); }
    });
}
function updateFilterDropdowns() {
    if(!tomCategory || !tomBrand) return;
    const validCats = new Set();
    const validBrands = new Set();
    const yFilter = filters.year;
    const bFilter = filters.brand;
    const cFilter = filters.category;
    const mpOnly = filters.mp === "only";
    const yActive = yFilter !== "all";
    const bActive = bFilter !== "all";
    const cActive = cFilter !== "all";

    for (let i = 0, n = globalData.length; i < n; i++) {
        const d = globalData[i];
        const passYear = !yActive || d.year == yFilter;
        if (!passYear) continue;
        const passMp = !mpOnly || d.isMP;
        if (!passMp) continue;
        const passBrand = !bActive || d.brand === bFilter;
        const passCat = !cActive || d.category === cFilter;
        if (passBrand) validCats.add(d.category);
        if (passCat) validBrands.add(d.brand);
    }

    const currCat = tomCategory.getValue();
    tomCategory.clearOptions();
    Array.from(validCats).sort().forEach(c => tomCategory.addOption({value: c, text: c}));
    if (currCat && currCat !== 'all' && !validCats.has(currCat)) {
        filters.category = 'all';
        tomCategory.setValue('', true);
    } else {
        tomCategory.setValue(currCat, true);
    }

    const currBrand = tomBrand.getValue();
    tomBrand.clearOptions();
    Array.from(validBrands).sort().forEach(b => tomBrand.addOption({value: b, text: b}));
    if (currBrand && currBrand !== 'all' && !validBrands.has(currBrand)) {
        filters.brand = 'all';
        tomBrand.setValue('', true);
    } else {
        tomBrand.setValue(currBrand, true);
    }
}

function updateDashboard(){
    updateFilterDropdowns();
    const { current: dataCurrent, withPrev: dataWithPrev } = getFilteredDataPair();
    if(dataCurrent.length===0) return;
    renderKPIs(dataWithPrev);
    renderMPAnalysis(dataWithPrev);
    renderCharts(dataWithPrev);
    renderTables(dataWithPrev);
    generateRecommendations(dataWithPrev);
    lucide.createIcons();
}

function renderKPIs(data){
    const years=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear=years[years.length-1];const prevYear=years.length>1?years[years.length-2]:null;
    const skusCount=new Set(data.filter(d => filters.year === "all" || d.year == filters.year).map(d=>d.article)).size;

    const lya=data.filter(d=>d.year===lastYear&&d.units>0);
    const maxMonth=lya.length>0?Math.max(...lya.map(d=>d.month)):11;
    const monthsElapsed=maxMonth+1;

    const ytdL=data.filter(d=>d.year===lastYear&&d.month<=maxMonth).reduce((s,d)=>s+d.units,0);
    const ytdP=prevYear?data.filter(d=>d.year===prevYear&&d.month<=maxMonth).reduce((s,d)=>s+d.units,0):0;

    const gv=ytdP>0?((ytdL/ytdP)-1)*100:null;
    const vColor=gv===null?"#6B7280":(gv>=0?"#10B981":"#EF4444");
    const vText=gv!==null?`${gv>=0?"+":""}${gv.toFixed(1)}%`:"N/D";
    const vSub=prevYear?`Ene\u2013${MESES[maxMonth]} ${lastYear} vs ${prevYear}`:`sin a\u00f1o anterior`;

    const projection=monthsElapsed>0?Math.round((ytdL/monthsElapsed)*12):null;
    const prevFull=prevYear?data.filter(d=>d.year===prevYear).reduce((s,d)=>s+d.units,0):null;
    const projVar=projection!==null&&prevFull&&prevFull>0?((projection/prevFull)-1)*100:null;

    const mm={};data.filter(d=>d.year===lastYear).forEach(d=>{mm[d.monthName]=(mm[d.monthName]||0)+d.units;});
    let bestMonth="-",maxU=-1;for(const k in mm){if(mm[k]>maxU){maxU=mm[k];bestMonth=k;}}

    const periodTxt=`Datos Acumulados: Ene\u2013${MESES[maxMonth]} ${lastYear} (${monthsElapsed} meses)${prevYear?" vs "+prevYear:""}`;
    document.getElementById("period-info").innerText=periodTxt;
    const topInfo=document.getElementById("top-period-info");if(topInfo)topInfo.innerText=periodTxt;

    const isAllYears = (filters.year === "all");
    let kpi1Title = `Ventas Acumuladas ${lastYear}`;
    let kpi1Value = ytdL.toLocaleString("es-ES");
    let kpi1Sub = `Ene\u2013${MESES[maxMonth]} \u00b7 ${monthsElapsed} meses reales`;

    if (isAllYears) {
        const totalAll = data.reduce((sum, d) => sum + d.units, 0);
        kpi1Title = `Ventas Totales Históricas`;
        kpi1Value = totalAll.toLocaleString("es-ES");
        kpi1Sub = `Suma acumulada de todos los años`;
    }

    document.getElementById("kpi-container").innerHTML=`
        <div class="glass-panel border-l-2 border-l-blue-500 relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-blue-500/50 transition-all audit-trigger" data-audit="totalUnits" data-tooltip="Clic para ver detalle">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="bar-chart-2" class="w-16 h-16 text-blue-500"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">${kpi1Title}</p>
            <p class="text-3xl font-bold text-accent tracking-tight">${kpi1Value}</p>
            <p class="text-xs text-muted mt-2">${kpi1Sub}</p>
        </div>
        <div class="glass-panel border-l-2 ${gv===null?'border-l-gray-400':(gv>=0?'border-l-emerald-500':'border-l-red-500')} relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-emerald-500/50 transition-all audit-trigger" data-audit="varYtd" data-tooltip="Clic para ver detalle">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="${gv===null||gv>=0?'trending-up':'trending-down'}" class="w-16 h-16 ${gv===null?'text-gray-400':(gv>=0?'text-emerald-500':'text-red-500')}"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">Variaci\u00f3n Acumulada</p>
            <p class="text-3xl font-bold tracking-tight" style="color:${vColor}">${vText}</p>
            <p class="text-xs text-muted mt-2">${vSub}</p>
        </div>
        <div class="glass-panel border-l-2 border-l-violet-500 relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-violet-500/50 transition-all audit-trigger" data-audit="proj" data-tooltip="Clic para ver detalle">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="target" class="w-16 h-16 text-violet-500"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">Proyecci\u00f3n Cierre ${lastYear}</p>
            <p class="text-3xl font-bold text-accent tracking-tight">${projection?projection.toLocaleString("es-ES"):"N/D"}</p>
            <p class="text-xs text-muted mt-2">Run-rate \u00b7 ${projVar!==null?`${projVar>=0?"+":""}${projVar.toFixed(1)}% vs a\u00f1o ant.`:"Sin ref."}</p>
        </div>
        <div class="glass-panel border-l-2 border-l-amber-500 relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-amber-500/50 transition-all audit-trigger" data-audit="bestMonth">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="star" class="w-16 h-16 text-amber-500"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">Mejor Mes</p>
            <p class="text-2xl font-bold text-accent tracking-tight truncate mt-1" title="${bestMonth} ${lastYear}">${bestMonth}</p>
            <p class="text-xs text-muted mt-2">Pico de ventas</p>
        </div>
        <div class="glass-panel border-l-2 border-l-emerald-400 relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-emerald-400/50 transition-all audit-trigger" data-audit="activeSkus">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="box" class="w-16 h-16 text-emerald-400"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">SKUs Activos</p>
            <p class="text-2xl font-bold text-accent tracking-tight mt-1">${skusCount.toLocaleString("es-ES")}</p>
            <p class="text-xs text-muted mt-2">Con venta en el a\u00f1o</p>
        </div>`;
}


function renderCharts(data){
    const years=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear=years[years.length-1];
    const prevYear=years.length>1?years[years.length-2]:null;
    const mode=(document.getElementById("chart-mode")||{}).value||"units";
    const lya2=data.filter(d=>d.year===lastYear&&d.units>0);
    const maxMonthC=lya2.length>0?Math.max(...lya2.map(d=>d.month)):11;

    const byYear={};
    years.forEach(y=>{byYear[y]=new Array(12).fill(0);});
    data.forEach(d=>{if(byYear[d.year])byYear[d.year][d.month]+=d.units;});
    const lastData=byYear[lastYear];
    const prevData=prevYear?byYear[prevYear]:null;

    const subtitle=document.getElementById("chart-subtitle");
    const legendEl=document.getElementById("chart-legend-ytd");
    if(mode==="units"){
        if(subtitle)subtitle.innerText=prevYear?`Unidades \u2014 ${lastYear} vs ${prevYear}`:`Unidades \u2014 ${lastYear}`;
    }else{
        if(subtitle)subtitle.innerText=prevYear?`Variaci\u00f3n % mes a mes vs ${prevYear}`:"Variaci\u00f3n % (sin a\u00f1o anterior)";
    }
    if(legendEl){
        const ytdUnits=lastData.slice(0,maxMonthC+1).reduce((a,b)=>a+b,0);
        const prevYtd=prevData?prevData.slice(0,maxMonthC+1).reduce((a,b)=>a+b,0):null;
        const ytdVar=prevYtd&&prevYtd>0?((ytdUnits/prevYtd)-1)*100:null;
        const varBg=ytdVar===null?"#F1F5F9":(ytdVar>=0?"#ECFDF5":"#FEF2F2");
        const varColor=ytdVar===null?"#6B7280":(ytdVar>=0?"#059669":"#DC2626");
        legendEl.innerHTML=`<div class="flex flex-col items-end gap-1">
            <span class="font-bold text-accent text-xs tracking-tight">${ytdUnits.toLocaleString("es-ES")} un.</span>
            <span class="text-[0.6rem] text-muted">${MESES[0]}\u2013${MESES[maxMonthC]} ${lastYear}</span>
            </div>`
            +(ytdVar!==null?`<span style="background:${varBg};color:${varColor}" class="font-bold text-xs px-2.5 py-1 rounded-lg ml-2 inline-flex items-center gap-1">${ytdVar>=0?"\u25B2":"\u25BC"} ${ytdVar>=0?"+":""}${ytdVar.toFixed(1)}%</span>`:"");
    }

    const ctxT=document.getElementById("trendChart").getContext("2d");
    let datasets=[];
    if(mode==="units"){
        if(prevData){
            datasets.push({
                label:`${prevYear}`,
                data:prevData.map(v=>v),
                borderColor:"#94A3B8",
                backgroundColor:"transparent",
                pointBackgroundColor:"#94A3B8",
                borderWidth:2,borderDash:[6,5],tension:0.4,
                pointRadius:0,pointHoverRadius:5,
                pointHoverBackgroundColor:"#94A3B8",
                pointHoverBorderColor:"#fff",pointHoverBorderWidth:2,
                fill:false,order:3
            });
        }
        datasets.push({
            label:`${lastYear} (Real)`,
            data:lastData.map((v,i)=>i<=maxMonthC?v:null),
            borderColor:"#2563EB",
            backgroundColor:(ctx)=>{
                const chart=ctx.chart;const {ctx:c,chartArea}=chart;
                if(!chartArea)return"transparent";
                const gradient=c.createLinearGradient(0,chartArea.top,0,chartArea.bottom);
                gradient.addColorStop(0,"rgba(37,99,235,0.18)");
                gradient.addColorStop(0.6,"rgba(37,99,235,0.06)");
                gradient.addColorStop(1,"rgba(37,99,235,0.0)");
                return gradient;
            },
            borderWidth:3,tension:0.4,fill:true,
            pointRadius:lastData.map((_,i)=>i===maxMonthC?6:3),
            pointHoverRadius:7,
            pointBackgroundColor:lastData.map((_,i)=>i===maxMonthC?"#2563EB":"#fff"),
            pointBorderColor:"#2563EB",pointBorderWidth:2,
            pointHoverBackgroundColor:"#2563EB",
            pointHoverBorderColor:"#fff",pointHoverBorderWidth:3,
            order:1
        });
        if(maxMonthC<11){
            const monthsElapsed=maxMonthC+1;
            const ytdTotal=lastData.slice(0,monthsElapsed).reduce((a,b)=>a+b,0);
            const monthlyAvg=ytdTotal/monthsElapsed;
            const projData=lastData.map((v,i)=>{
                if(i<maxMonthC) return null;
                if(i===maxMonthC) return v;
                return Math.round(monthlyAvg);
            });
            datasets.push({
                label:`${lastYear} (Proyección)`,
                data:projData,
                borderColor:"#F59E0B",
                backgroundColor:(ctx)=>{
                    const chart=ctx.chart;const {ctx:c,chartArea}=chart;
                    if(!chartArea)return"transparent";
                    const gradient=c.createLinearGradient(0,chartArea.top,0,chartArea.bottom);
                    gradient.addColorStop(0,"rgba(245,158,11,0.12)");
                    gradient.addColorStop(0.7,"rgba(245,158,11,0.03)");
                    gradient.addColorStop(1,"rgba(245,158,11,0.0)");
                    return gradient;
                },
                borderWidth:2.5,borderDash:[7,5],tension:0.4,fill:true,
                pointRadius:projData.map((v,i)=>v!==null&&i>maxMonthC?4:0),
                pointHoverRadius:6,
                pointBackgroundColor:"#F59E0B",
                pointBorderColor:"#fff",pointBorderWidth:2,
                pointHoverBackgroundColor:"#F59E0B",
                pointHoverBorderColor:"#fff",pointHoverBorderWidth:3,
                order:2
            });
        }
    }else{
        if(prevData){
            const varData=lastData.map((v,i)=>{
                const p=prevData[i];
                if(i>maxMonthC)return null;
                return(p&&p>0)?parseFloat(((v/p-1)*100).toFixed(1)):null;
            });
            datasets.push({
                label:`Var % ${lastYear} vs ${prevYear}`,
                data:varData,
                borderColor:"#2563EB",
                backgroundColor:(ctx)=>{
                    const chart=ctx.chart;const {ctx:c,chartArea}=chart;
                    if(!chartArea)return"transparent";
                    const gradient=c.createLinearGradient(0,chartArea.top,0,chartArea.bottom);
                    gradient.addColorStop(0,"rgba(37,99,235,0.15)");
                    gradient.addColorStop(1,"rgba(37,99,235,0.0)");
                    return gradient;
                },
                borderWidth:2.5,tension:0.4,fill:true,
                pointBackgroundColor:"#fff",
                pointBorderColor:"#2563EB",pointBorderWidth:2,
                pointRadius:4,pointHoverRadius:7,
                pointHoverBackgroundColor:"#2563EB",
                pointHoverBorderColor:"#fff",pointHoverBorderWidth:3,
            });
        }
    }

    if(charts["trend"])charts["trend"].destroy();
    charts["trend"]=new Chart(ctxT,{
        type:"line",data:{labels:MESES,datasets},
        options:{
            responsive:true,maintainAspectRatio:false,
            animation:{ duration:800, easing:'easeOutQuart' },
            interaction:{mode:"index",intersect:false},
            plugins:{
                legend:{
                    position:"top",
                    align:"end",
                    labels:{
                        color:"#334155",
                        usePointStyle:true,
                        pointStyle:"circle",
                        boxWidth:8,
                        boxHeight:8,
                        padding:16,
                        font:{size:11, weight:"600", family:"Inter"}
                    }
                },
                tooltip:{
                    backgroundColor:"#FFFFFF",
                    borderColor:"#E2E8F0",
                    borderWidth:1,
                    titleColor:"#0F172A",
                    titleFont:{size:13, weight:"700", family:"Inter"},
                    bodyColor:"#475569",
                    bodyFont:{size:12, weight:"500", family:"Inter"},
                    padding:{top:12,bottom:12,left:14,right:14},
                    cornerRadius:10,
                    boxPadding:6,
                    usePointStyle:true,
                    callbacks:{
                        title:(items)=>{
                            const idx=items[0].dataIndex;
                            const label=MESES[idx];
                            return idx>maxMonthC?`${label} (Proyección run-rate)`:label;
                        },
                        label:(ctx)=>{
                            if(ctx.raw===null)return null;
                            if(mode==="units"){
                                const isProj=ctx.dataset.label&&ctx.dataset.label.includes("Proyecci");
                                const suffix=isProj?" (est.)":"";
                                return` ${ctx.dataset.label.replace(" (Real)","").replace(" (Proyección)","")}: ${(ctx.raw||0).toLocaleString("es-ES")} un.${suffix}`;
                            }
                            return` ${ctx.dataset.label}: ${ctx.raw!==null?(ctx.raw>=0?"+":"")+ctx.raw+"%":"N/D"}`;
                        },
                        afterBody:(items)=>{
                            if(mode==="units"&&prevData){
                                const idx=items[0].dataIndex;
                                if(idx>maxMonthC)return["","  ⚠ Período proyectado (run-rate)"];
                                const cur=lastData[idx],prev=prevData[idx];
                                if(prev>0){const v=((cur/prev)-1)*100;return["",`  Var. vs ${prevYear}: ${v>=0?"+":""}${v.toFixed(1)}%`];}
                            }
                            return"";
                        }
                    }
                }
            },
            scales:{
                x:{
                    ticks:{color:"#64748B",font:{size:10.5,weight:"500",family:"Inter"}},
                    grid:{color:"rgba(0,0,0,0.03)",drawBorder:false}
                },
                y:{
                    ticks:{color:"#64748B",callback:(v)=>mode==="variation"?v+"%":v.toLocaleString("es-ES"),font:{size:10.5,weight:"500",family:"Inter"},padding:8},
                    grid:{color:"rgba(0,0,0,0.04)",drawBorder:false},
                    beginAtZero:true
                }
            }
        }
    });

    // Pareto Chart
    const ctxP=document.getElementById("paretoChart").getContext("2d");
    const pm2={};let ta=0;data.forEach(d=>{pm2[d.article]=(pm2[d.article]||0)+d.units;ta+=d.units;});
    const sp=Object.keys(pm2).map(k=>({name:k,units:pm2[k]})).sort((a,b)=>b.units-a.units);
    const t20=sp.slice(0,20);let accP=0;
    const pl=t20.map(p=>{accP+=p.units;return+((accP/ta)*100).toFixed(1);});
    if(charts["pareto"])charts["pareto"].destroy();
    charts["pareto"]=new Chart(ctxP,{
        type:"bar",
        data:{labels:t20.map(p=>p.name.substring(0,18)+"\u2026"),datasets:[
            {type:"line",label:"% Acumulado",data:pl,borderColor:"#F59E0B",backgroundColor:"transparent",yAxisID:"y1",tension:0.2,pointRadius:2,borderWidth:1.5},
            {type:"bar",label:"Unidades",data:t20.map(p=>p.units),backgroundColor:t20.map((_,i)=>pl[i]<=80?"#2563EB":"#374151"),yAxisID:"y",borderRadius:2}
        ]},
        options:{
            responsive:true,maintainAspectRatio:false,
            plugins:{
                legend:{position:"bottom",labels:{color:"#71717A",usePointStyle:true,font:{size:10}}},
                tooltip:{backgroundColor:"#FFFFFF",borderColor:"#E4E4E7",borderWidth:1,titleColor:"#0F172A",bodyColor:"#71717A"}
            },
            scales:{
                x:{ticks:{display:false}},
                y:{type:"linear",position:"left",ticks:{color:"#71717A",callback:v=>v.toLocaleString("es-ES"),font:{size:10}},grid:{color:"rgba(0,0,0,0.04)"}},
                y1:{type:"linear",position:"right",min:0,max:100,ticks:{color:"#F59E0B",callback:v=>v+"%",font:{size:10}},grid:{drawOnChartArea:false}}
            }
        }
    });
    let s80=0,a80=0;for(const p of sp){a80+=p.units;s80++;if(a80/ta>=0.8)break;}
    document.getElementById("pareto-insight").innerText=`El 80% de las ventas est\u00e1 generado por ${s80} SKUs de ${sp.length} (${((s80/sp.length)*100).toFixed(1)}% del portafolio).`;

    // Doughnut Chart
    const ctxD = document.getElementById("doughnutChart");
    if(ctxD){
        const catMap = {}; data.forEach(d => { catMap[d.category] = (catMap[d.category] || 0) + d.units; });
        const cats = Object.keys(catMap).sort((a,b) => catMap[b] - catMap[a]);
        if (cats.length === 0) {
            if(charts["doughnut"]) charts["doughnut"].destroy();
            return;
        }
        const dColors = ["#3B82F6", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#64748B"];
        if(charts["doughnut"]) charts["doughnut"].destroy();
        charts["doughnut"] = new Chart(ctxD.getContext("2d"), {
            type: "doughnut",
            plugins: [ChartDataLabels],
            data: {
                labels: cats.slice(0, 6),
                datasets: [{
                    data: cats.slice(0, 6).map(c => catMap[c]),
                    backgroundColor: dColors,
                    borderWidth: 2,
                    borderColor: "#ffffff"
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                cutout: "70%",
                plugins: {
                    legend: { position: "right", labels: { color: "#71717A", usePointStyle: true, boxWidth: 8, font: { size: 10 } } },
                    tooltip: { backgroundColor: "#FFFFFF", borderColor: "#E4E4E7", borderWidth: 1, titleColor: "#0F172A", bodyColor: "#71717A" },
                    datalabels: {
                        color: "#fff",
                        font: { weight: "bold", size: 10 },
                        formatter: (value, ctx) => {
                            let sum = 0;
                            let dataArr = ctx.chart.data.datasets[0].data;
                            dataArr.forEach(data => { sum += data; });
                            let percentage = (value * 100 / sum).toFixed(0) + "%";
                            return percentage;
                        }
                    }
                }
            }
        });
    }

    // YoY Chart
    const ctxY = document.getElementById("yoyChart");
    if(ctxY && prevData){
        if(charts["yoy"]) charts["yoy"].destroy();
        charts["yoy"] = new Chart(ctxY.getContext("2d"), {
            type: "bar",
            data: {
                labels: MESES.map(m=>m.substring(0,3)),
                datasets: [
                    { label: String(prevYear), data: prevData, backgroundColor: "#E2E8F0", borderRadius: 4 },
                    { label: String(lastYear), data: lastData.map((v,i) => i <= maxMonthC ? v : null), backgroundColor: "#3B82F6", borderRadius: 4 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { position: "top", align: "end", labels: { color: "#71717A", usePointStyle: true, font: { size: 10 } } },
                    tooltip: { backgroundColor: "#FFFFFF", borderColor: "#E4E4E7", borderWidth: 1, titleColor: "#0F172A", bodyColor: "#71717A", mode: "index" }
                },
                scales: {
                    x: { ticks: { color: "#71717A", font: { size: 10 } }, grid: { display: false } },
                    y: { ticks: { color: "#71717A", font: { size: 10 }, callback: v=>v.toLocaleString("es-ES") }, grid: { color: "rgba(0,0,0,0.05)", drawBorder: false } }
                }
            }
        });
    }
}

function renderTables(data){
    const years=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear=years[years.length-1];const prevYear=years.length>1?years[years.length-2]:null;

    const lya=data.filter(d=>d.year===lastYear&&d.units>0);
    const maxMonth=lya.length>0?Math.max(...lya.map(d=>d.month)):11;

    const isAllYears = (filters.year === "all");
    const totalLabel = isAllYears ? `Total Histórico` : `Total ${lastYear}`;
    const varLabel = isAllYears ? (prevYear ? `Var. YTD ${lastYear} vs ${prevYear}` : `Var. YTD`) : (prevYear ? `Var. YTD vs ${prevYear}` : `Var. YTD`);

    const thTopTotal = document.getElementById("th-top-total");
    if(thTopTotal) thTopTotal.innerText = totalLabel;
    const thTopAcum = document.getElementById("th-top-acum");
    if(thTopAcum) thTopAcum.innerText = varLabel;

    const thFullTotal = document.getElementById("th-full-total");
    if(thFullTotal) thFullTotal.innerText = totalLabel;
    const thFullVar = document.getElementById("th-full-var");
    if(thFullVar) thFullVar.innerText = varLabel;

    const pm={};
    data.forEach(d=>{
        if(!pm[d.article])pm[d.article]={article:d.article,category:d.category,total:0,ytdLast:0,ytdPrev:0,lastSaleMonth:-1,lastSaleYear:0};
        const p=pm[d.article];
        if (filters.year === "all" || d.year == filters.year) p.total+=d.units;
        if(d.month<=maxMonth){if(d.year===lastYear)p.ytdLast+=d.units;if(prevYear&&d.year===prevYear)p.ytdPrev+=d.units;}
        if(d.units>0){const n=d.year>p.lastSaleYear||(d.year===p.lastSaleYear&&d.month>p.lastSaleMonth);if(n){p.lastSaleYear=d.year;p.lastSaleMonth=d.month;}}
    });
    const products=Object.values(pm);
    const tp=products.reduce((s,p)=>s+p.total,0);
    products.forEach(p=>{
        p.share=tp>0?(p.total/tp)*100:0;
        if(p.ytdPrev===0&&p.ytdLast>0)p.ytdVar=null;
        else if(p.ytdPrev===0)p.ytdVar=-100;
        else p.ytdVar=((p.ytdLast/p.ytdPrev)-1)*100;
    });
    const top10=[...products].sort((a,b)=>b.total-a.total).slice(0,10);
    document.getElementById("top-products-body").innerHTML=top10.map((p,idx)=>{
        const isN=p.ytdVar===null;const vs=isN?"Nuevo":`${p.ytdVar>=0?"+":""}${p.ytdVar.toFixed(1)}%`;
        const badgeBg=isN?"bg-blue-50 text-blue-600 border-blue-200":(p.ytdVar>=0?"bg-emerald-50 text-emerald-700 border-emerald-200":"bg-red-50 text-red-700 border-red-200");
        const medal=idx===0?"\ud83e\udd47":idx===1?"\ud83e\udd48":idx===2?"\ud83e\udd49":`<span class="text-muted font-bold">${idx+1}</span>`;
        const varArrow=isN?"":p.ytdVar>=0?`<svg class="w-2.5 h-2.5 inline" fill="currentColor" viewBox="0 0 20 20"><path d="M10 3l7 7H3l7-7z"/></svg>`:`<svg class="w-2.5 h-2.5 inline" fill="currentColor" viewBox="0 0 20 20"><path d="M10 17l-7-7h14l-7 7z"/></svg>`;
        return`<tr class="hover:bg-violet-50/40 transition-colors">
          <td class="py-2.5 px-1 text-center font-bold text-sm w-8">${medal}</td>
          <td class="py-2.5 px-2 font-semibold text-accent truncate max-w-[220px]" title="${p.article}">${p.article}</td>
          <td class="py-2.5 px-2 text-muted text-[0.7rem] truncate max-w-[100px]" title="${p.category}">${p.category}</td>
          <td class="py-2.5 px-2 text-right font-bold text-accent tabular-nums">${p.total.toLocaleString("es-ES")}</td>
          <td class="py-2.5 px-2 text-right"><span class="inline-flex items-center gap-0.5 text-[0.7rem] font-bold border rounded-lg px-1.5 py-0.5 ${badgeBg}">${varArrow}${vs}</span></td>
        </tr>`;
    }).join("");

    const catMap = {};
    data.forEach(d => {
        if(!catMap[d.category]) catMap[d.category] = { category: d.category, total: 0, ytdLast: 0, ytdPrev: 0 };
        const c = catMap[d.category];
        if (isAllYears && prevYear) {
            if (d.year === prevYear) c.total += d.units;
        } else {
            if (filters.year === "all" || d.year == filters.year) c.total += d.units;
        }
        if(d.month<=maxMonth){
            if(d.year===lastYear) c.ytdLast+=d.units;
            if(prevYear&&d.year===prevYear) c.ytdPrev+=d.units;
        }
    });
    const cats = Object.values(catMap).filter(c => c.total > 0).sort((a,b) => b.total - a.total);
    if(cats.length > 0) {
        const topCat = cats[0];
        const catVar = topCat.ytdPrev > 0 ? ((topCat.ytdLast / topCat.ytdPrev) - 1) * 100 : null;
        let varHtml = "";
        if(catVar !== null) {
            const vColor = catVar >= 0 ? "text-emerald-500" : "text-red-500";
            const vIcon = catVar >= 0 ? "trending-up" : "trending-down";
            varHtml = `<div class="flex items-center gap-1 ${vColor}"><i data-lucide="${vIcon}" class="w-3 h-3"></i><span class="text-xs font-bold">${Math.abs(catVar).toFixed(1)}%</span></div>`;
        }
        const catLeader = document.getElementById("cat-leader-badge");
        if(catLeader) catLeader.classList.remove("hidden");
        document.getElementById("extra-kpi-content").innerHTML = `
            <span class="text-xs font-extrabold text-accent truncate max-w-[120px]" title="${topCat.category}">${topCat.category}</span>
            <span class="text-xs font-black text-brand">${topCat.total.toLocaleString("es-ES")}</span>
            ${catVar!==null?`<span class="text-[0.65rem] font-bold px-1.5 py-0.5 rounded-md ${catVar>=0?"bg-emerald-100 text-emerald-700":"bg-red-100 text-red-700"}">${catVar>=0?"+":""}${catVar.toFixed(1)}%</span>`:""}
        `;
    } else {
        document.getElementById("extra-kpi-content").innerHTML = `<p class="text-xs text-muted">No hay datos</p>`;
    }

    const bm = {};
    data.forEach(d => {
        if(!bm[d.brand]) bm[d.brand] = { brand: d.brand, total: 0 };
        const b = bm[d.brand];
        if (filters.year === "all" || d.year == filters.year) b.total += d.units;
    });
    const brands = Object.values(bm).sort((a,b) => b.total - a.total);
    const top10B = brands.slice(0, 10);

    const labels = top10B.map(b => b.brand);
    const chartData = top10B.map(b => b.total);

    const ctxB = document.getElementById("brandsChart");
    if(ctxB) {
        if(charts["brands"]) charts["brands"].destroy();
        charts["brands"] = new Chart(ctxB, {
            type: 'bar',
            plugins: [ChartDataLabels],
            data: {
                labels: labels,
                datasets: [{
                    data: chartData,
                    backgroundColor: ['#3B82F6', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#14B8A6', '#F43F5E', '#06B6D4', '#84CC16', '#D946EF'],
                    borderRadius: 4,
                    barPercentage: 0.7,
                    categoryPercentage: 0.9
                }]
            },
            options: {
                onClick: (e, activeElements) => {
                    if (activeElements.length > 0) {
                        const idx = activeElements[0].index;
                        const brandName = labels[idx];
                        openAuditModal("brandTop", brandName);
                    }
                },
                onHover: (e, activeElements) => {
                    e.native.target.style.cursor = activeElements.length > 0 ? 'pointer' : 'default';
                },
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { right: 50, top: 5, bottom: 5, left: 0 } },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: "#FFFFFF", borderColor: "#E4E4E7", borderWidth: 1, titleColor: "#0F172A", bodyColor: "#71717A",
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const pct = ((val / total) * 100).toFixed(1) + '%';
                                return ` ${val.toLocaleString("es-ES")} uni. (${pct})`;
                            }
                        }
                    },
                    datalabels: {
                        anchor: 'end',
                        align: 'right',
                        color: '#71717A',
                        font: { weight: '600', size: 10, family: 'Inter' },
                        formatter: (value, context) => {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const pct = Math.round((value / total) * 100);
                            return `${value.toLocaleString("es-ES")} (${pct}%)`;
                        }
                    }
                },
                scales: {
                    x: { display: false, max: chartData[0] ? chartData[0] * 1.25 : 100 },
                    y: {
                        grid: { display: false, drawBorder: false },
                        ticks: { color: '#3F3F46', font: { family: 'Inter', size: 11, weight: '500' } }
                    }
                }
            }
        });
    }
    const fullSorted=[...products].filter(p=>p.total>0).sort((a,b)=>b.total-a.total);
    let accShare=0;fullSorted.forEach(p=>{accShare+=p.share;p.inPareto=accShare<=80;});
    const monthsElapsed = maxMonth >= 0 ? maxMonth + 1 : 0;
    const renderFull=(list)=>{
        document.getElementById("full-ranking-body").innerHTML=list.map((p,idx)=>{
            const isN=p.ytdVar===null;const vt=isN?"Nuevo":`${p.ytdVar>=0?"+":""}${p.ytdVar.toFixed(1)}%`;
            const vc=isN?"text-blue-600":(p.ytdVar>=0?"text-emerald-600":"text-red-600");
            let cl="Consolidado",cc="badge-gray";
            if(isN){cl="Nuevo";cc="badge-blue";}else if(p.ytdVar>=20){cl="Crecimiento";cc="badge-green";}else if(p.ytdVar<-30){cl="Declive";cc="badge-red";}else if(idx<list.length*0.15&&!isN&&p.ytdVar>=0){cl="L\u00edder";cc="badge-green";}
            const sw=Math.min(p.share*4,100);
            const paretoColor=p.inPareto?"bg-emerald-500":"bg-blue-500";
            const paretoText=p.inPareto?"text-emerald-600 font-bold":"text-muted";
            const sb=`<div class="w-full bg-gray-200 rounded-full h-1.5"><div class="${paretoColor} h-1.5 rounded-full" style="width:${sw}%"></div></div><p class="text-xs ${paretoText} text-center mt-0.5">${p.share.toFixed(1)}%</p>`;
            const avgYTD = monthsElapsed > 0 ? Math.round(p.ytdLast / monthsElapsed) : 0;
            return`<tr><td class="text-muted">${idx+1}</td><td class="font-medium text-accent" title="${p.article}">${p.article}</td><td class="text-muted">${p.category}</td><td class="text-right font-semibold text-accent">${p.total.toLocaleString("es-ES")}</td><td class="text-right text-muted font-medium">${avgYTD.toLocaleString("es-ES")}</td><td class="text-right ${vc} font-semibold">${vt}</td><td class="text-center"><span class="badge ${cc}">${cl}</span></td><td class="min-w-[80px]">${sb}</td></tr>`;
        }).join("");
    };
    renderFull(fullSorted);

    const searchInput = document.getElementById("search-input");
    if (searchInput) {
        searchInput.oninput = (e) => {
            const q = e.target.value.toLowerCase();
            renderFull(fullSorted.filter(p =>
                p.article.toLowerCase().includes(q) ||
                p.category.toLowerCase().includes(q)
            ));
        };
    }
}


function generateRecommendations(data){
    const R = buildAnalytics(data);
    let mkt = R.insights.filter(i=>i.tipo==="Oportunidad"||i.tipo==="Global");
    if(mkt.length===0) mkt=[{titulo:"Estabilidad",texto:"Sin insights destacados de crecimiento."}];
    let port = R.insights.filter(i=>i.tipo==="Alerta"||i.tipo==="Estrat\u00e9gico");
    if(port.length===0) port=[{titulo:"Salud del Portafolio",texto:"Sin alertas severas registradas."}];

    const iconMap={Alerta:"alert-circle",Oportunidad:"trending-up","Estrat\u00e9gico":"target",Global:"bar-chart-2"};
    const colorMap={Alerta:"#f87171",Oportunidad:"#34d399","Estrat\u00e9gico":"#fbbf24",Global:"#60a5fa"};

    const execEl = document.getElementById("executive-summary");
    if (execEl) execEl.innerHTML = R.insights.map(i=>`
        <li class="insight-item">
            <div class="flex-shrink-0 mt-0.5" style="color:${colorMap[i.tipo]||"#9CA3AF"}">
                <i data-lucide="${iconMap[i.tipo]||'info'}" class="w-4 h-4"></i>
            </div>
            <div>
                <p class="font-semibold text-accent text-xs tracking-wide">${i.titulo}</p>
                <p class="text-muted mt-1 leading-relaxed" style="font-size:0.75rem">${i.texto}</p>
            </div>
        </li>`).join("");

    const mktEl = document.getElementById("mkt-recommendations");
    if (mktEl) mktEl.innerHTML = mkt.map(i=>`
        <li class="flex items-start gap-2">
            <i data-lucide="arrow-right" class="w-3 h-3 text-blue-500 flex-shrink-0 mt-0.5"></i>
            <div><span class="text-accent font-medium">${i.titulo}:</span> ${i.texto}</div>
        </li>`).join("");
    const portEl = document.getElementById("portfolio-recommendations");
    if (portEl) portEl.innerHTML = port.map(i=>`
        <li class="flex items-start gap-2">
            <i data-lucide="arrow-right" class="w-3 h-3 text-amber-500 flex-shrink-0 mt-0.5"></i>
            <div><span class="text-accent font-medium">${i.titulo}:</span> ${i.texto}</div>
        </li>`).join("");
}

function addSecHdr(ws,r,c1,c2,title,color){
    ws.mergeCells(r,c1,r,c2);const c=ws.getCell(r,c1);c.value=title;
    c.font=xlFont(true,"FFFFFFFF",11);c.fill=xlBg(color);c.alignment=xlAlign("left","middle");
}

function xlFont(b,c,s){return{name:"Calibri",bold:b,color:{argb:c},size:s};}
function xlBg(c){return{type:"pattern",pattern:"solid",fgColor:{argb:c}};}
function xlAlign(h,v,w){return{horizontal:h,vertical:v,wrapText:w||false};}
function xlBorder(s="thin",c="FFE2E8F0"){return{top:{style:s,color:{argb:c}},left:{style:s,color:{argb:c}},bottom:{style:s,color:{argb:c}},right:{style:s,color:{argb:c}}};}
function varColorFn(v){return v===null?"FF94A3B8":(v>=0?"FF059669":"FFDC2626");}

function buildAnalytics(data){
    const years=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear=years[years.length-1];
    const prevYear=years.length>1?years[years.length-2]:null;
    const lya=data.filter(d=>d.year===lastYear&&d.units>0);
    const maxMonth=lya.length>0?Math.max(...lya.map(d=>d.month)):11;
    const ytdLabel=`Ene\u2013${MESES[maxMonth]}`;
    const monthsElapsed = maxMonth + 1;

    const skuMap={};
    data.forEach(d=>{
        if(!skuMap[d.article])skuMap[d.article]={article:d.article,category:d.category,brand:d.brand,total:0,ytdLast:0,ytdPrev:0,lastSaleYear:0,lastSaleMonth:-1, monthlyLast: new Array(12).fill(0)};
        const s=skuMap[d.article];s.total+=d.units;
        if(d.month<=maxMonth){
            if(d.year===lastYear){ s.ytdLast+=d.units; s.monthlyLast[d.month]+=d.units; }
            if(prevYear&&d.year===prevYear)s.ytdPrev+=d.units;
        }
        if(d.units>0){const n=d.year>s.lastSaleYear||(d.year===s.lastSaleYear&&d.month>s.lastSaleMonth);if(n){s.lastSaleYear=d.year;s.lastSaleMonth=d.month;}}
    });
    const skus=Object.values(skuMap);
    const totalUnitsLast=skus.reduce((s,p)=>s+p.ytdLast,0);
    const totalUnitsPrev=skus.reduce((s,p)=>s+p.ytdPrev,0);

    skus.forEach(p=>{
        p.varAbs = p.ytdLast - p.ytdPrev;
        p.ytdVar=p.ytdPrev>0?((p.ytdLast/p.ytdPrev)-1)*100:(p.ytdLast>0?null:-100);
        p.shareLast=totalUnitsLast>0?(p.ytdLast/totalUnitsLast)*100:0;
        p.sharePrev=totalUnitsPrev>0?(p.ytdPrev/totalUnitsPrev)*100:0;
        p.shareVar = p.shareLast - p.sharePrev;
        p.runRate = (p.ytdLast / monthsElapsed) * 12;

        let sumX=0,sumY=0,sumXY=0,sumX2=0,n=monthsElapsed;
        for(let i=0;i<monthsElapsed;i++){
            sumX+=i; sumY+=p.monthlyLast[i]; sumXY+=i*p.monthlyLast[i]; sumX2+=i*i;
        }
        p.tendencia = (n*sumXY - sumX*sumY) / (n*sumX2 - sumX*sumX);

        if(p.ytdVar===null)p.clasif="Nuevo";
        else if(p.ytdVar>=20)p.clasif="Fuerte Crecimiento";
        else if(p.ytdVar>=-5)p.clasif="Estable";
        else if(p.ytdVar>=-30)p.clasif="Declive Moderado";
        else p.clasif="Declive Severo";
        p.semaforo=(p.ytdVar===null||p.ytdVar>=-5)?"VERDE":(p.ytdVar>=-30?"AMARILLO":"ROJO");
        p.lastSaleLabel=p.lastSaleMonth>=0?`${MESES[p.lastSaleMonth]} ${p.lastSaleYear}`:"Sin venta";
    });
    const ss=[...skus].sort((a,b)=>b.ytdLast-a.ytdLast);
    let acc=0,s80=0;for(const p of ss){acc+=p.ytdLast;s80++;if(acc/totalUnitsLast>=0.8)break;}
    const pct80=((s80/skus.length)*100).toFixed(1);

    const catMap={};
    skus.forEach(p=>{
        if(!catMap[p.category])catMap[p.category]={category:p.category,total:0,ytdLast:0,ytdPrev:0, varAbs:0, monthlyLast: new Array(12).fill(0)};
        const c=catMap[p.category];c.total+=p.total;c.ytdLast+=p.ytdLast;c.ytdPrev+=p.ytdPrev;c.varAbs+=p.varAbs;
        for(let i=0;i<12;i++) c.monthlyLast[i]+=p.monthlyLast[i];
    });
    const cats=Object.values(catMap).map(c=>{
        c.ytdVar=c.ytdPrev>0?((c.ytdLast/c.ytdPrev)-1)*100:null;
        c.share=totalUnitsLast>0?(c.ytdLast/totalUnitsLast)*100:0;
        let sumX=0,sumY=0,sumXY=0,sumX2=0,n=monthsElapsed;
        for(let i=0;i<monthsElapsed;i++){sumX+=i;sumY+=c.monthlyLast[i];sumXY+=i*c.monthlyLast[i];sumX2+=i*i;}
        c.tendencia = (n*sumXY - sumX*sumY) / (n*sumX2 - sumX*sumX);
        return c;
    }).sort((a,b)=>b.ytdLast-a.ytdLast);
    const cg=cats.filter(c=>c.ytdVar!==null&&c.ytdVar>0).sort((a,b)=>b.varAbs-a.varAbs);
    const cd=cats.filter(c=>c.ytdVar!==null&&c.ytdVar<0).sort((a,b)=>a.varAbs-b.varAbs);

    const globalVar=totalUnitsPrev>0?((totalUnitsLast/totalUnitsPrev)-1)*100:null;
    const globalVarAbs = totalUnitsLast - totalUnitsPrev;
    const globalRunRate = (totalUnitsLast / monthsElapsed) * 12;

    const mT=new Array(12).fill(0);data.filter(d=>d.year===lastYear).forEach(d=>{mT[d.month]+=d.units;});
    let bIdx=0, wIdx=0;
    for(let i=0;i<monthsElapsed;i++){ if(mT[i]>mT[bIdx])bIdx=i; if(mT[i]<mT[wIdx])wIdx=i; }

    const insights=[];
    const nf = (n) => n.toLocaleString("es-ES");

    if(globalVar!==null){
        const gTxt = globalVar>=0?"crecimiento":"contracci\u00f3n";
        insights.push({
            tipo:"Global", titulo:`${globalVar>=0?"\ud83d\ude80":"\u26a0\ufe0f"} Tendencia de Portafolio`,
            texto: `Durante ${ytdLabel} ${lastYear}, las ventas muestran un ${gTxt} del ${Math.abs(globalVar).toFixed(1)}% (${globalVarAbs>0?'+':''}${nf(globalVarAbs)} unidades) en comparaci\u00f3n al a\u00f1o anterior. Si esta tendencia se mantiene, se proyecta un cierre de a\u00f1o cercano a ${nf(globalRunRate.toFixed(0))} unidades. El mes de mayor actividad fue ${MESES[bIdx]}.`
        });
    }

    const drivers = ss.filter(p=>p.varAbs>0).sort((a,b)=>b.varAbs-a.varAbs);
    if(drivers.length>0){
        const d1 = drivers[0];
        insights.push({
            tipo:"Oportunidad", titulo:`\ud83d\udcc8 Motor de Crecimiento: ${d1.article.substring(0,30)}`,
            texto: `Este SKU impuls\u00f3 el volumen agregando ${nf(d1.varAbs)} unidades extras vs. el a\u00f1o anterior (+${d1.ytdVar?d1.ytdVar.toFixed(1):100}%). Esto significa una excelente aceptaci\u00f3n o efectividad en PDV. Acci\u00f3n: Proteger el inventario de este c\u00f3digo y evaluar replicar su estrategia promocional en otros formatos.`
        });
    }

    const bleeders = ss.filter(p=>p.varAbs<0).sort((a,b)=>a.varAbs-b.varAbs);
    if(bleeders.length>0){
        const b1 = bleeders[0];
        insights.push({
            tipo:"Alerta", titulo:`\ud83d\udea8 Riesgo de Volumen: ${b1.article.substring(0,30)}`,
            texto: `Este SKU sufri\u00f3 la mayor contracci\u00f3n absoluta, perdiendo ${nf(Math.abs(b1.varAbs))} unidades (${b1.ytdVar.toFixed(1)}%). Esta ca\u00edda afecta severamente la participaci\u00f3n de su categor\u00eda. Acci\u00f3n: Investigar inmediatamente si obedece a quiebre de stock, cambio de precio o p\u00e9rdida de cliente clave.`
        });
    }

    if(cg.length>0){
        insights.push({
            tipo:"Estrat\u00e9gico", titulo:`\ud83c\udfaf Desplazamiento de Mix: ${cg[0].category}`,
            texto: `La categor\u00eda "${cg[0].category}" lidera la expansi\u00f3n con ${nf(cg[0].varAbs)} unidades adicionales. Su tendencia interna es ${cg[0].tendencia>0?"acelerada":"desacelerada"} en meses recientes. Acci\u00f3n: Maximizar exhibici\u00f3n cruzada de estos productos mientras mantengan momento de venta.`
        });
    }

    const alertas=ss.filter(p=>p.varAbs< -500 || (p.ytdVar !== null && p.ytdVar < -30 && p.ytdLast > 0)).slice(0,10);
    const oportunidades=ss.filter(p=>p.tendencia > 0 && p.ytdVar > 20 && p.ytdLast > 50).slice(0,10);

    return{years,lastYear,prevYear,ytdLabel,maxMonth,monthsElapsed,totalUnitsLast,totalUnitsPrev,globalVar,globalVarAbs,globalRunRate,skusSorted:ss,skus80:s80,pct80,cats,catsGrowth:cg,catsDecline:cd,alertas,oportunidades,insights,bestMonth:MESES[bIdx],worstMonth:MESES[wIdx]};
}

async function exportToExcel(){
    const data=getFilteredData();if(data.length===0){alert("No hay datos para exportar.");return;}
    const btn=document.getElementById("btn-export-analysis");const origText=btn.innerHTML;
    btn.innerHTML='<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> Generando...';lucide.createIcons();
    try{
        const R=buildAnalytics(data);
        const nf=n=>typeof n==="number"?n.toLocaleString("es-ES"):n;
        const pct=n=>n!==null?`${n>=0?"+":""}${n.toFixed(1)}%`:"N/D";
        const wb=new ExcelJS.Workbook();wb.creator="Sales Analytics BI";wb.created=new Date();

        const HDR_DARK="FF0F172A";const HDR_BLUE="FF1E3A5F";const HDR_GREEN="FF065F46";
        const HDR_RED="FF7F1D1D";const HDR_VIOLET="FF4C1D95";const HDR_AMBER="FF78350F";
        const CELL_ALT="FFF8FAFC";const CELL_WHITE="FFFFFFFF";

        const setHdr=(ws,row,col1,col2,text,bg,fSize=11)=>{
            if(col2>col1)ws.mergeCells(row,col1,row,col2);
            const c=ws.getCell(row,col1);
            c.value=text;c.font=xlFont(true,"FFFFFFFF",fSize);c.fill=xlBg(bg);
            c.alignment=xlAlign("left","middle");ws.getRow(row).height=22;
        };
        const setColHdrs=(ws,row,cols,bg)=>{
            cols.forEach(([h,sc,ec])=>{
                if(ec>sc)ws.mergeCells(row,sc,row,ec);
                const c=ws.getCell(row,sc);c.value=h;
                c.font=xlFont(true,"FFFFFFFF",9);c.fill=xlBg(bg);
                c.alignment=xlAlign("center","middle");c.border=xlBorder("thin","FF1F2937");
            });ws.getRow(row).height=18;
        };

        // HOJA 1 – RESUMEN EJECUTIVO
        const ws1=wb.addWorksheet("1. Resumen Ejecutivo",{views:[{showGridLines:false}]});
        [3,28,16,16,16,16,16,3].forEach((w,i)=>ws1.getColumn(i+1).width=w);
        ws1.getRow(1).height=8;
        ws1.mergeCells("B2:G4");
        const t1=ws1.getCell("B2");
        t1.value="REPORTE EJECUTIVO DE PORTAFOLIO DE VENTAS";
        t1.font=xlFont(true,"FFFFFFFF",22);t1.fill=xlBg(HDR_DARK);t1.alignment=xlAlign("center","middle");
        ws1.getRow(2).height=25;ws1.getRow(3).height=25;ws1.getRow(4).height=25;

        const af=[];if(filters.year!=="all")af.push(`A\u00f1o: ${filters.year}`);
        if(filters.category!=="all")af.push(`Categor\u00eda: ${filters.category}`);
        if(filters.brand!=="all")af.push(`Marca: ${filters.brand}`);
        ws1.mergeCells("B5:G5");
        const sub1=ws1.getCell("B5");
        sub1.value=`Generado: ${new Date().toLocaleDateString("es-ES",{weekday:"long",day:"2-digit",month:"long",year:"numeric"})}  |  Per\u00edodo analizado: ${R.ytdLabel} ${R.lastYear}  |  ${af.length>0?"Filtros: "+af.join(" · "):"Sin filtros aplicados"}`;
        sub1.font={name:"Calibri",size:9,italic:true,color:{argb:"FF94A3B8"}};sub1.fill=xlBg("FF0F1629");sub1.alignment=xlAlign("center","middle");
        ws1.getRow(5).height=14;ws1.getRow(6).height=8;

        setHdr(ws1,7,2,7,"  \u25a0  INDICADORES CLAVE DE RENDIMIENTO",HDR_BLUE,11);ws1.getRow(8).height=6;

        const prevFull=R.prevYear?data.filter(d=>d.year===R.prevYear).reduce((s,d)=>s+d.units,0):null;
        const projVar=prevFull&&prevFull>0?((R.globalRunRate/prevFull)-1)*100:null;
        const mesesRestantes=12-R.monthsElapsed;
        const proyFaltante=Math.round(R.globalRunRate-R.totalUnitsLast);

        const kpiBlocks=[
            {label:`VENTAS YTD ${R.lastYear}`,val:nf(R.totalUnitsLast),sub1:`${R.ytdLabel} (${R.monthsElapsed} meses reales)`,sub2:`Prom. mensual: ${nf(Math.round(R.totalUnitsLast/R.monthsElapsed))} un/mes`,col:"FF2563EB"},
            {label:`AÑO ANTERIOR YTD`,val:R.prevYear?nf(R.totalUnitsPrev):"N/D",sub1:R.prevYear?`${R.ytdLabel} ${R.prevYear}`:"Sin año anterior",sub2:R.prevYear?`Diferencia: ${R.globalVarAbs>=0?"+":""}${nf(R.globalVarAbs)} un.`:"",col:"FF4B5563"},
            {label:`VARIACIÓN YTD`,val:pct(R.globalVar),sub1:`Comparación período idéntico`,sub2:R.globalVar!==null?(R.globalVar>=0?"Tendencia positiva ▲":"Tendencia negativa ▼"):"Sin año base",col:R.globalVar===null?"FF6B7280":(R.globalVar>=0?"FF059669":"FFDC2626")},
            {label:`PROYECCIÓN CIERRE ${R.lastYear}`,val:nf(Math.round(R.globalRunRate)),sub1:`Run-rate: ${nf(Math.round(R.totalUnitsLast/R.monthsElapsed))} un/mes × 12`,sub2:projVar!==null?`vs ${R.prevYear} anual: ${pct(projVar)}`:"Sin referencia anual",col:"FF7C3AED"},
            {label:`UNIDADES FALTANTES`,val:nf(proyFaltante),sub1:`Para completar proyección`,sub2:`${mesesRestantes} meses restantes`,col:"FFF59E0B"},
            {label:`PARETO 80/20`,val:`${R.skus80} SKUs`,sub1:`Generan el 80% del volumen`,sub2:`${R.pct80}% del portafolio total`,col:"FF0891B2"},
        ];

        let kr=9;
        kpiBlocks.forEach((k,i)=>{
            const col=2+i;
            ws1.mergeCells(kr,col,kr,col);ws1.getCell(kr,col).fill=xlBg(k.col);ws1.getRow(kr).height=5;
            ws1.mergeCells(kr+1,col,kr+1,col);
            const kc=ws1.getCell(kr+1,col);kc.value=k.label;kc.font=xlFont(true,"FF94A3B8",8);kc.fill=xlBg("FF1E293B");kc.alignment=xlAlign("center","middle");ws1.getRow(kr+1).height=14;
            ws1.mergeCells(kr+2,col,kr+2,col);
            const kv=ws1.getCell(kr+2,col);kv.value=k.val;kv.font={name:"Calibri",bold:true,size:16,color:{argb:"FFFFFFFF"}};kv.fill=xlBg("FF1E293B");kv.alignment=xlAlign("center","middle");ws1.getRow(kr+2).height=24;
            ws1.mergeCells(kr+3,col,kr+3,col);
            const ks1=ws1.getCell(kr+3,col);ks1.value=k.sub1;ks1.font={name:"Calibri",size:8,color:{argb:"FF64748B"},italic:true};ks1.fill=xlBg("FF1E293B");ks1.alignment=xlAlign("center","middle");ws1.getRow(kr+3).height=12;
            ws1.mergeCells(kr+4,col,kr+4,col);
            const ks2=ws1.getCell(kr+4,col);ks2.value=k.sub2;ks2.font={name:"Calibri",size:8,color:{argb:k.col},bold:true};ks2.fill=xlBg("FF1E293B");ks2.alignment=xlAlign("center","middle");ws1.getRow(kr+4).height=12;
            [kr,kr+1,kr+2,kr+3,kr+4].forEach(r=>{ws1.getCell(r,col).border={left:{style:"thin",color:{argb:"FF0F172A"}},right:{style:"thin",color:{argb:"FF0F172A"}},bottom:{style:"thin",color:{argb:"FF0F172A"}}};});
        });

        ws1.getRow(kr+5).height=10;

        let mr=kr+6;
        setHdr(ws1,mr,2,7,"  \u25a0  EVOLUCIÓN MENSUAL COMPARATIVA — Unidades por Mes",HDR_BLUE,11);mr++;
        const mHdrs=MESES.map((m,i)=>[m,3+i,3+i]);
        setColHdrs(ws1,mr,[["Métrica",2,2],...mHdrs,["Total",15,15]].slice(0,14),HDR_DARK);
        const byYearM={};const expYears=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
        expYears.forEach(y=>{byYearM[y]=new Array(12).fill(0);});
        data.forEach(d=>{if(byYearM[d.year])byYearM[d.year][d.month]+=d.units;});
        mr++;
        const addMonthRow=(ws,r,label,dataArr,total,bg,fontColor="FF1E293B",bold=false)=>{
            const lc=ws.getCell(r,2);lc.value=label;lc.font=xlFont(bold,fontColor,9);lc.fill=xlBg(bg);lc.alignment=xlAlign("left","middle");lc.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
            dataArr.forEach((v,i)=>{
                const cc=ws.getCell(r,3+i);cc.value=v;cc.numFmt="#,##0";
                cc.font=xlFont(bold,fontColor,9);cc.fill=xlBg(bg);cc.alignment=xlAlign("center","middle");cc.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
            });
            const tc=ws.getCell(r,15);tc.value=total;tc.numFmt="#,##0";
            tc.font=xlFont(true,fontColor,9);tc.fill=xlBg(bg);tc.alignment=xlAlign("center","middle");tc.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}},left:{style:"thin",color:{argb:"FFD1D5DB"}}};
            ws.getRow(r).height=16;
        };
        if(R.prevYear&&byYearM[R.prevYear]){
            addMonthRow(ws1,mr,`${R.prevYear} (Real)`,byYearM[R.prevYear],R.totalUnitsPrev,CELL_ALT,"FF374155");mr++;
        }
        const monthlyAvg=Math.round(R.totalUnitsLast/R.monthsElapsed);
        addMonthRow(ws1,mr,`${R.lastYear} (Real YTD)`,byYearM[R.lastYear].map((v,i)=>i<=R.maxMonth?v:""),R.totalUnitsLast,"FFEff6FF","FF1E3A5F",true);mr++;
        addMonthRow(ws1,mr,`${R.lastYear} (Proyectado)`,byYearM[R.lastYear].map((v,i)=>i<=R.maxMonth?v:monthlyAvg),Math.round(R.globalRunRate),"FFFFFBEB","FF78350F",false);mr++;
        if(R.prevYear&&byYearM[R.prevYear]){
            const varRow=byYearM[R.lastYear].map((v,i)=>{
                if(i>R.maxMonth)return "";
                const p=byYearM[R.prevYear][i];
                return p>0?`${((v/p-1)*100).toFixed(1)}%`:"N/D";
            });
            const lc=ws1.getCell(mr,2);lc.value="Var. % YoY (Real)";lc.font=xlFont(false,"FF374155",9);lc.fill=xlBg("FFF8FAFC");
            varRow.forEach((v,i)=>{
                const cc=ws1.getCell(mr,3+i);cc.value=v;
                const isPos=typeof v==="string"&&v.startsWith("+");const isNeg=typeof v==="string"&&v.startsWith("-");
                cc.font={name:"Calibri",bold:true,size:9,color:{argb:isPos?"FF059669":isNeg?"FFDC2626":"FF94A3B8"}};
                cc.fill=xlBg("FFF8FAFC");cc.alignment=xlAlign("center","middle");cc.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
            });
            ws1.getRow(mr).height=14;mr++;
        }
        ws1.getRow(mr).height=10;mr++;

        setHdr(ws1,mr,2,7,"  \u25a0  INSIGHTS AUTOMÁTICOS (Narrativa Ejecutiva)",HDR_GREEN,11);mr++;
        R.insights.forEach(ins=>{
            const bgTitle=ins.tipo==="Alerta"?"FFFFF1F2":ins.tipo==="Oportunidad"?"FFF0FDF4":ins.tipo==="Estrat\u00e9gico"?"FFFFFBEB":"FFEFF6FF";
            const accentC=ins.tipo==="Alerta"?"FFDC2626":ins.tipo==="Oportunidad"?"FF059669":ins.tipo==="Estrat\u00e9gico"?"FFF59E0B":"FF3B82F6";
            ws1.mergeCells(mr,2,mr,7);const ht=ws1.getCell(mr,2);
            ht.value=`  ${ins.tipo.toUpperCase()}  \u2014  ${ins.titulo}`;
            ht.font=xlFont(true,"FF1E293B",10);ht.fill=xlBg(bgTitle);ht.alignment=xlAlign("left","middle");
            ht.border={top:{style:"thin",color:{argb:accentC}},left:{style:"medium",color:{argb:accentC}},right:{style:"hair",color:{argb:accentC}}};
            ws1.getRow(mr).height=20;mr++;
            ws1.mergeCells(mr,2,mr,7);const bd=ws1.getCell(mr,2);
            bd.value="  "+ins.texto;
            bd.font=xlFont(false,"FF334155",10);bd.fill=xlBg(CELL_WHITE);bd.alignment=xlAlign("left","top",true);
            bd.border={left:{style:"medium",color:{argb:accentC}},bottom:{style:"thin",color:{argb:accentC}},right:{style:"hair",color:{argb:accentC}}};
            ws1.getRow(mr).height=48;mr++;ws1.getRow(mr).height=4;mr++;
        });

        // HOJA 2 – ANÁLISIS POR CATEGORÍA
        const ws2=wb.addWorksheet("2. An\u00e1lisis por Categor\u00eda",{views:[{showGridLines:false,state:"frozen",ySplit:2}]});
        ws2.columns=[
            {header:"Categor\u00eda",key:"cat",width:28},
            {header:`Vol. YTD ${R.prevYear||"Ant."}`,key:"prev",width:16},
            {header:`Vol. YTD ${R.lastYear}`,key:"last",width:16},
            {header:"Var. Absoluta",key:"vabs",width:14},
            {header:"Var. %",key:"vpct",width:12},
            {header:"Mix % YTD",key:"share",width:12},
            {header:"Proyecci\u00f3n Cierre",key:"proj",width:16},
            {header:`vs ${R.prevYear||"Ant."} Anual`,key:"projvar",width:14},
            {header:"Tendencia",key:"tend",width:14},
            {header:"SKUs Activos",key:"skus",width:12},
            {header:"Mejor SKU",key:"top",width:32},
        ];
        const hw2=ws2.getRow(1);hw2.font=xlFont(true,"FFFFFFFF",10);hw2.fill=xlBg(HDR_BLUE);hw2.alignment=xlAlign("center","middle");hw2.height=22;ws2.autoFilter="A1:K1";

        const catSkuMap={};R.skusSorted.forEach(s=>{if(!catSkuMap[s.category])catSkuMap[s.category]=[];catSkuMap[s.category].push(s);});

        R.cats.forEach((c,idx)=>{
            const skuCount=(catSkuMap[c.category]||[]).length;
            const topSku=(catSkuMap[c.category]||[]).sort((a,b)=>b.ytdLast-a.ytdLast)[0];
            const catRunRate=(c.ytdLast/R.monthsElapsed)*12;
            const catPrevFull=R.prevYear?data.filter(d=>d.year===R.prevYear&&d.category===c.category).reduce((s,d)=>s+d.units,0):null;
            const catProjVar=catPrevFull&&catPrevFull>0?((catRunRate/catPrevFull)-1)*100:null;
            const tc=c.tendencia>0?"\u25b2 Ascendente":(c.tendencia<0?"\u25bc Descendente":"\u25b6 Estable");
            const bg=idx%2===0?CELL_ALT:CELL_WHITE;
            const row=ws2.addRow({cat:c.category,prev:c.ytdPrev,last:c.ytdLast,vabs:c.varAbs,vpct:pct(c.ytdVar),share:c.share/100,proj:Math.round(catRunRate),projvar:catProjVar!==null?pct(catProjVar):"N/D",tend:tc,skus:skuCount,top:topSku?topSku.article:""});
            row.eachCell(cell=>cell.fill=xlBg(bg));row.font=xlFont(false,"FF1E293B",10);row.height=16;
            row.getCell("prev").numFmt="#,##0";row.getCell("last").numFmt="#,##0";row.getCell("vabs").numFmt="#,##0";
            row.getCell("share").numFmt="0.0%";row.getCell("proj").numFmt="#,##0";
            row.getCell("vabs").font=xlFont(true,varColorFn(c.varAbs),10);
            row.getCell("vpct").font=xlFont(true,varColorFn(c.ytdVar),10);
            row.getCell("tend").font=xlFont(true,c.tendencia>0?"FF059669":(c.tendencia<0?"FFDC2626":"FF94A3B8"),10);
            if(catProjVar!==null)row.getCell("projvar").font=xlFont(true,varColorFn(catProjVar),10);
        });

        // HOJA 3 – IMPULSORES Y RIESGOS
        const ws3=wb.addWorksheet("3. Impulsores y Riesgos",{views:[{showGridLines:false}]});
        [3,38,18,14,14,14,12,14,14,3].forEach((w,i)=>ws3.getColumn(i+1).width=w);
        ws3.getRow(1).height=8;ws3.mergeCells("B2:I3");
        const t3=ws3.getCell("B2");t3.value="IMPULSORES DE CRECIMIENTO Y FOCOS DE RIESGO";
        t3.font=xlFont(true,"FFFFFFFF",16);t3.fill=xlBg(HDR_DARK);t3.alignment=xlAlign("center","middle");
        ws3.getRow(2).height=22;ws3.getRow(3).height=22;ws3.getRow(4).height=8;

        const fmtImpRow=(ws,p,idx,r,isGrowth)=>{
            const bg=idx%2===0?(isGrowth?"FFF0FDF4":"FFFFF5F5"):CELL_WHITE;
            const vpct=p.ytdVar!==null?pct(p.ytdVar):"Nuevo";
            const tc=p.tendencia>0?"\u25b2 Asc.":(p.tendencia<0?"\u25bc Desc.":"\u25b6 Est.");
            const projStr=nf(Math.round(p.runRate));
            [[2,p.article],[3,p.category],[4,p.ytdPrev],[5,p.ytdLast],[6,p.varAbs],[7,vpct],[8,projStr],[9,tc]].forEach(([sc,val])=>{
                const c=ws.getCell(r,sc);c.value=val;c.fill=xlBg(bg);c.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
                const isNum=typeof val==="number";c.alignment=xlAlign(isNum||sc>3?"center":"left","middle");
                if(isNum)c.numFmt="#,##0";
                if(sc===6)c.font=xlFont(true,varColorFn(p.varAbs),10);
                else if(sc===7)c.font=xlFont(true,varColorFn(p.ytdVar),10);
                else if(sc===9)c.font=xlFont(true,p.tendencia>0?"FF059669":(p.tendencia<0?"FFDC2626":"FF94A3B8"),10);
                else c.font=xlFont(false,"FF1E293B",10);
            });ws.getRow(r).height=17;
        };

        let r3=5;
        setHdr(ws3,r3,2,9,"  \u25b2  TOP 15 IMPULSORES DE CRECIMIENTO (Mayor ganancia de volumen YTD)",HDR_GREEN);r3++;
        setColHdrs(ws3,r3,[["Art\u00edculo",2,2],["Categor\u00eda",3,3],[`Vol. ${R.prevYear||"Ant."}`,4,4],[`Vol. ${R.lastYear}`,5,5],["Var. Abs.",6,6],["Var. %",7,7],["Proy. Cierre",8,8],["Tendencia",9,9]],HDR_GREEN);r3++;
        R.skusSorted.filter(p=>p.varAbs>0).slice(0,15).forEach((p,idx)=>fmtImpRow(ws3,p,idx,r3++,true));

        r3++;
        setHdr(ws3,r3,2,9,"  \u25bc  TOP 15 FOCOS DE RIESGO (Mayor pérdida de volumen YTD)",HDR_RED);r3++;
        setColHdrs(ws3,r3,[["Art\u00edculo",2,2],["Categor\u00eda",3,3],[`Vol. ${R.prevYear||"Ant."}`,4,4],[`Vol. ${R.lastYear}`,5,5],["Var. Abs.",6,6],["Var. %",7,7],["Proy. Cierre",8,8],["Tendencia",9,9]],HDR_RED);r3++;
        [...R.skusSorted].filter(p=>p.varAbs<0).sort((a,b)=>a.varAbs-b.varAbs).slice(0,15).forEach((p,idx)=>fmtImpRow(ws3,p,idx,r3++,false));

        r3+=2;
        setHdr(ws3,r3,2,9,"  \u25cf  SKUs NUEVOS EN EL PERÍODO (Sin venta el año anterior)",HDR_VIOLET);r3++;
        setColHdrs(ws3,r3,[["Art\u00edculo",2,2],["Categor\u00eda",3,3],["Vol. YTD",5,5],["Mix %",6,6],["Proy. Cierre",8,8],["Tendencia",9,9]],HDR_VIOLET);r3++;
        R.skusSorted.filter(p=>p.ytdVar===null&&p.ytdLast>0).slice(0,10).forEach((p,idx)=>{
            const bg=idx%2===0?"FFF5F3FF":CELL_WHITE;
            [[2,p.article],[3,p.category],[5,p.ytdLast],[6,`${p.shareLast.toFixed(2)}%`],[8,Math.round(p.runRate)],[9,p.tendencia>0?"\u25b2 Asc.":"\u25bc Desc."]].forEach(([sc,val])=>{
                const c=ws3.getCell(r3,sc);c.value=val;c.fill=xlBg(bg);c.font=xlFont(false,"FF1E293B",10);
                c.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
                if(typeof val==="number")c.numFmt="#,##0";
            });ws3.getRow(r3).height=17;r3++;
        });

        // HOJA 4 – PROYECCIONES
        const ws4=wb.addWorksheet("4. Proyecciones de Cierre",{views:[{showGridLines:false,state:"frozen",ySplit:3}]});
        [3,38,14,14,14,14,14,14,14,3].forEach((w,i)=>ws4.getColumn(i+1).width=w);
        ws4.getRow(1).height=8;ws4.mergeCells("B2:I3");
        const t4=ws4.getCell("B2");t4.value=`PROYECCIONES DE CIERRE DE A\u00d1O ${R.lastYear} — Metodolog\u00eda Run-Rate`;
        t4.font=xlFont(true,"FFFFFFFF",16);t4.fill=xlBg(HDR_VIOLET);t4.alignment=xlAlign("center","middle");
        ws4.getRow(2).height=22;ws4.getRow(3).height=22;
        ws4.getRow(4).height=6;
        ws4.mergeCells("B5:I6");const nota=ws4.getCell("B5");
        nota.value=`\u26a0 METODOLOG\u00cdA: Proyecci\u00f3n Cierre = (Ventas YTD ${R.lastYear} / ${R.monthsElapsed} meses) \u00d7 12. Los meses proyectados (${MESES[R.maxMonth+1]||"N/A"} en adelante) usan el promedio mensual del per\u00edodo real.`;
        nota.font={name:"Calibri",bold:false,italic:true,size:9,color:{argb:"FF78350F"}};nota.fill=xlBg("FFFFFBEB");nota.alignment=xlAlign("left","middle",true);nota.border={left:{style:"medium",color:{argb:"FFF59E0B"}},top:{style:"thin",color:{argb:"FFF59E0B"}},bottom:{style:"thin",color:{argb:"FFF59E0B"}}};
        ws4.getRow(5).height=18;ws4.getRow(6).height=18;ws4.getRow(7).height=8;

        setHdr(ws4,8,2,9,"  PROYECCIÓN GLOBAL DEL PORTAFOLIO",HDR_VIOLET);
        ws4.getRow(9).height=6;
        const projSummary=[
            ["Métrica","Valor","Referencia"],
            [`Ventas YTD Real (${R.ytdLabel})`,R.totalUnitsLast,"Unidades confirmadas"],
            ["Promedio mensual real",Math.round(R.totalUnitsLast/R.monthsElapsed),"Unidades/mes"],
            [`Proyección Cierre ${R.lastYear}`,Math.round(R.globalRunRate),"Run-rate anualizado"],
            [R.prevYear?`${R.prevYear} Año Completo`:"Año anterior","N/D",R.prevYear?nf(prevFull||0):"Sin referencia"],
            ["Variación proyectada vs año ant.",projVar!==null?pct(projVar):"N/D",projVar!==null?(projVar>=0?"Superaría el año anterior":"Quedaría por debajo del año anterior"):"Sin año base"],
            [`Meses restantes para proyectar`,mesesRestantes,`${MESES[R.maxMonth+1]||"N/A"} – Diciembre`],
            ["Unidades proyectadas faltantes",proyFaltante,"Período estimado"],
        ];
        projSummary.forEach((row,ri)=>{
            const r=10+ri;const bg=ri===0?"FF0F172A":(ri%2===0?CELL_ALT:CELL_WHITE);
            [[2,row[0]],[4,row[1]],[7,row[2]]].forEach(([sc,val])=>{
                ws4.mergeCells(r,sc,r,sc===2?3:sc===4?6:9);
                const c=ws4.getCell(r,sc);c.value=val;c.fill=xlBg(bg);
                c.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
                if(ri===0){c.font=xlFont(true,"FFFFFFFF",10);c.alignment=xlAlign("center","middle");}
                else{c.font=xlFont(sc===4,ri%2===0?"FF1E293B":"FF374155",10);c.alignment=xlAlign(sc===4?"center":"left","middle");}
                if(typeof val==="number")c.numFmt="#,##0";
            });ws4.getRow(r).height=17;
        });

        ws4.getRow(18).height=10;
        setHdr(ws4,19,2,9,"  PROYECCIÓN POR SKU (Top 50 por Volumen Proyectado)",HDR_VIOLET,10);
        setColHdrs(ws4,20,[["Art\u00edculo",2,2],["Categor\u00eda",3,3],[`YTD Real ${R.lastYear}`,4,4],[`YTD ${R.prevYear||"Ant."}`,5,5],["Var. %",6,6],["Proy. Cierre",7,7],[`vs ${R.prevYear||"Ant."} Anual`,8,8],["Semáforo",9,9]],HDR_VIOLET);
        let r4=21;
        [...R.skusSorted].sort((a,b)=>b.runRate-a.runRate).slice(0,50).forEach((p,idx)=>{
            const bg=idx%2===0?CELL_ALT:CELL_WHITE;
            const sku_prev_full=R.prevYear?data.filter(d=>d.year===R.prevYear&&d.article===p.article).reduce((s,d)=>s+d.units,0):null;
            const skuProjVar=sku_prev_full&&sku_prev_full>0?((p.runRate/sku_prev_full)-1)*100:null;
            const semaforoBg=p.semaforo==="VERDE"?"FF166534":p.semaforo==="AMARILLO"?"FF78350F":"FF7F1D1D";
            [[2,p.article],[3,p.category],[4,p.ytdLast],[5,p.ytdPrev],[6,pct(p.ytdVar)],[7,Math.round(p.runRate)],[8,skuProjVar!==null?pct(skuProjVar):"N/D"],[9,p.semaforo]].forEach(([sc,val])=>{
                const c=ws4.getCell(r4,sc);c.value=val;
                if(sc===9){c.fill=xlBg(semaforoBg);c.font=xlFont(true,"FFFFFFFF",9);}else{c.fill=xlBg(bg);c.font=xlFont(false,"FF1E293B",9);}
                c.border={bottom:{style:"hair",color:{argb:"FFE2E8F0"}}};
                c.alignment=xlAlign(sc>3?"center":"left","middle");
                if(typeof val==="number")c.numFmt="#,##0";
                if(sc===6)c.font=xlFont(true,varColorFn(p.ytdVar),9);
                if(sc===8&&skuProjVar!==null)c.font=xlFont(true,varColorFn(skuProjVar),9);
            });ws4.getRow(r4).height=16;r4++;
        });

        // HOJA 5 – BASE COMPLETA BI
        const ws5=wb.addWorksheet("5. Base Completa BI",{views:[{showGridLines:true,state:"frozen",ySplit:1}]});
        ws5.columns=[
            {header:"Ranking",key:"rank",width:10},{header:"Art\u00edculo",key:"art",width:48},
            {header:"Categor\u00eda",key:"cat",width:22},{header:"Marca",key:"brand",width:16},
            {header:`YTD ${R.lastYear}`,key:"ytd",width:16},{header:`YTD ${R.prevYear||"Ant."}`,key:"yprev",width:14},
            {header:"Var. Absoluta",key:"varabs",width:14},{header:"Var. %",key:"var",width:12},
            {header:"Mix % YTD",key:"share",width:11},{header:"\u0394 Mix %",key:"sharevar",width:11},
            {header:"Proy. Cierre",key:"runrate",width:14},{header:"Prom. Mensual",key:"mavg",width:14},
            {header:"Tendencia",key:"tend",width:14},{header:"Clasificaci\u00f3n",key:"clasif",width:18},
            {header:"Sem\u00e1foro",key:"sem",width:12},{header:"\u00dalt. Venta",key:"lastsale",width:16},
        ];
        const hw5=ws5.getRow(1);hw5.font=xlFont(true,"FFFFFFFF",10);hw5.fill=xlBg(HDR_DARK);hw5.alignment=xlAlign("center","middle");hw5.height=22;ws5.autoFilter="A1:P1";
        R.skusSorted.forEach((p,idx)=>{
            const vs=pct(p.ytdVar);const tc=p.tendencia>0?"Ascendente":(p.tendencia<0?"Descendente":"Estable");
            const mavg=Math.round(p.ytdLast/R.monthsElapsed);
            const bg=idx%2===0?CELL_ALT:CELL_WHITE;
            const semBg=p.semaforo==="VERDE"?"FF166534":p.semaforo==="AMARILLO"?"FF78350F":"FF7F1D1D";
            const row=ws5.addRow({rank:idx+1,art:p.article,cat:p.category,brand:p.brand,ytd:p.ytdLast,yprev:p.ytdPrev,varabs:p.varAbs,var:vs,share:p.shareLast/100,sharevar:p.shareVar/100,runrate:Math.round(p.runRate),mavg,tend:tc,clasif:p.clasif,sem:p.semaforo,lastsale:p.lastSaleLabel});
            row.eachCell(c=>c.fill=xlBg(bg));row.font=xlFont(false,"FF1E293B",10);row.height=16;
            ["ytd","yprev","varabs","runrate","mavg"].forEach(k=>row.getCell(k).numFmt="#,##0");
            ["share","sharevar"].forEach(k=>row.getCell(k).numFmt="0.0%");
            row.getCell("varabs").font=xlFont(true,varColorFn(p.varAbs),10);
            row.getCell("var").font=xlFont(true,varColorFn(p.ytdVar),10);row.getCell("var").alignment={horizontal:"center",vertical:"middle"};
            row.getCell("sharevar").font=xlFont(true,varColorFn(p.shareVar),10);
            row.getCell("tend").font=xlFont(true,p.tendencia>0?"FF059669":(p.tendencia<0?"FFDC2626":"FF94A3B8"),10);
            row.getCell("sem").fill=xlBg(semBg);row.getCell("sem").font=xlFont(true,"FFFFFFFF",9);row.getCell("sem").alignment={horizontal:"center",vertical:"middle"};
            row.getCell("rank").alignment={horizontal:"center",vertical:"middle"};
        });

        const buf=await wb.xlsx.writeBuffer();
        saveAs(new Blob([buf]),`Reporte_BI_${R.lastYear}_${new Date().toISOString().split("T")[0]}.xlsx`);
    }catch(e){console.error(e);alert("Error al generar el reporte. Verifica la consola del navegador.");}
    finally{btn.innerHTML=origText;lucide.createIcons();}
}// ---- Exportaci\u00f3n L\u00edderes ----
function extractFormatInfo(articleName){
    const m=articleName.match(/(?:X\s*)?(\d+(?:\.\d+)?)\s*(ML|LT|L|GR|G|KG)\b/i)||articleName.match(/\b(\d+(?:\.\d+)?)\s*(ML|LT|L|GR|G|KG)\b/i);
    if(m)return{raw:m[0].toUpperCase(),number:parseFloat(m[1]),unit:m[2].toUpperCase()};return{raw:"N/D",number:0,unit:"N/D"};
}
function extractFlavor(articleName,brand,formatRaw){
    let s=articleName.toUpperCase();if(brand&&brand!=="Desconocida")s=s.replace(brand.toUpperCase(),"");if(formatRaw&&formatRaw!=="N/D")s=s.replace(formatRaw,"");
    ["BEBIDA","ISOTONICA","HIDRATANTE","REHIDRATANTE","PACK","CAJA","BOTELLA","BOT","PAGA","LLEVA","SABOR","SPORT","ZERO","MUNDIAL"].forEach(w=>{s=s.replace(new RegExp(`\\b${w}\\b`,"g"),"");});
    s=s.replace(/[0-9]/g,"").replace(/[^\w\s-]/g,"").trim().replace(/\s+/g," ");return s||"N/D";
}

async function exportCategoryLeaders(){
    const data=getFilteredData();if(data.length===0){alert("No hay datos para exportar.");return;}
    const btn=document.getElementById("btn-export-leaders");const origText=btn.innerHTML;
    btn.innerHTML='<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> Generando...';lucide.createIcons();
    try{
        const years=[...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
        const lastYear=years[years.length-1];const prevYear=years.length>1?years[years.length-2]:null;
        const lya=data.filter(d=>d.year===lastYear&&d.units>0);const maxMonth=lya.length>0?Math.max(...lya.map(d=>d.month)):11;
        const skuMap={};
        data.forEach(d=>{
            if(!skuMap[d.article]){const f=extractFormatInfo(d.article);skuMap[d.article]={article:d.article,category:d.category,brand:d.brand,formatNum:f.number||"N/D",formatUnit:f.unit,flavor:extractFlavor(d.article,d.brand,f.raw),total:0,ytdLast:0,ytdPrev:0,monthlySales:{}};}
            const s=skuMap[d.article];s.total+=d.units;
            if(d.month<=maxMonth){if(d.year===lastYear)s.ytdLast+=d.units;if(prevYear&&d.year===prevYear)s.ytdPrev+=d.units;}
            s.monthlySales[`${d.year}_${d.month}`]=(s.monthlySales[`${d.year}_${d.month}`]||0)+d.units;
        });
        Object.values(skuMap).forEach(p=>{p.ytdVar=p.ytdPrev>0?((p.ytdLast/p.ytdPrev)-1)*100:(p.ytdLast>0?null:-100);});
        const list=Object.values(skuMap).sort((a,b)=>{if(a.category<b.category)return -1;if(a.category>b.category)return 1;return b.total-a.total;});
        const wb=new ExcelJS.Workbook();wb.creator="Sales Analytics BI";wb.created=new Date();
        const ws=wb.addWorksheet("L\u00edderes por Categor\u00eda",{views:[{showGridLines:true,state:"frozen",ySplit:2}]});
        const baseH=[{header:"Ranking Cat.",key:"rank",width:13},{header:"ARTICULO",key:"art",width:45},{header:"CAT 3",key:"cat",width:22},{header:"Marca",key:"brand",width:18},{header:"Gramaje",key:"formatNum",width:12},{header:"Medida",key:"formatUnit",width:10},{header:"Sabor/Variante",key:"flavor",width:22},{header:"Total General",key:"total",width:14},{header:`YTD ${prevYear||"Ant."}`,key:"yprev",width:14},{header:`YTD ${lastYear}`,key:"ylast",width:14},{header:"Var. YTD %",key:"ytdvar",width:12}];
        const mH=[];years.forEach(y=>{MESES.forEach((mn,mi)=>{mH.push({header:`${mn} ${y}`,key:`m_${y}_${mi}`,width:11});});});
        ws.columns=[...baseH,...mH];const totalCols=baseH.length+mH.length;
        ws.mergeCells(1,1,1,totalCols);const tc=ws.getCell("A1");
        tc.value=`L\u00cdDERES POR CATEGOR\u00cdA \u2014 Exportado: ${new Date().toLocaleDateString("es-ES",{day:"2-digit",month:"long",year:"numeric"})} | Per\u00edodo: ${lastYear}`;
        tc.font={name:"Calibri",bold:true,size:13,color:{argb:"FFFFFFFF"}};tc.fill=xlBg("FF059669");tc.alignment=xlAlign("center","middle");ws.getRow(1).height=22;
        ws.getRow(2).height=18;[...baseH,...mH].forEach((h,i)=>{const c=ws.getCell(2,i+1);c.value=h.header;c.font=xlFont(true,"FFFFFFFF",10);c.fill=xlBg("FF065F46");c.alignment=xlAlign("center","middle");c.border=xlBorder("thin","FF059669");});
        let currentCat=null,rankCounter=1;
        list.forEach((item,rowIdx)=>{
            if(item.category!==currentCat){currentCat=item.category;rankCounter=1;}
            const vs=item.ytdVar!==null?`${item.ytdVar>=0?"+":""}${item.ytdVar.toFixed(1)}%`:"Nuevo";
            const rowData={rank:rankCounter,art:item.article,cat:item.category,brand:item.brand,formatNum:item.formatNum,formatUnit:item.formatUnit,flavor:item.flavor,total:item.total,yprev:item.ytdPrev,ylast:item.ytdLast,ytdvar:vs};
            years.forEach(y=>{MESES.forEach((_,mi)=>{rowData[`m_${y}_${mi}`]=item.monthlySales[`${y}_${mi}`]||0;});});
            const r=ws.getRow(rowIdx+3);[...baseH,...mH].forEach((h,i)=>{r.getCell(i+1).value=rowData[h.key];});
            const isL=rankCounter===1;r.fill=xlBg(isL?"FFE0F7EC":(rowIdx%2===0?"FFF8FAFC":"FFFFFFFF"));
            r.font=xlFont(isL,"FF1E293B",10);r.height=18;
            r.getCell(8).numFmt="#,##0";r.getCell(9).numFmt="#,##0";r.getCell(10).numFmt="#,##0";
            r.getCell(11).font=xlFont(true,varColorFn(item.ytdVar),10);r.getCell(11).alignment=xlAlign("center","middle");
            for(let mi=0;mi<mH.length;mi++){r.getCell(baseH.length+1+mi).numFmt="#,##0";r.getCell(baseH.length+1+mi).alignment=xlAlign("center","middle");}
            if(isL){for(let ci=1;ci<=totalCols;ci++){r.getCell(ci).border={top:{style:"medium",color:{argb:"FF059669"}}};}}
            rankCounter++;
        });
        ws.autoFilter={from:{row:2,column:1},to:{row:2,column:totalCols}};
        const buf=await wb.xlsx.writeBuffer();saveAs(new Blob([buf]),`Lideres_Categoria_${lastYear}_${new Date().toISOString().split("T")[0]}.xlsx`);
    }catch(e){console.error(e);alert("Error al exportar l\u00edderes.");}finally{btn.innerHTML=origText;lucide.createIcons();}
}

// Event delegation: works even before dashboard is shown
document.addEventListener('click', function(e){
    const t = e.target.closest('button');
    if(!t) return;
    if(t.id==='btn-export-analysis') exportToExcel();
    if(t.id==='btn-export-leaders') exportCategoryLeaders();
    if(t.id==='btn-new-file') {
        globalData = [];
        // ✅ FIX: reset completo de los datasets multi-hoja
        globalDataNati = [];
        globalDataAleUnits = [];
        globalDataAleBS = [];
        _aleFiltersPopulatedFor = null;
        filters = { year: "all", month: "all", category: "all", brand: "all", mp: "all" };
        Object.values(charts).forEach(c => { if(c && c.destroy) c.destroy(); });
        charts = {};
        dashboardLayout.classList.add('hidden');
        dashboardLayout.classList.remove('flex');
        uploadView.classList.remove('hidden');
        fileInput.value = '';
    }
});

// ==========================================
// Módulo: Marca Propia (MP)
// ==========================================
function renderMPAnalysis(data) {
    const years = [...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length-1];
    const prevYear = years.length>1 ? years[years.length-2] : null;

    const lya = data.filter(d=>d.year===lastYear && d.units>0);
    const maxMonth = lya.length>0 ? Math.max(...lya.map(d=>d.month)) : 11;

    const ytdAllLast = data.filter(d=>d.year===lastYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0);
    const ytdAllPrev = prevYear ? data.filter(d=>d.year===prevYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0) : 0;

    const mpData = data.filter(d=>d.isMP);
    const restData = data.filter(d=>!d.isMP);

    const ytdMpLast = mpData.filter(d=>d.year===lastYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0);
    const ytdMpPrev = prevYear ? mpData.filter(d=>d.year===prevYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0) : 0;

    const ytdRestLast = restData.filter(d=>d.year===lastYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0);
    const ytdRestPrev = prevYear ? restData.filter(d=>d.year===prevYear && d.month<=maxMonth).reduce((s,d)=>s+d.units,0) : 0;

    const mpShareLast = ytdAllLast > 0 ? (ytdMpLast / ytdAllLast) * 100 : 0;
    const kpiContainer = document.getElementById("kpi-container");
    
    const mpExists = mpData.length > 0;
    let kpiHtml = '';
    if(!mpExists || ytdMpLast === 0) {
        kpiHtml = `
        <div class="glass-panel border-l-2 border-l-slate-400 relative overflow-hidden group">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="shield-off" class="w-16 h-16 text-slate-400"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">Penetración MP</p>
            <p class="text-2xl font-bold text-slate-500 tracking-tight mt-1">N/A</p>
            <p class="text-xs text-muted mt-2">Sin marca propia</p>
        </div>`;
    } else {
        const mpSharePrev = ytdAllPrev > 0 ? (ytdMpPrev / ytdAllPrev) * 100 : 0;
        const shareDelta = mpShareLast - mpSharePrev;
        const color = shareDelta >= 0 ? 'emerald-500' : 'red-500';
        kpiHtml = `
        <div class="glass-panel border-l-2 border-l-emerald-500 relative overflow-hidden group cursor-pointer hover:ring-2 hover:ring-emerald-500/50 transition-all audit-trigger" data-audit="mpShare">
            <div class="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                <i data-lucide="shield-check" class="w-16 h-16 text-emerald-500"></i>
            </div>
            <p class="text-[0.65rem] text-muted font-semibold uppercase tracking-widest mb-3">Penetración MP</p>
            <p class="text-3xl font-bold text-accent tracking-tight">${mpShareLast.toFixed(1)}%</p>
            <p class="text-xs text-muted mt-2 flex items-center gap-1">
                <span class="text-${color} font-medium">${shareDelta >= 0 ? '+' : ''}${shareDelta.toFixed(1)} pp</span> vs año ant.
            </p>
        </div>`;
    }
    
    if(kpiContainer) {
        kpiContainer.innerHTML += kpiHtml;
    }

    const panel = document.getElementById("mp-comparative-container");
    if(!panel) return;
    if(!mpExists || ytdMpLast === 0) {
        panel.innerHTML = `<div class="col-span-1 md:col-span-3 text-sm text-muted p-4 text-center">No hay artículos de Marca Propia en esta selección.</div>`;
        return;
    }

    const deltaTotal = ytdAllLast - ytdAllPrev;
    const deltaMp = ytdMpLast - ytdMpPrev;
    const deltaRest = ytdRestLast - ytdRestPrev;

    const mpSharePrev = ytdAllPrev > 0 ? (ytdMpPrev / ytdAllPrev) * 100 : 0;
    const shareDelta = mpShareLast - mpSharePrev;

    let contributionMp = "N/A";
    if (deltaTotal > 0) {
        if (deltaMp > 0) contributionMp = ((deltaMp / deltaTotal) * 100).toFixed(1) + "%";
        else contributionMp = "0%";
    }

    const varMp = ytdMpPrev > 0 ? ((ytdMpLast / ytdMpPrev) - 1) * 100 : null;
    const varRest = ytdRestPrev > 0 ? ((ytdRestLast / ytdRestPrev) - 1) * 100 : null;

    const formatVar = (v) => v === null ? "N/D" : `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`;
    const getColor = (v) => v === null ? "text-slate-400" : (v >= 0 ? "text-emerald-400" : "text-red-400");

    let winnerText = "";
    if(varMp !== null && varRest !== null) {
        if(varMp > varRest) winnerText = "🔥 Marca Propia lidera el crecimiento relativo";
        else if(varRest > varMp) winnerText = "⚠️ Resto de marcas impulsan el desempeño";
        else winnerText = "Igualdad en tracción de ventas";
    } else {
        winnerText = "Sin datos previos para análisis YTD";
    }

    panel.innerHTML = `
        <div class="glass-panel p-5 flex flex-col justify-between border-t-2 border-t-blue-500">
            <div class="flex items-center justify-between mb-3">
                <p class="text-[0.65rem] text-muted font-bold uppercase tracking-widest">Penetración MP (Mix)</p>
                <i data-lucide="pie-chart" class="w-4 h-4 text-blue-500"></i>
            </div>
            <div class="flex items-end gap-2 mb-4">
                <span class="text-3xl font-bold text-accent">${mpShareLast.toFixed(1)}%</span>
                <span class="text-sm font-medium ${shareDelta >= 0 ? 'text-emerald-600' : 'text-red-600'} mb-1">
                    ${shareDelta >= 0 ? '+' : ''}${shareDelta.toFixed(1)} pp
                </span>
            </div>
            <div class="space-y-1">
                <div class="flex justify-between text-xs text-muted"><span class="text-emerald-600 font-medium">MP: ${mpShareLast.toFixed(1)}%</span><span>Resto: ${(100-mpShareLast).toFixed(1)}%</span></div>
                <div class="w-full bg-gray-200 flex h-2 rounded-full overflow-hidden">
                    <div class="bg-emerald-500 h-full" style="width: ${mpShareLast}%"></div>
                    <div class="bg-gray-300 h-full" style="width: ${100-mpShareLast}%"></div>
                </div>
            </div>
        </div>

        <div class="glass-panel p-5 flex flex-col justify-between border-t-2 border-t-emerald-500">
            <div class="flex items-center justify-between mb-3">
                <p class="text-[0.65rem] text-muted font-bold uppercase tracking-widest">Velocidad (Var. YTD)</p>
                <i data-lucide="rocket" class="w-4 h-4 text-emerald-500"></i>
            </div>
            <div class="space-y-4 mt-1">
                <div>
                    <div class="flex justify-between text-xs mb-1"><span class="font-medium text-emerald-600">Marca Propia</span><span class="font-bold ${getColor(varMp)}">${formatVar(varMp)}</span></div>
                    <div class="w-full bg-gray-200 rounded-full h-1.5"><div class="bg-emerald-500 h-1.5 rounded-full" style="width: ${Math.min(Math.max((varMp||0)+50, 0), 100)}%"></div></div>
                </div>
                <div>
                    <div class="flex justify-between text-xs mb-1"><span class="font-medium text-muted">Resto Marcas</span><span class="font-bold ${getColor(varRest)}">${formatVar(varRest)}</span></div>
                    <div class="w-full bg-gray-200 rounded-full h-1.5"><div class="bg-gray-400 h-1.5 rounded-full" style="width: ${Math.min(Math.max((varRest||0)+50, 0), 100)}%"></div></div>
                </div>
            </div>
        </div>

        <div class="glass-panel p-5 flex flex-col justify-between bg-blue-50/50 border border-brand/15">
            <div class="flex items-center justify-between mb-3">
                <p class="text-[0.65rem] text-brand font-bold uppercase tracking-widest">Insights Gerenciales</p>
                <i data-lucide="award" class="w-4 h-4 text-brand"></i>
            </div>
            <p class="text-sm font-bold text-accent leading-snug mb-3">${winnerText}</p>
            <div class="border-t border-border pt-3 space-y-2">
                <p class="text-xs text-muted flex justify-between">Volumen MP: <span class="font-semibold text-accent">${ytdMpLast.toLocaleString("es-ES")}</span></p>
                <p class="text-xs text-muted flex justify-between">Volumen Resto: <span class="font-semibold text-accent">${ytdRestLast.toLocaleString("es-ES")}</span></p>
                ${deltaTotal > 0 && deltaMp > 0 ? `<p class="text-xs flex justify-between mt-1 text-emerald-600 font-medium">Aporte Crecimiento Cat.: <span>${contributionMp}</span></p>` : ''}
            </div>
        </div>
    `;
}

// ==========================================
// Módulo: Auditoría de Cálculos
// ==========================================
document.addEventListener('click', function(e){
    const auditCard = e.target.closest('.audit-trigger');
    if(auditCard) {
        const type = auditCard.getAttribute('data-audit');
        openAuditModal(type);
    }
    const closeBtn = e.target.closest('#close-audit');
    if(closeBtn) {
        closeAuditModal();
    }
});

const auditModal = document.getElementById('audit-modal');
if(auditModal) {
    auditModal.addEventListener('click', function(e){
        if(e.target === this) closeAuditModal();
    });
}

function closeAuditModal() {
    const modal = document.getElementById('audit-modal');
    const content = document.getElementById('audit-modal-content');
    content.classList.replace('scale-100', 'scale-95');
    modal.classList.replace('opacity-100', 'opacity-0');
    setTimeout(() => { modal.classList.add('hidden'); modal.classList.remove('flex'); }, 200);
}

function openAuditModal(type, extraParam) {
    const data = getFilteredData(true);
    if(data.length === 0) return;
    
    const years = [...new Set(data.map(d=>d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length-1];
    const prevYear = years.length>1 ? years[years.length-2] : null;

    const lya = data.filter(d=>d.year===lastYear && d.units>0);
    const maxMonth = lya.length>0 ? Math.max(...lya.map(d=>d.month)) : 11;
    const monthsElapsed = maxMonth + 1;
    const isAllYears = (filters.year === "all");

    let title = "Auditoría";
    let formula = "";
    let headers = [];
    let rows = [];

    const skuMap = {};
    data.forEach(d => {
        if(!skuMap[d.article]) skuMap[d.article] = { art: d.article, cat: d.category, brand: d.brand, ytdL: 0, ytdP: 0, totalL: 0, totalL_all: 0, isMP: d.isMP };
        const s = skuMap[d.article];
        if(d.year === lastYear) s.totalL += d.units;
        s.totalL_all += d.units;
        if(d.month <= maxMonth) {
            if(d.year === lastYear) s.ytdL += d.units;
            if(prevYear && d.year === prevYear) s.ytdP += d.units;
        }
    });

    const skus = Object.values(skuMap);

    switch(type) {
        case "totalUnits":
            title = isAllYears ? `Auditoría: Total Histórico` : `Auditoría: Volumen Acumulado ${lastYear}`;
            formula = isAllYears 
                ? `Suma de todas las unidades vendidas en todos los años disponibles.` 
                : `Suma de todas las unidades vendidas en ${lastYear} (Enero a ${MESES[maxMonth]}).`;
            headers = ["Artículo", "Categoría", "Marca", "Total Unidades"];
            rows = skus.filter(s => (isAllYears ? s.totalL_all : s.ytdL) > 0).sort((a,b)=>(isAllYears ? b.totalL_all - a.totalL_all : b.ytdL - a.ytdL)).map(s => [
                s.art, s.cat, s.brand, `<span class="font-bold text-blue-600">${(isAllYears ? s.totalL_all : s.ytdL).toLocaleString("es-ES")}</span>`
            ]);
            break;
            
        case "varYtd":
            title = `Variación Acumulada YTD`;
            formula = `Fórmula: ((Unidades YTD ${lastYear} / Unidades YTD ${prevYear||'N/A'}) - 1) * 100. Periodo: Ene-${MESES[maxMonth]}.`;
            headers = ["Artículo", "Marca", `YTD ${prevYear||'-'}`, `YTD ${lastYear}`, "Variación"];
            rows = skus.filter(s => s.ytdL > 0 || s.ytdP > 0).sort((a,b)=>b.ytdL - a.ytdL).map(s => {
                const isNew = s.ytdP === 0;
                const v = isNew ? null : ((s.ytdL/s.ytdP)-1)*100;
                const vT = isNew ? "Nuevo" : `${v>=0?"+":""}${v.toFixed(1)}%`;
                const vC = isNew ? "text-blue-600" : (v>=0?"text-emerald-600":"text-red-600");
                return [s.art, s.brand, s.ytdP.toLocaleString("es-ES"), s.ytdL.toLocaleString("es-ES"), `<span class="font-medium ${vC}">${vT}</span>`];
            });
            break;
            
        case "activeSkus":
            title = "SKUs Activos";
            formula = `Conteo de artículos distintos que registraron ventas mayores a 0 durante ${lastYear}.`;
            headers = ["Artículo", "Categoría", "Marca", "Total Año"];
            rows = skus.filter(s => s.totalL > 0).sort((a,b)=>b.totalL - a.totalL).map(s => [
                s.art, s.cat, s.brand, s.totalL.toLocaleString("es-ES")
            ]);
            break;
            
        case "proj":
            title = "Proyección Cierre Anual";
            formula = `Fórmula Run-Rate: (Unidades Acumuladas YTD / ${monthsElapsed} meses transcurridos) * 12 meses.`;
            headers = ["Artículo", "Acumulado YTD", "Promedio Mensual", "Proyección Cierre"];
            rows = skus.filter(s => s.ytdL > 0).sort((a,b)=>b.ytdL - a.ytdL).map(s => {
                const avg = s.ytdL / monthsElapsed;
                const proj = Math.round(avg * 12);
                return [s.art, s.ytdL.toLocaleString("es-ES"), Math.round(avg).toLocaleString("es-ES"), `<span class="font-bold text-violet-600">${proj.toLocaleString("es-ES")}</span>`];
            });
            break;

        case "mpShare":
            title = "Penetración Marca Propia (Mix %)";
            formula = `Fórmula: (Unidades Marca Propia YTD / Total Unidades YTD) * 100.`;
            headers = ["Artículo", "Clasificación", "Unidades YTD", "Peso sobre Total"];
            const totalL = skus.reduce((a,b) => a + b.ytdL, 0);
            rows = skus.filter(s => s.ytdL > 0).sort((a,b)=>b.ytdL - a.ytdL).map(s => {
                const w = totalL > 0 ? (s.ytdL / totalL) * 100 : 0;
                return [
                    s.art, 
                    s.isMP ? '<span class="text-emerald-600 font-bold"><i data-lucide="shield-check" class="w-3 h-3 inline mr-1"></i>Marca Propia</span>' : 'Resto Marcas', 
                    s.ytdL.toLocaleString("es-ES"), 
                    `${w.toFixed(2)}%`
                ];
            });
            break;
            
        case "bestMonth":
            title = "Mejor Mes";
            formula = `Mes con el mayor volumen de unidades totales vendidas en ${lastYear}.`;
            headers = ["Artículo", "Marca", "Ventas en el Mejor Mes"];
            
            const mm={}; data.filter(d=>d.year===lastYear).forEach(d=>{mm[d.monthName]=(mm[d.monthName]||0)+d.units;});
            let bestMonth="-",maxU=-1;for(const k in mm){if(mm[k]>maxU){maxU=mm[k];bestMonth=k;}}
            
            const bestMonthData = {};
            data.filter(d=>d.year===lastYear && d.monthName===bestMonth).forEach(d => {
                if(!bestMonthData[d.article]) bestMonthData[d.article] = {art: d.article, brand: d.brand, units: 0};
                bestMonthData[d.article].units += d.units;
            });
            
            rows = Object.values(bestMonthData).filter(s => s.units > 0).sort((a,b)=>b.units - a.units).map(s => [
                s.art, s.brand, `<span class="font-bold text-amber-600">${s.units.toLocaleString("es-ES")}</span>`
            ]);
            break;

        case "brandTop":
            title = `Top Sellers: ${extraParam}`;
            formula = `Ranking de artículos de la marca "${extraParam}" ordenados por volumen YTD en ${lastYear}.`;
            headers = ["#", "Artículo", "Categoría", "Unidades Acumuladas"];
            let brandSkus = skus.filter(s => s.brand === extraParam && s.ytdL > 0);
            brandSkus.sort((a,b) => b.ytdL - a.ytdL);
            rows = brandSkus.map((s, idx) => [
                idx + 1, s.art, s.cat, `<span class="font-bold text-brand">${s.ytdL.toLocaleString("es-ES")}</span>`
            ]);
            break;

        default:
            return;
    }

    document.getElementById("audit-title").innerText = title;
    document.getElementById("audit-formula").innerHTML = formula;
    
    const thead = document.getElementById("audit-thead-row");
    thead.innerHTML = headers.map((h, j) => `<th class="px-4 py-3 font-semibold text-muted uppercase tracking-wider text-xs ${j >= headers.length - 2 ? 'text-right' : 'text-left'}">${h}</th>`).join("");
    
    const tbody = document.getElementById("audit-tbody");
    tbody.innerHTML = rows.map((r, i) => `
        <tr class="hover:bg-blue-50 transition-colors">
            ${r.map((c, j) => `<td class="px-4 py-2.5 ${j===0?'text-gray-800 font-medium':'text-gray-600'} ${j >= r.length - 2 ? 'text-right' : 'text-left'}">${c}</td>`).join("")}
        </tr>
    `).join("");

    lucide.createIcons();

    const modal = document.getElementById('audit-modal');
    const content = document.getElementById('audit-modal-content');
    modal.classList.remove('hidden');
    modal.classList.add('flex');
    setTimeout(() => {
        modal.classList.replace('opacity-0', 'opacity-100');
        content.classList.replace('scale-95', 'scale-100');
    }, 10);
}

// ============================================================
// MULTI-SHEET AUTO-LOAD
// ============================================================
let globalDataNati = [];
let globalDataAleUnits = [];
let globalDataAleBS = [];
let activeTeam = 'nati';
let aleMode = 'units'; // 'units' | 'bs'
let aleFilters = { year: 'all', month: 'all', brand: 'all', category: 'all', tipo: 'all', sku: '' };
let aleCharts = {};

async function autoLoadData() {
    const statusText = document.getElementById("status-text");
    const loadingBar = document.querySelector(".loading-bar");

    try {
        const res = await fetch('nuevos_datos.xlsx');
        if (!res.ok) throw new Error("No se encontró el archivo");

        const arrayBuffer = await res.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const wb = XLSX.read(data, { type: "array" });

        if(statusText) statusText.innerText = "Procesando hojas de datos...";

        const sheetNames = wb.SheetNames;
        const natiSheetName = sheetNames.find(s => s.toUpperCase().includes('NATI'));
        const aleUnitsSheetName = sheetNames.find(s => s.toUpperCase().includes('UNIDADES') && s.toUpperCase().includes('ALE'));
        const aleBSSheetName = sheetNames.find(s => s.toUpperCase().includes('BS') && s.toUpperCase().includes('ALE'));

        if (natiSheetName) {
            const ws = wb.Sheets[natiSheetName];
            const jd = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
            globalDataNati = parseNatiSheet(jd);
            globalData = globalDataNati;
        }

        if (aleUnitsSheetName) {
            const ws = wb.Sheets[aleUnitsSheetName];
            const jd = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
            globalDataAleUnits = parseAleSheet(jd);
        }

        if (aleBSSheetName) {
            const ws = wb.Sheets[aleBSSheetName];
            const jd = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
            globalDataAleBS = parseAleSheet(jd);
        }

        if (!natiSheetName && !aleUnitsSheetName && !aleBSSheetName) {
            throw new Error(`No se encontró ninguna hoja esperada. Hojas presentes en el Excel: ${sheetNames.join(', ')}`);
        }

        if(statusText) statusText.innerText = `Listo: ${globalDataNati.length} reg. NATI · ${globalDataAleUnits.length} reg. ALE`;
        setTimeout(() => initAllDashboards(), 500);

    } catch (err) {
        console.error(err);
        if(loadingBar) {
            loadingBar.style.animationPlayState = 'paused';
        }
        if (window.location.protocol === 'file:') {
            if(statusText) statusText.innerHTML = "<span class='text-red-500 font-semibold'>Error de seguridad local:</span><br>Los navegadores bloquean la carga automática al abrir el HTML con doble clic.<br>Sube los archivos a GitHub Pages y funcionará perfectamente.";
        } else {
            if(statusText) statusText.innerHTML = "<span class='text-red-500 font-semibold'>Error:</span> No se encontró el archivo Excel.<br>Asegúrate de que se llame exactamente <b>nuevos_datos.xlsx</b> y esté subido a GitHub.";
        }
    }
}

// ============================================================
// NATI PARSER
// ============================================================
function parseNatiSheet(matrix) {
    if (matrix.length < 2) return [];
    let hIdx = 0;
    for (let i = 0; i < Math.min(15, matrix.length); i++) {
        if (matrix[i].some(c => typeof c === 'string' && (c.toLowerCase().includes('articulo') || c.toLowerCase().includes('artículo')))) {
            hIdx = i; break;
        }
    }
    const headers = matrix[hIdx].map(h => String(h).trim());
    const colYears = [];
    let cy = new Date().getFullYear();
    for (let col = 0; col < headers.length; col++) {
        let yf = null;
        for (let r = 0; r < hIdx; r++) {
            const v = String(matrix[r][col] || '').trim();
            const m = v.match(/\b(20\d{2})\b/);
            if (m) { yf = parseInt(m[1]); break; }
        }
        if (yf) cy = yf;
        colYears[col] = cy;
    }
    const artIdx = headers.findIndex(h => h.toLowerCase().includes('articulo') || h.toLowerCase().includes('artículo'));
    const marcaIdx = headers.findIndex(h => h.toUpperCase() === 'MARCA');
    const cat3Idx = headers.findIndex(h => h.toUpperCase() === 'CAT 3');
    const tipoIdx = headers.findIndex(h => h.toUpperCase().includes('TIPO_NEGOCIO_1') || h.toUpperCase() === 'TIPO_NEGOCIO');
    const mc = [];
    headers.forEach((h, idx) => { if (isMonthColumn(h)) mc.push({ idx, ...extractMonthYear(h, colYears[idx]) }); });

    const result = [];
    for (let i = hIdx + 1; i < matrix.length; i++) {
        const row = matrix[i];
        const art = row[artIdx >= 0 ? artIdx : 1];
        if (!art || String(art).toLowerCase().includes('total')) continue;
        const brand = marcaIdx >= 0 && row[marcaIdx] ? String(row[marcaIdx]).trim().toUpperCase() : guessBrand(String(art));
        const cat = cat3Idx >= 0 && row[cat3Idx] ? String(row[cat3Idx]).trim() : 'Sin Categoría';
        const tipo = tipoIdx >= 0 && row[tipoIdx] ? String(row[tipoIdx]).trim() : '';
        const isMP = tipo.toLowerCase().includes('propia') || isMarcaPropia(String(art));
        mc.forEach(col => {
            const us = row[col.idx];
            const units = (us === '' || us === undefined || us === null) ? 0 : normalizeNumber(us);
            result.push({ article: String(art).trim(), category: cat, brand, tipo, year: col.year, month: col.month, monthName: MESES[col.month], units, isMP });
        });
    }
    return result;
}

// ============================================================
// ALE PARSER
// ============================================================
function parseAleSheet(matrix) {
    if (matrix.length < 2) return [];
    let hIdx = 0;
    for (let i = 0; i < Math.min(15, matrix.length); i++) {
        if (matrix[i].some(c => typeof c === 'string' && (c.toLowerCase().includes('articulo') || c.toLowerCase().includes('artículo')))) {
            hIdx = i; break;
        }
    }
    const headers = matrix[hIdx].map(h => String(h).trim());
    const colYears = [];
    let cy = new Date().getFullYear();
    for (let col = 0; col < headers.length; col++) {
        let yf = null;
        for (let r = 0; r < hIdx; r++) {
            const v = String(matrix[r][col] || '').trim();
            const m = v.match(/\b(20\d{2})\b/);
            if (m) { yf = parseInt(m[1]); break; }
        }
        if (yf) cy = yf;
        colYears[col] = cy;
    }

    const artIdx = headers.findIndex(h => h.toLowerCase().includes('articulo') || h.toLowerCase().includes('artículo'));
    const marcaRawIdx = headers.findIndex(h => h.toUpperCase() === 'MARCA');
    const tipoRawIdx = headers.findIndex(h => h.toUpperCase() === 'TIPO_NEGOCIO_2');
    const cat3Idx = headers.findIndex(h => h.toUpperCase() === 'CAT 3');

    const businessTypes = new Set(['marca propia', 'franquicia', 'marca propia adoptada', 'representaciones']);
    let marcaIdx = marcaRawIdx;
    let tipoIdx = tipoRawIdx;

    // ✅ FIX: muestreo robusto sobre varias filas (no solo la primera)
    if (marcaRawIdx >= 0 && tipoRawIdx >= 0) {
        let marcaHits = 0, tipoHits = 0;
        const end = Math.min(matrix.length, hIdx + 30);
        for (let i = hIdx + 1; i < end; i++) {
            const r = matrix[i];
            if (!r) continue;
            const mv = String(r[marcaRawIdx] || '').trim().toLowerCase();
            const tv = String(r[tipoRawIdx] || '').trim().toLowerCase();
            if (businessTypes.has(mv)) marcaHits++;
            if (businessTypes.has(tv)) tipoHits++;
        }
        if (marcaHits > tipoHits) {
            marcaIdx = tipoRawIdx;
            tipoIdx = marcaRawIdx;
        }
    }

    const mc = [];
    headers.forEach((h, idx) => { if (isMonthColumn(h)) mc.push({ idx, ...extractMonthYear(h, colYears[idx]) }); });

    const result = [];
    for (let i = hIdx + 1; i < matrix.length; i++) {
        const row = matrix[i];
        const art = row[artIdx >= 0 ? artIdx : 1];
        if (!art || String(art).toLowerCase().includes('total')) continue;
        const brand = marcaIdx >= 0 && row[marcaIdx] ? String(row[marcaIdx]).trim().toUpperCase() : 'SIN MARCA';
        const tipo = tipoIdx >= 0 && row[tipoIdx] ? String(row[tipoIdx]).trim() : '';
        const cat = cat3Idx >= 0 && row[cat3Idx] ? String(row[cat3Idx]).trim() : 'Sin Categoría';
        mc.forEach(col => {
            const us = row[col.idx];
            const units = (us === '' || us === undefined || us === null) ? 0 : normalizeNumber(us);
            result.push({ article: String(art).trim(), category: cat, brand, tipo, year: col.year, month: col.month, monthName: MESES[col.month], units });
        });
    }
    return result;
}

// ============================================================
// INIT ALL DASHBOARDS
// ============================================================
function initAllDashboards() {
    const uploadView = document.getElementById('upload-view');
    const dashboardLayout = document.getElementById('dashboard-layout');

    try {
        // Init Nati dashboard
        globalData = globalDataNati;
        populateFilters();
        updateDashboard();
        // ✅ FIX: ya NO se agrega aquí el listener de #chart-mode (lo hace populateFilters con .onchange)

        // Init Ale dashboard filters (una vez)
        populateAleFilters(true);

        uploadView.classList.add('hidden');
        dashboardLayout.classList.remove('hidden');
        dashboardLayout.classList.add('flex');
        dashboardLayout.classList.add('flex-col');

        switchTeam('nati');

        if (typeof lucide !== 'undefined') { lucide.createIcons(); }
    } catch (err) {
        console.error('Error al inicializar el dashboard:', err);
        const statusText = document.getElementById('status-text');
        const loadingBar = document.querySelector('.loading-bar');
        if (loadingBar) loadingBar.style.animationPlayState = 'paused';
        if (statusText) {
            statusText.innerHTML = `<span class="text-red-500 font-semibold">Error al construir el dashboard:</span><br>${(err && err.message) ? err.message : err}<br><span class="text-[0.65rem]">Revisa la consola (F12) para el detalle técnico.</span>`;
        }
    }
}

// ============================================================
// TEAM SWITCHING
// ============================================================
function switchTeam(team) {
    activeTeam = team;
    const aleView = document.getElementById('team-ale-view');
    const natiView = document.getElementById('team-nati-view');
    const btnAle = document.getElementById('btn-team-ale');
    const btnNati = document.getElementById('btn-team-nati');

    if (team === 'ale') {
        aleView.classList.remove('hidden');
        aleView.classList.add('flex');
        natiView.classList.add('hidden');
        natiView.classList.remove('flex');
        btnAle.classList.add('active');
        btnNati.classList.remove('active');
        showAleAnalysisSelector();
    } else {
        natiView.classList.remove('hidden');
        natiView.classList.add('flex');
        aleView.classList.add('hidden');
        aleView.classList.remove('flex');
        btnNati.classList.add('active');
        btnAle.classList.remove('active');
    }
    lucide.createIcons();
}

// ============================================================
// ALE MODE (Unidades / BS)
// ============================================================
function setAleMode(mode) {
    aleMode = mode;
    populateAleFilters(true); // ✅ FIX: forzar repoblado solo cuando cambia el dataset
    document.getElementById('btn-ale-units').classList.toggle('ale-mode-active', mode === 'units');
    document.getElementById('btn-ale-bs').classList.toggle('ale-mode-active', mode === 'bs');
    updateAleDashboard();
}

function getAleData() {
    return aleMode === 'units' ? globalDataAleUnits : globalDataAleBS;
}

function getAleFilteredData() {
    const data = getAleData();
    // ✅ OPT: lowercasing del SKU una sola vez
    const skuLower = aleFilters.sku ? aleFilters.sku.toLowerCase().trim() : '';
    return data.filter(d => {
        if (aleFilters.year !== 'all' && String(d.year) !== aleFilters.year) return false;
        if (aleFilters.month !== 'all' && d.monthName !== aleFilters.month) return false;
        if (aleFilters.brand !== 'all' && d.brand !== aleFilters.brand) return false;
        if (aleFilters.category !== 'all' && d.category !== aleFilters.category) return false;
        if (aleFilters.tipo !== 'all' && d.tipo !== aleFilters.tipo) return false;
        if (skuLower && !d.article.toLowerCase().includes(skuLower)) return false;
        return true;
    });
}

// ✅ FIX/OPT: memoizado por modo (evita reconstruir los <select> en cada cambio)
let _aleFiltersPopulatedFor = null;
function populateAleFilters(force = false) {
    if (!force && _aleFiltersPopulatedFor === aleMode) return;
    _aleFiltersPopulatedFor = aleMode;

    const allData = [...globalDataAleUnits, ...globalDataAleBS];
    const years = [...new Set(allData.map(d => d.year))].sort((a,b)=>a-b);
    const brands = [...new Set(allData.map(d => d.brand))].filter(Boolean).sort();
    const cats = [...new Set(allData.map(d => d.category))].filter(Boolean).sort();
    const tipos = [...new Set(allData.map(d => d.tipo))].filter(Boolean).sort();
    const months = MESES.filter(m => allData.some(d => d.monthName === m));

    const yEl = document.getElementById('ale-filter-year');
    if (yEl) { yEl.innerHTML = '<option value="all">Todos</option>'; years.forEach(y => yEl.innerHTML += `<option value="${y}">${y}</option>`); }

    const mEl = document.getElementById('ale-filter-month');
    if (mEl) { mEl.innerHTML = '<option value="all">Todos</option>'; months.forEach(m => mEl.innerHTML += `<option value="${m}">${m}</option>`); }

    const bEl = document.getElementById('ale-filter-brand');
    if (bEl) { bEl.innerHTML = '<option value="all">Todas</option>'; brands.forEach(b => bEl.innerHTML += `<option value="${b}">${b}</option>`); }

    const cEl = document.getElementById('ale-filter-category');
    if (cEl) { cEl.innerHTML = '<option value="all">Todas</option>'; cats.forEach(c => cEl.innerHTML += `<option value="${c}">${c}</option>`); }

    const tEl = document.getElementById('ale-filter-tipo');
    if (tEl) { tEl.innerHTML = '<option value="all">Todos</option>'; tipos.forEach(t => tEl.innerHTML += `<option value="${t}">${t}</option>`); }
}

function resetAleFilters() {
    aleFilters = { year: 'all', month: 'all', brand: 'all', category: 'all', tipo: 'all', sku: '' };
    ['ale-filter-year','ale-filter-month','ale-filter-brand','ale-filter-category','ale-filter-tipo'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = 'all';
    });
    const skuEl = document.getElementById('ale-filter-sku');
    if (skuEl) skuEl.value = '';
    updateAleContextLabel('all', 'all');
    updateAleDashboard();
}

// ============================================================
// ALE — ANALYSIS SELECTOR
// ============================================================
function getAleSelectorOptions() {
    const allData = [...globalDataAleUnits, ...globalDataAleBS];
    const categories = [...new Set(allData.map(d => d.category))].filter(Boolean).sort();
    const brands = [...new Set(allData.map(d => d.brand))].filter(Boolean).sort();
    return { categories, brands };
}

function showAleAnalysisSelector() {
    const selector = document.getElementById('ale-analysis-selector');
    const dashboard = document.getElementById('ale-dashboard-content');
    if (selector) { selector.classList.remove('hidden'); selector.classList.add('flex'); }
    if (dashboard) { dashboard.classList.add('hidden'); dashboard.classList.remove('flex'); }

    const searchEl = document.getElementById('ale-selector-search');
    if (searchEl) searchEl.value = '';
    const dropdown = document.getElementById('ale-selector-dropdown');
    if (dropdown) { dropdown.classList.add('hidden'); dropdown.innerHTML = ''; }

    if (window.lucide) lucide.createIcons();
}

function renderAleSelectorDropdown(query) {
    const dropdown = document.getElementById('ale-selector-dropdown');
    if (!dropdown) return;
    const q = String(query || '').trim().toLowerCase();

    if (!q) {
        dropdown.classList.add('hidden');
        dropdown.innerHTML = '';
        return;
    }

    const { categories, brands } = getAleSelectorOptions();
    const filteredCats = categories.filter(c => c.toLowerCase().includes(q));
    const filteredBrands = brands.filter(b => b.toLowerCase().includes(q));

    if (filteredCats.length === 0 && filteredBrands.length === 0) {
        dropdown.innerHTML = `<div class="ale-selector-empty">Sin resultados para "${escapeAleHtml(query)}"</div>`;
        dropdown.classList.remove('hidden');
        return;
    }

    let html = '';
    if (filteredCats.length > 0) {
        html += `<div class="ale-selector-group-label">Categorías</div>`;
        filteredCats.forEach(c => {
            html += `<div class="ale-selector-option" onclick="openAleAnalysis('category', '${escapeAleAttr(c)}')">
                <i data-lucide="layers" class="w-3.5 h-3.5"></i><span>${escapeAleHtml(c)}</span>
            </div>`;
        });
    }
    if (filteredBrands.length > 0) {
        html += `<div class="ale-selector-group-label">Marcas</div>`;
        filteredBrands.forEach(b => {
            html += `<div class="ale-selector-option" onclick="openAleAnalysis('brand', '${escapeAleAttr(b)}')">
                <i data-lucide="tag" class="w-3.5 h-3.5"></i><span>${escapeAleHtml(b)}</span>
            </div>`;
        });
    }
    dropdown.innerHTML = html;
    dropdown.classList.remove('hidden');
    if (window.lucide) lucide.createIcons();
}

function openAleAnalysis(type, value) {
    if (type === 'category') {
        aleFilters.category = value;
        aleFilters.brand = 'all';
    } else if (type === 'brand') {
        aleFilters.brand = value;
        aleFilters.category = 'all';
    } else {
        aleFilters.category = 'all';
        aleFilters.brand = 'all';
    }

    const selector = document.getElementById('ale-analysis-selector');
    const dashboard = document.getElementById('ale-dashboard-content');
    if (selector) { selector.classList.add('hidden'); selector.classList.remove('flex'); }
    if (dashboard) { dashboard.classList.remove('hidden'); dashboard.classList.add('flex'); }

    populateAleFilters(true);
    ['year', 'month', 'brand', 'category', 'tipo'].forEach(f => {
        const el = document.getElementById(`ale-filter-${f}`);
        if (el) el.value = aleFilters[f];
    });

    updateAleContextLabel(type, value);
    updateAleDashboard();
    if (window.lucide) lucide.createIcons();
}

function updateAleContextLabel(type, value) {
    const label = document.getElementById('ale-context-label');
    if (!label) return;
    if (type === 'category') label.textContent = `📁 Categoría: ${value}`;
    else if (type === 'brand') label.textContent = `🏷️ Marca: ${value}`;
    else label.textContent = '📊 Dashboard General';
}

function escapeAleHtml(str) {
    return String(str).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

function escapeAleAttr(str) {
    return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

// ============================================================
// ALE DASHBOARD — MAIN UPDATE
// ============================================================
function updateAleDashboard() {
    // Lee valores actuales de los filtros
    ['year','month','brand','category','tipo'].forEach(f => {
        const el = document.getElementById(`ale-filter-${f}`);
        if (el) aleFilters[f] = el.value;
    });
    const skuEl = document.getElementById('ale-filter-sku');
    if (skuEl) aleFilters.sku = skuEl.value;

    // ✅ FIX: ya NO se repueblan los filtros aquí (memoizado por modo)

    const data = getAleFilteredData();
    if (data.length === 0) return;

    renderAleKPIs(data);
    renderAleTrendChart(data);
    renderAlePeriodComparison(data);
    renderAleBrandRanking(data);
    renderAleFullRanking(data);
    renderAleTopBottom(data);
    renderAleAlerts(data);
    renderAleMatrix(data);
    renderAleExecutiveSummary();
    lucide.createIcons();
}

// ============================================================
// ALE — HELPER: get brand stats
// ============================================================
function getAleBrandStats(data) {
    const allData = getAleData();
    const years = [...new Set(allData.map(d => d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length - 1];
    const prevYear = years.length > 1 ? years[years.length - 2] : null;

    const lastYearActive = allData.filter(d => d.year === lastYear && d.units > 0);
    const maxMonth = lastYearActive.length > 0 ? Math.max(...lastYearActive.map(d => d.month)) : 11;
    const curMonth = maxMonth;

    const brandFiltered = aleFilters.brand !== 'all';
    const groupKey = brandFiltered ? 'article' : 'brand';
    const groupLabel = brandFiltered ? 'sku' : 'brand';

    const entityMap = {};
    data.forEach(d => {
        const key = d[groupKey];
        if (!key) return;
        if (!entityMap[key]) {
            entityMap[key] = {
                brand: d.brand,
                name: key,
                label: key,
                tipo: d.tipo,
                category: d.category,
                current: 0,
                previous: 0,
                prevMonthVal: 0,
                curMonthVal: 0,
                monthlyHistory: {}
            };
        }
        const bm = entityMap[key];
        if (d.year === lastYear && d.month <= curMonth) bm.current += d.units;
        if (prevYear && d.year === prevYear && d.month <= curMonth) bm.previous += d.units;
        const hKey = `${d.year}-${d.month}`;
        bm.monthlyHistory[hKey] = (bm.monthlyHistory[hKey] || 0) + d.units;
        if (d.year === lastYear && d.month === curMonth) bm.curMonthVal += d.units;
        if (d.year === lastYear && curMonth > 0 && d.month === curMonth - 1) bm.prevMonthVal += d.units;
    });

    const allBrands = Object.values(entityMap).filter(b => b.current > 0 || b.previous > 0);
    // ✅ OPT: fusionado (antes eran 3 loops sobre allBrands)
    let total = 0, totalPrev = 0;
    allBrands.forEach(b => { total += b.current; totalPrev += b.previous; });
    allBrands.forEach(b => {
        b.share = total > 0 ? (b.current / total) * 100 : 0;
        b.prevShare = totalPrev > 0 ? (b.previous / totalPrev) * 100 : 0;
        b.shareDelta = b.share - b.prevShare;
        b.varPct = b.previous > 0 ? ((b.current / b.previous) - 1) * 100 : (b.current > 0 ? null : null);
        b.varAbs = b.current - b.previous;
        b.total = b.current;
    });

    return { entityMap, allBrands, total, lastYear, prevYear, curMonth, maxMonth, brandFiltered, groupLabel };
}

// ============================================================
// ALE KPIs (6 cards)
// ============================================================
function renderAleKPIs(data) {
    const { allBrands, lastYear, prevYear } = getAleBrandStats(data);
    const isBs = aleMode === 'bs';
    const unit = isBs ? 'Bs' : 'un.';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES', {minimumFractionDigits: 0, maximumFractionDigits: 0})}` : v.toLocaleString('es-ES');

    const allData = getAleData();
    const lastYearActive = allData.filter(d => d.year === lastYear && d.units > 0);
    const maxMonth = lastYearActive.length > 0 ? Math.max(...lastYearActive.map(d => d.month)) : 0;

    // ✅ OPT: una sola pasada
    let totalVal = 0, curMonthTotal = 0, prevMonthTotal = 0;
    const activeMonthsSet = new Set();
    data.forEach(d => {
        if (d.year !== lastYear) return;
        totalVal += d.units;
        if (d.units > 0) activeMonthsSet.add(d.month);
        if (d.month === maxMonth) curMonthTotal += d.units;
        else if (maxMonth > 0 && d.month === maxMonth - 1) prevMonthTotal += d.units;
    });
    const varMes = prevMonthTotal > 0 ? ((curMonthTotal / prevMonthTotal) - 1) * 100 : null;
    const monthsWithData = activeMonthsSet.size;
    const monthlyAvg = monthsWithData > 0 ? Math.round(totalVal / monthsWithData) : 0;

    const totalUniverse = allData.filter(d => d.year === lastYear).reduce((s, d) => s + d.units, 0);
    const participation = totalUniverse > 0 ? (totalVal / totalUniverse) * 100 : 100;

    const growing = allBrands.filter(b => b.varPct !== null && b.varPct > 0).length;
    const declining = allBrands.filter(b => b.varPct !== null && b.varPct < 0).length;

    const curMonthName = MESES[maxMonth] || '';

    const kpis = [
        {
            title: `Venta Total ${lastYear}`,
            value: fmt(totalVal),
            sub: `${unit} acumulados`,
            icon: 'bar-chart-2',
            borderColor: '#0EA5E9',
            bg: 'linear-gradient(135deg, #FFFFFF, #F0F9FF)',
            textColor: '#0C4A6E'
        },
        {
            title: `Var. vs ${maxMonth > 0 ? MESES[maxMonth-1] : 'mes ant.'}`,
            value: varMes !== null ? `${varMes >= 0 ? '+' : ''}${varMes.toFixed(1)}%` : 'N/D',
            sub: `${curMonthName}: ${fmt(curMonthTotal)}`,
            icon: varMes === null || varMes >= 0 ? 'trending-up' : 'trending-down',
            borderColor: varMes === null ? '#E2E8F0' : varMes >= 0 ? '#10B981' : '#EF4444',
            bg: varMes === null ? 'linear-gradient(135deg, #FFFFFF, #F8FAFC)' : varMes >= 0 ? 'linear-gradient(135deg, #FFFFFF, #ECFDF5)' : 'linear-gradient(135deg, #FFFFFF, #FEF2F2)',
            textColor: varMes === null ? '#64748B' : varMes >= 0 ? '#065F46' : '#991B1B'
        },
        {
            title: 'Promedio Mensual',
            value: fmt(monthlyAvg),
            sub: `${monthsWithData} meses con datos`,
            icon: 'activity',
            borderColor: '#8B5CF6',
            bg: 'linear-gradient(135deg, #FFFFFF, #F5F3FF)',
            textColor: '#4C1D95'
        },
        {
            title: 'Participación',
            value: `${participation.toFixed(1)}%`,
            sub: aleFilters.brand !== 'all' || aleFilters.category !== 'all' ? 'del total seleccionado' : 'del portafolio total',
            icon: 'pie-chart',
            borderColor: '#F59E0B',
            bg: 'linear-gradient(135deg, #FFFFFF, #FFFBEB)',
            textColor: '#78350F'
        },
        {
            title: 'Marcas Creciendo',
            value: String(growing),
            sub: prevYear ? `vs ${prevYear}` : 'sin período anterior',
            icon: 'trending-up',
            borderColor: '#10B981',
            bg: 'linear-gradient(135deg, #FFFFFF, #ECFDF5)',
            textColor: '#065F46'
        },
        {
            title: 'Marcas en Caída',
            value: String(declining),
            sub: prevYear ? `vs ${prevYear}` : 'sin período anterior',
            icon: 'trending-down',
            borderColor: '#EF4444',
            bg: 'linear-gradient(135deg, #FFFFFF, #FEF2F2)',
            textColor: '#991B1B'
        }
    ];

    document.getElementById('ale-kpi-container').innerHTML = kpis.map(k => `
        <div class="ale-kpi-card" style="background: ${k.bg}; border-color: ${k.borderColor};">
            <div class="kpi-icon"><i data-lucide="${k.icon}" style="width:48px;height:48px;color:${k.borderColor};"></i></div>
            <p class="text-[0.6rem] font-black uppercase tracking-widest mb-2" style="color:${k.textColor};opacity:0.7">${k.title}</p>
            <p class="text-2xl font-black tracking-tight" style="color:${k.textColor}">${k.value}</p>
            <p class="text-[0.65rem] font-semibold mt-1 opacity-60" style="color:${k.textColor}">${k.sub}</p>
        </div>
    `).join('');
}

// ============================================================
// ALE YoY CHART
// ============================================================
function renderAleTrendChart(data) {
    const allData = getAleData();
    const years = [...new Set(allData.map(d => d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length - 1];
    const prevYear = years.length > 1 ? years[years.length - 2] : null;
    const isBs = aleMode === 'bs';
    const fmtTick = v => {
        if (v >= 1000000) return (v/1000000).toFixed(1) + 'M';
        if (v >= 1000) return (v/1000).toFixed(0) + 'K';
        return v;
    };
    const fmtFull = v => isBs
        ? `Bs ${v.toLocaleString('es-ES', {minimumFractionDigits:0, maximumFractionDigits:0})}`
        : v.toLocaleString('es-ES');

    // ✅ OPT: una sola pasada (antes eran 24 filtrados)
    const curYearData = new Array(12).fill(0);
    const prevYearData = prevYear ? new Array(12).fill(0) : [];
    data.forEach(d => {
        if (d.year === lastYear) curYearData[d.month] += d.units;
        else if (prevYear && d.year === prevYear) prevYearData[d.month] += d.units;
    });

    const lastMonthWithData = curYearData.reduce((last, v, i) => v > 0 ? i : last, -1);
    const shortMeses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

    const datasets = [];

    if (prevYear) {
        datasets.push({
            label: String(prevYear),
            data: prevYearData,
            backgroundColor: 'rgba(148, 163, 184, 0.45)',
            hoverBackgroundColor: 'rgba(148, 163, 184, 0.7)',
            borderRadius: 6,
            borderSkipped: false,
            barPercentage: 0.75,
            categoryPercentage: 0.8,
            order: 2
        });
    }

    datasets.push({
        label: String(lastYear),
        data: curYearData,
        backgroundColor: curYearData.map((v, i) => {
            if (v === 0 && i > lastMonthWithData) return 'rgba(14, 165, 233, 0.08)';
            return '#0EA5E9';
        }),
        hoverBackgroundColor: '#0284C7',
        borderRadius: 6,
        borderSkipped: false,
        barPercentage: 0.75,
        categoryPercentage: 0.8,
        order: 1
    });

    const ctx = document.getElementById('ale-trend-chart');
    if (!ctx) return;
    if (aleCharts['trend']) aleCharts['trend'].destroy();

    aleCharts['trend'] = new Chart(ctx, {
        type: 'bar',
        plugins: [ChartDataLabels],
        data: {
            labels: shortMeses,
            datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: { duration: 700, easing: 'easeOutQuart' },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    align: 'end',
                    labels: {
                        font: { family: 'Nunito', size: 11, weight: '700' },
                        color: '#475569',
                        usePointStyle: true,
                        pointStyle: 'circle',
                        padding: 16,
                        boxWidth: 8
                    }
                },
                datalabels: {
                    display: (ctx) => ctx.dataset.data[ctx.dataIndex] > 0,
                    anchor: 'end',
                    align: 'end',
                    color: (ctx) => {
                        const isCurrentYear = prevYear
                            ? ctx.datasetIndex === datasets.length - 1
                            : ctx.datasetIndex === 0;
                        return isCurrentYear ? '#0369A1' : '#64748B';
                    },
                    font: (ctx) => {
                        const isCurrentYear = prevYear
                            ? ctx.datasetIndex === datasets.length - 1
                            : ctx.datasetIndex === 0;
                        return { weight: isCurrentYear ? '800' : '700', size: 9, family: 'Nunito' };
                    },
                    formatter: val => fmtTick(val)
                },
                tooltip: {
                    backgroundColor: '#FFFFFF',
                    borderColor: '#E2E8F0',
                    borderWidth: 1,
                    titleColor: '#0F172A',
                    bodyColor: '#475569',
                    cornerRadius: 12,
                    padding: 12,
                    titleFont: { family: 'Nunito', size: 12, weight: 'bold' },
                    bodyFont: { family: 'Nunito', size: 11 },
                    callbacks: {
                        title: items => shortMeses[items[0].dataIndex],
                        label: ctx => {
                            const val = ctx.raw;
                            return ` ${ctx.dataset.label}: ${fmtFull(val)}`;
                        },
                        afterBody: items => {
                            if (!prevYear || items.length < 2) return [];
                            const idx = items[0].dataIndex;
                            const cur = curYearData[idx];
                            const prev = prevYearData[idx];
                            if (prev === 0 || cur === 0) return [];
                            const varPct = ((cur / prev) - 1) * 100;
                            const arrow = varPct >= 0 ? '▲' : '▼';
                            return [``, ` YoY: ${arrow} ${varPct >= 0 ? '+' : ''}${varPct.toFixed(1)}%`];
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: '#64748B',
                        font: { family: 'Nunito', size: 10, weight: '700' },
                        maxRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    ticks: {
                        color: '#64748B',
                        font: { family: 'Nunito', size: 10 },
                        callback: v => (isBs ? 'Bs ' : '') + fmtTick(v)
                    },
                    grid: { color: 'rgba(0,0,0,0.04)', drawBorder: false },
                    beginAtZero: true
                }
            }
        }
    });

    const sub = document.getElementById('ale-chart-subtitle');
    if (sub) {
        sub.textContent = prevYear
            ? `${prevYear} vs ${lastYear} · ${isBs ? 'Bolivianos' : 'Unidades'} por mes`
            : `${lastYear} · ${isBs ? 'Bolivianos vendidos' : 'Unidades vendidas'} por mes`;
    }
}

// ============================================================
// ALE PERIOD COMPARISON
// ============================================================
function renderAlePeriodComparison(data) {
    const allData = getAleData();
    const years = [...new Set(allData.map(d => d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length - 1];
    const prevYear = years.length > 1 ? years[years.length - 2] : null;
    const isBs = aleMode === 'bs';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES', {minimumFractionDigits:0,maximumFractionDigits:0})}` : `${v.toLocaleString('es-ES')} un.`;

    const lastYearActive = allData.filter(d => d.year === lastYear && d.units > 0);
    const maxMonth = lastYearActive.length > 0 ? Math.max(...lastYearActive.map(d => d.month)) : 0;

    // ✅ OPT: una sola pasada
    let curMonthVal = 0, prevMonthVal = 0, sameMonthLastYear = 0, ytdCur = 0, ytdPrev = 0;
    data.forEach(d => {
        if (d.year === lastYear) {
            if (d.month <= maxMonth) ytdCur += d.units;
            if (d.month === maxMonth) curMonthVal += d.units;
            else if (maxMonth > 0 && d.month === maxMonth - 1) prevMonthVal += d.units;
        } else if (prevYear && d.year === prevYear) {
            if (d.month <= maxMonth) ytdPrev += d.units;
            if (d.month === maxMonth) sameMonthLastYear += d.units;
        }
    });

    function compCard(title, cur, prev, periodLabel) {
        const varAbs = cur - prev;
        const varPct = prev > 0 ? ((cur / prev) - 1) * 100 : null;
        const isPos = varAbs >= 0;
        const color = varPct === null ? '#64748B' : isPos ? '#10B981' : '#EF4444';
        const bg = varPct === null ? '#F8FAFC' : isPos ? '#ECFDF5' : '#FEF2F2';
        const border = varPct === null ? '#E2E8F0' : isPos ? '#A7F3D0' : '#FECACA';
        return `<div class="ale-period-card" style="border-color:${border};background:${bg};">
            <p class="text-[0.6rem] font-black uppercase tracking-widest text-slate-500 mb-1">${title}</p>
            <p class="text-xl font-black text-slate-800">${fmt(cur)}</p>
            <p class="text-[0.65rem] text-slate-500 mt-1 font-semibold">${periodLabel}: <span class="font-black text-slate-700">${fmt(prev)}</span></p>
            <div class="mt-2 flex items-center gap-2">
                <span class="text-sm font-black" style="color:${color}">${varPct !== null ? (isPos ? '+' : '') + varPct.toFixed(1) + '%' : 'N/D'}</span>
                <span class="text-xs font-bold" style="color:${color}">(${isPos ? '+' : ''}${fmt(varAbs)})</span>
            </div>
        </div>`;
    }

    const el = document.getElementById('ale-period-comparison');
    if (!el) return;
    el.innerHTML = [
        compCard(`${MESES[maxMonth]} vs ${maxMonth > 0 ? MESES[maxMonth-1] : 'Mes anterior'}`, curMonthVal, prevMonthVal, maxMonth > 0 ? MESES[maxMonth-1] : 'Mes ant.'),
        prevYear ? compCard(`${MESES[maxMonth]} vs mismo mes ${prevYear}`, curMonthVal, sameMonthLastYear, `${MESES[maxMonth]} ${prevYear}`) : '',
        prevYear ? compCard(`YTD ${lastYear} vs ${prevYear}`, ytdCur, ytdPrev, `YTD ${prevYear}`) : ''
    ].join('');
}

// ============================================================
// ALE BRAND RANKING TABLE
// ============================================================
function renderAleBrandRanking(data) {
    const { allBrands, brandFiltered } = getAleBrandStats(data);
    const isBs = aleMode === 'bs';
    const sortEl = document.getElementById('ale-brand-sort');
    const sortBy = sortEl ? sortEl.value : 'total';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES',{minimumFractionDigits:0,maximumFractionDigits:0})}` : v.toLocaleString('es-ES');

    let sorted = [...allBrands];
    if (sortBy === 'total') sorted.sort((a, b) => b.current - a.current);
    else if (sortBy === 'growth') sorted = sorted.filter(b => b.varPct !== null && b.varPct > 0).sort((a, b) => b.varPct - a.varPct);
    else if (sortBy === 'decline') sorted = sorted.filter(b => b.varPct !== null && b.varPct < 0).sort((a, b) => a.varPct - b.varPct);
    else if (sortBy === 'share') sorted.sort((a, b) => b.share - a.share);
    else if (sortBy === 'share-loss') sorted = sorted.filter(b => b.shareDelta < 0).sort((a, b) => a.shareDelta - b.shareDelta);

    const body = document.getElementById('ale-brand-ranking-body');
    if (!body) return;

    const tableEl = body.closest('table');
    if (tableEl) {
        const ths = tableEl.querySelectorAll('thead th');
        if (ths.length >= 3) {
            ths[1].textContent = brandFiltered ? 'SKU / Producto' : 'Marca';
            ths[2].textContent = brandFiltered ? 'Categoría' : 'Tipo';
        }
    }

    body.innerHTML = sorted.slice(0, 30).map((b, i) => {
        const varPct = b.varPct;
        const varBg = varPct === null ? 'bg-slate-100 text-slate-600' : varPct >= 0 ? 'bg-emerald-50 text-emerald-700' : 'bg-red-50 text-red-700';
        const shareBar = Math.min(100, (b.share / (sorted[0]?.share || 1)) * 100);
        const shareDeltaStr = b.shareDelta !== undefined && b.shareDelta !== 0
            ? (b.shareDelta >= 0 ? `+${b.shareDelta.toFixed(1)}` : b.shareDelta.toFixed(1)) + ' pp'
            : '';
        const shareDeltaColor = b.shareDelta >= 0 ? 'text-emerald-600' : 'text-red-600';
        const displayName = b.label || b.brand;
        const secondaryInfo = brandFiltered ? (b.category || '—') : (b.tipo || '—');
        const varDisplay = varPct !== null ? (varPct >= 0 ? '+' : '') + varPct.toFixed(1) + '%' : '—';
        return `<tr>
            <td class="py-2.5 px-2 font-black text-center w-8 text-slate-400">${i+1}</td>
            <td class="py-2.5 px-2 font-bold text-slate-800 max-w-[200px]">
                <span class="block truncate" title="${displayName}">${displayName}</span>
            </td>
            <td class="py-2.5 px-2">
                <span class="text-[0.65rem] font-bold bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded-md">${secondaryInfo}</span>
            </td>
            <td class="py-2.5 px-2 text-right font-black text-slate-800 tabular-nums">${fmt(b.current)}</td>
            <td class="py-2.5 px-2 text-right text-slate-500 tabular-nums">${b.previous > 0 ? fmt(b.previous) : '—'}</td>
            <td class="py-2.5 px-2 text-right">
                <span class="inline-block text-xs font-black px-2 py-0.5 rounded-lg ${varBg}">${varDisplay}</span>
            </td>
            <td class="py-2.5 px-2 text-right">
                <div class="flex items-center justify-end gap-1.5">
                    <div class="w-16 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                        <div class="h-full bg-sky-500 rounded-full" style="width:${shareBar}%"></div>
                    </div>
                    <span class="text-xs font-bold text-slate-700 tabular-nums w-10 text-right">${b.share.toFixed(1)}%</span>
                    ${shareDeltaStr ? `<span class="text-[0.65rem] font-bold ${shareDeltaColor}">${shareDeltaStr}</span>` : ''}
                </div>
            </td>
            <td class="py-2.5 px-2 text-center">
                ${varPct === null ? '<span class="text-slate-400 text-xs">—</span>' : varPct >= 0 ? '<span class="text-emerald-500">▲</span>' : '<span class="text-red-500">▼</span>'}
            </td>
        </tr>`;
    }).join('');
}


// ============================================================
// ALE FULL PORTFOLIO RANKING TABLE
// ============================================================
let aleFullRankingCache = [];
function renderAleFullRanking(data) {
    const isBs = aleMode === 'bs';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES',{minimumFractionDigits:0,maximumFractionDigits:0})}` : v.toLocaleString('es-ES');

    const years = [...new Set(data.map(d => d.year))].sort((a,b)=>a-b);
    const lastYear = years[years.length - 1];
    const prevYear = years.length > 1 ? years[years.length - 2] : null;
    const isAllYears = aleFilters.year === 'all';

    const lastYearActive = data.filter(d => d.year === lastYear && d.units > 0);
    const maxMonth = lastYearActive.length > 0 ? Math.max(...lastYearActive.map(d => d.month)) : 11;
    const monthsElapsed = maxMonth >= 0 ? maxMonth + 1 : 0;

    const totalLabel = 'Total Histórico';
    const varLabel = prevYear ? `Var. YTD ${lastYear} vs ${prevYear}` : 'Var. YTD';
    const thTotal = document.getElementById('ale-th-full-total');
    if (thTotal) thTotal.textContent = totalLabel;
    const thVar = document.getElementById('ale-th-full-var');
    if (thVar) thVar.textContent = varLabel;

    const pm = {};
    data.forEach(d => {
        if (!pm[d.article]) pm[d.article] = { article: d.article, category: d.category, total: 0, ytdLast: 0, ytdPrev: 0 };
        const p = pm[d.article];
        if (isAllYears || String(d.year) === aleFilters.year) p.total += d.units;
        if (d.month <= maxMonth) {
            if (d.year === lastYear) p.ytdLast += d.units;
            if (prevYear && d.year === prevYear) p.ytdPrev += d.units;
        }
    });
    const products = Object.values(pm).filter(p => p.total > 0 || p.ytdLast > 0 || p.ytdPrev > 0);
    const tp = products.reduce((s, p) => s + p.total, 0);
    products.forEach(p => {
        p.share = tp > 0 ? (p.total / tp) * 100 : 0;
        if (p.ytdPrev === 0 && p.ytdLast > 0) p.ytdVar = null;
        else if (p.ytdPrev === 0) p.ytdVar = -100;
        else p.ytdVar = ((p.ytdLast / p.ytdPrev) - 1) * 100;
    });
    const fullSorted = [...products].sort((a, b) => b.total - a.total);
    let accShare = 0;
    fullSorted.forEach(p => { accShare += p.share; p.inPareto = accShare <= 80; });
    aleFullRankingCache = fullSorted;

    const renderFull = (list) => {
        const body = document.getElementById('ale-full-ranking-body');
        if (!body) return;
        body.innerHTML = list.map((p, idx) => {
            const isN = p.ytdVar === null;
            const vt = isN ? 'Nuevo' : `${p.ytdVar >= 0 ? '+' : ''}${p.ytdVar.toFixed(1)}%`;
            const vc = isN ? 'text-sky-600' : (p.ytdVar >= 0 ? 'text-emerald-600' : 'text-red-600');
            let cl = 'Consolidado', cc = 'bg-slate-100 text-slate-600';
            if (isN) { cl = 'Nuevo'; cc = 'bg-sky-50 text-sky-700'; }
            else if (p.ytdVar >= 20) { cl = 'Crecimiento'; cc = 'bg-emerald-50 text-emerald-700'; }
            else if (p.ytdVar < -30) { cl = 'Declive'; cc = 'bg-red-50 text-red-700'; }
            else if (idx < list.length * 0.15 && !isN && p.ytdVar >= 0) { cl = 'Líder'; cc = 'bg-emerald-50 text-emerald-700'; }
            const sw = Math.min(p.share * 4, 100);
            const paretoColor = p.inPareto ? 'bg-emerald-500' : 'bg-sky-500';
            const paretoText = p.inPareto ? 'text-emerald-600 font-bold' : 'text-slate-500';
            const avgYTD = monthsElapsed > 0 ? Math.round(p.ytdLast / monthsElapsed) : 0;
            return `<tr>
                <td class="py-2.5 px-2 text-slate-400 font-black text-center w-8">${idx + 1}</td>
                <td class="py-2.5 px-2 font-bold text-slate-800 max-w-[220px]"><span class="block truncate" title="${p.article}">${p.article}</span></td>
                <td class="py-2.5 px-2 text-slate-500 max-w-[120px]"><span class="block truncate" title="${p.category}">${p.category}</span></td>
                <td class="py-2.5 px-2 text-right font-black text-slate-800 tabular-nums">${fmt(p.total)}</td>
                <td class="py-2.5 px-2 text-right text-slate-500 tabular-nums">${fmt(avgYTD)}</td>
                <td class="py-2.5 px-2 text-right font-bold ${vc} tabular-nums">${vt}</td>
                <td class="py-2.5 px-2 text-center"><span class="text-[0.65rem] font-bold px-2 py-0.5 rounded-lg ${cc}">${cl}</span></td>
                <td class="py-2.5 px-2 min-w-[80px]">
                    <div class="w-full bg-slate-100 rounded-full h-1.5"><div class="${paretoColor} h-1.5 rounded-full" style="width:${sw}%"></div></div>
                    <p class="text-[0.65rem] ${paretoText} text-center mt-0.5">${p.share.toFixed(1)}%</p>
                </td>
            </tr>`;
        }).join('');
    };
    renderFull(fullSorted);

    // ✅ FIX: sin clonar el input (conserva foco y texto)
    const searchInput = document.getElementById('ale-search-input');
    if (searchInput) {
        searchInput.oninput = (e) => {
            const q = e.target.value.toLowerCase();
            renderFull(aleFullRankingCache.filter(p =>
                p.article.toLowerCase().includes(q) ||
                p.category.toLowerCase().includes(q)
            ));
        };
    }
}


// ============================================================
// ALE TOP 5 / BOTTOM 5
// ============================================================
function renderAleTopBottom(data) {
    const { allBrands, brandFiltered } = getAleBrandStats(data);
    const isBs = aleMode === 'bs';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES',{minimumFractionDigits:0,maximumFractionDigits:0})}` : `${v.toLocaleString('es-ES')} un.`;
    const sorted = [...allBrands].sort((a, b) => b.current - a.current);
    const entityLabel = brandFiltered ? 'SKUs' : 'Marcas';

    const top5Title = document.querySelector('#ale-top5')?.closest('.ale-card')?.querySelector('h3');
    const bot5Title = document.querySelector('#ale-bottom5')?.closest('.ale-card')?.querySelector('h3');
    if (top5Title) top5Title.textContent = `Top 5 ${entityLabel}`;
    if (bot5Title) bot5Title.textContent = `Bottom 5 ${entityLabel}`;

    function pill(b, rank, type) {
        const colors = {
            top:     ['#ECFDF5','#065F46','#A7F3D0'],
            bottom:  ['#FEF2F2','#991B1B','#FECACA'],
            growth:  ['#EFF6FF','#1D4ED8','#BFDBFE'],
            decline: ['#FFF7ED','#9A3412','#FED7AA']
        };
        const [bg, text, border] = colors[type];
        const val = (type === 'growth' || type === 'decline')
            ? (b.varPct !== null ? `${b.varPct >= 0 ? '+' : ''}${b.varPct.toFixed(1)}%` : '—')
            : fmt(b.current);
        const displayName = b.label || b.brand;
        return `<div class="ale-rank-pill" style="background:${bg};border-color:${border};color:${text};">
            <span class="font-black opacity-60 w-5 text-center">${rank}</span>
            <span class="font-bold truncate flex-grow" title="${displayName}">${displayName}</span>
            <span class="font-black text-xs tabular-nums">${val}</span>
        </div>`;
    }

    const top5El = document.getElementById('ale-top5');
    if (top5El) top5El.innerHTML = sorted.slice(0, 5).map((b, i) => pill(b, i+1, 'top')).join('');

    const bottom5El = document.getElementById('ale-bottom5');
    if (bottom5El) {
        if (sorted.length <= 1) {
            bottom5El.innerHTML = `<p class="text-xs text-slate-400 text-center py-2">Filtra un conjunto con múltiples ${entityLabel.toLowerCase()} para ver el Bottom 5.</p>`;
        } else {
            bottom5El.innerHTML = [...sorted].reverse().slice(0, 5).map((b, i) => pill(b, sorted.length - i, 'bottom')).join('');
        }
    }

    const growthSorted = allBrands.filter(b => b.varPct !== null && b.varPct > 0).sort((a, b) => b.varPct - a.varPct);
    const top5GEl = document.getElementById('ale-top5-growth');
    if (top5GEl) {
        const html = growthSorted.slice(0, 5).map((b, i) => pill(b, i+1, 'growth')).join('');
        top5GEl.innerHTML = html || '<p class="text-xs text-slate-400">Sin datos de período anterior</p>';
    }

    const declineSorted = allBrands.filter(b => b.varPct !== null && b.varPct < 0).sort((a, b) => a.varPct - b.varPct);
    const bottom5DEl = document.getElementById('ale-bottom5-decline');
    if (bottom5DEl) {
        const html = declineSorted.slice(0, 5).map((b, i) => pill(b, i+1, 'decline')).join('');
        bottom5DEl.innerHTML = html || '<p class="text-xs text-slate-400">Ninguna en caída</p>';
    }
}

// ============================================================
// ALE COMMERCIAL ALERTS
// ============================================================
function renderAleAlerts(data) {
    const { allBrands, brandFiltered } = getAleBrandStats(data);
    const alerts = [];

    allBrands.forEach(b => {
        const name = b.label || b.brand;
        if (b.varPct !== null) {
            if (b.varPct <= -15) {
                alerts.push({ level: 'red', msg: `<b>${name}</b> cayó <b>${b.varPct.toFixed(1)}%</b> en ventas vs período anterior.` });
            } else if (b.varPct < 0) {
                alerts.push({ level: 'yellow', msg: `<b>${name}</b> registra caída de <b>${b.varPct.toFixed(1)}%</b> vs período anterior.` });
            }
            if (b.varPct >= 15) {
                alerts.push({ level: 'green', msg: `<b>${name}</b> creció <b>+${b.varPct.toFixed(1)}%</b> vs período anterior.` });
            }
        }

        if (allBrands.length > 1 && b.shareDelta !== undefined) {
            if (b.shareDelta <= -2) {
                alerts.push({ level: 'red', msg: `<b>${name}</b> perdió <b>${Math.abs(b.shareDelta).toFixed(1)} pp</b> de participación.` });
            } else if (b.shareDelta < -0.5) {
                alerts.push({ level: 'yellow', msg: `<b>${name}</b> perdió <b>${Math.abs(b.shareDelta).toFixed(1)} pp</b> de participación.` });
            } else if (b.shareDelta >= 1.5) {
                alerts.push({ level: 'green', msg: `<b>${name}</b> ganó <b>+${b.shareDelta.toFixed(1)} pp</b> de participación.` });
            }
        }
    });

    const order = { red: 0, yellow: 1, green: 2 };
    alerts.sort((a, b) => order[a.level] - order[b.level]);

    const el = document.getElementById('ale-alerts');
    if (!el) return;
    if (alerts.length === 0) {
        el.innerHTML = '<p class="text-xs text-slate-400 text-center py-4">No hay alertas para el período seleccionado.</p>';
        return;
    }
    el.innerHTML = alerts.slice(0, 15).map(a => {
        const dotColor = a.level === 'red' ? '#EF4444' : a.level === 'yellow' ? '#F59E0B' : '#10B981';
        return `<div class="ale-alert ale-alert-${a.level}">
            <span class="ale-alert-dot" style="background:${dotColor}"></span>
            <span>${a.msg}</span>
        </div>`;
    }).join('');
}


// ============================================================
// ALE BRAND MATRIX (Bubble chart)
// ============================================================
function renderAleMatrix(data) {
    const { allBrands, total } = getAleBrandStats(data);
    const validBrands = allBrands.filter(b => b.varPct !== null);
    const avgGrowth = validBrands.reduce((s, b) => s + b.varPct, 0) / (validBrands.length || 1);
    const avgShare = 100 / Math.max(allBrands.length, 1);

    const quadrantColors = { leaders: '#10B981', potential: '#0EA5E9', monitor: '#F59E0B', low: '#94A3B8' };

    const datasets = [
        { label: 'Líderes', data: [], backgroundColor: quadrantColors.leaders + 'CC', borderColor: quadrantColors.leaders },
        { label: 'Con Potencial', data: [], backgroundColor: quadrantColors.potential + 'CC', borderColor: quadrantColors.potential },
        { label: 'A Monitorear', data: [], backgroundColor: quadrantColors.monitor + 'CC', borderColor: quadrantColors.monitor },
        { label: 'Bajo Desempeño', data: [], backgroundColor: quadrantColors.low + 'CC', borderColor: quadrantColors.low }
    ];

    allBrands.filter(b => b.varPct !== null && b.current > 0).forEach(b => {
        const isHighGrowth = b.varPct >= avgGrowth;
        const isHighShare = b.share >= avgShare;
        const displayName = b.label || b.brand;
        const point = { x: b.share, y: b.varPct, r: Math.max(5, Math.min(20, Math.sqrt(b.current / total) * 80)), label: displayName };
        if (isHighGrowth && isHighShare) datasets[0].data.push(point);
        else if (isHighGrowth && !isHighShare) datasets[1].data.push(point);
        else if (!isHighGrowth && isHighShare) datasets[2].data.push(point);
        else datasets[3].data.push(point);
    });


    const ctx = document.getElementById('ale-matrix-chart');
    if (!ctx) return;
    if (aleCharts['matrix']) aleCharts['matrix'].destroy();

    aleCharts['matrix'] = new Chart(ctx, {
        type: 'bubble',
        data: { datasets },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { font: { family: 'Nunito', size: 11, weight: '700' }, padding: 16, usePointStyle: true } },
                tooltip: {
                    backgroundColor: '#FFFFFF', borderColor: '#E2E8F0', borderWidth: 1,
                    titleColor: '#0F172A', bodyColor: '#475569',
                    cornerRadius: 12, padding: 12,
                    titleFont: { family: 'Nunito', size: 13, weight: 'bold' },
                    bodyFont: { family: 'Nunito', size: 12 },
                    callbacks: {
                        title: items => items[0].raw.label,
                        label: item => [
                            ` Participación: ${item.raw.x.toFixed(1)}%`,
                            ` Crecimiento: ${item.raw.y >= 0 ? '+' : ''}${item.raw.y.toFixed(1)}%`
                        ]
                    }
                },
                datalabels: { display: false }
            },
            scales: {
                x: {
                    title: { display: true, text: 'Participación de mercado (%)', font: { family: 'Nunito', size: 11, weight: '700' }, color: '#64748B' },
                    ticks: { font: { family: 'Nunito' }, callback: v => v + '%' },
                    grid: { color: 'rgba(0,0,0,0.05)' }
                },
                y: {
                    title: { display: true, text: 'Crecimiento (%)', font: { family: 'Nunito', size: 11, weight: '700' }, color: '#64748B' },
                    ticks: { font: { family: 'Nunito' }, callback: v => (v >= 0 ? '+' : '') + v + '%' },
                    grid: { color: 'rgba(0,0,0,0.05)' }
                }
            }
        }
    });
}

// ============================================================
// ALE EXECUTIVE SUMMARY
// ============================================================
function getAleClosedMonthsPeriod() {
    const allData = getAleData();
    const now = new Date();
    const sysMonth = now.getMonth();
    const sysYear = now.getFullYear();

    let targetMonth = sysMonth - 1;
    let targetYear = sysYear;
    if (targetMonth < 0) { targetMonth = 11; targetYear -= 1; }

    const hasData = (y, m) => allData.some(d => d.year === y && d.month === m && d.units > 0);

    if (!hasData(targetYear, targetMonth)) {
        const closedMonths = allData
            .filter(d => d.units > 0 && !(d.year === sysYear && d.month === sysMonth))
            .map(d => ({ year: d.year, month: d.month }));
        if (closedMonths.length > 0) {
            closedMonths.sort((a, b) => (a.year - b.year) || (a.month - b.month));
            const last = closedMonths[closedMonths.length - 1];
            targetYear = last.year;
            targetMonth = last.month;
        }
    }

    let compMonth = targetMonth - 1;
    let compYear = targetYear;
    if (compMonth < 0) { compMonth = 11; compYear -= 1; }

    return { targetYear, targetMonth, compYear, compMonth };
}

function getAleExecSummaryData() {
    const data = getAleData();
    return data.filter(d => {
        if (aleFilters.brand !== 'all' && d.brand !== aleFilters.brand) return false;
        if (aleFilters.category !== 'all' && d.category !== aleFilters.category) return false;
        if (aleFilters.tipo !== 'all' && d.tipo !== aleFilters.tipo) return false;
        if (aleFilters.sku && !d.article.toLowerCase().includes(aleFilters.sku.toLowerCase())) return false;
        return true;
    });
}

function renderAleExecutiveSummary() {
    const isBs = aleMode === 'bs';
    const unit = isBs ? 'Bs' : 'unidades';
    const fmt = v => isBs ? `Bs ${v.toLocaleString('es-ES',{minimumFractionDigits:0,maximumFractionDigits:0})}` : v.toLocaleString('es-ES');

    const brandFiltered = aleFilters.brand !== 'all';
    const groupKey = brandFiltered ? 'article' : 'brand';
    const entityWord = brandFiltered ? 'SKU' : 'marca';
    const entityWordP = brandFiltered ? 'producto' : 'portafolio';

    const { targetYear, targetMonth, compYear, compMonth } = getAleClosedMonthsPeriod();
    const execData = getAleExecSummaryData();

    const curMonthData = execData.filter(d => d.year === targetYear && d.month === targetMonth);
    const prevMonthData = execData.filter(d => d.year === compYear && d.month === compMonth);

    const curMonthTotal = curMonthData.reduce((s, d) => s + d.units, 0);
    const prevMonthTotal = prevMonthData.reduce((s, d) => s + d.units, 0);
    const varMes = prevMonthTotal > 0 ? ((curMonthTotal / prevMonthTotal) - 1) * 100 : null;

    const entityMap = {};
    const addEntity = (d, field) => {
        const key = d[groupKey];
        if (!key) return;
        if (!entityMap[key]) entityMap[key] = { label: key, current: 0, previous: 0 };
        entityMap[key][field] += d.units;
    };
    curMonthData.forEach(d => addEntity(d, 'current'));
    prevMonthData.forEach(d => addEntity(d, 'previous'));

    const entities = Object.values(entityMap).filter(e => e.current > 0 || e.previous > 0);
    entities.forEach(e => {
        e.varPct = e.previous > 0 ? ((e.current / e.previous) - 1) * 100 : null;
        e.share = curMonthTotal > 0 ? (e.current / curMonthTotal) * 100 : 0;
        e.prevShare = prevMonthTotal > 0 ? (e.previous / prevMonthTotal) * 100 : 0;
        e.shareDelta = e.share - e.prevShare;
    });

    const topItem = [...entities].sort((a, b) => b.current - a.current)[0];
    const topGrowth = [...entities].filter(e => e.varPct !== null && e.varPct > 0).sort((a, b) => b.varPct - a.varPct)[0];
    const topDecline = [...entities].filter(e => e.varPct !== null && e.varPct < 0).sort((a, b) => a.varPct - b.varPct)[0];
    const topShareLoss = entities.length > 1 ? [...entities].filter(e => e.shareDelta < 0).sort((a, b) => a.shareDelta - b.shareDelta)[0] : null;

    let text = `Las ventas de ${MESES[targetMonth]} ${targetYear} alcanzaron <strong>${fmt(curMonthTotal)} ${unit}</strong>`;
    if (varMes !== null) {
        text += `, representando un ${varMes >= 0 ? '<span style="color:#10B981">crecimiento</span>' : '<span style="color:#EF4444">retroceso</span>'} de <strong>${varMes >= 0 ? '+' : ''}${varMes.toFixed(1)}%</strong> vs ${MESES[compMonth]} ${compYear}`;
    }
    text += '.';
    if (topItem) {
        text += ` El ${entityWord} <strong>${topItem.label}</strong> lidera el ${entityWordP} con una participación del <strong>${topItem.share.toFixed(1)}%</strong>.`;
    }
    if (topDecline) {
        text += ` <span style="color:#EF4444"><strong>${topDecline.label}</strong> presenta la mayor caída (${topDecline.varPct.toFixed(1)}%)</span>.`;
    }
    if (topGrowth) {
        text += ` <span style="color:#10B981"><strong>${topGrowth.label}</strong> destaca por crecimiento (+${topGrowth.varPct.toFixed(1)}%)</span>.`;
    }
    if (topShareLoss) {
        text += ` <strong>${topShareLoss.label}</strong> registró pérdida de participación (${topShareLoss.shareDelta.toFixed(1)} pp).`;
    }
    if (curMonthData.length === 0) {
        text = `No hay datos disponibles para un periodo mensual completamente cerrado todavía.`;
    }

    const el = document.getElementById('ale-executive-summary');
    if (el) el.innerHTML = text;

    const periodEl = document.getElementById('ale-exec-period');
    if (periodEl) periodEl.textContent = `${MESES[targetMonth]} ${targetYear} vs ${MESES[compMonth]} ${compYear} · ${isBs ? 'Bs Vendidos' : 'Unidades'}`;
}


// ============================================================
// ALE PRESENTATION MODE
// ============================================================
let alePresentationActive = false;
function toggleAlePresentation() {
    alePresentationActive = !alePresentationActive;
    const aleView = document.getElementById('team-ale-view');
    const btn = document.getElementById('ale-btn-presentation');
    if (alePresentationActive) {
        aleView.classList.add('ale-presentation-mode');
        if (btn) btn.innerHTML = '<i data-lucide="x" class="w-3 h-3"></i> Salir Presentación';
    } else {
        aleView.classList.remove('ale-presentation-mode');
        if (btn) btn.innerHTML = '<i data-lucide="presentation" class="w-3 h-3"></i> Modo Presentación';
    }
    lucide.createIcons();
}

// ============================================================
// ALE EXPORT PNG
// ============================================================
function exportAleChart(canvasId, filename) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = `${filename}-${new Date().toISOString().slice(0,10)}.png`;
    link.href = canvas.toDataURL('image/png', 1.0);
    link.click();
}

// ============================================================
// BOOTSTRAP
// ============================================================
document.addEventListener('DOMContentLoaded', autoLoadData);