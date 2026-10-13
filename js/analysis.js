(() => {
  "use strict";
  const $ = s => document.querySelector(s);
  const dataInput = $("#analysisDataFile"), mapInput = $("#analysisMappingFile"), runButton = $("#runAnalysis");
  const dataStatus = $("#dataStatus"), mappingStatus = $("#mappingStatus"), summary = $("#analysisSummary");
  const resultsRoot = $("#analysisResults"), downloads = $("#overallDownloads");
  const saveButton=$("#saveAnalysisWorkspace"), resumeButton=$("#resumeAnalysisWorkspace"), deleteButton=$("#deleteAnalysisWorkspace"), saveStatus=$("#analysisSaveStatus");
  let dataModel = null, mappingModel = null, results = [];
  let dataFileBuffer=null, mappingFileBuffer=null, dataFileName="", mappingFileName="";
  const ANALYSIS_DB="survey-analysis-workspace-v1", ANALYSIS_STORE="workspace", ANALYSIS_KEY="current";

  const text = v => String(v ?? "").trim();
  const esc = v => text(v).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const codeKey = v => {
    if (v === null || v === undefined) return "";
    const s = text(v);
    if (/^-?\d+(?:\.0+)?$/.test(s)) return String(Number(s));
    return s;
  };
  const isMissing = v => v === null || v === undefined || text(v) === "" || text(v) === ".";
  const isOne = v => codeKey(v) === "1";
  const baseV = h => (text(h).match(/^(V\d+)/i) || [])[1]?.toUpperCase() || "";
  const branchCode = h => (text(h).match(/^V\d+-(.+)$/i) || [])[1] || "";
  const sortV = (a,b) => Number(a.variable.slice(1)) - Number(b.variable.slice(1));
  const setStatus = (el, type, message) => { el.className = `analysis-status-${type}`; el.textContent = message; };

  async function workbookFromBuffer(buffer){
    if (!window.ExcelJS) throw new Error("Excel読込ライブラリが読み込まれていません。");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer.slice(0));
    return wb;
  }
  async function workbookFromFile(file){return workbookFromBuffer(await file.arrayBuffer());}
  function cellValue(cell){
    const v = cell?.value;
    if (v && typeof v === "object") {
      if ("result" in v) return v.result;
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text).join("");
      if ("text" in v) return v.text;
    }
    return v;
  }
  function rowValues(ws,rowNumber,maxColumn){
    const values=[];
    for(let c=1;c<=maxColumn;c++) values.push(cellValue(ws.getCell(rowNumber,c)));
    return values;
  }

  function openAnalysisDB(){
    return new Promise((resolve,reject)=>{
      const request=indexedDB.open(ANALYSIS_DB,1);
      request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(ANALYSIS_STORE))request.result.createObjectStore(ANALYSIS_STORE);};
      request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error||new Error("途中保存領域を開けませんでした。"));
    });
  }
  async function dbPut(value){const db=await openAnalysisDB();return new Promise((resolve,reject)=>{const tx=db.transaction(ANALYSIS_STORE,"readwrite");tx.objectStore(ANALYSIS_STORE).put(value,ANALYSIS_KEY);tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>{db.close();reject(tx.error);};});}
  async function dbGet(){const db=await openAnalysisDB();return new Promise((resolve,reject)=>{const tx=db.transaction(ANALYSIS_STORE,"readonly"),req=tx.objectStore(ANALYSIS_STORE).get(ANALYSIS_KEY);req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error);tx.oncomplete=()=>db.close();});}
  async function dbDelete(){const db=await openAnalysisDB();return new Promise((resolve,reject)=>{const tx=db.transaction(ANALYSIS_STORE,"readwrite");tx.objectStore(ANALYSIS_STORE).delete(ANALYSIS_KEY);tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>{db.close();reject(tx.error);};});}
  function savedChartTypes(){return Object.fromEntries(results.map(r=>[r.variable,r.chartType||"bar"]));}
  function formatSavedAt(value){try{return new Intl.DateTimeFormat("ja-JP",{dateStyle:"medium",timeStyle:"short"}).format(new Date(value));}catch{return text(value);}}
  async function refreshSaveStatus(){try{const saved=await dbGet();if(saved){saveStatus.textContent=`保存済み：${formatSavedAt(saved.savedAt)}（${saved.dataFileName||"データExcel"}／${saved.mappingFileName||"対応表Excel"}）`;resumeButton.disabled=false;deleteButton.disabled=false;}else{saveStatus.textContent="保存データはありません。";resumeButton.disabled=true;deleteButton.disabled=true;}}catch(error){saveStatus.textContent=`途中保存を確認できません：${error.message}`;}}
  async function saveAnalysisWorkspace(){
    if(!dataFileBuffer||!mappingFileBuffer){alert("データExcelと対応表Excelを読み込んでから途中保存してください。");return;}
    saveButton.disabled=true;
    try{
      const payload={version:1,savedAt:new Date().toISOString(),dataFileName,mappingFileName,dataBuffer:dataFileBuffer.slice(0),mappingBuffer:mappingFileBuffer.slice(0),chartTypes:savedChartTypes(),scrollY:window.scrollY};
      await dbPut(payload);
      const verified=await dbGet();
      if(!verified||!verified.dataBuffer||!verified.mappingBuffer)throw new Error("保存内容を読み戻せませんでした。");
      saveStatus.textContent=`保存済み：${formatSavedAt(verified.savedAt)}（${verified.dataFileName}／${verified.mappingFileName}）`;
      resumeButton.disabled=false;deleteButton.disabled=false;
    }catch(error){alert(`途中保存できませんでした。\n${error.message}`);saveStatus.textContent="途中保存に失敗しました。";}finally{saveButton.disabled=false;}
  }
  async function restoreAnalysisWorkspace(){
    resumeButton.disabled=true;
    try{
      const saved=await dbGet();if(!saved)throw new Error("保存データがありません。");
      dataFileBuffer=saved.dataBuffer.slice(0);mappingFileBuffer=saved.mappingBuffer.slice(0);dataFileName=saved.dataFileName||"保存済みデータ.xlsx";mappingFileName=saved.mappingFileName||"保存済み対応表.xlsx";
      setStatus(dataStatus,"warn","保存データを復元中…");setStatus(mappingStatus,"warn","保存データを復元中…");
      dataModel=parseDataWorkbook(await workbookFromBuffer(dataFileBuffer));mappingModel=parseMappingWorkbook(await workbookFromBuffer(mappingFileBuffer));
      setStatus(dataStatus,"ok",`${dataFileName}：${dataModel.rows.length}行 × ${dataModel.headers.length}列を復元しました。`);
      setStatus(mappingStatus,"ok",`${mappingFileName}：${mappingModel.maps.size}設問を復元しました。`);
      runAnalysis(saved.chartTypes||{});
      saveStatus.textContent=`再開済み：${formatSavedAt(saved.savedAt)}`;
      requestAnimationFrame(()=>window.scrollTo({top:Number(saved.scrollY)||0,behavior:"auto"}));
    }catch(error){alert(`保存データから再開できませんでした。\n${error.message}`);}finally{resumeButton.disabled=false;}
  }
  async function deleteAnalysisWorkspace(){if(!confirm("簡易分析の途中保存を削除しますか？"))return;try{await dbDelete();await refreshSaveStatus();}catch(error){alert(`途中保存を削除できませんでした。\n${error.message}`);}}

  function parseDataWorkbook(wb){
    const ws = wb.getWorksheet("加工済みデータ") || wb.worksheets[0];
    if (!ws) throw new Error("データシートがありません。");
    const headers = rowValues(ws,1,ws.columnCount).map(text);
    if (!headers.some(h => /^V\d+(?:-\d+)?$/i.test(h))) throw new Error("1行目にV番号が見つかりません。");
    const rows=[];
    for(let r=2;r<=ws.rowCount;r++){
      const row=rowValues(ws,r,headers.length);
      if(row.some(v => !isMissing(v))) rows.push(row);
    }
    const index = new Map(headers.map((h,i)=>[h.toUpperCase(),i]));
    const metadata = new Map();
    const metaSheet = wb.getWorksheet("加工結果対応表");
    if(metaSheet){
      for(let r=2;r<=metaSheet.rowCount;r++){
        const variable=text(cellValue(metaSheet.getCell(r,1))).toUpperCase();
        const question=text(cellValue(metaSheet.getCell(r,2)));
        const method=text(cellValue(metaSheet.getCell(r,3)));
        if(/^V\d+$/.test(variable) && !metadata.has(variable)) metadata.set(variable,{variable,question,method});
      }
    }
    return {sheetName:ws.name,headers,rows,index,metadata};
  }

  function optionScore(code,label){
    const c=codeKey(code), l=text(label);
    if(l===c) return 4;
    if(new RegExp(`^${c.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}\\s*[.．]`).test(l)) return 3;
    if(/^\d+\s*[.．]/.test(l)) return 2;
    return 1;
  }
  function dedupeOptions(raw){
    const best=new Map();
    raw.forEach((o,order)=>{
      const k=codeKey(o.code); if(!k || !text(o.label)) return;
      const item={code:k,label:text(o.label),order,score:optionScore(k,o.label)};
      const prior=best.get(k);
      if(!prior || item.score>prior.score) best.set(k,item);
    });
    return [...best.values()].sort((a,b)=>a.order-b.order).map(({code,label})=>({code,label}));
  }
  function parseMappingWorkbook(wb){
    const maps=new Map(), warnings=[];
    wb.worksheets.forEach(ws=>{
      let variable="", question="", headerRow=0;
      for(let r=1;r<=Math.min(ws.rowCount,20);r++){
        const a=text(cellValue(ws.getCell(r,1))), b=text(cellValue(ws.getCell(r,2)));
        if(a==="V番号") variable=(b||ws.name).toUpperCase();
        if(a==="設問文") question=b;
        if(a==="加工済の番号" && b==="回答選択肢") headerRow=r;
      }
      if(!/^V\d+$/i.test(variable) || !headerRow) return;
      const raw=[];
      for(let r=headerRow+1;r<=ws.rowCount;r++){
        const code=cellValue(ws.getCell(r,1)), label=cellValue(ws.getCell(r,2));
        if(!isMissing(code) && !isMissing(label)) raw.push({code,label});
      }
      const options=dedupeOptions(raw);
      if(raw.length!==options.length) warnings.push(`${variable}: 重複コード${raw.length-options.length}件を統合`);
      maps.set(variable,{variable,question,options,sheetName:ws.name});
    });
    if(!maps.size) throw new Error("「V番号」「設問文」「加工済の番号」「回答選択肢」の形式を持つシートがありません。");
    return {maps,warnings};
  }

  function numericValues(variable){
    const ci=dataModel.index.get(variable);
    if(ci===undefined)return [];
    return dataModel.rows.map(row=>row[ci]).filter(v=>!isMissing(v)).map(v=>Number(v)).filter(Number.isFinite);
  }
  function numericBins(values){
    if(!values.length)return [];
    const sorted=[...values].sort((a,b)=>a-b), min=sorted[0], max=sorted.at(-1);
    const unique=[...new Set(sorted)];
    if(unique.length<=12){
      return unique.map(v=>({code:String(v),label:String(v),count:values.filter(x=>x===v).length}));
    }
    if(min===max)return [{code:String(min),label:String(min),count:values.length}];
    const binCount=Math.min(10,Math.max(5,Math.ceil(Math.sqrt(values.length))));
    const width=(max-min)/binCount, bins=[];
    for(let i=0;i<binCount;i++){
      const lo=min+i*width, hi=i===binCount-1?max:min+(i+1)*width;
      bins.push({code:String(i+1),label:`${Math.round(lo*10)/10} - ${Math.round(hi*10)/10}`,count:0,lo,hi});
    }
    values.forEach(v=>{const i=Math.min(binCount-1,Math.floor((v-min)/width));bins[i].count++;});
    return bins;
  }
  function analyzeNumeric(meta){
    const values=numericValues(meta.variable); if(!values.length)return null;
    const items=numericBins(values), valid=values.length;
    const ci=dataModel.index.get(meta.variable);
    const missing=dataModel.rows.reduce((n,row)=>n+(isMissing(row[ci])?1:0),0);
    const mean=values.reduce((a,b)=>a+b,0)/valid;
    const sorted=[...values].sort((a,b)=>a-b), median=sorted.length%2?sorted[(sorted.length-1)/2]:(sorted[sorted.length/2-1]+sorted[sorted.length/2])/2;
    items.forEach(x=>x.percent=valid?x.count/valid*100:0);
    return {variable:meta.variable,question:meta.question,type:"NUMERIC",items,valid,missing,invalid:0,unexpected:0,mean,median,min:sorted[0],max:sorted.at(-1),values:[...values],numericMode:"bins",binCount:10,rangeMin:sorted[0],rangeMax:sorted.at(-1),labelMode:"percent"};
  }

  function analyzeOne(map){
    const directIndex=dataModel.index.get(map.variable);
    const branchHeaders=dataModel.headers.filter(h=>baseV(h)===map.variable && branchCode(h));
    if(branchHeaders.length){
      const byCode=new Map(map.options.map(o=>[codeKey(o.code),o]));
      const items=[]; let valid=0, missing=0, invalid=0;
      const available=branchHeaders.filter(h=>byCode.has(codeKey(branchCode(h))));
      const unexpectedColumns=branchHeaders.filter(h=>!byCode.has(codeKey(branchCode(h))));
      dataModel.rows.forEach(row=>{
        const vals=available.map(h=>row[dataModel.index.get(h.toUpperCase())]);
        if(vals.every(isMissing)){missing++;return;}
        const bad=vals.some(v=>!isMissing(v) && !["0","1"].includes(codeKey(v)));
        if(bad) invalid++;
        else valid++;
      });
      available.forEach(h=>{
        const code=codeKey(branchCode(h)), option=byCode.get(code), ci=dataModel.index.get(h.toUpperCase());
        const count=dataModel.rows.reduce((n,row)=>n+(isOne(row[ci])?1:0),0);
        items.push({code,label:option.label,count,percent:valid?count/valid*100:0,column:h});
      });
      return {variable:map.variable,question:map.question,type:"MA",items,valid,missing,invalid,unexpected:unexpectedColumns.length};
    }
    if(directIndex!==undefined){
      const allowed=new Map(map.options.map(o=>[codeKey(o.code),o]));
      const counts=new Map(map.options.map(o=>[codeKey(o.code),0]));
      let valid=0,missing=0,invalid=0;
      dataModel.rows.forEach(row=>{
        const v=row[directIndex];
        if(isMissing(v)){missing++;return;}
        const k=codeKey(v);
        if(!allowed.has(k)){invalid++;return;}
        counts.set(k,(counts.get(k)||0)+1);valid++;
      });
      const items=map.options.map(o=>({code:codeKey(o.code),label:o.label,count:counts.get(codeKey(o.code))||0,percent:valid?(counts.get(codeKey(o.code))||0)/valid*100:0}));
      return {variable:map.variable,question:map.question,type:"SA",items,valid,missing,invalid,unexpected:0};
    }
    return null;
  }
  function runAnalysis(chartTypes={}){
    if(!dataModel || !mappingModel){alert("回答データと設問別対応表の両方を読み込んでください。");return;}
    const mapped=[...mappingModel.maps.values()].map(analyzeOne).filter(Boolean);
    const mappedVariables=new Set(mapped.map(r=>r.variable));
    const numeric=[...dataModel.metadata.values()]
      .filter(meta=>meta.variable!=="V0" && !mappedVariables.has(meta.variable) && dataModel.index.has(meta.variable))
      .map(analyzeNumeric).filter(Boolean);
    results=[...mapped,...numeric].sort(sortV);
    results.forEach(r=>{
      r.chartType=chartTypes[r.variable]||r.chartType||"bar";r.labelMode=(r.labelMode&&r.labelMode!=="none")?r.labelMode:"percent";
      if(r.type==="SA"||r.type==="MA"){
        r.categoryItems=(r.items||[]).map(x=>({...x}));r.numericMode="raw";
        const nums=categoricalNumericValues(r);const values=nums.map(x=>x.value);
        r.binCount=10;r.rangeMin=values.length?Math.min(...values):0;r.rangeMax=values.length?Math.max(...values):1;
      }
    });
    renderResults();
    const invalid=results.reduce((n,r)=>n+r.invalid,0), skipped=[...mappingModel.maps.keys()].filter(v=>!results.some(r=>r.variable===v)).length;
    summary.textContent=`${results.length}設問を集計しました。対応表に一致しない値は${invalid}件、データ列がない対応表は${skipped}件です。`;
    downloads.hidden=!results.length;
  }

  function numericItems(r){
    const vals=r.values||[]; if(r.numericMode==="raw"){
      const counts=new Map();vals.forEach(v=>counts.set(v,(counts.get(v)||0)+1));
      return [...counts.entries()].sort((a,b)=>a[0]-b[0]).map(([v,count],i)=>({code:String(v),label:String(v),count,percent:r.valid?count/r.valid*100:0}));
    }
    const n=Math.max(2,Math.min(50,Number(r.binCount)||10)),lo=Number.isFinite(Number(r.rangeMin))?Number(r.rangeMin):r.min,hi=Number.isFinite(Number(r.rangeMax))?Number(r.rangeMax):r.max;
    if(hi<=lo)return numericBins(vals);
    const width=(hi-lo)/n,bins=Array.from({length:n},(_,i)=>({code:String(i+1),label:`${round(lo+i*width)}～${round(lo+(i+1)*width)}`,count:0}));
    let under=0,over=0;vals.forEach(v=>{if(v<lo)under++;else if(v>hi)over++;else bins[Math.min(n-1,Math.floor((v-lo)/width))].count++;});
    const out=[];if(under)out.push({code:"<",label:`${round(lo)}未満`,count:under});out.push(...bins);if(over)out.push({code:">",label:`${round(hi)}超`,count:over});
    out.forEach(x=>x.percent=r.valid?x.count/r.valid*100:0);return out;
  }
  function round(v){return Math.round(v*100)/100;}
  function categoricalNumericValues(r){
    return (r.categoryItems||r.items||[]).map(x=>({value:Number(x.code),count:Number(x.count)||0})).filter(x=>Number.isFinite(x.value));
  }
  function categoricalBins(r){
    const vals=categoricalNumericValues(r); if(!vals.length)return r.categoryItems||r.items||[];
    const loDefault=Math.min(...vals.map(x=>x.value)),hiDefault=Math.max(...vals.map(x=>x.value));
    const n=Math.max(2,Math.min(50,Number(r.binCount)||10)),lo=Number.isFinite(Number(r.rangeMin))?Number(r.rangeMin):loDefault,hi=Number.isFinite(Number(r.rangeMax))?Number(r.rangeMax):hiDefault;
    if(hi<=lo)return r.categoryItems||r.items||[];
    const width=(hi-lo)/n,bins=Array.from({length:n},(_,i)=>({code:String(i+1),label:`${round(lo+i*width)}～${round(lo+(i+1)*width)}`,count:0}));
    let under=0,over=0; vals.forEach(x=>{if(x.value<lo)under+=x.count;else if(x.value>hi)over+=x.count;else bins[Math.min(n-1,Math.floor((x.value-lo)/width))].count+=x.count;});
    const out=[];if(under)out.push({code:"<",label:`${round(lo)}未満`,count:under});out.push(...bins);if(over)out.push({code:">",label:`${round(hi)}超`,count:over});
    out.forEach(x=>x.percent=r.valid?x.count/r.valid*100:0);return out;
  }
  function refreshNumeric(r){
    if(r.type==="NUMERIC")r.items=numericItems(r);
    else if((r.type==="SA"||r.type==="MA")&&r.numericMode==="bins")r.items=categoricalBins(r);
    else if(r.categoryItems)r.items=r.categoryItems;
  }
  function valueLabel(r,x){return r.labelMode==="count"?String(x.count):r.labelMode==="percent"?`${x.percent.toFixed(1)}%`:"";}
  function statsText(r){return r.type==="NUMERIC"?`有効 ${r.valid}件 ／ 欠測 ${r.missing}件 ／ 平均 ${r.mean.toFixed(1)} ／ 中央値 ${r.median.toFixed(1)} ／ 最小 ${r.min} ／ 最大 ${r.max}`:`有効 ${r.valid}件 ／ 欠測 ${r.missing}件`;}
  function controlsHtml(r,i){
    const common=`<label>グラフ<select onchange="changeChartType(${i},this.value)"><option value="bar" ${r.chartType!=="pie"?"selected":""}>棒グラフ</option><option value="pie" ${r.chartType==="pie"?"selected":""}>円グラフ</option></select></label><label>グラフ内表示<select onchange="changeLabelMode(${i},this.value)"><option value="none" ${r.labelMode==="none"?"selected":""}>表示しない</option><option value="count" ${r.labelMode==="count"?"selected":""}>実数</option><option value="percent" ${r.labelMode==="percent"?"selected":""}>割合（%）</option></select></label>`;
    const canBin=r.type==="NUMERIC"||r.type==="SA"||r.type==="MA";
    if(!canBin)return common;
    const modeLabel=r.type==="NUMERIC"?"数値のまとめ方":"数値コードのまとめ方";
    const current=r.numericMode||(r.type==="NUMERIC"?"bins":"raw");
    return common+`<label>${modeLabel}<select onchange="changeNumericMode(${i},this.value)"><option value="bins" ${current!=="raw"?"selected":""}>階級でまとめる</option><option value="raw" ${current==="raw"?"selected":""}>${r.type==="NUMERIC"?"値ごとの件数":"通常の選択肢表示"}</option></select></label><span class="numeric-bin-settings ${current==="raw"?"is-hidden":""}"><label>階級数<input type="number" min="2" max="50" value="${r.binCount}" onchange="changeNumericSetting(${i},'binCount',this.value)"></label><label>範囲<input type="number" value="${r.rangeMin}" onchange="changeNumericSetting(${i},'rangeMin',this.value)">～<input type="number" value="${r.rangeMax}" onchange="changeNumericSetting(${i},'rangeMax',this.value)"></label></span>`;
  }
  function renderResults(){
    results.forEach(refreshNumeric);
    resultsRoot.innerHTML=results.map((r,i)=>{const rows=r.items.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.label)}</td><td>${x.count}</td><td>${x.percent.toFixed(1)}%</td></tr>`).join("");const col=r.type==="NUMERIC"?(r.numericMode==="raw"?"値":"値・階級"):"回答選択肢";return `<article class="analysis-result-card" data-result="${i}"><header class="analysis-result-head"><div><span class="analysis-v-number">${esc(r.variable)} / ${r.type}</span><h2 class="analysis-question">${esc(r.question)}</h2></div><div class="analysis-meta">${statsText(r)}</div></header><div class="analysis-chart-control">${controlsHtml(r,i)}</div><div class="analysis-result-grid"><div><table class="analysis-table"><thead><tr><th>番号</th><th>${col}</th><th>件数</th><th>割合</th></tr></thead><tbody>${rows}</tbody></table></div><div class="analysis-chart-host" id="analysisChart${i}">${r.chartType==="pie"?pieChartHtml(r):barChartHtml(r)}</div></div><div class="analysis-export-groups"><section><strong>1. 表単体</strong><div><button class="button" onclick="downloadTablePNG(${i})">PNG</button><button class="button" onclick="downloadResultExcel(${i})">Excel</button><button class="button" onclick="downloadResultCSV(${i})">CSV</button></div></section><section><strong>2. グラフ単体</strong><div><button class="button" onclick="downloadChartPNG(${i})">PNG</button></div></section><section><strong>3. 表＋グラフ</strong><div><button class="button primary" onclick="downloadCombinedPNG(${i})">PNG</button><button class="button" onclick="downloadCombinedExcel(${i})">Excel</button><button class="button" onclick="downloadCombinedCSV(${i})">CSV</button></div></section></div></article>`}).join("");
  }
  window.changeLabelMode=(i,v)=>{results[i].labelMode=v;renderResults();};
  function showAnalysisWarning(message){
    const modal=document.getElementById("analysisWarningModal"),textEl=document.getElementById("analysisWarningMessage"),accept=document.getElementById("analysisWarningAccept"),cancel=document.getElementById("analysisWarningCancel");
    if(!modal||!textEl||!accept||!cancel)return Promise.resolve(false);
    textEl.textContent=message;modal.classList.add("is-open");modal.setAttribute("aria-hidden","false");document.body.classList.add("analysis-modal-open");
    return new Promise(resolve=>{const close=result=>{modal.classList.remove("is-open");modal.setAttribute("aria-hidden","true");document.body.classList.remove("analysis-modal-open");accept.removeEventListener("click",yes);cancel.removeEventListener("click",no);modal.removeEventListener("click",backdrop);document.removeEventListener("keydown",key);resolve(result);};const yes=()=>close(true),no=()=>close(false),backdrop=e=>{if(e.target.hasAttribute("data-modal-cancel"))close(false);},key=e=>{if(e.key==="Escape")close(false);};accept.addEventListener("click",yes);cancel.addEventListener("click",no);modal.addEventListener("click",backdrop);document.addEventListener("keydown",key);requestAnimationFrame(()=>accept.focus());});
  }
  window.changeNumericMode=async(i,v)=>{
    const r=results[i];
    if(v==="bins"&&(r.type==="SA"||r.type==="MA")){
      const ok=await showAnalysisWarning(`${r.type}の回答を数値階級にまとめます。通常は選択肢別に集計する設問です。本当に階級でまとめてよろしいですか？`);
      if(!ok){renderResults();return;}
      if(!categoricalNumericValues(r).length){alert("数値として扱える選択肢コードがないため、階級にまとめられません。");renderResults();return;}
    }
    r.numericMode=v;refreshNumeric(r);renderResults();
  };
  window.changeNumericSetting=(i,k,v)=>{results[i][k]=Number(v);refreshNumeric(results[i]);renderResults();};

  function renderResultsLegacy(){
    if(!results.length){resultsRoot.innerHTML='<div class="analysis-empty">集計対象がありません。</div>';return;}
    resultsRoot.innerHTML=results.map((r,i)=>{
      const max=Math.max(1,...r.items.map(x=>x.count));
      const rows=r.items.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.label)}</td><td>${x.count}</td><td>${x.percent.toFixed(1)}%</td></tr>`).join("");
      const bars=r.items.map(x=>`<div class="analysis-bar-row"><span class="analysis-bar-label">${esc(x.label)}</span><span class="analysis-bar-track"><span class="analysis-bar" style="width:${x.count/max*100}%"></span></span><span class="analysis-bar-value">${x.count}</span></div>`).join("");
      const warnings=[]; if(r.invalid)warnings.push(`対応表にない値 ${r.invalid}件は除外`); if(r.unexpected)warnings.push(`対応表にないMA列 ${r.unexpected}列は除外`);
      const numericMeta=r.type==="NUMERIC"?`<br>平均 ${r.mean.toFixed(1)} ／ 中央値 ${r.median.toFixed(1)} ／ 最小 ${r.min} ／ 最大 ${r.max}`:"";
      const column2=r.type==="NUMERIC"?"値・階級":"回答選択肢";
      return `<article class="analysis-result-card" data-result="${i}"><header class="analysis-result-head"><div><span class="analysis-v-number">${esc(r.variable)} / ${r.type}</span><h2 class="analysis-question">${esc(r.question||"設問文なし")}</h2></div><div class="analysis-meta">有効 ${r.valid}件 ／ 欠測 ${r.missing}件${numericMeta}${warnings.length?`<br><span class="analysis-status-warn">${esc(warnings.join("・"))}</span>`:""}</div></header><div class="analysis-chart-control"><label>グラフ<select onchange="changeChartType(${i},this.value)"><option value="bar" ${r.chartType!=="pie"?"selected":""}>棒グラフ</option><option value="pie" ${r.chartType==="pie"?"selected":""}>円グラフ</option></select></label></div><div class="analysis-result-grid"><div><table class="analysis-table"><thead><tr><th>番号</th><th>${column2}</th><th>件数</th><th>割合</th></tr></thead><tbody>${rows}</tbody></table></div><div class="analysis-chart-host" id="analysisChart${i}">${r.chartType==="pie"?pieChartHtml(r):barChartHtml(r)}</div></div><div class="analysis-export-groups"><section><strong>1. 表単体</strong><div><button class="button" onclick="downloadTablePNG(${i})">PNG</button><button class="button" onclick="downloadResultExcel(${i})">Excel</button><button class="button" onclick="downloadResultCSV(${i})">CSV</button></div></section><section><strong>2. グラフ単体</strong><div><button class="button" onclick="downloadChartPNG(${i})">PNG</button></div></section><section><strong>3. 表＋グラフ</strong><div><button class="button primary" onclick="downloadCombinedPNG(${i})">PNG</button></div></section></div></article>`;
    }).join("");
  }
  const PASTEL_COLORS=["#8FB3E8","#86D1C5","#F4C08A","#EE9CAA","#B69BDD","#8CC8E2","#A8D58A","#F1D477","#D59ACA","#95D0B1","#F2A47F","#A8B9DC"];
  function polarPoint(cx,cy,radius,angle){const rad=(angle-90)*Math.PI/180;return {x:cx+radius*Math.cos(rad),y:cy+radius*Math.sin(rad)};}
  function pieSlicePath(cx,cy,radius,start,end){const a=polarPoint(cx,cy,radius,end),b=polarPoint(cx,cy,radius,start),large=end-start>180?1:0;return `M ${cx} ${cy} L ${a.x} ${a.y} A ${radius} ${radius} 0 ${large} 0 ${b.x} ${b.y} Z`;}
  function barChartHtml(r){
    const W=820,plotX=300,plotW=440,rightPad=70,rowH=48,top=42,bottom=38,H=Math.max(190,top+r.items.length*rowH+bottom),max=Math.max(1,...r.items.map(x=>x.count));
    const grid=Array.from({length:11},(_,i)=>{const x=plotX+plotW*i/10;return `<line x1="${x}" y1="${top-8}" x2="${x}" y2="${H-bottom}" stroke="${i===0||i===10?'#aeb7c4':'#d8dee8'}" stroke-width="${i===0||i===10?1.4:1}"/><text x="${x}" y="${H-12}" text-anchor="middle" font-size="15" font-weight="700" fill="#354052">${i*10}%</text>`;}).join("");
    const rows=r.items.map((x,i)=>{const y=top+i*rowH,bw=Math.max(0,x.percent/100*plotW),label=valueLabel(r,x),inside=bw>=84,tx=inside?plotX+bw-10:Math.min(plotX+bw+10,plotX+plotW+rightPad-8),anchor=inside?"end":"start";return `<text x="8" y="${y+22}" font-size="18" font-weight="750" fill="#202124">${esc(text(x.label).slice(0,24))}</text><rect x="${plotX}" y="${y}" width="${plotW}" height="30" rx="0" fill="#f2f4f7"/><rect x="${plotX}" y="${y}" width="${bw}" height="30" rx="0" fill="${PASTEL_COLORS[i%PASTEL_COLORS.length]}"/>${label?`<text x="${tx}" y="${y+21}" text-anchor="${anchor}" font-size="14" font-weight="850" fill="#172033" stroke="rgba(255,255,255,.92)" stroke-width="3" paint-order="stroke">${esc(label)}</text>`:""}`;}).join("");
    return `<svg class="analysis-chart-svg analysis-bar-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="棒グラフ">${grid}${rows}</svg>`;
  }
  function pieChartHtml(r){
    const W=760,H=390,cx=205,cy=195,radius=145,total=r.items.reduce((n,x)=>n+x.count,0)||1;let angle=0;
    const slices=r.items.map((x,i)=>{const sweep=x.count/total*360,start=angle,end=angle+sweep,mid=start+sweep/2;angle=end;const label=valueLabel(r,x),small=x.percent<8&&x.count>0,inner=polarPoint(cx,cy,radius*.62,mid),edge=polarPoint(cx,cy,radius*.98,mid),elbow=polarPoint(cx,cy,radius+22,mid),right=elbow.x>=cx,tx=right?Math.min(elbow.x+34,350):Math.max(elbow.x-34,20),anchor=right?"start":"end";const mark=!label||!x.count?"":small?`<polyline points="${edge.x},${edge.y} ${elbow.x},${elbow.y} ${tx+(right?-6:6)},${elbow.y}" fill="none" stroke="#596273" stroke-width="1.5"/><text x="${tx}" y="${elbow.y}" text-anchor="${anchor}" dominant-baseline="middle" font-size="14" font-weight="850" fill="#202124" stroke="white" stroke-width="4" paint-order="stroke">${esc(label)}</text>`:`<text x="${inner.x}" y="${inner.y}" text-anchor="middle" dominant-baseline="middle" font-size="14" font-weight="850" fill="#202124" stroke="white" stroke-width="4" paint-order="stroke">${esc(label)}</text>`;return `<path d="${pieSlicePath(cx,cy,radius,start,end)}" fill="${PASTEL_COLORS[i%PASTEL_COLORS.length]}" stroke="#fff" stroke-width="1"/>${mark}`;}).join("");
    const legend=r.items.map((x,i)=>`<g transform="translate(430 ${52+i*32})"><rect width="13" height="13" rx="3" fill="${PASTEL_COLORS[i%PASTEL_COLORS.length]}"/><text x="22" y="12" font-size="18" font-weight="750" fill="#202124">${esc(text(x.label).slice(0,22))}：${x.count}${r.labelMode==="percent"?`（${x.percent.toFixed(1)}%）`:""}</text></g>`).join("");
    return `<svg class="analysis-chart-svg analysis-pie-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="円グラフ">${slices}${legend}</svg>`;
  }
  window.changeChartType=(i,type)=>{const host=document.querySelector(`#analysisChart${i}`),r=results[i];r.chartType=type;if(host)host.innerHTML=type==="pie"?pieChartHtml(r):barChartHtml(r);};

  function csvForResult(r){
    const lines=[["V番号",r.variable],["設問文",r.question],...(r.type==="NUMERIC"?[["有効",r.valid],["欠測",r.missing],["平均",r.mean],["中央値",r.median],["最小",r.min],["最大",r.max],["表示形式",r.numericMode==="raw"?"値ごとの件数":`階級（${r.binCount}段階、${r.rangeMin}～${r.rangeMax}）`]]:[]),[],[r.type==="NUMERIC"?"階級番号":"加工済の番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"]];
    r.items.forEach(x=>lines.push([x.code,x.label,x.count,`${x.percent.toFixed(1)}%`]));
    return lines.map(row=>row.map(v=>`"${text(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");
  }
  function downloadBlob(blob,name){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
  window.downloadResultCSV=i=>downloadBlob(new Blob(["\ufeff"+csvForResult(results[i])],{type:"text/csv;charset=utf-8"}),`${results[i].variable}_単純集計.csv`);
  window.downloadResultExcel=async i=>{
    const r=results[i], wb=new ExcelJS.Workbook(), ws=wb.addWorksheet(r.variable.slice(0,31));
    ws.addRow(["V番号",r.variable]);ws.addRow(["設問文",r.question]);if(r.type==="NUMERIC"){ws.addRow(["有効",r.valid]);ws.addRow(["欠測",r.missing]);ws.addRow(["平均",r.mean]);ws.addRow(["中央値",r.median]);ws.addRow(["最小",r.min]);ws.addRow(["最大",r.max]);ws.addRow(["表示形式",r.numericMode==="raw"?"値ごとの件数":`階級（${r.binCount}段階、${r.rangeMin}～${r.rangeMax}）`]);}ws.addRow([]);
    ws.addRow([r.type==="NUMERIC"?"階級番号":"加工済の番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"]);
    r.items.forEach(x=>ws.addRow([x.code,x.label,x.count,x.percent/100]));
    ws.getRow(4).font={bold:true,color:{argb:"FFFFFFFF"}};ws.getRow(4).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF673AB7"}};
    ws.getColumn(1).width=18;ws.getColumn(2).width=55;ws.getColumn(3).width=12;ws.getColumn(4).width=12;ws.getColumn(4).numFmt="0.0%";
    ws.getCell("A1").font={bold:true};ws.getCell("A2").font={bold:true};ws.getCell("B2").alignment={wrapText:true,vertical:"top"};
    downloadBlob(new Blob([await wb.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),`${r.variable}_単純集計.xlsx`);
  };
  const PNG_COLORS=PASTEL_COLORS;
  function setupCanvas(width,height){const canvas=document.createElement("canvas");const scale=Math.max(2,window.devicePixelRatio||1);canvas.width=width*scale;canvas.height=height*scale;canvas.style.width=`${width}px`;canvas.style.height=`${height}px`;const ctx=canvas.getContext("2d");ctx.scale(scale,scale);ctx.fillStyle="#ffffff";ctx.fillRect(0,0,width,height);ctx.textBaseline="middle";return{canvas,ctx,width,height};}
  function wrappedLines(ctx,value,maxWidth){const chars=[...text(value)],lines=[];let line="";for(const ch of chars){if(ctx.measureText(line+ch).width>maxWidth&&line){lines.push(line);line=ch;}else line+=ch;}if(line)lines.push(line);return lines.length?lines:[""];}
  function drawHeader(ctx,r,width){ctx.fillStyle="#673ab7";ctx.font="700 18px 'Noto Sans JP',sans-serif";ctx.fillText(`${r.variable} / ${r.type}`,28,28);ctx.fillStyle="#202124";ctx.font="700 17px 'Noto Sans JP',sans-serif";const lines=wrappedLines(ctx,r.question||"設問文なし",width-56);lines.slice(0,3).forEach((line,i)=>ctx.fillText(line,28,62+i*24));const base=62+Math.min(lines.length,3)*24;ctx.fillStyle="#5f6470";ctx.font="13px 'Noto Sans JP',sans-serif";ctx.fillText(statsText(r),28,base+18);ctx.fillStyle="#202124";ctx.font="700 15px 'Noto Sans JP',sans-serif";ctx.fillText(`N=${r.valid}`,28,base+42);return base+66;}
  function tableCanvas(r){const width=1000,rowH=42,headExtra=130,height=headExtra+rowH*(r.items.length+1)+30,{canvas,ctx}=setupCanvas(width,height);let y=drawHeader(ctx,r,width);const cols=[28,120,650,790,970];ctx.fillStyle="#f2edf9";ctx.fillRect(20,y,width-40,rowH);ctx.fillStyle="#202124";ctx.font="700 14px 'Noto Sans JP',sans-serif";["番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"].forEach((v,i)=>ctx.fillText(v,cols[i]+8,y+rowH/2));y+=rowH;ctx.font="14px 'Noto Sans JP',sans-serif";r.items.forEach(item=>{ctx.strokeStyle="#ded8e6";ctx.beginPath();ctx.moveTo(20,y+rowH);ctx.lineTo(width-20,y+rowH);ctx.stroke();ctx.fillStyle="#202124";ctx.fillText(text(item.code),cols[0]+8,y+rowH/2);ctx.fillText(text(item.label).slice(0,52),cols[1]+8,y+rowH/2);ctx.textAlign="right";ctx.fillText(String(item.count),cols[3]-18,y+rowH/2);ctx.fillText(`${item.percent.toFixed(1)}%`,cols[4]-8,y+rowH/2);ctx.textAlign="left";y+=rowH;});return canvas;}
  function chartCanvas(r){
    const width=1000,rowH=48,height=Math.max(460,155+r.items.length*rowH),{canvas,ctx}=setupCanvas(width,height);let y=drawHeader(ctx,r,width);
    if((r.chartType||"bar")==="pie"){
      const total=r.items.reduce((n,x)=>n+x.count,0)||1,cx=290,cy=y+165,radius=135;let start=-Math.PI/2;
      r.items.forEach((item,i)=>{const angle=item.count/total*Math.PI*2;ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,radius,start,start+angle);ctx.closePath();ctx.fillStyle=PNG_COLORS[i%PNG_COLORS.length];ctx.fill();const label=valueLabel(r,item);if(label&&item.count>0){const mid=start+angle/2,small=item.percent<8;ctx.font="700 14px 'Noto Sans JP',sans-serif";ctx.textBaseline="middle";if(small){const ex=cx+Math.cos(mid)*radius,ey=cy+Math.sin(mid)*radius,mx=cx+Math.cos(mid)*(radius+28),my=cy+Math.sin(mid)*(radius+28),right=Math.cos(mid)>=0,tx=mx+(right?34:-34);ctx.strokeStyle="#596273";ctx.lineWidth=2;ctx.beginPath();ctx.moveTo(ex,ey);ctx.lineTo(mx,my);ctx.lineTo(tx+(right?-6:6),my);ctx.stroke();ctx.textAlign=right?"left":"right";ctx.fillStyle="#202124";ctx.fillText(label,tx,my);}else{const lx=cx+Math.cos(mid)*radius*.64,ly=cy+Math.sin(mid)*radius*.64;ctx.textAlign="center";const tw=ctx.measureText(label).width;ctx.fillStyle="rgba(255,255,255,.9)";ctx.fillRect(lx-tw/2-5,ly-10,tw+10,20);ctx.fillStyle="#202124";ctx.fillText(label,lx,ly);}ctx.textAlign="left";}start+=angle;});
      ctx.font="13px 'Noto Sans JP',sans-serif";r.items.forEach((item,i)=>{const ly=y+22+i*29;ctx.fillStyle=PNG_COLORS[i%PNG_COLORS.length];ctx.fillRect(520,ly-6,12,12);ctx.fillStyle="#202124";ctx.fillText(`${item.label}：${item.count}${r.labelMode==="percent"?`（${item.percent.toFixed(1)}%）`:""}`,542,ly);});
    }else{
      const max=Math.max(1,...r.items.map(x=>x.count)),barX=355,barMax=465,labelArea=105;ctx.font="13px 'Noto Sans JP',sans-serif";
      r.items.forEach((item,i)=>{const yy=y+i*rowH;ctx.fillStyle="#202124";ctx.fillText(text(item.label).slice(0,30),28,yy+18);ctx.fillStyle="#eeeaf3";ctx.fillRect(barX,yy+7,barMax+labelArea,22);const bw=item.count?Math.max(4,item.count/max*barMax):0;ctx.fillStyle=PNG_COLORS[i%PNG_COLORS.length];ctx.fillRect(barX,yy+7,bw,22);const vl=valueLabel(r,item);if(vl){ctx.fillStyle="#24303f";ctx.font="700 12px 'Noto Sans JP',sans-serif";ctx.fillText(vl,barX+bw+8,yy+18);}ctx.fillStyle="#202124";ctx.font="13px 'Noto Sans JP',sans-serif";ctx.fillText(String(item.count),940,yy+18);});
    }return canvas;
  }
  function combinedCanvas(r){const table=tableCanvas(r),chart=chartCanvas(r),width=Math.max(table.width,chart.width)/Math.max(2,window.devicePixelRatio||1),tableH=table.height/Math.max(2,window.devicePixelRatio||1),chartH=chart.height/Math.max(2,window.devicePixelRatio||1),out=setupCanvas(width,tableH+chartH+20);out.ctx.drawImage(table,0,0,width,tableH);out.ctx.strokeStyle="#ded8e6";out.ctx.beginPath();out.ctx.moveTo(20,tableH+10);out.ctx.lineTo(width-20,tableH+10);out.ctx.stroke();out.ctx.drawImage(chart,0,tableH+20,width,chartH);return out.canvas;}
  function downloadCanvas(canvas,name){canvas.toBlob(blob=>{if(blob)downloadBlob(blob,name);},"image/png");}
  window.downloadTablePNG=i=>downloadCanvas(tableCanvas(results[i]),`${results[i].variable}_表.png`);
  window.downloadChartPNG=i=>downloadCanvas(chartCanvas(results[i]),`${results[i].variable}_グラフ.png`);
  window.downloadCombinedPNG=i=>downloadCanvas(combinedCanvas(results[i]),`${results[i].variable}_表とグラフ.png`);
  window.downloadCombinedCSV=i=>downloadBlob(new Blob(["\ufeff"+csvForResult(results[i])],{type:"text/csv;charset=utf-8"}),`${results[i].variable}_表とグラフ用データ.csv`);
  window.downloadCombinedExcel=async i=>{
    const r=results[i],wb=new ExcelJS.Workbook(),ws=wb.addWorksheet(safeSheetName(r.variable));
    fillResultSheet(ws,r);
    const canvas=combinedCanvas(r),base64=canvas.toDataURL("image/png"),imageId=wb.addImage({base64,extension:"png"});
    ws.addImage(imageId,{tl:{col:5,row:1},ext:{width:960,height:Math.min(1500,canvas.height*960/canvas.width)}});
    ws.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0};
    downloadBlob(new Blob([await wb.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),`${r.variable}_表とグラフ.xlsx`);
  };

  async function resultWorkbook(){
    const wb=new ExcelJS.Workbook();
    results.forEach(r=>{
      const ws=wb.addWorksheet(r.variable.slice(0,31));
      ws.addRow(["V番号",r.variable]);ws.addRow(["設問文",r.question]);ws.addRow([]);ws.addRow([r.type==="NUMERIC"?"階級番号":"加工済の番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"]);
      r.items.forEach(x=>ws.addRow([x.code,x.label,x.count,x.percent/100]));
      ws.getRow(4).font={bold:true,color:{argb:"FFFFFFFF"}};ws.getRow(4).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF6D28D9"}};
      ws.getColumn(1).width=18;ws.getColumn(2).width=55;ws.getColumn(3).width=12;ws.getColumn(4).width=12;
      ws.getColumn(4).numFmt="0.0%";ws.getCell("A1").font={bold:true};ws.getCell("A2").font={bold:true};ws.getCell("B2").alignment={wrapText:true,vertical:"top"};
    });
    return wb;
  }
  window.downloadAllExcel=async()=>{
    if(!requireResults())return;
    try{
      if(!window.ExcelJS)throw new Error("Excel生成ライブラリを読み込めませんでした。");
      const wb=await resultWorkbook(),buffer=await wb.xlsx.writeBuffer();
      downloadBlob(new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),"簡易分析_全表.xlsx");
    }catch(error){saveError("Excel",error);}
  };
  window.downloadAllCSVs=async()=>{
    if(!requireResults())return;
    try{
      const blocks=results.map(r=>csvForResult(r).trimEnd());
      downloadBlob(new Blob(["\ufeff"+blocks.join("\r\n\r\n")+"\r\n"],{type:"text/csv;charset=utf-8"}),"簡易分析_全表.csv");
    }catch(error){saveError("CSV",error);}
  };
  function svgForResult(r){
    const w=1000,rowH=46,h=125+r.items.length*rowH,max=Math.max(1,...r.items.map(x=>x.count));
    if((r.chartType||"bar")==="pie"){
      const total=r.items.reduce((n,x)=>n+x.count,0)||1,cx=280,cy=240,radius=125;let start=-Math.PI/2;
      const slices=r.items.map((x,i)=>{const angle=x.count/total*Math.PI*2,end=start+angle,x1=cx+Math.cos(start)*radius,y1=cy+Math.sin(start)*radius,x2=cx+Math.cos(end)*radius,y2=cy+Math.sin(end)*radius,large=angle>Math.PI?1:0,path=`M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z`;const mid=start+angle/2,label=valueLabel(r,x),small=x.percent<8&&x.count>0,ex=cx+Math.cos(mid)*radius,ey=cy+Math.sin(mid)*radius,mx=cx+Math.cos(mid)*(radius+30),my=cy+Math.sin(mid)*(radius+30),right=Math.cos(mid)>=0,tx=mx+(right?42:-42),lx=cx+Math.cos(mid)*radius*.64,ly=cy+Math.sin(mid)*radius*.64;start=end;const mark=!label?"":small?`<polyline points="${ex},${ey} ${mx},${my} ${tx+(right?-7:7)},${my}" fill="none" stroke="#596273" stroke-width="2"/><text x="${tx}" y="${my}" text-anchor="${right?'start':'end'}" dominant-baseline="middle" font-size="15" font-weight="700" fill="#202124">${esc(label)}</text>`:`<text x="${lx}" y="${ly}" text-anchor="middle" dominant-baseline="middle" font-size="14" font-weight="700" fill="#202124" stroke="white" stroke-width="4" paint-order="stroke">${esc(label)}</text>`;return `<path d="${path}" fill="${PNG_COLORS[i%PNG_COLORS.length]}" stroke="white"/>${mark}`;}).join("");
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="500" viewBox="0 0 ${w} 500"><rect width="100%" height="100%" fill="white"/><text x="20" y="28" font-size="18" font-weight="bold" fill="#6d28d9">${esc(r.variable)}</text><text x="20" y="56" font-size="15">${esc(r.question)}</text><text x="20" y="80" font-size="15" font-weight="700">N=${r.valid}</text>${slices}</svg>`;
    }
    const bars=r.items.map((x,i)=>{const y=90+i*rowH,bw=460*x.count/max,label=valueLabel(r,x);return `<text x="20" y="${y+15}" font-size="14">${esc(x.label)}</text><rect x="380" y="${y}" width="565" height="22" rx="6" fill="#eef1f5"/><rect x="380" y="${y}" width="${bw}" height="22" rx="6" fill="${PNG_COLORS[i%PNG_COLORS.length]}"/>${label?`<text x="${388+bw}" y="${y+16}" font-size="13" font-weight="700">${esc(label)}</text>`:""}`;}).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="100%" height="100%" fill="white"/><text x="20" y="28" font-size="18" font-weight="bold" fill="#6d28d9">${esc(r.variable)}</text><text x="20" y="56" font-size="15">${esc(r.question)}</text><text x="20" y="80" font-size="15" font-weight="700">N=${r.valid}</text>${bars}</svg>`;
  }
  window.downloadAllCharts=async()=>{const zip=new JSZip();results.forEach(r=>zip.file(`${r.variable}_グラフ.svg`,svgForResult(r)));downloadBlob(await zip.generateAsync({type:"blob"}),"簡易分析_グラフ一式.zip");};
  window.downloadAllReports=async()=>{const zip=new JSZip();const wb=await resultWorkbook();zip.file("簡易分析_全設問.xlsx",await wb.xlsx.writeBuffer());results.forEach(r=>{zip.file(`${r.variable}_単純集計.csv`,"\ufeff"+csvForResult(r));zip.file(`${r.variable}_グラフ.svg`,svgForResult(r));});downloadBlob(await zip.generateAsync({type:"blob"}),"簡易分析_全表とグラフ.zip");};

  window.downloadAllPDF=async()=>{
    if(!requireResults())return;
    try{
      const PDF=window.jspdf?.jsPDF;
      if(!PDF)throw new Error("PDF生成ライブラリを読み込めませんでした。");
      const pdf=new PDF({orientation:"landscape",unit:"mm",format:"a4",compress:true});
      for(let i=0;i<results.length;i++){
        if(i)pdf.addPage("a4","landscape");
        const cv=combinedCanvas(results[i]),img=cv.toDataURL("image/jpeg",0.92),pw=281,ph=194;
        const ratio=Math.min(pw/cv.width,ph/cv.height);
        pdf.addImage(img,"JPEG",8,8,cv.width*ratio,cv.height*ratio,undefined,"FAST");
        await new Promise(resolve=>setTimeout(resolve,0));
      }
      pdf.save("簡易分析_全表とグラフ.pdf");
    }catch(error){saveError("PDF",error);}
  };
  dataInput.addEventListener("change",async e=>{try{const file=e.target.files[0];if(!file)return;setStatus(dataStatus,"warn","読み込み中…");dataFileBuffer=await file.arrayBuffer();dataFileName=file.name;dataModel=parseDataWorkbook(await workbookFromBuffer(dataFileBuffer));setStatus(dataStatus,"ok",`${file.name}：${dataModel.rows.length}行 × ${dataModel.headers.length}列を読み込みました。`);}catch(err){dataModel=null;dataFileBuffer=null;setStatus(dataStatus,"error",err.message);}});
  mapInput.addEventListener("change",async e=>{try{const file=e.target.files[0];if(!file)return;setStatus(mappingStatus,"warn","読み込み中…");mappingFileBuffer=await file.arrayBuffer();mappingFileName=file.name;mappingModel=parseMappingWorkbook(await workbookFromBuffer(mappingFileBuffer));const extra=mappingModel.warnings.length?` 重複コードは規定選択肢を優先して${mappingModel.warnings.length}設問で統合しました。`:"";setStatus(mappingStatus,"ok",`${file.name}：${mappingModel.maps.size}設問を読み込みました。${extra}`);}catch(err){mappingModel=null;mappingFileBuffer=null;setStatus(mappingStatus,"error",err.message);}});
  runButton.addEventListener("click",()=>runAnalysis(savedChartTypes()));
  saveButton.addEventListener("click",saveAnalysisWorkspace);
  resumeButton.addEventListener("click",restoreAnalysisWorkspace);
  deleteButton.addEventListener("click",deleteAnalysisWorkspace);
  refreshSaveStatus();
})();
