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
    return {variable:meta.variable,question:meta.question,type:"NUMERIC",items,valid,missing,invalid:0,unexpected:0,mean,median,min:sorted[0],max:sorted.at(-1)};
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
    results.forEach(r=>r.chartType=chartTypes[r.variable]||r.chartType||"bar");
    renderResults();
    const invalid=results.reduce((n,r)=>n+r.invalid,0), skipped=[...mappingModel.maps.keys()].filter(v=>!results.some(r=>r.variable===v)).length;
    summary.textContent=`${results.length}設問を集計しました。対応表に一致しない値は${invalid}件、データ列がない対応表は${skipped}件です。`;
    downloads.hidden=!results.length;
  }
  function renderResults(){
    if(!results.length){resultsRoot.innerHTML='<div class="analysis-empty">集計対象がありません。</div>';return;}
    resultsRoot.innerHTML=results.map((r,i)=>{
      const max=Math.max(1,...r.items.map(x=>x.count));
      const rows=r.items.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.label)}</td><td>${x.count}</td><td>${x.percent.toFixed(1)}%</td></tr>`).join("");
      const bars=r.items.map(x=>`<div class="analysis-bar-row"><span class="analysis-bar-label">${esc(x.label)}</span><span class="analysis-bar-track"><span class="analysis-bar" style="width:${x.count/max*100}%"></span></span><span class="analysis-bar-value">${x.count}</span></div>`).join("");
      const warnings=[]; if(r.invalid)warnings.push(`対応表にない値 ${r.invalid}件は除外`); if(r.unexpected)warnings.push(`対応表にないMA列 ${r.unexpected}列は除外`);
      const numericMeta=r.type==="NUMERIC"?`<br>平均 ${r.mean.toFixed(1)} ／ 中央値 ${r.median.toFixed(1)} ／ 最小 ${r.min} ／ 最大 ${r.max}`:"";
      const column2=r.type==="NUMERIC"?"値・階級":"回答選択肢";
      return `<article class="analysis-result-card" data-result="${i}"><header class="analysis-result-head"><div><span class="analysis-v-number">${esc(r.variable)} / ${r.type}</span><h2 class="analysis-question">${esc(r.question||"設問文なし")}</h2></div><div class="analysis-meta">有効 ${r.valid}件 ／ 欠測 ${r.missing}件${numericMeta}${warnings.length?`<br><span class="analysis-status-warn">${esc(warnings.join("・"))}</span>`:""}</div></header><div class="analysis-chart-control"><label>グラフ<select onchange="changeChartType(${i},this.value)" ${r.type==="NUMERIC"?"disabled":""}><option value="bar" ${r.chartType!=="pie"?"selected":""}>棒グラフ</option><option value="pie" ${r.chartType==="pie"?"selected":""}>円グラフ</option></select></label></div><div class="analysis-result-grid"><div><table class="analysis-table"><thead><tr><th>番号</th><th>${column2}</th><th>件数</th><th>割合</th></tr></thead><tbody>${rows}</tbody></table></div><div class="analysis-chart-host" id="analysisChart${i}">${r.chartType==="pie"&&r.type!=="NUMERIC"?pieChartHtml(r):barChartHtml(r)}</div></div><div class="analysis-card-actions"><button class="button" onclick="downloadResultCSV(${i})">CSV</button><button class="button" onclick="downloadResultExcel(${i})">Excel</button><button class="button" onclick="downloadTablePNG(${i})">表PNG</button><button class="button" onclick="downloadChartPNG(${i})">グラフPNG</button><button class="button primary" onclick="downloadCombinedPNG(${i})">表＋グラフPNG</button></div></article>`;
    }).join("");
  }
  function barChartHtml(r){
    const max=Math.max(1,...r.items.map(x=>x.count));
    return `<div class="analysis-bar-list">${r.items.map(x=>{
      const width=x.count>0?Math.max(2,x.count/max*100):0;
      return `<div class="analysis-bar-row ${x.count>0?"has-value":"is-zero"}"><span class="analysis-bar-label">${esc(x.label)}</span><span class="analysis-bar-track"><span class="analysis-bar" style="width:${width}%"></span></span><span class="analysis-bar-value">${x.count}</span></div>`;
    }).join("")}</div>`;
  }
  function pieChartHtml(r){
    const total=r.items.reduce((n,x)=>n+x.count,0)||1, colors=["#673ab7","#8559c7","#9d78d3","#b39ddb","#512da8","#7e57c2","#9575cd","#5e35b1","#8e6bbf","#c5b4e3"];
    let acc=0;const stops=r.items.map((x,i)=>{const from=acc,to=acc+x.count/total*100;acc=to;return `${colors[i%colors.length]} ${from}% ${to}%`;}).join(",");
    const legend=r.items.map((x,i)=>`<span><i style="background:${colors[i%colors.length]}"></i>${esc(x.label)}：${x.count}</span>`).join("");
    return `<div class="analysis-pie-wrap"><div class="analysis-pie" style="background:conic-gradient(${stops})"></div><div class="analysis-pie-legend">${legend}</div></div>`;
  }
  window.changeChartType=(i,type)=>{const host=document.querySelector(`#analysisChart${i}`),r=results[i];r.chartType=type;if(host)host.innerHTML=type==="pie"?pieChartHtml(r):barChartHtml(r);};

  function csvForResult(r){
    const lines=[["V番号",r.variable],["設問文",r.question],[],[r.type==="NUMERIC"?"階級番号":"加工済の番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"]];
    r.items.forEach(x=>lines.push([x.code,x.label,x.count,`${x.percent.toFixed(1)}%`]));
    return lines.map(row=>row.map(v=>`"${text(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");
  }
  function downloadBlob(blob,name){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
  window.downloadResultCSV=i=>downloadBlob(new Blob(["\ufeff"+csvForResult(results[i])],{type:"text/csv;charset=utf-8"}),`${results[i].variable}_単純集計.csv`);
  window.downloadResultExcel=async i=>{
    const r=results[i], wb=new ExcelJS.Workbook(), ws=wb.addWorksheet(r.variable.slice(0,31));
    ws.addRow(["V番号",r.variable]);ws.addRow(["設問文",r.question]);ws.addRow([]);
    ws.addRow([r.type==="NUMERIC"?"階級番号":"加工済の番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"]);
    r.items.forEach(x=>ws.addRow([x.code,x.label,x.count,x.percent/100]));
    ws.getRow(4).font={bold:true,color:{argb:"FFFFFFFF"}};ws.getRow(4).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF673AB7"}};
    ws.getColumn(1).width=18;ws.getColumn(2).width=55;ws.getColumn(3).width=12;ws.getColumn(4).width=12;ws.getColumn(4).numFmt="0.0%";
    ws.getCell("A1").font={bold:true};ws.getCell("A2").font={bold:true};ws.getCell("B2").alignment={wrapText:true,vertical:"top"};
    downloadBlob(new Blob([await wb.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),`${r.variable}_単純集計.xlsx`);
  };
  const PNG_COLORS=["#673ab7","#8559c7","#9d78d3","#b39ddb","#512da8","#7e57c2","#9575cd","#5e35b1","#8e6bbf","#c5b4e3"];
  function setupCanvas(width,height){const canvas=document.createElement("canvas");const scale=Math.max(2,window.devicePixelRatio||1);canvas.width=width*scale;canvas.height=height*scale;canvas.style.width=`${width}px`;canvas.style.height=`${height}px`;const ctx=canvas.getContext("2d");ctx.scale(scale,scale);ctx.fillStyle="#ffffff";ctx.fillRect(0,0,width,height);ctx.textBaseline="middle";return{canvas,ctx,width,height};}
  function wrappedLines(ctx,value,maxWidth){const chars=[...text(value)],lines=[];let line="";for(const ch of chars){if(ctx.measureText(line+ch).width>maxWidth&&line){lines.push(line);line=ch;}else line+=ch;}if(line)lines.push(line);return lines.length?lines:[""];}
  function drawHeader(ctx,r,width){ctx.fillStyle="#673ab7";ctx.font="700 18px 'Noto Sans JP',sans-serif";ctx.fillText(`${r.variable} / ${r.type}`,28,28);ctx.fillStyle="#202124";ctx.font="700 17px 'Noto Sans JP',sans-serif";const lines=wrappedLines(ctx,r.question||"設問文なし",width-56);lines.slice(0,3).forEach((line,i)=>ctx.fillText(line,28,62+i*24));return 62+Math.min(lines.length,3)*24+18;}
  function tableCanvas(r){const width=1000,rowH=42,headExtra=130,height=headExtra+rowH*(r.items.length+1)+30,{canvas,ctx}=setupCanvas(width,height);let y=drawHeader(ctx,r,width);const cols=[28,120,650,790,970];ctx.fillStyle="#f2edf9";ctx.fillRect(20,y,width-40,rowH);ctx.fillStyle="#202124";ctx.font="700 14px 'Noto Sans JP',sans-serif";["番号",r.type==="NUMERIC"?"値・階級":"回答選択肢","件数","割合"].forEach((v,i)=>ctx.fillText(v,cols[i]+8,y+rowH/2));y+=rowH;ctx.font="14px 'Noto Sans JP',sans-serif";r.items.forEach(item=>{ctx.strokeStyle="#ded8e6";ctx.beginPath();ctx.moveTo(20,y+rowH);ctx.lineTo(width-20,y+rowH);ctx.stroke();ctx.fillStyle="#202124";ctx.fillText(text(item.code),cols[0]+8,y+rowH/2);ctx.fillText(text(item.label).slice(0,52),cols[1]+8,y+rowH/2);ctx.textAlign="right";ctx.fillText(String(item.count),cols[3]-18,y+rowH/2);ctx.fillText(`${item.percent.toFixed(1)}%`,cols[4]-8,y+rowH/2);ctx.textAlign="left";y+=rowH;});return canvas;}
  function chartCanvas(r){const width=1000,rowH=44,height=Math.max(430,140+r.items.length*rowH),{canvas,ctx}=setupCanvas(width,height);let y=drawHeader(ctx,r,width);if((r.chartType||"bar")==="pie"&&r.type!=="NUMERIC"){const total=r.items.reduce((n,x)=>n+x.count,0)||1,cx=285,cy=y+150,radius=125;let start=-Math.PI/2;r.items.forEach((item,i)=>{const angle=item.count/total*Math.PI*2;ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,radius,start,start+angle);ctx.closePath();ctx.fillStyle=PNG_COLORS[i%PNG_COLORS.length];ctx.fill();start+=angle;});ctx.font="13px 'Noto Sans JP',sans-serif";r.items.forEach((item,i)=>{const ly=y+20+i*28;ctx.fillStyle=PNG_COLORS[i%PNG_COLORS.length];ctx.fillRect(500,ly-6,12,12);ctx.fillStyle="#202124";ctx.fillText(`${item.label}：${item.count}`,522,ly);});}else{const max=Math.max(1,...r.items.map(x=>x.count));ctx.font="13px 'Noto Sans JP',sans-serif";r.items.forEach((item,i)=>{const yy=y+i*rowH;ctx.fillStyle="#202124";ctx.fillText(text(item.label).slice(0,30),28,yy+16);ctx.fillStyle="#eeeaf3";ctx.fillRect(360,yy+5,520,20);const bw=item.count?Math.max(4,item.count/max*520):0;ctx.fillStyle="#673ab7";ctx.fillRect(360,yy+5,bw,20);ctx.fillStyle="#202124";ctx.fillText(String(item.count),900,yy+16);});}return canvas;}
  function combinedCanvas(r){const table=tableCanvas(r),chart=chartCanvas(r),width=Math.max(table.width,chart.width)/Math.max(2,window.devicePixelRatio||1),tableH=table.height/Math.max(2,window.devicePixelRatio||1),chartH=chart.height/Math.max(2,window.devicePixelRatio||1),out=setupCanvas(width,tableH+chartH+20);out.ctx.drawImage(table,0,0,width,tableH);out.ctx.strokeStyle="#ded8e6";out.ctx.beginPath();out.ctx.moveTo(20,tableH+10);out.ctx.lineTo(width-20,tableH+10);out.ctx.stroke();out.ctx.drawImage(chart,0,tableH+20,width,chartH);return out.canvas;}
  function downloadCanvas(canvas,name){canvas.toBlob(blob=>{if(blob)downloadBlob(blob,name);},"image/png");}
  window.downloadTablePNG=i=>downloadCanvas(tableCanvas(results[i]),`${results[i].variable}_表.png`);
  window.downloadChartPNG=i=>downloadCanvas(chartCanvas(results[i]),`${results[i].variable}_グラフ.png`);
  window.downloadCombinedPNG=i=>downloadCanvas(combinedCanvas(results[i]),`${results[i].variable}_表とグラフ.png`);

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
  window.downloadAllExcel=async()=>{const wb=await resultWorkbook();downloadBlob(new Blob([await wb.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),"簡易分析_全設問.xlsx");};
  window.downloadAllCSVs=async()=>{const zip=new JSZip();results.forEach(r=>zip.file(`${r.variable}_単純集計.csv`,"\ufeff"+csvForResult(r)));downloadBlob(await zip.generateAsync({type:"blob"}),"簡易分析_CSV一式.zip");};
  function svgForResult(r){
    const w=960,rowH=42,h=110+r.items.length*rowH,max=Math.max(1,...r.items.map(x=>x.count));
    const bars=r.items.map((x,i)=>{const y=85+i*rowH,bw=520*x.count/max;return `<text x="20" y="${y+14}" font-size="14">${esc(x.label)}</text><rect x="380" y="${y}" width="${bw}" height="20" rx="6" fill="#6d28d9"/><text x="${390+bw}" y="${y+15}" font-size="13">${x.count}</text>`;}).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="100%" height="100%" fill="white"/><text x="20" y="28" font-size="18" font-weight="bold" fill="#6d28d9">${esc(r.variable)}</text><text x="20" y="56" font-size="15">${esc(r.question)}</text>${bars}</svg>`;
  }
  window.downloadAllCharts=async()=>{const zip=new JSZip();results.forEach(r=>zip.file(`${r.variable}_グラフ.svg`,svgForResult(r)));downloadBlob(await zip.generateAsync({type:"blob"}),"簡易分析_グラフ一式.zip");};
  window.downloadAllReports=async()=>{const zip=new JSZip();const wb=await resultWorkbook();zip.file("簡易分析_全設問.xlsx",await wb.xlsx.writeBuffer());results.forEach(r=>{zip.file(`${r.variable}_単純集計.csv`,"\ufeff"+csvForResult(r));zip.file(`${r.variable}_グラフ.svg`,svgForResult(r));});downloadBlob(await zip.generateAsync({type:"blob"}),"簡易分析_全表とグラフ.zip");};

  dataInput.addEventListener("change",async e=>{try{const file=e.target.files[0];if(!file)return;setStatus(dataStatus,"warn","読み込み中…");dataFileBuffer=await file.arrayBuffer();dataFileName=file.name;dataModel=parseDataWorkbook(await workbookFromBuffer(dataFileBuffer));setStatus(dataStatus,"ok",`${file.name}：${dataModel.rows.length}行 × ${dataModel.headers.length}列を読み込みました。`);}catch(err){dataModel=null;dataFileBuffer=null;setStatus(dataStatus,"error",err.message);}});
  mapInput.addEventListener("change",async e=>{try{const file=e.target.files[0];if(!file)return;setStatus(mappingStatus,"warn","読み込み中…");mappingFileBuffer=await file.arrayBuffer();mappingFileName=file.name;mappingModel=parseMappingWorkbook(await workbookFromBuffer(mappingFileBuffer));const extra=mappingModel.warnings.length?` 重複コードは規定選択肢を優先して${mappingModel.warnings.length}設問で統合しました。`:"";setStatus(mappingStatus,"ok",`${file.name}：${mappingModel.maps.size}設問を読み込みました。${extra}`);}catch(err){mappingModel=null;mappingFileBuffer=null;setStatus(mappingStatus,"error",err.message);}});
  runButton.addEventListener("click",()=>runAnalysis(savedChartTypes()));
  saveButton.addEventListener("click",saveAnalysisWorkspace);
  resumeButton.addEventListener("click",restoreAnalysisWorkspace);
  deleteButton.addEventListener("click",deleteAnalysisWorkspace);
  refreshSaveStatus();
})();
