const $ = selector => document.querySelector(selector);
const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
const isMissing = value => value === null || value === undefined || String(value).trim() === "" || String(value).trim() === ".";
const colors = ["#6b3fa0","#1f78b4","#2a9d8f","#e9a23b","#c94f7c","#5c6b73","#8e5bb7","#4f8f48","#d66b2c","#3f65a6","#ad4f91","#7d6b55"];
let dataHeaders = [], dataRows = [], mappings = new Map(), analyses = [];

function cellValue(value) {
  if (value && typeof value === "object") {
    if ("result" in value) return value.result ?? "";
    if ("text" in value) return value.text;
    if (Array.isArray(value.richText)) return value.richText.map(item => item.text).join("");
    if (value instanceof Date) return value.toISOString();
  }
  return value ?? "";
}
async function readWorkbook(file) {
  if (!file) throw new Error("ファイルを選択してください。");
  if (file.size > 30 * 1024 * 1024) throw new Error("30MBを超えるファイルは読み込めません。");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return workbook.worksheets.map(sheet => ({
    name: sheet.name,
    rows: sheet.getSheetValues().slice(1).map(row => (row || []).slice(1).map(cellValue))
  }));
}
function variableParts(value) {
  const match = String(value ?? "").trim().match(/^(V\d+)(?:-(\d+))?$/i);
  return match ? { base:match[1].toUpperCase(), branch:match[2] ? Number(match[2]) : null } : null;
}
function parseEmbeddedMapping(rows) {
  const result = new Map();
  for (const row of rows.slice(1)) {
    const variable = variableParts(row[0]);
    if (!variable || variable.base === "V0") continue;
    if (!result.has(variable.base)) result.set(variable.base, {question:"", choices:[]});
    const item = result.get(variable.base);
    if (!item.question && row[1]) item.question = String(row[1]);
    if (row[2]) item.choices.push({code:variable.branch ?? item.choices.length + 1, label:String(row[2])});
  }
  return result;
}
async function loadData(file) {
  const sheets = await readWorkbook(file);
  const sheet = sheets.find(item => item.name === "加工済みデータ") || sheets[0];
  if (!sheet || sheet.rows.length < 2) throw new Error("加工済みデータの見出し行と回答行を確認できません。");
  dataHeaders = sheet.rows[0].map(value => String(value ?? "").trim());
  dataRows = sheet.rows.slice(1);
  const embedded = sheets.find(item => item.name === "対応表");
  if (embedded) {
    const parsed = parseEmbeddedMapping(embedded.rows);
    if (parsed.size) mappings = parsed;
  }
  $("#dataStatus").innerHTML = `<div class="analysis-status success">${escapeHtml(file.name)}：${dataRows.length}回答、${dataHeaders.length}列を読み込みました。</div>`;
}
async function loadMapping(file) {
  const sheets = await readWorkbook(file);
  const embedded = sheets.find(item => item.name === "対応表");
  const parsed = embedded ? parseEmbeddedMapping(embedded.rows) : new Map();
  if (!parsed.size) {
    for (const sheet of sheets) {
      const variable = variableParts(sheet.name);
      if (!variable || variable.base === "V0") continue;
      const existing = mappings.get(variable.base) || {question:"", choices:[]};
      const start = /設問番号|番号/.test(String(sheet.rows[0]?.[0] ?? "")) ? 1 : 0;
      existing.choices = sheet.rows.slice(start).filter(row => !isMissing(row[1])).map((row,index) => ({code:Number(row[0]) || index + 1,label:String(row[1])}));
      parsed.set(variable.base, existing);
    }
  }
  if (!parsed.size) throw new Error("V番号と回答選択肢の対応表を確認できません。");
  for (const [key,value] of parsed) {
    const current = mappings.get(key);
    if (current?.question && !value.question) value.question = current.question;
  }
  mappings = parsed;
  $("#mappingStatus").innerHTML = `<div class="analysis-status success">${escapeHtml(file.name)}：${mappings.size}設問の対応表を読み込みました。</div>`;
}
function buildAnalyses() {
  const groups = new Map();
  dataHeaders.forEach((header,index) => {
    const parts = variableParts(header);
    if (!parts || parts.base === "V0") return;
    if (!groups.has(parts.base)) groups.set(parts.base, []);
    groups.get(parts.base).push({index,branch:parts.branch,header});
  });
  analyses = [...groups.entries()].map(([variable,columns]) => {
    const map = mappings.get(variable) || {question:"設問文未登録",choices:[]};
    const isMulti = columns.some(column => column.branch !== null);
    let items = [], valid = 0, missing = 0;
    if (isMulti) {
      columns.sort((a,b) => a.branch - b.branch);
      items = columns.map(column => {
        const choice = map.choices.find(item => Number(item.code) === column.branch) || map.choices[column.branch - 1];
        return {code:column.branch,label:choice?.label || column.header,count:dataRows.filter(row => Number(row[column.index]) === 1).length};
      });
      for (const row of dataRows) {
        const values = columns.map(column => row[column.index]);
        values.every(isMissing) ? missing++ : valid++;
      }
    } else {
      const column = columns[0], counts = new Map();
      for (const row of dataRows) {
        const value = row[column.index];
        if (isMissing(value)) { missing++; continue; }
        valid++;
        const key = String(value).trim();
        counts.set(key,(counts.get(key) || 0) + 1);
      }
      const choices = map.choices.length ? map.choices : [...counts.keys()].map((label,index) => ({code:index+1,label}));
      items = choices.map(choice => ({code:choice.code,label:choice.label,count:counts.get(String(choice.code)) || counts.get(String(choice.label)) || 0}));
      for (const [value,count] of counts) if (!items.some(item => String(item.code) === value || item.label === value)) items.push({code:value,label:value,count});
    }
    const denominator = isMulti ? valid : items.reduce((sum,item) => sum + item.count,0);
    items.forEach(item => item.percent = denominator ? item.count / denominator * 100 : 0);
    return {variable,question:map.question || "設問文未登録",isMulti,items,valid,missing};
  }).sort((a,b) => Number(a.variable.slice(1)) - Number(b.variable.slice(1)));
}
function cardHtml(analysis,index) {
  const options = analysis.isMulti
    ? '<option value="horizontal">横棒グラフ</option><option value="bar">棒グラフ</option><option value="pie">円グラフ</option><option value="doughnut">ドーナツグラフ</option>'
    : '<option value="pie">円グラフ</option><option value="doughnut">ドーナツグラフ</option><option value="bar">棒グラフ</option><option value="horizontal">横棒グラフ</option>';
  return `<article class="analysis-card"><header class="analysis-card-title"><div><span class="analysis-variable">${analysis.variable}</span><h2>${escapeHtml(analysis.question)}</h2></div><div>有効 ${analysis.valid}件｜欠測 ${analysis.missing}件</div></header><div class="mapping-strip"><strong>番号と選択肢：</strong>${analysis.items.map(item=>`<span>${escapeHtml(item.code)} = ${escapeHtml(item.label)}</span>`).join("")}</div><div class="chart-controls"><label>グラフ<select id="chartType${index}" onchange="redraw(${index})">${options}</select></label><label><input id="showCount${index}" type="checkbox" checked onchange="redraw(${index})">件数</label><label><input id="showPercent${index}" type="checkbox" checked onchange="redraw(${index})">割合</label><label>背景<select id="chartBg${index}" onchange="redraw(${index})"><option value="white">白</option><option value="transparent">透明</option></select></label></div><div class="analysis-grid"><div class="analysis-table-wrap"><table class="analysis-table"><thead><tr><th>番号</th><th>選択肢</th><th>件数</th><th>割合</th></tr></thead><tbody>${analysis.items.map(item=>`<tr><td>${escapeHtml(item.code)}</td><td>${escapeHtml(item.label)}</td><td>${item.count}</td><td>${item.percent.toFixed(1)}%</td></tr>`).join("")}</tbody></table></div><div class="analysis-chart-wrap"><canvas id="chart${index}" width="900" height="520"></canvas></div></div><div class="analysis-downloads"><button class="button" onclick="downloadQuestionExcel(${index})">表をExcelでDL</button><button class="button" onclick="downloadQuestionCSV(${index})">表をCSVでDL</button><button class="button" onclick="downloadQuestionPNG(${index})">グラフをPNGでDL</button><button class="button primary" onclick="downloadQuestionReport(${index})">表＋グラフをPNGでDL</button></div></article>`;
}
function render() {
  if (!dataRows.length) throw new Error("数値化されたデータを読み込んでください。");
  if (!mappings.size) throw new Error("回答選択肢の対応表を読み込んでください。");
  buildAnalyses();
  $("#analysisResults").innerHTML = analyses.map(cardHtml).join("");
  analyses.forEach((_,index) => redraw(index));
  $("#overallDownloads").hidden = false;
  $("#analysisSummary").textContent = `${analyses.length}設問、${dataRows.length}回答を集計しました。`;
}
window.redraw = index => drawChart(index);
function drawChart(index) {
  const analysis=analyses[index],canvas=$("#chart"+index),ctx=canvas.getContext("2d"),type=$("#chartType"+index).value;
  const showCount=$("#showCount"+index).checked,showPercent=$("#showPercent"+index).checked,bg=$("#chartBg"+index).value;
  ctx.clearRect(0,0,canvas.width,canvas.height);if(bg==="white"){ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);}
  ctx.fillStyle="#2c2730";ctx.textAlign="center";ctx.font="700 25px sans-serif";ctx.fillText(`${analysis.variable} ${analysis.question}`,450,36);
  const valueLabel=item=>[showCount?`${item.count}件`:"",showPercent?`${item.percent.toFixed(1)}%`:""].filter(Boolean).join(" ");
  if(type==="pie"||type==="doughnut"){
    const total=Math.max(analysis.items.reduce((sum,item)=>sum+item.count,0),1);let angle=-Math.PI/2;
    analysis.items.forEach((item,i)=>{const slice=item.count/total*Math.PI*2;ctx.beginPath();ctx.moveTo(320,285);ctx.arc(320,285,165,angle,angle+slice);ctx.closePath();ctx.fillStyle=colors[i%colors.length];ctx.fill();angle+=slice;});
    if(type==="doughnut"){ctx.fillStyle=bg==="transparent"?"rgba(0,0,0,0)":"#fff";ctx.globalCompositeOperation="destination-out";ctx.beginPath();ctx.arc(320,285,82,0,Math.PI*2);ctx.fill();ctx.globalCompositeOperation="source-over";}
    ctx.textAlign="left";ctx.font="15px sans-serif";analysis.items.forEach((item,i)=>{ctx.fillStyle=colors[i%colors.length];ctx.fillRect(585,95+i*34,18,18);ctx.fillStyle="#2c2730";ctx.fillText(`${item.label} ${valueLabel(item)}`,612,110+i*34);});
  } else if(type==="horizontal") {
    const max=Math.max(...analysis.items.map(item=>item.count),1),slot=390/Math.max(analysis.items.length,1);
    analysis.items.forEach((item,i)=>{const y=82+i*slot,length=item.count/max*560;ctx.fillStyle=colors[i%colors.length];ctx.fillRect(205,y,length,Math.min(48,slot*.62));ctx.fillStyle="#2c2730";ctx.textAlign="right";ctx.font="14px sans-serif";ctx.fillText(item.label,195,y+24);ctx.textAlign="left";ctx.fillText(valueLabel(item),215+length,y+24);});
  } else {
    const max=Math.max(...analysis.items.map(item=>item.count),1),slot=760/Math.max(analysis.items.length,1);
    analysis.items.forEach((item,i)=>{const height=item.count/max*315,x=70+i*slot,y=420-height;ctx.fillStyle=colors[i%colors.length];ctx.fillRect(x,y,Math.min(67,slot*.65),height);ctx.fillStyle="#2c2730";ctx.textAlign="center";ctx.font="14px sans-serif";ctx.fillText(valueLabel(item),x+30,y-8);ctx.fillText(item.label.slice(0,12),x+30,450);});
  }
}
function rowsFor(analysis){return [["V番号",analysis.variable],["設問文",analysis.question],[],["番号","選択肢","件数","割合"],...analysis.items.map(item=>[item.code,item.label,item.count,item.percent/100]),[],["有効回答",analysis.valid],["欠測",analysis.missing]];}
async function excelBlob(list){const workbook=new ExcelJS.Workbook();for(const analysis of list){const sheet=workbook.addWorksheet(analysis.variable);sheet.addRows(rowsFor(analysis));sheet.getColumn(2).width=45;sheet.getColumn(4).numFmt="0.0%";}return new Blob([await workbook.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});}
function csvText(analysis){return rowsFor(analysis).map(row=>row.map(value=>`"${String(value??"").replace(/"/g,'""')}"`).join(",")).join("\r\n");}
function download(blob,name){const link=document.createElement("a");link.href=URL.createObjectURL(blob);link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1000);}
window.downloadQuestionExcel=async index=>download(await excelBlob([analyses[index]]),`${analyses[index].variable}_簡易分析.xlsx`);
window.downloadQuestionCSV=index=>download(new Blob(["\ufeff"+csvText(analyses[index])],{type:"text/csv;charset=utf-8"}),`${analyses[index].variable}_簡易分析.csv`);
window.downloadQuestionPNG=index=>$("#chart"+index).toBlob(blob=>download(blob,`${analyses[index].variable}_グラフ.png`));
function reportCanvas(index){const analysis=analyses[index],source=$("#chart"+index),canvas=document.createElement("canvas");canvas.width=1050;canvas.height=675+analysis.items.length*34;const ctx=canvas.getContext("2d");ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);ctx.fillStyle="#2c2730";ctx.font="700 25px sans-serif";ctx.fillText(`${analysis.variable} ${analysis.question}`,35,42);analysis.items.forEach((item,i)=>{ctx.fillStyle=i%2?"#f6f2fa":"#fff";ctx.fillRect(35,62+i*34,980,34);ctx.fillStyle="#2c2730";ctx.font="15px sans-serif";ctx.fillText(`${item.code}　${item.label}　${item.count}件　${item.percent.toFixed(1)}%`,50,85+i*34);});ctx.drawImage(source,75,105+analysis.items.length*34,900,520);return canvas;}
window.downloadQuestionReport=index=>reportCanvas(index).toBlob(blob=>download(blob,`${analyses[index].variable}_表とグラフ.png`));
window.downloadAllExcel=async()=>download(await excelBlob(analyses),"簡易分析_全設問.xlsx");
window.downloadAllCSVs=async()=>{const zip=new JSZip();analyses.forEach(item=>zip.file(`${item.variable}_簡易分析.csv`,`\ufeff${csvText(item)}`));download(await zip.generateAsync({type:"blob"}),"簡易分析_CSV一式.zip");};
window.downloadAllCharts=async()=>{const zip=new JSZip();for(let i=0;i<analyses.length;i++){const blob=await new Promise(resolve=>$("#chart"+i).toBlob(resolve));zip.file(`${analyses[i].variable}_グラフ.png`,blob);}download(await zip.generateAsync({type:"blob"}),"簡易分析_グラフ一式.zip");};
window.downloadAllReports=async()=>{const zip=new JSZip();for(let i=0;i<analyses.length;i++){const blob=await new Promise(resolve=>reportCanvas(i).toBlob(resolve));zip.file(`${analyses[i].variable}_表とグラフ.png`,blob);}download(await zip.generateAsync({type:"blob"}),"簡易分析_表とグラフ一式.zip");};
document.addEventListener("DOMContentLoaded",()=>{$("#analysisDataFile").onchange=async event=>{try{await loadData(event.target.files[0]);}catch(error){alert(error.message);}};$("#analysisMappingFile").onchange=async event=>{try{await loadMapping(event.target.files[0]);}catch(error){alert(error.message);}};$("#runAnalysis").onclick=()=>{try{render();}catch(error){alert(error.message);}};});
