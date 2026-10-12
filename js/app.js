
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const STORAGE_KEY = "survey-coding-final-work-final17-final-flow";
const TYPES = { single: "単一選択（SA）", multi: "複数選択（MA）", text: "自由回答", numeric: "数値", timestamp: "日時" };
let headers=[], rows=[], questions=[], currentStep=1, previewOpen=false, includeId=true;

const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const stripAnswerNumber = value => String(value ?? "").replace(/^\s*\d+\s*[.．]\s*/, "").trim();
const normalize = value => stripAnswerNumber(value).replace(/[\s　。．、,，？?「」『』（）()\[\]]/g, "").toLowerCase();
const unique = (values) => [...new Set(values.map(v => String(v ?? "").trim()).filter(Boolean))];
const splitMulti = (value) => String(value ?? "").split(/[,，]\s*/).map(v => v.trim()).filter(Boolean);
const leadingInteger = (value) => { const match=String(value ?? "").match(/^\s*(\d+)\s*[.．]/); return match ? Number(match[1]) : null; };
const isOtherLabel = (value) => /^その他(?:\s*[:：]|$)/.test(String(value ?? "").trim());
const otherFreeText = (value) => {
  const text=String(value ?? "").trim();
  const match=text.match(/^その他\s*[:：]?\s*(.*)$/);
  return match ? match[1].trim() : text;
};
function officialCount(question){
  const max=Math.max(1,(question.options||[]).length);
  const value=Number(question.officialOptionCount);
  return Number.isFinite(value) ? Math.min(max,Math.max(1,Math.floor(value))) : max;
}
function groupedOtherEnabled(question){return question.method==="multi" && question.groupOutsideAsOther===true;}
function effectiveOptions(question){
  if(!groupedOtherEnabled(question)) return question.options||[];
  const official=(question.options||[]).slice(0,officialCount(question)).map(option=>({...option,isOther:false}));
  return [...official,{label:question.groupedOtherLabel||"その他",code:official.length+1,isOther:true,synthetic:true}];
}
function makeOtherCandidate(value){
  return {value:String(value ?? "").trim(),action:"other",targetCode:"",newLabel:""};
}
function normaliseOtherCandidates(question){
  question.otherCandidates=(question.otherCandidates||[]).map(candidate=>({
    value:String(candidate.value ?? "").trim(),
    action:candidate.action || (candidate.accepted===false ? "missing" : "other"),
    targetCode:String(candidate.targetCode ?? ""),
    newLabel:String(candidate.newLabel ?? "")
  })).filter(candidate=>candidate.value);
}
function refreshOtherCandidates(question){
  const options=effectiveOptions(question);
  const known=new Set(options.filter(option=>!option.isOther).map(option=>normalize(option.label)));
  const current=new Map((question.otherCandidates||[]).map(candidate=>[normalize(candidate.value),candidate]));
  const values=observedOptions(question).filter(value=>{
    const norm=normalize(value);
    return isOtherLabel(value) || !known.has(norm);
  });
  question.otherCandidates=unique(values).map(value=>{
    const existing=current.get(normalize(value));
    return existing ? {...makeOtherCandidate(value),...existing,value} : makeOtherCandidate(value);
  });
  normaliseOtherCandidates(question);
}
const isBlank = (value) => value === null || value === undefined || String(value).trim() === "";

function spreadsheetSafe(value) {
  const text=String(value ?? "");
  return /^[=+\-@]/.test(text) ? `'${text}` : value;
}

function parseCSV(text) {
  const output=[]; let row=[], value="", quoted=false;
  for (let i=0;i<text.length;i++) {
    const char=text[i], next=text[i+1];
    if (char==='"' && quoted && next==='"') { value+='"'; i++; }
    else if (char==='"') quoted=!quoted;
    else if (char===',' && !quoted) { row.push(value); value=""; }
    else if ((char==='\n'||char==='\r') && !quoted) {
      if (char==='\r' && next==='\n') i++;
      row.push(value); output.push(row); row=[]; value="";
    } else value+=char;
  }
  if (value || row.length) { row.push(value); output.push(row); }
  return output;
}

function cellValue(value) {
  if (value && typeof value === "object") {
    if ("result" in value) return value.result ?? "";
    if ("text" in value) return value.text;
    if (Array.isArray(value.richText)) return value.richText.map(x=>x.text).join("");
    if (value instanceof Date) return value.toISOString();
  }
  return value ?? "";
}

async function readSpreadsheet(file) {
  if (!file) throw new Error("ファイルが選択されていません。");
  if (file.size > MAX_FILE_SIZE) throw new Error("20MBを超えるファイルは読み込めません。");
  if (file.name.toLowerCase().endsWith(".csv")) return [parseCSV(await file.text())];
  const workbook=new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return workbook.worksheets.map(sheet => sheet.getSheetValues().slice(1).map(row => (row||[]).slice(1).map(cellValue)));
}

function detectType(header, values) {
  const clean=unique(values);
  if (/タイムスタンプ|timestamp|送信日時|回答日時/i.test(header)) return "timestamp";
  if (clean.some(v=>/[,，]\s*/.test(v))) return "multi";
  if (clean.length && clean.every(v=>v!=="" && !Number.isNaN(Number(v)))) return "numeric";
  if (clean.length>25 || clean.some(v=>v.length>90)) return "text";
  return "single";
}

function observedOptions(question) {
  const values=rows.map(row=>row[question.columnIndex]);
  return question.responseType==="multi" ? unique(values.flatMap(splitMulti)) : unique(values);
}

function makeQuestions() {
  questions=headers.map((header,columnIndex)=>{
    const values=rows.map(row=>row[columnIndex]);
    const responseType=detectType(header,values);
    const labels=responseType==="multi" ? unique(values.flatMap(splitMulti)) : unique(values);
    return { columnIndex, header, include:responseType!=="timestamp", variableName:"", responseType,
      method:responseType==="multi"?"multi":responseType==="single"?"manual":"none",
      options:labels.map((label,index)=>({label,code:leadingInteger(label) ?? index+1,isOther:isOtherLabel(label)})).sort((a,b)=>{const x=leadingInteger(a.label),y=leadingInteger(b.label);if(x===null&&y===null)return 0;if(x===null)return 1;if(y===null)return-1;return x-y;}),
      hasOther:labels.some(isOtherLabel), otherMode:"branch", otherCandidates:[], groupOutsideAsOther:false, officialOptionCount:Math.max(1,labels.length), groupedOtherLabel:"その他", completed:false };
  });
}

async function loadAnswers(file) {
  try {
    const matrix=(await readSpreadsheet(file))[0]||[];
    if (matrix.length<2) throw new Error("見出し行と回答行が必要です。");
    headers=matrix[0].map((v,i)=>String(v||`列${i+1}`));
    rows=matrix.slice(1).filter(row=>row.some(v=>v!==""));
    makeQuestions(); renumber();
    const status=$("#answerStatus");
    if(status) status.innerHTML=`<div class="notice">読み込み完了：${escapeHtml(file.name)}（${rows.length}件・${headers.length}列）</div>`;
    go(2);
  } catch(error) { alert(`回答データを読み込めませんでした。\n${error.message}`); }
}

function renumber(){let number=1;questions.forEach(q=>q.variableName=q.include?`V${number++}`:"");}
function renderVariables(){
  if(!questions.length)return;
  const idToggle=$("#includeId");
  if(idToggle) idToggle.checked=includeId;
  $("#varRows").innerHTML=questions.map((q,i)=>`<tr><td><input type="checkbox" ${q.include?"checked":""} onchange="setInclude(${i},this.checked)"></td><td>${escapeHtml(q.header)}</td><td><input value="${escapeHtml(q.variableName)}" ${q.include?"":"disabled"} onchange="setName(${i},this.value)"></td><td><select ${q.include?"":"disabled"} onchange="setType(${i},this.value)">${Object.entries(TYPES).map(([value,label])=>`<option value="${value}" ${q.responseType===value?"selected":""}>${label}</option>`).join("")}</select></td></tr>`).join("");
  updateMetrics(); updatePreview();
}
function updateMetrics(){const out=buildOutput();const el=$("#metrics");if(el)el.innerHTML=`<div class="metric"><strong>${rows.length}</strong>回答</div><div class="metric"><strong>${headers.length}</strong>元の列</div><div class="metric"><strong>${questions.filter(q=>q.include).length}</strong>使用設問</div><div class="metric"><strong>${out.headers.length}</strong>加工後列</div>`;}
window.setInclude=(i,v)=>{questions[i].include=v;renumber();renderVariables();save(false);};
window.setIncludeId=(value)=>{
  includeId=value;
  updateMetrics();
  updatePreview();
  save(false);
};
window.setName=(i,v)=>{questions[i].variableName=v.trim();updatePreview();save(false);};
window.setType=(i,v)=>{const q=questions[i];q.responseType=v;q.method=v==="multi"?"multi":v==="single"?"manual":"none";q.options=observedOptions(q).map((label,j)=>({label,code:j+1,isOther:isOtherLabel(label)}));q.hasOther=q.options.some(o=>o.isOther);refreshOtherCandidates(q);renderVariables();save(false);};

function processingCard(q){
  return `<article class="card setting-card"><div class="setting-head"><strong>${escapeHtml(q.variableName)}　${escapeHtml(q.header)} ${q.completed?'<span class="complete-mark">✓ 完了</span>':''}</strong><select onchange="setMethod(${q.columnIndex},this.value)"><option value="none" ${q.method==="none"?"selected":""}>1. 加工しない</option><option value="prefix" ${q.method==="prefix"?"selected":""}>2. 先頭の整数を利用</option><option value="manual" ${q.method==="manual"?"selected":""}>3. 回答ごとに番号を手動設定</option><option value="multi" ${q.method==="multi"?"selected":""}>4. 複数選択を選択肢別の有無に分ける</option></select></div>${q.method==="none"?'<p class="help">元の値を残します。</p>':processingDetails(q)}</article>`;
}
function mappedAnswersHtml(q,o){
  const list=(q.otherCandidates||[]).filter(c=>(c.action==="map"&&String(c.targetCode)===String(o.code))||(o.isOther&&c.action==="other"));
  return list.length?`<div class="mapped-answers"><small>編入された自由回答</small>${list.map(c=>`<span>${escapeHtml(otherFreeText(c.value)||c.value)} → ${escapeHtml(o.label)}へ編入</span>`).join("")}</div>`:"";
}
function processingDetails(q){
  refreshOtherCandidates(q);
  return `${q.method==="multi"?multiGroupingSettings(q):""}<div class="notice">${q.method==="multi"?"選択肢ごとに1/0列を作成します。V枝番号は現在の表示順で再生成します。":"抽出した番号は手動修正できます。"}</div>${q.hasOther||q.otherCandidates.length||groupedOtherEnabled(q)?otherSettings(q):""}<div class="toolbar"><button class="button" onclick="downloadMap(${q.columnIndex})">この設問の対応表を保存</button><button class="icon-button reload-map" onclick="reloadQuestion(${q.columnIndex})" title="最新の編集内容を再読み込み" aria-label="最新の編集内容を再読み込み">↻</button><label class="button">対応表を読込<input hidden type="file" accept=".xlsx,.csv" onchange="importMap(${q.columnIndex},this.files[0])"></label></div><div class="option-list" data-question="${q.columnIndex}">${effectiveOptions(q).map((o,j)=>`<div class="option-card"><span class="drag-handle">☰</span><span>${escapeHtml(o.label)}${o.isOther?'<span class="badge">その他</span>':""}</span>${q.method==="multi"?`<strong>${escapeHtml(q.variableName)}-${j+1}</strong>`:`<input value="${escapeHtml(o.code)}" onchange="setCode(${q.columnIndex},${j},this.value)">`}${mappedAnswersHtml(q,o)}</div>`).join("")}</div><div class="question-complete-row"><button class="button primary complete-button" onclick="completeQuestion(${q.columnIndex})">${q.completed?"再編集":"完了"}</button></div>`;
}
function multiGroupingSettings(q){
  const count=officialCount(q), total=(q.options||[]).length;
  return `<section class="multi-group-setting"><div><strong>正式な選択肢の範囲</strong><p class="help">上から何件までを個別の選択肢として出力するか指定できます。範囲外の回答は、すべて1つの「その他」列にまとめます。</p></div><label class="group-toggle"><input type="checkbox" ${groupedOtherEnabled(q)?"checked":""} onchange="setGroupOutside(${q.columnIndex},this.checked)"> 指定範囲外をすべて「その他」にまとめる</label>${groupedOtherEnabled(q)?`<div class="multi-group-controls"><label>正式な選択肢は上から<input type="number" min="1" max="${total}" value="${count}" onchange="setOfficialCount(${q.columnIndex},this.value)">件</label><label>「その他」の表示名<input value="${escapeHtml(q.groupedOtherLabel||"その他")}" onchange="setGroupedOtherLabel(${q.columnIndex},this.value)"></label><div class="group-result">出力: ${escapeHtml(q.variableName)}-1 ～ ${escapeHtml(q.variableName)}-${count} ＋ <strong>${escapeHtml(q.variableName)}-${count+1} ${escapeHtml(q.groupedOtherLabel||"その他")}</strong></div></div>`:""}</section>`;
}
function recodeResultText(q,c){if(c.action==="map"){const t=effectiveOptions(q).find(o=>String(o.code)===String(c.targetCode));return t?`${q.variableName}-${t.code} ${t.label}へ集約`:"振り分け先未選択";}if(c.action==="missing")return "欠測値「.」として処理";if(c.action==="new")return `新規選択肢「${c.newLabel||otherFreeText(c.value)}」として追加`;const o=effectiveOptions(q).find(x=>x.isOther);return o?`${q.variableName}-${o.code} ${o.label}へ集約`:"その他として保持";}
function otherSettings(q){
  const choices=effectiveOptions(q).filter(option=>!option.isOther);
  const rowsHtml=q.otherCandidates.map((candidate,index)=>{const action=candidate.action||"other";const resultText=recodeResultText(q,candidate);return `<div class="other-recode-row"><div class="other-recode-value"><span class="other-recode-label">自由回答</span><strong>${escapeHtml(otherFreeText(candidate.value)||candidate.value)}</strong><small>元の値: ${escapeHtml(candidate.value)}</small><span class="recode-result">${escapeHtml(resultText)}</span></div><div class="other-recode-controls"><label>扱い<select onchange="setCandidateAction(${q.columnIndex},${index},this.value)"><option value="other" ${action==="other"?"selected":""}>「その他」として扱う</option><option value="map" ${action==="map"?"selected":""}>既存選択肢として数える</option><option value="missing" ${action==="missing"?"selected":""}>欠測値「.」として扱う</option><option value="new" ${action==="new"?"selected":""}>新規選択肢として追加</option></select></label>${action==="map"?`<label>振り分け先<select onchange="setCandidateTarget(${q.columnIndex},${index},this.value)"><option value="">選択してください</option>${choices.map(option=>`<option value="${escapeHtml(option.code)}" ${String(candidate.targetCode)===String(option.code)?"selected":""}>${escapeHtml(option.code)}. ${escapeHtml(option.label)}</option>`).join("")}</select></label>`:""}${action==="new"?`<label>新規選択肢名<input value="${escapeHtml(candidate.newLabel||otherFreeText(candidate.value))}" onchange="setCandidateNewLabel(${q.columnIndex},${index},this.value)"></label>`:""}</div></div>`;}).join("");
  return `<section class="other-setting"><div class="other-setting-head"><div><strong>「その他・自由回答」の振り分け</strong><p class="help">既存、新規、その他、欠測値を選び、完了でExcelと対応表へ反映します。</p></div><select onchange="setOtherMode(${q.columnIndex},this.value)"><option value="branch" ${q.otherMode==="branch"?"selected":""}>未振り分けは「その他」として扱う</option><option value="missing" ${q.otherMode==="missing"?"selected":""}>未振り分けは欠測値「.」にする</option></select></div>${rowsHtml?`<div class="other-recode-list">${rowsHtml}</div>`:'<p class="help">現在、振り分け対象の自由回答はありません。</p>'}</section>`;
}
function renderProcessing(){if(!questions.length)return;$("#processing").innerHTML=`<div class="toolbar"><button class="button" onclick="downloadAllMaps()">全設問の対応表を保存</button><label class="button">全設問の対応表を読込<input hidden type="file" accept=".xlsx" onchange="importAllMaps(this.files[0])"></label></div>`+questions.filter(q=>q.include).map(processingCard).join("");initDrag();updatePreview();}
window.setMethod=(i,v)=>{const q=questions[i];q.method=v;if(v==="prefix"){q.options.forEach(o=>{const n=leadingInteger(o.label);o.code=n??".";});q.options.sort((a,b)=>{const an=leadingInteger(a.label),bn=leadingInteger(b.label);if(an===null&&bn===null)return 0;if(an===null)return 1;if(bn===null)return-1;return an-bn;});}renderProcessing();save(false);};
window.setCode=(i,j,v)=>{questions[i].options[j].code=v||".";updatePreview();save(false);};
window.setGroupOutside=(i,v)=>{const q=questions[i];q.groupOutsideAsOther=v;if(v){q.officialOptionCount=Math.min(Number(q.officialOptionCount)||Math.max(1,(q.options||[]).length-1),(q.options||[]).length);q.groupedOtherLabel=q.groupedOtherLabel||"その他";}refreshOtherCandidates(q);renderProcessing();save(false);};
window.setOfficialCount=(i,v)=>{const q=questions[i];q.officialOptionCount=Math.min((q.options||[]).length,Math.max(1,Number(v)||1));refreshOtherCandidates(q);renderProcessing();save(false);};
window.setGroupedOtherLabel=(i,v)=>{questions[i].groupedOtherLabel=v.trim()||"その他";updatePreview();save(false);};
window.setOtherMode=(i,v)=>{questions[i].otherMode=v;renderProcessing();save(false);};
window.setCandidateAction=(i,j,v)=>{const candidate=questions[i].otherCandidates[j];candidate.action=v;if(v!=="map")candidate.targetCode="";renderProcessing();save(false);};
window.setCandidateTarget=(i,j,v)=>{questions[i].otherCandidates[j].targetCode=v;questions[i].completed=false;renderProcessing();save(false);};
window.setCandidateNewLabel=(i,j,v)=>{questions[i].otherCandidates[j].newLabel=v.trim();questions[i].completed=false;save(false);};
function rebuildQuestion(q){
  refreshOtherCandidates(q);
  const base=(q.options||[]).filter(o=>!o.generatedNew);
  const maxCode=base.reduce((m,o)=>Math.max(m,Number(o.code)||0),0);
  const labels=[];
  (q.otherCandidates||[]).filter(c=>c.action==="new").forEach(c=>{const label=(c.newLabel||otherFreeText(c.value)||c.value).trim();if(label&&!labels.some(x=>normalize(x)===normalize(label)))labels.push(label);});
  q.options=[...base,...labels.map((label,k)=>({label,code:maxCode+k+1,isOther:false,generatedNew:true}))];
  if(q.method==="multi") q.options.forEach((o,k)=>o.code=k+1);
  q.completed=true;
}
window.reloadQuestion=i=>{rebuildQuestion(questions[i]);renderProcessing();updatePreview();save(false);};
window.completeQuestion=i=>{const q=questions[i];if(q.completed){q.completed=false;renderProcessing();save(false);return;}rebuildQuestion(q);renderProcessing();updatePreview();save(false);};
window.reloadAll=()=>{questions.filter(q=>q.include).forEach(rebuildQuestion);const out=buildOutput();const summary=$("#summary");if(summary)summary.innerHTML=`加工後は <strong>${out.rows.length}</strong> 行 × <strong>${out.headers.length}</strong> 列です。欠測値は「.」です。`;updatePreview();save(false);};
function initDrag(){$$(".option-list").forEach(list=>list.querySelectorAll(".option-card").forEach(card=>{card.querySelector(".drag-handle").onpointerdown=event=>startDrag(event,list,card);}));}
function startDrag(event,list,card){event.preventDefault();const oldIndex=[...list.children].indexOf(card),placeholder=document.createElement("div");placeholder.className="option-card ghost";placeholder.style.height=`${card.offsetHeight}px`;card.after(placeholder);card.style.cssText=`position:fixed;z-index:500;width:${placeholder.offsetWidth}px;pointer-events:none;opacity:.92`;document.body.append(card);const move=e=>{card.style.left=`${e.clientX+12}px`;card.style.top=`${e.clientY+12}px`;const target=document.elementFromPoint(e.clientX,e.clientY)?.closest?.(".option-card");if(target&&target!==placeholder&&target.parentElement===list){const box=target.getBoundingClientRect();e.clientY<box.top+box.height/2?list.insertBefore(placeholder,target):target.after(placeholder);}};const up=()=>{window.removeEventListener("pointermove",move);window.removeEventListener("pointerup",up);placeholder.replaceWith(card);card.removeAttribute("style");const options=questions[Number(list.dataset.question)].options,[moved]=options.splice(oldIndex,1);options.splice([...list.children].indexOf(card),0,moved);renderProcessing();save(false);};window.addEventListener("pointermove",move);window.addEventListener("pointerup",up);move(event);}

function classify(q,value){
  const options=effectiveOptions(q), selected=new Set(), tokens=splitMulti(value), other=options.find(o=>o.isOther);
  let hasOther=false,invalid=false,blank=isBlank(value);
  for(const token of tokens){
    const known=options.find(o=>!o.isOther&&normalize(o.label)===normalize(token));
    if(known){selected.add(normalize(known.label));continue;}
    const candidate=(q.otherCandidates||[]).find(c=>normalize(c.value)===normalize(token));
    if(candidate?.action==="map"){
      const target=options.find(option=>String(option.code)===String(candidate.targetCode) && !option.isOther);
      if(target){selected.add(normalize(target.label));continue;}
      invalid=true;continue;
    }
    if(candidate?.action==="missing"){invalid=true;continue;}
    if(candidate?.action==="new"){const label=candidate.newLabel||otherFreeText(candidate.value)||candidate.value;const target=options.find(o=>normalize(o.label)===normalize(label));if(target){selected.add(normalize(target.label));continue;}invalid=true;continue;}
    if(other&&(groupedOtherEnabled(q)||candidate?.action==="other"||(!candidate&&q.otherMode==="branch"))){hasOther=true;selected.add(normalize(other.label));continue;}
    invalid=true;
  }
  return{selected,hasOther,invalid,blank};
}
function singleValue(q,raw){
  const found=q.options.find(item=>normalize(item.label)===normalize(raw));
  if(found)return found.code;
  const candidate=(q.otherCandidates||[]).find(item=>normalize(item.value)===normalize(raw));
  if(candidate?.action==="map"){
    const target=effectiveOptions(q).find(option=>String(option.code)===String(candidate.targetCode));
    return target?target.code:".";
  }
  if(candidate?.action==="missing")return ".";
  const other=q.options.find(option=>option.isOther);
  if(other&&(candidate?.action==="other"||q.otherMode==="branch"))return other.code;
  return ".";
}
function buildOutput() {
  const outHeaders=[];
  const metadata=[];

  // IDを付与する場合は、必ず先頭列をV0とする。
  // ID値は回答データの表示順に1、2、3…と連番で付与する。
  if (includeId) {
    outHeaders.push("V0");
    metadata.push({isId:true,q:null,option:null});
  }

  for (const q of questions.filter(question => question.include)) {
    if (q.method === "multi") {
      effectiveOptions(q).forEach((option,index) => {
        outHeaders.push(`${q.variableName}-${option.code}`);
        metadata.push({isId:false,q,option});
      });
    } else {
      outHeaders.push(q.variableName);
      metadata.push({isId:false,q,option:null});
    }
  }

  const outRows=rows.map((row,rowIndex) => {
    const cache=new Map();

    return metadata.map(({isId,q,option}) => {
      if (isId) return rowIndex+1;

      const raw=row[q.columnIndex];

      // SA、自由回答、数値などを問わず、未回答は欠測値「.」に統一する。
      if (!option) {
        if (isBlank(raw)) return ".";

        if (q.method !== "manual" && q.method !== "prefix") {
          return spreadsheetSafe(raw);
        }

        return singleValue(q,raw);
      }

      if (!cache.has(q)) cache.set(q,classify(q,raw));
      const result=cache.get(q);

      // MA全体が空白の場合は、各枝番を0ではなく欠測値「.」にする。
      if (result.blank) return ".";

      // 質問票にない事故値、認めていない自由記述は欠測値にする。
      if (result.invalid) return ".";

      // 「その他」を欠測扱いに設定した場合は全枝番を欠測値にする。
      if (result.hasOther && q.otherMode === "missing") return ".";

      // 有効なMA回答では、選択した枝番を1、選択しなかった枝番を0にする。
      // 例: 選択肢1～4のうち1と2を選択 → [1, 1, 0, 0]
      return result.selected.has(normalize(option.label)) ? 1 : 0;
    });
  });

  return {headers:outHeaders,rows:outRows,metadata};
}
function updatePreview(){if(!questions.length)return;const out=buildOutput();$("#livePreview").innerHTML=`<table><thead><tr>${out.headers.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>${out.rows.slice(0,8).map(row=>`<tr>${row.map(v=>`<td>${escapeHtml(v)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;const info=$("#previewInfo");if(info)info.textContent=`先頭8件　${rows.length}回答／${out.headers.length}列`;}
window.togglePreview=()=>{previewOpen=!previewOpen;$("#previewDock").classList.toggle("collapsed",!previewOpen);};

async function exportWorkbook(sheets,name){const workbook=new ExcelJS.Workbook();for(const [sheetName,data] of sheets){const sheet=workbook.addWorksheet(sheetName.slice(0,31));data.forEach(row=>sheet.addRow(row.map(spreadsheetSafe)));sheet.getRow(1).font={bold:true};}download(new Blob([await workbook.xlsx.writeBuffer()],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),name);}
function mapRows(q){
  const rows=[["設問番号","回答選択肢","元の自由回答","振り分け処理","振り分け先番号","振り分け先選択肢"]];
  effectiveOptions(q).forEach((option,index)=>rows.push([String(option.code??"").trim()===""?index+1:option.code,option.label,"","","",""]));
  (q.otherCandidates||[]).forEach(candidate=>{
    const target=effectiveOptions(q).find(option=>String(option.code)===String(candidate.targetCode));
    const action=candidate.action==="map"?"既存選択肢へ再分類":candidate.action==="missing"?"欠測値":candidate.action==="new"?"新規選択肢として追加":"その他として保持";
    rows.push(["","",candidate.value,action,target?.code||"",target?.label||""]);
  });
  return rows;
}
window.downloadMap=i=>exportWorkbook([[questions[i].variableName,mapRows(questions[i])]],`${questions[i].variableName}_回答選択肢.xlsx`);
window.downloadAllMaps=()=>exportWorkbook(questions.filter(q=>q.include).map(q=>[q.variableName,mapRows(q)]),"全設問_回答選択肢.xlsx");
window.importMap=async(i,file)=>{if(!file)return;try{const data=(await readSpreadsheet(file))[0].slice(1).filter(r=>r[1]!=="").sort((a,b)=>Number(a[0])-Number(b[0]));questions[i].options=data.map((r,j)=>({code:(String(r[0]??"").trim()===""?j+1:String(r[0]).trim()),label:String(r[1]),isOther:isOtherLabel(r[1])}));questions[i].hasOther=questions[i].options.some(o=>o.isOther);refreshOtherCandidates(questions[i]);renderProcessing();save(false);}catch(e){alert(e.message);}};
window.importAllMaps=async file=>{if(!file)return;try{const sheets=await readSpreadsheet(file);questions.filter(q=>q.include).forEach((q,i)=>{const data=(sheets[i]||[]).slice(1).filter(r=>r[1]!=="");if(data.length){q.options=data.sort((a,b)=>Number(a[0])-Number(b[0])).map((r,j)=>({code:(String(r[0]??"").trim()===""?j+1:String(r[0]).trim()),label:String(r[1]),isOther:isOtherLabel(r[1])}));q.hasOther=q.options.some(o=>o.isOther);refreshOtherCandidates(q);}});renderProcessing();save(false);}catch(e){alert(e.message);}};
function download(blob,name){const link=document.createElement("a");link.href=URL.createObjectURL(blob);link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1000);}
function processedMapRows(){const t=[["変数名","元の質問文","加工方法","出力列・コード","加工後の回答選択肢","元の自由回答","振り分け処理"]];if(includeId)t.push(["V0","ID番号","連番付与","V0","1から始まる連番","",""]);questions.filter(q=>q.include).forEach(q=>{if(q.method==="none"){t.push([q.variableName,q.header,"無加工",q.variableName,"元の値をそのまま出力","","無加工"]);return;}effectiveOptions(q).forEach(o=>t.push([q.variableName,q.header,q.method==="multi"?"複数選択を1/0化":"回答値をコード化",q.method==="multi"?`${q.variableName}-${o.code}`:o.code,o.label,"","選択肢"]));(q.otherCandidates||[]).forEach(c=>{const o=c.action==="map"?effectiveOptions(q).find(x=>String(x.code)===String(c.targetCode)):null;t.push([q.variableName,q.header,"自由回答の振り分け",o?.code||".",o?.label||"欠測値",c.value,c.action==="map"?"既存選択肢へ再分類":c.action==="missing"?"欠測値":"その他として保持"]);});});return t;}
window.downloadProcessedMap=()=>exportWorkbook([["加工結果対応表",processedMapRows()]],"加工結果_対応表.xlsx");
window.exportExcel=()=>{const out=buildOutput();return exportWorkbook([["加工済みデータ",[out.headers,...out.rows]],["加工結果対応表",processedMapRows()]],"加工済みアンケートデータ.xlsx");};
window.exportCSV=()=>{const out=buildOutput(),csv=[out.headers,...out.rows].map(row=>row.map(v=>`"${String(spreadsheetSafe(v)??"").replace(/"/g,'""')}"`).join(",")).join("\r\n");download(new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"}),"加工済みデータ_UTF8.csv");};

function save(show){if(!headers.length)return;try{localStorage.setItem(STORAGE_KEY,JSON.stringify({headers,rows,questions,currentStep,includeId}));if(show)alert("一時保存しました。");}catch(e){alert("一時保存できませんでした。データ量が大きすぎる可能性があります。");}}
window.saveNow=()=>save(true); window.clearSave=()=>{localStorage.removeItem(STORAGE_KEY);alert("一時保存を削除しました。");};
window.go=go;function go(step){currentStep=step;$$('.panel').forEach(panel=>panel.classList.toggle('active',Number(panel.dataset.step)===step));$$('.step-nav button').forEach(button=>button.classList.toggle('active',Number(button.dataset.step)===step));if(step===2)renderVariables();if(step===3)renderProcessing();if(step===4){reloadAll();}updatePreview();save(false);}
document.addEventListener("DOMContentLoaded",()=>{$("#answerFile").onchange=e=>e.target.files[0]&&loadAnswers(e.target.files[0]);$$('.step-nav button').forEach(button=>button.onclick=()=>{const step=Number(button.dataset.step);if(step===1||headers.length)go(step);});const saved=localStorage.getItem(STORAGE_KEY);if(saved&&confirm("一時保存された作業を再開しますか？")){try{const state=JSON.parse(saved);headers=state.headers||[];rows=state.rows||[];questions=state.questions||[];questions.forEach(q=>{q.groupOutsideAsOther=q.groupOutsideAsOther===true;q.officialOptionCount=Number(q.officialOptionCount)||Math.max(1,(q.options||[]).length);q.groupedOtherLabel=q.groupedOtherLabel||"その他";normaliseOtherCandidates(q);refreshOtherCandidates(q);});includeId=state.includeId!==false;go(Math.max(2,state.currentStep||2));}catch{localStorage.removeItem(STORAGE_KEY);}}});
