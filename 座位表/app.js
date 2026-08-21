// ===== 教室座位表生成器 =====
// 本機端運作，不需要伺服器；資料以 *.json 保存/加載。

const DESK_W = 112, DESK_H = 72, GAP = 16, GROUP_GAP = 36, CANVAS_PAD = 20, PAIR_GROUP_GAP = 30;
const NEIGHBOR_THRESHOLD = 125;
const GROUP_BORDER = ['#e07ca0','#4f9fe0','#5cb85c','#e0b34f','#9b6fe0','#3fbfa8','#e08a5c','#6f7fe0'];
const GROUP_BG = ['#fdeef2','#eaf4fd','#eefaee','#fdf6e6','#f3ecfd','#e6fbf7','#fdeee6','#eceffd'];
const CHINESE_NUM = ['一','二','三','四','五','六','七','八','九','十','十一','十二','十三','十四','十五'];
const CIRCLED_NUMBERS = Array.from({length:15}, (_,i)=> String.fromCodePoint(0x2460+i)); // ①..⑮
// 已知身分名稱的預設代號（若該身分尚未自訂圖示/代號才會套用）
const KNOWN_ROLE_ICONS = { '教練':'⑤', '助教':'②', '經理':'③', '經理A':'③', '經理B':'④', '球員':'①' };
// 已知的身分名稱：即使還沒被建立成「身分類別」項目，姓名解析／修正時也會當作身分辨識，
// 用來打破「身分要先存在才能拆分姓名、姓名要先拆分身分才能建立」的雞生蛋問題
const KNOWN_ROLE_NAMES = ['教練','助教','經理','經理A','經理B','球員','球員A','球員B'];

function defaultRoles(){
  return [
    {id: uid('r'), name: '教練', max: 1, icon: ''},
    {id: uid('r'), name: '助教', max: 1, icon: ''},
    {id: uid('r'), name: '球員', max: null, icon: ''},
    {id: uid('r'), name: '經理', max: 1, icon: ''}
  ];
}

function chineseNumeral(n){ return CHINESE_NUM[n-1] || String(n); }

function loadCustomScenarios(){
  try{
    const raw = localStorage.getItem('seatingCustomScenarios');
    return raw ? JSON.parse(raw) : [];
  } catch(err){ return []; }
}
function saveCustomScenariosToStorage(){
  try{ localStorage.setItem('seatingCustomScenarios', JSON.stringify(state.customScenarios)); } catch(err){}
}

// 身分設定（名稱／每組上限／圖示代號）存在瀏覽器本機，下次開啟沿用上次的設定
function loadRolesFromStorage(){
  try{
    const raw = localStorage.getItem('seatingRoles');
    if(raw){
      const parsed = JSON.parse(raw);
      if(Array.isArray(parsed) && parsed.length) return parsed;
    }
  } catch(err){}
  return defaultRoles();
}
function saveRolesToStorage(){
  try{ localStorage.setItem('seatingRoles', JSON.stringify(state.roles)); } catch(err){}
}

function loadShowRoleIcons(){
  try{
    const raw = localStorage.getItem('seatingShowRoleIcons');
    return raw===null ? true : raw==='1';
  } catch(err){ return true; }
}
function saveShowRoleIcons(){
  try{ localStorage.setItem('seatingShowRoleIcons', state.showRoleIcons ? '1' : '0'); } catch(err){}
}

let state = {
  students: [],           // {id, number, name, role|null}
  desks: [],              // {id, x, y, studentId|null, locked, groupIndex|null}
  mode: 'group',
  settings: { groupCount: 5, perRow: 6, pairsPerRow: 3 },
  constraints: { apart: [], front: [], groupApart: [] }, // apart/groupApart: [[id,id],...], front: [id,...]
  roles: loadRolesFromStorage(),  // {id, name, max|null, icon}  max=null 代表不限每組人數，icon 為小圖示/代號文字（可留空）；存在瀏覽器本機
  customScenarios: loadCustomScenarios(), // {id, name, desks:[{x,y,groupIndex}]} 使用者自存的座位情境，存在瀏覽器本機
  showRoleIcons: loadShowRoleIcons(),    // 座位表上是否以圖示/代號取代身分文字；存在瀏覽器本機
  swapSourceId: null,
  groupPaintMode: false,  // 是否處於「點選座位設定分組」模式
  activeGroupPaint: 0      // 目前選取要塗上的組別（null 代表清除標記）
};

function uid(prefix){
  return prefix + '_' + Math.random().toString(36).slice(2,9) + Date.now().toString(36).slice(-4);
}

function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

let toastTimer;
function toast(msg){
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=> el.classList.remove('show'), 2600);
}

function studentLabelObj(s){ return `${s.number} ${s.name}`; }
function studentLabel(sid){
  const s = state.students.find(x=>x.id===sid);
  return s ? studentLabelObj(s) : '(未知學生)';
}
function isAssigned(sid){ return state.desks.some(d=>d.studentId===sid); }
function roleDef(name){ return state.roles.find(r=>r.name===name); }
function isCappedRole(name){ const r = roleDef(name); return !!(r && r.max!=null); }
// 依「顯示圖示」開關，回傳身分要顯示的文字：開啟且有設定圖示/代號時顯示圖示，否則顯示原本文字
function roleDisplay(name){
  const r = roleDef(name);
  if(!r) return name;
  if(state.showRoleIcons && r.icon) return r.icon;
  return r.name;
}
// 為尚未自訂圖示的身分，依名稱套用已知的預設代號（只在 icon 是空的時候才填入，不覆蓋使用者自訂內容）
function applyDefaultRoleIcons(){
  let changed = false;
  state.roles.forEach(r=>{
    if(!r.icon && KNOWN_ROLE_ICONS[r.name]){ r.icon = KNOWN_ROLE_ICONS[r.name]; changed = true; }
  });
  return changed;
}
function shuffleArr(arr){
  const a=[...arr];
  for(let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
}

function syncStudentTextarea(){
  document.getElementById('studentInput').value = state.students.slice().sort((a,b)=>a.number-b.number)
    .map(s=>`${s.number} ${s.name}${s.role?(' '+s.role):''}`).join('\n');
}

// 修正「身分是後來才新增/改名，導致之前匯入時身分文字被誤當成姓名一部分」的資料
// 例如姓名存成了「莊宇恩 球員A」、身分卻是預設的「球員」：一旦「球員A」成為合法身分名稱，就把它從姓名搬到身分欄位
function reconcileStudentRoles(){
  const roleNames = new Set([...state.roles.map(r=>r.name), ...KNOWN_ROLE_NAMES]);
  let count = 0;
  state.students.forEach(s=>{
    const tokens = s.name.split(/\s+/).filter(Boolean);
    if(tokens.length>1 && roleNames.has(tokens[tokens.length-1])){
      const role = tokens.pop();
      const newName = tokens.join(' ');
      if(newName){
        s.name = newName;
        s.role = role;
        count++;
      }
    }
  });
  if(count>0) syncStudentTextarea();
  return count;
}

// 確保每個學生目前的身分值，在「身分類別」清單裡都有對應項目；
// 找不到的話自動新增一個（例如貼名單時直接打了清單裡還沒有的身分文字）
function ensureRoleDefinitionsExist(){
  let count = 0;
  const existingNames = new Set(state.roles.map(r=>r.name));
  state.students.forEach(s=>{
    if(s.role && !existingNames.has(s.role)){
      state.roles.push({id: uid('r'), name: s.role, max: null, icon: ''});
      existingNames.add(s.role);
      count++;
    }
  });
  if(count>0) saveRolesToStorage();
  return count;
}

// ---------- 復原 / 重做 ----------
let historyStack = [];
let redoStack = [];
const HISTORY_LIMIT = 50;

function snapshotState(){
  return JSON.stringify({
    students: state.students,
    desks: state.desks,
    mode: state.mode,
    settings: state.settings,
    constraints: state.constraints,
    roles: state.roles
  });
}

function pushHistory(){
  historyStack.push(snapshotState());
  if(historyStack.length>HISTORY_LIMIT) historyStack.shift();
  redoStack = [];
  updateHistoryButtons();
}

function restoreSnapshot(json){
  const data = JSON.parse(json);
  state.students = data.students;
  state.desks = data.desks;
  state.mode = data.mode;
  state.settings = data.settings;
  state.constraints = data.constraints;
  state.roles = data.roles;
  state.swapSourceId = null;
}

function undo(){
  if(historyStack.length===0) return;
  redoStack.push(snapshotState());
  restoreSnapshot(historyStack.pop());
  syncStudentTextarea();
  render();
  updateHistoryButtons();
}

function redo(){
  if(redoStack.length===0) return;
  historyStack.push(snapshotState());
  restoreSnapshot(redoStack.pop());
  syncStudentTextarea();
  render();
  updateHistoryButtons();
}

function updateHistoryButtons(){
  const u = document.getElementById('btnUndo'), r = document.getElementById('btnRedo');
  if(u) u.disabled = historyStack.length===0;
  if(r) r.disabled = redoStack.length===0;
}

// ---------- 學生名單 ----------
function parseStudents(){
  pushHistory();
  const raw = document.getElementById('studentInput').value;
  const lines = raw.split('\n').map(l=>l.trim()).filter(Boolean);
  const roleNames = state.roles.map(r=>r.name);
  const recognizedRoleNames = new Set([...roleNames, ...KNOWN_ROLE_NAMES]);
  const defaultRoleName = roleNames.includes('球員') ? '球員' : null;
  const parsed = [];
  lines.forEach(line=>{
    const numMatch = line.match(/^(\d+)[\s\.\、]*/);
    const number = numMatch ? parseInt(numMatch[1],10) : null;
    const rest = numMatch ? line.slice(numMatch[0].length).trim() : line;
    const tokens = rest.split(/\s+/).filter(Boolean);
    let role = null;
    if(tokens.length>1 && recognizedRoleNames.has(tokens[tokens.length-1])){
      role = tokens.pop();
    }
    const name = tokens.join(' ') || rest;
    parsed.push({number, name, role: role || defaultRoleName});
  });
  const used = new Set();
  parsed.forEach(p=>{ if(p.number!=null) used.add(p.number); });
  let counter=1;
  parsed.forEach(p=>{
    if(p.number==null){
      while(used.has(counter)) counter++;
      p.number = counter; used.add(counter);
    }
  });
  const key = p => `${p.number}|${p.name}`;
  const existingByKey = new Map(state.students.map(s=>[key(s), s]));
  const newStudents = parsed.map(p=>{
    const k = key(p);
    if(existingByKey.has(k)){
      const ex = existingByKey.get(k);
      ex.role = p.role;
      return ex;
    }
    return {id: uid('s'), number: p.number, name: p.name, role: p.role};
  });
  const newIds = new Set(newStudents.map(s=>s.id));
  state.desks.forEach(d=>{ if(d.studentId && !newIds.has(d.studentId)) d.studentId = null; });
  state.constraints.apart = state.constraints.apart.filter(([a,b])=> newIds.has(a) && newIds.has(b));
  state.constraints.front = state.constraints.front.filter(id=> newIds.has(id));
  state.constraints.groupApart = state.constraints.groupApart.filter(([a,b])=> newIds.has(a) && newIds.has(b));
  state.students = newStudents;
  const newRolesCount = ensureRoleDefinitionsExist();
  const prunedRolesCount = pruneUnusedRoles();
  saveRolesToStorage();
  render();
  const msgs = [`已匯入 ${state.students.length} 位學生`];
  if(newRolesCount>0) msgs.push(`自動新增 ${newRolesCount} 個身分類別`);
  if(prunedRolesCount>0) msgs.push(`移除 ${prunedRolesCount} 個此次名單未使用的身分類別`);
  toast(msgs.join('，'));
}

// 匯入新名單時，把「這份名單完全沒人使用」的身分類別移除，讓身分類別清單跟目前名單保持一致
function pruneUnusedRoles(){
  const usedNames = new Set(state.students.map(s=>s.role).filter(Boolean));
  const before = state.roles.length;
  state.roles = state.roles.filter(r=> usedNames.has(r.name));
  return before - state.roles.length;
}

// 手動按鈕：直接依「目前名單」重設身分類別清單（新增缺少的、移除沒人用的），不需要重新貼一次名單
function resetRolesFromStudents(){
  if(state.students.length===0){ toast('目前沒有學生資料，請先匯入學生名單'); return; }
  pushHistory();
  const newRolesCount = ensureRoleDefinitionsExist();
  const prunedRolesCount = pruneUnusedRoles();
  saveRolesToStorage();
  render();
  const msgs = [];
  if(newRolesCount>0) msgs.push(`新增 ${newRolesCount} 個身分類別`);
  if(prunedRolesCount>0) msgs.push(`移除 ${prunedRolesCount} 個未使用的身分類別`);
  toast(msgs.length ? `已依目前名單重設：${msgs.join('、')}` : '身分類別已經跟目前名單一致，沒有變動');
}

// ---------- 佈局產生器 ----------
// 把 count 個座位排進寬度固定為 cols 欄的區塊：每排排滿 cols 個就換行，
// 最後一排若不足 cols 個，會置中對齊（例如5人、2欄 → 2、2、1，最後的1置中）
function layoutZoneRows(positions, count, cols, baseX, baseY, groupIndex){
  const rowWidth = cols*(DESK_W+GAP) - GAP;
  let seat = 0, row = 0;
  while(seat < count){
    const remaining = Math.min(cols, count-seat);
    const usedWidth = remaining*(DESK_W+GAP) - GAP;
    const xOffset = remaining<cols ? (rowWidth-usedWidth)/2 : 0;
    for(let i=0;i<remaining;i++){
      positions.push({x: baseX+xOffset+i*(DESK_W+GAP), y: baseY+row*(DESK_H+GAP), groupIndex});
    }
    seat += remaining;
    row++;
  }
  return row;
}

function genGroup(n, groupCount){
  groupCount = Math.max(1, groupCount|0 || 1);
  const base = Math.floor(n/groupCount), rem = n % groupCount;
  const groupSizes = Array.from({length: groupCount}, (_,g)=> base + (g<rem?1:0));
  const groupsPerRow = 3, cols = 2, maxRowsInGroup = 3;
  const groupBoxW = cols*DESK_W + (cols-1)*GAP;
  const groupBoxH = maxRowsInGroup*DESK_H + (maxRowsInGroup-1)*GAP;
  const positions = [];
  groupSizes.forEach((size, g)=>{
    const gRow = Math.floor(g/groupsPerRow), gCol = g % groupsPerRow;
    const baseX = CANVAS_PAD + gCol*(groupBoxW+GROUP_GAP);
    const baseY = CANVAS_PAD + gRow*(groupBoxH+GROUP_GAP);
    layoutZoneRows(positions, size, cols, baseX, baseY, g);
  });
  return positions;
}

function genRows(n, perRow){
  perRow = Math.max(1, perRow|0 || 1);
  const positions = [];
  for(let i=0;i<n;i++){
    const r = Math.floor(i/perRow), c = i % perRow;
    positions.push({x: CANVAS_PAD + c*(DESK_W+GAP), y: CANVAS_PAD + r*(DESK_H+GAP)});
  }
  return positions;
}

function genPairs(n, pairsPerRow){
  pairsPerRow = Math.max(1, pairsPerRow|0 || 1);
  const colsPerRow = pairsPerRow*2;
  const positions = [];
  for(let i=0;i<n;i++){
    const r = Math.floor(i/colsPerRow), cInRow = i % colsPerRow;
    const pairIndex = Math.floor(cInRow/2), within = cInRow % 2;
    const x = CANVAS_PAD + pairIndex*(2*DESK_W+GAP+PAIR_GROUP_GAP) + within*(DESK_W+GAP);
    const y = CANVAS_PAD + r*(DESK_H+GAP);
    positions.push({x,y});
  }
  return positions;
}

// U 字型（ㄇ字型）：左右各分上下兩區（各2欄），中間只在下方排一區（2欄），中間上方留空對著黑板，五區各自上色
function genUShape(n){
  const zoneCount = 5;
  const base = Math.floor(n/zoneCount), rem = n % zoneCount;
  const sizes = Array.from({length:zoneCount}, (_,z)=> base + (z<rem?1:0));
  const [leftTopN, leftBottomN, rightTopN, rightBottomN, midBottomN] = sizes;

  const sideCols = 2;
  const sideW = sideCols*DESK_W + (sideCols-1)*GAP;
  const midX = CANVAS_PAD + sideW + GROUP_GAP;
  const rightX = midX + sideW + GROUP_GAP;
  const maxTopRows = 3; // 固定上方區塊高度，讓下排（含中間）左右對齊
  const bottomY = CANVAS_PAD + maxTopRows*(DESK_H+GAP) + GROUP_GAP;

  const positions = [];
  layoutZoneRows(positions, leftTopN, sideCols, CANVAS_PAD, CANVAS_PAD, 0);
  layoutZoneRows(positions, rightTopN, sideCols, rightX, CANVAS_PAD, 2);
  layoutZoneRows(positions, leftBottomN, sideCols, CANVAS_PAD, bottomY, 1);
  layoutZoneRows(positions, midBottomN, sideCols, midX, bottomY, 4);
  layoutZoneRows(positions, rightBottomN, sideCols, rightX, bottomY, 3);

  return positions;
}

// 分區排排坐：先用「兩兩配對」排好座位（沿用 genPairs 的幾何排法），
// 再依組數把整排學生依座位順序分成幾個顏色區塊，不改變桌子本身的排法
function genZoneRows(n, zoneCount, pairsPerRow){
  const positions = genPairs(n, pairsPerRow);
  zoneCount = Math.max(1, zoneCount|0 || 1);
  const base = Math.floor(n/zoneCount), rem = n % zoneCount;
  let idx = 0;
  for(let z=0; z<zoneCount; z++){
    const size = base + (z<rem?1:0);
    for(let i=0; i<size && idx<positions.length; i++, idx++){
      positions[idx].groupIndex = z;
    }
  }
  return positions;
}

// 學思達分組：固定每組4人（2欄2列兩兩併坐），並依 U 字型原則排列——
// 左右各分成上下疊放的小組（各2欄），中間只從第二格開始排（也是2欄），中間最上方留空對著黑板
function genXueSiDa(n){
  const groupSize = 4;
  const groupCount = Math.max(1, Math.ceil(n/groupSize));
  const base = Math.floor(n/groupCount), rem = n % groupCount;
  const groupSizes = Array.from({length: groupCount}, (_,g)=> base + (g<rem?1:0));

  const cols = 2, slotRows = 2; // 每組最多4人＝2列，固定slot高度方便同欄對齊
  const colW = cols*DESK_W + (cols-1)*GAP;
  const slotH = slotRows*DESK_H + (slotRows-1)*GAP;
  const midX = CANVAS_PAD + colW + GROUP_GAP;
  const rightX = midX + colW + GROUP_GAP;

  let leftCount, rightCount, midCount;
  if(groupCount < 3){
    midCount = 0;
    leftCount = Math.ceil(groupCount/2);
    rightCount = groupCount - leftCount;
  } else {
    midCount = Math.max(1, Math.round(groupCount/5));
    const sideTotal = groupCount - midCount;
    leftCount = Math.ceil(sideTotal/2);
    rightCount = sideTotal - leftCount;
  }

  const positions = [];
  let gi = 0;
  function placeColumn(count, colX, startSlot){
    for(let i=0;i<count;i++){
      const y = CANVAS_PAD + (startSlot+i)*(slotH+GROUP_GAP);
      layoutZoneRows(positions, groupSizes[gi], cols, colX, y, gi);
      gi++;
    }
  }
  placeColumn(leftCount, CANVAS_PAD, 0);
  placeColumn(rightCount, rightX, 0);
  // 中間對齊左右兩欄「較高那欄」的最底部，讓底部連成一條線（真正的 U 型），上方全部留空
  const maxSideRows = Math.max(leftCount, rightCount);
  placeColumn(midCount, midX, Math.max(0, maxSideRows - midCount));

  return positions;
}

function applyPositions(positions){
  const prevSorted = [...state.desks].sort((a,b)=> a.y-b.y || a.x-b.x);
  const prevAssignedOrder = prevSorted.filter(d=>d.studentId).map(d=>d.studentId);
  state.desks = positions.map(p=>({
    id: uid('d'), x: p.x, y: p.y,
    groupIndex: (p.groupIndex!=null ? p.groupIndex : null),
    studentId: null, locked: false
  }));
  const newSorted = [...state.desks].sort((a,b)=> a.y-b.y || a.x-b.x);
  for(let i=0;i<Math.min(newSorted.length, prevAssignedOrder.length); i++){
    newSorted[i].studentId = prevAssignedOrder[i];
  }
}

function generateLayout(){
  if(state.mode.startsWith('custom:')){
    const sc = state.customScenarios.find(x=> 'custom:'+x.id===state.mode);
    if(!sc){ toast('找不到此自訂座位情境'); return; }
    if(state.desks.length>0){
      if(!confirm('套用已儲存的座位情境將依範本重排所有桌子（鎖定會被清除），是否繼續？')) return;
    }
    pushHistory();
    applyPositions(sc.desks.map(d=>({x:d.x, y:d.y, groupIndex:d.groupIndex})));
    render();
    toast(`已套用座位情境「${sc.name}」`);
    return;
  }
  if(state.students.length===0){ toast('請先在左側匯入學生名單'); return; }
  if(state.mode==='free'){
    if(state.desks.length===0){
      pushHistory();
      applyPositions(genRows(state.students.length, 6));
      toast('已建立初始桌子，之後可自由拖曳、新增或刪除');
    } else {
      toast('自由模式：請用「新增桌子」或直接拖曳桌子上緣調整');
    }
    render();
    return;
  }
  if(state.desks.length>0){
    if(!confirm('重新產生佈局將依新排法重排所有桌子（鎖定會被清除），是否繼續？')) return;
  }
  pushHistory();
  const n = state.students.length;
  let positions;
  if(state.mode==='group') positions = genGroup(n, state.settings.groupCount);
  else if(state.mode==='rows') positions = genRows(n, state.settings.perRow);
  else if(state.mode==='pairs') positions = genPairs(n, state.settings.pairsPerRow);
  else if(state.mode==='ushape') positions = genUShape(n);
  else if(state.mode==='zonerows') positions = genZoneRows(n, state.settings.groupCount, state.settings.pairsPerRow);
  else if(state.mode==='xuesida') positions = genXueSiDa(n);
  applyPositions(positions);
  render();
  toast('已產生新的座位佈局');
}

// 將目前的桌子位置與分組標記，另存為可重複套用的自訂座位情境（存在瀏覽器本機）
function saveCurrentAsScenario(){
  if(state.desks.length===0){ toast('目前沒有桌子可儲存'); return; }
  const name = window.prompt('請輸入座位情境名稱：', '');
  if(name===null) return;
  const trimmed = name.trim();
  if(!trimmed){ toast('名稱不能為空'); return; }
  const desks = state.desks.map(d=>({x: d.x, y: d.y, groupIndex: d.groupIndex}));
  state.customScenarios.push({id: uid('cs'), name: trimmed, desks});
  saveCustomScenariosToStorage();
  render();
  toast(`已儲存座位情境「${trimmed}」`);
}

function deleteCustomScenario(id){
  const sc = state.customScenarios.find(x=>x.id===id);
  if(!sc) return;
  if(!confirm(`確定刪除座位情境「${sc.name}」嗎？`)) return;
  state.customScenarios = state.customScenarios.filter(x=>x.id!==id);
  saveCustomScenariosToStorage();
  if(state.mode==='custom:'+id) state.mode = 'group';
  render();
}

function addDesk(){
  pushHistory();
  const maxY = state.desks.length ? Math.max(...state.desks.map(d=>d.y)) : (CANVAS_PAD - (DESK_H+GAP));
  state.desks.push({id: uid('d'), x: CANVAS_PAD, y: maxY + DESK_H + GAP, studentId:null, locked:false, groupIndex:null});
  render();
}

function removeDesk(id){
  pushHistory();
  state.desks = state.desks.filter(d=>d.id!==id);
  render();
}

// ---------- 指派 / 隨機排列 ----------
function autoFillInOrder(){
  const emptyDesks = state.desks.filter(d=>!d.studentId && !d.locked).sort((a,b)=> a.y-b.y || a.x-b.x);
  const poolStudents = state.students.filter(s=>!isAssigned(s.id)).sort((a,b)=> a.number-b.number);
  const n = Math.min(emptyDesks.length, poolStudents.length);
  if(n===0){ toast('沒有可填入的空位或學生'); return; }
  pushHistory();
  for(let i=0;i<n;i++) emptyDesks[i].studentId = poolStudents[i].id;
  render();
  toast(`已依序填入 ${n} 位學生`);
}

function getNeighborPairs(desks){
  const pairs = [];
  for(let i=0;i<desks.length;i++){
    for(let j=i+1;j<desks.length;j++){
      const dx = desks[i].x - desks[j].x, dy = desks[i].y - desks[j].y;
      if(Math.sqrt(dx*dx+dy*dy) < NEIGHBOR_THRESHOLD) pairs.push([desks[i].id, desks[j].id]);
    }
  }
  return pairs;
}

function isApartPair(a,b){
  return state.constraints.apart.some(([x,y])=> (x===a&&y===b) || (x===b&&y===a));
}

function isNotSameGroupPair(a,b){
  return state.constraints.groupApart.some(([x,y])=> (x===a&&y===b) || (x===b&&y===a));
}

// 在指定的一批桌子裡，把指定的一批學生排進去，盡量避開「不可相鄰」並讓「前排優先」名單坐前面
// 回傳無法避開「不可相鄰」的組數（最佳嘗試後仍有的違規數）
function assignDesksLocal(desks, students, useApart, useFront){
  if(desks.length===0) return 0;
  const neighborPairs = useApart ? getNeighborPairs(desks) : [];
  const frontIds = useFront ? state.constraints.front.filter(id=> students.some(s=>s.id===id)) : [];
  let best = null, bestViolations = Infinity;
  const maxAttempts = 200;
  for(let attempt=0; attempt<maxAttempts && bestViolations>0; attempt++){
    const shuffledFront = shuffleArr(frontIds);
    const others = shuffleArr(students.filter(s=>!frontIds.includes(s.id)).map(s=>s.id));
    const deskSorted = shuffleArr(desks).sort((a,b)=> a.y-b.y || a.x-b.x);
    const frontSlotCount = Math.min(shuffledFront.length, deskSorted.length);
    const frontDesks = deskSorted.slice(0, frontSlotCount);
    const restDesks = shuffleArr(deskSorted.slice(frontSlotCount));
    const assignment = new Map();
    frontDesks.forEach((d,i)=> assignment.set(d.id, shuffledFront[i]));
    restDesks.forEach((d,i)=>{ if(i<others.length) assignment.set(d.id, others[i]); });
    let violations = 0;
    neighborPairs.forEach(([a,b])=>{
      const da = assignment.get(a), db = assignment.get(b);
      if(da && db && isApartPair(da, db)) violations++;
    });
    if(violations < bestViolations){ bestViolations = violations; best = assignment; }
  }
  desks.forEach(d=>{ d.studentId = (best && best.get(d.id)) || null; });
  return bestViolations;
}

// 統一的「自動排列」：依勾選的限制執行。
// - 不可相鄰／前排優先：以座位距離為準，作用在所有要排的桌子上。
// - 身分每組上限／不可同組：以分組標記（groupIndex）為準，只要桌子有分組標記就能套用，
//   跟桌子是用哪種情境產生的無關；沒有任何桌子被標記分組時，這兩項會自動略過。
function smartArrange(options){
  if(state.desks.length===0){ toast('請先產生座位佈局'); return; }
  const unlockedDesks = state.desks.filter(d=>!d.locked);
  if(unlockedDesks.length===0){ toast('所有座位皆已鎖定，無法重新排列'); return; }
  const lockedStudentIds = new Set(state.desks.filter(d=>d.locked && d.studentId).map(d=>d.studentId));
  const assignableStudents = state.students.filter(s=>!lockedStudentIds.has(s.id));

  const groupIndexes = [...new Set(state.desks.filter(d=>d.groupIndex!=null).map(d=>d.groupIndex))].sort((a,b)=>a-b);
  const useGrouping = groupIndexes.length>0 && (options.useRoleCap || options.useGroupApart);

  pushHistory();

  let apartViolationsTotal = 0;
  const capViolationRoles = new Set();
  let groupApartViolations = 0;

  if(useGrouping){
    const groups = groupIndexes.map(gIdx=>{
      const deskInGroup = state.desks.filter(d=>d.groupIndex===gIdx);
      const roleCounts = {};
      const studentIds = [];
      deskInGroup.forEach(d=>{
        if(d.locked && d.studentId){
          studentIds.push(d.studentId);
          const s = state.students.find(x=>x.id===d.studentId);
          if(s && s.role) roleCounts[s.role] = (roleCounts[s.role]||0)+1;
        }
      });
      const unlockedInGroup = deskInGroup.filter(d=>!d.locked);
      const avgY = deskInGroup.length ? deskInGroup.reduce((sum,d)=>sum+d.y,0)/deskInGroup.length : 0;
      return { index: gIdx, desks: unlockedInGroup, capacityLeft: unlockedInGroup.length, roleCounts, studentIds, totalCount: studentIds.length, avgY };
    });

    // 前排優先名單的學生，分組時優先考慮桌子平均位置較靠前（y 較小）的組別，
    // 這樣「前排優先」在有分組時才是真的比較靠近黑板，而不只是「所在組自己的前排」
    const frontIdSet = new Set(options.useFront ? state.constraints.front : []);

    function hasConflict(s, g){
      if(!options.useGroupApart) return false;
      return g.studentIds.some(id2=> isNotSameGroupPair(s.id, id2));
    }
    function place(s, g){
      g.capacityLeft--; g.totalCount++; g.studentIds.push(s.id);
      if(s.role) g.roleCounts[s.role] = (g.roleCounts[s.role]||0)+1;
      s._placed = true;
    }
    function assignList(list, roleMax){
      list.forEach(s=>{
        let candidates = groups.filter(g=> g.capacityLeft>0 && (roleMax==null || (g.roleCounts[s.role]||0) < roleMax) && !hasConflict(s,g));
        if(candidates.length===0) candidates = groups.filter(g=> g.capacityLeft>0 && !hasConflict(s,g));
        if(candidates.length===0) candidates = groups.filter(g=> g.capacityLeft>0);
        if(candidates.length===0) return;
        const isFrontStudent = frontIdSet.has(s.id);
        candidates.sort((a,b)=>
          (isFrontStudent ? (a.avgY-b.avgY) : 0) ||
          (a.roleCounts[s.role]||0)-(b.roleCounts[s.role]||0) || a.totalCount-b.totalCount || a.index-b.index);
        place(s, candidates[0]);
      });
    }

    const pool = shuffleArr(assignableStudents);
    if(options.useRoleCap){
      const cappedRoles = state.roles.filter(r=> r.max!=null);
      cappedRoles.forEach(r=>{
        const list = shuffleArr(pool.filter(s=> s.role===r.name && !s._placed));
        assignList(list, r.max);
      });
    }
    const remaining = shuffleArr(pool.filter(s=>!s._placed));
    assignList(remaining, null);

    const placedIds = new Set(pool.filter(s=>s._placed).map(s=>s.id));
    pool.forEach(s=> delete s._placed);

    // 每組內部，再依「不可相鄰」／「前排優先」（前排＝該組自己桌子裡最靠前的）排定實際座位
    groups.forEach(g=>{
      const students = g.studentIds
        .filter(sid=> !lockedStudentIds.has(sid))
        .map(sid=> state.students.find(s=>s.id===sid))
        .filter(Boolean);
      apartViolationsTotal += assignDesksLocal(g.desks, students, options.useApart, options.useFront);
    });

    if(options.useRoleCap){
      state.roles.filter(r=>r.max!=null).forEach(r=>{
        groups.forEach(g=>{ if((g.roleCounts[r.name]||0) > r.max) capViolationRoles.add(r.name); });
      });
    }
    if(options.useGroupApart){
      state.constraints.groupApart.forEach(([a,b])=>{
        const ga = groups.find(g=> g.studentIds.includes(a));
        const gb = groups.find(g=> g.studentIds.includes(b));
        if(ga && gb && ga.index===gb.index) groupApartViolations++;
      });
    }

    // 沒有分組標記的桌子（若混用），及分組容量放不下的學生，另外排
    const leftoverDesks = unlockedDesks.filter(d=>d.groupIndex==null);
    if(leftoverDesks.length>0){
      const leftoverStudents = assignableStudents.filter(s=> !placedIds.has(s.id));
      apartViolationsTotal += assignDesksLocal(leftoverDesks, leftoverStudents, options.useApart, options.useFront);
    }
  } else {
    apartViolationsTotal = assignDesksLocal(unlockedDesks, assignableStudents, options.useApart, options.useFront);
  }

  render();

  const unplacedCount = assignableStudents.filter(s=> !isAssigned(s.id)).length;
  const problems = [];
  if(options.useApart && apartViolationsTotal>0) problems.push(`${apartViolationsTotal} 組座位無法避開「不可相鄰」限制`);
  if(capViolationRoles.size>0) problems.push(`「${[...capViolationRoles].join('、')}」人數多於已標記組數，部分組別超出每組上限（紅框標示）`);
  if(groupApartViolations>0) problems.push(`${groupApartViolations} 組「不可同組」限制無法避免`);
  if(unplacedCount>0) problems.push(`${unplacedCount} 位學生因座位不足留在未分配`);

  if(problems.length===0){
    renderArrangeResult('✅ 已完成排列，符合所有勾選的限制', 'success');
    toast('已完成自動排列');
  } else {
    renderArrangeResult('⚠ 已完成排列，但以下限制未能完全滿足：<ul>'+problems.map(p=>`<li>${escapeHtml(p)}</li>`).join('')+'</ul>', 'warn');
    toast('已完成排列，但有限制無法完全滿足，詳見下方說明');
  }
}

function renderArrangeResult(html, level){
  const el = document.getElementById('arrangeResult');
  el.innerHTML = html;
  el.className = 'arrange-result' + (level ? ' '+level : '') + (html ? '' : ' empty');
}

function runArrange(){
  const roleCapCb = document.getElementById('chkRoleCap');
  const groupApartCb = document.getElementById('chkGroupApart');
  smartArrange({
    useApart: document.getElementById('chkApart').checked,
    useFront: document.getElementById('chkFront').checked,
    useRoleCap: roleCapCb.checked && !roleCapCb.disabled,
    useGroupApart: groupApartCb.checked && !groupApartCb.disabled
  });
}

// ---------- 點選座位設定分組 ----------
function handleGroupPaintClick(desk){
  pushHistory();
  if(state.activeGroupPaint===null){
    desk.groupIndex = null;
  } else if(desk.groupIndex === state.activeGroupPaint){
    desk.groupIndex = null;
  } else {
    desk.groupIndex = state.activeGroupPaint;
  }
  render();
}

// ---------- 渲染 ----------
function fitCanvasSize(){
  const el = document.getElementById('canvas');
  if(state.desks.length===0){ el.style.width='100%'; el.style.height='260px'; return; }
  const maxX = Math.max(...state.desks.map(d=>d.x+DESK_W)) + CANVAS_PAD;
  const maxY = Math.max(...state.desks.map(d=>d.y+DESK_H)) + CANVAS_PAD;
  el.style.width = maxX + 'px';
  el.style.height = Math.max(maxY, 260) + 'px';
}

const SNAP_THRESHOLD = 14;

// 磁吸對齊：優先吸附到其他桌子的邊緣（同列/同欄），沒有鄰近桌子時則吸附到標準網格
function snapDeskPosition(desk, nx, ny){
  let snapX = null, snapY = null, bestXDiff = SNAP_THRESHOLD, bestYDiff = SNAP_THRESHOLD;
  state.desks.forEach(o=>{
    if(o.id===desk.id) return;
    const dx = Math.abs(o.x-nx);
    if(dx < bestXDiff){ bestXDiff = dx; snapX = o.x; }
    const dy = Math.abs(o.y-ny);
    if(dy < bestYDiff){ bestYDiff = dy; snapY = o.y; }
  });
  if(snapX===null){
    const relX = nx - CANVAS_PAD, stepX = DESK_W+GAP;
    const nearestX = Math.round(relX/stepX)*stepX;
    if(Math.abs(relX-nearestX) <= SNAP_THRESHOLD) snapX = CANVAS_PAD + nearestX;
  }
  if(snapY===null){
    const relY = ny - CANVAS_PAD, stepY = DESK_H+GAP;
    const nearestY = Math.round(relY/stepY)*stepY;
    if(Math.abs(relY-nearestY) <= SNAP_THRESHOLD) snapY = CANVAS_PAD + nearestY;
  }
  return { x: snapX!==null ? snapX : nx, y: snapY!==null ? snapY : ny };
}

function startDeskDrag(e, desk, deskEl){
  e.preventDefault(); e.stopPropagation();
  const preSnapshot = snapshotState();
  const startX = e.clientX, startY = e.clientY;
  const origX = desk.x, origY = desk.y;
  let moved = false;
  function onMove(ev){
    let nx = Math.max(0, origX + (ev.clientX-startX));
    let ny = Math.max(0, origY + (ev.clientY-startY));
    const snapped = snapDeskPosition(desk, nx, ny);
    nx = snapped.x; ny = snapped.y;
    if(nx!==origX || ny!==origY) moved = true;
    deskEl.style.left = nx+'px'; deskEl.style.top = ny+'px';
    desk._tx = nx; desk._ty = ny;
  }
  function onUp(){
    document.removeEventListener('pointermove', onMove);
    document.removeEventListener('pointerup', onUp);
    if(desk._tx!=null){ desk.x = desk._tx; desk.y = desk._ty; delete desk._tx; delete desk._ty; }
    if(moved){
      historyStack.push(preSnapshot);
      if(historyStack.length>HISTORY_LIMIT) historyStack.shift();
      redoStack = [];
      updateHistoryButtons();
    }
    fitCanvasSize();
  }
  document.addEventListener('pointermove', onMove);
  document.addEventListener('pointerup', onUp);
}

function onDeskSlotClick(desk, slotEl){
  if(state.groupPaintMode){ handleGroupPaintClick(desk); return; }
  if(state.swapSourceId){
    if(state.swapSourceId===desk.id){ state.swapSourceId=null; render(); toast('已取消交換'); return; }
    const src = state.desks.find(d=>d.id===state.swapSourceId);
    state.swapSourceId = null;
    if(!src || src.locked || desk.locked){ render(); toast('鎖定的座位無法交換'); return; }
    pushHistory();
    const tmp = src.studentId; src.studentId = desk.studentId; desk.studentId = tmp;
    render();
    toast('已交換座位');
    return;
  }
  openDeskPopup(desk, slotEl);
}

function closePopup(){
  document.getElementById('popup').style.display='none';
  document.getElementById('overlay').style.display='none';
}

function openDeskPopup(desk, anchorEl){
  closePopup();
  const popup = document.getElementById('popup');
  const rect = anchorEl.getBoundingClientRect();
  popup.style.left = (rect.left + window.scrollX) + 'px';
  popup.style.top = (rect.bottom + window.scrollY + 4) + 'px';

  let html = '';
  if(desk.studentId){
    html += `<h4>${escapeHtml(studentLabel(desk.studentId))}</h4>`;
    html += `<div class="popup-action" data-action="unassign">↩ 移回未分配</div>`;
    html += `<div class="popup-action" data-action="lock">${desk.locked?'🔓 解除鎖定':'🔒 鎖定座位'}</div>`;
    html += `<div class="popup-action" data-action="swap">⇄ 與其他座位交換</div>`;
    html += `<hr><div class="popup-action" data-action="cancel">取消</div>`;
  } else {
    html += `<h4>選擇學生入座</h4>`;
    const unassigned = state.students.filter(s=>!isAssigned(s.id)).sort((a,b)=>a.number-b.number);
    if(unassigned.length===0) html += `<div class="hint" style="padding:6px;">沒有未分配的學生</div>`;
    unassigned.forEach(s=>{
      const roleSuffix = s.role ? `（${roleDisplay(s.role)}）` : '';
      html += `<div class="popup-item" data-assign="${s.id}">${escapeHtml(studentLabelObj(s)+roleSuffix)}</div>`;
    });
    html += `<hr><div class="popup-action" data-action="cancel">取消</div>`;
  }
  popup.innerHTML = html;
  popup.style.display = 'block';
  document.getElementById('overlay').style.display = 'block';

  popup.querySelectorAll('[data-assign]').forEach(elm=>{
    elm.addEventListener('click', ()=>{
      pushHistory();
      desk.studentId = elm.getAttribute('data-assign');
      closePopup(); render();
    });
  });
  popup.querySelectorAll('[data-action]').forEach(elm=>{
    elm.addEventListener('click', ()=>{
      const action = elm.getAttribute('data-action');
      if(action==='unassign'){ pushHistory(); desk.studentId=null; closePopup(); render(); }
      else if(action==='lock'){ pushHistory(); desk.locked=!desk.locked; closePopup(); render(); }
      else if(action==='swap'){ state.swapSourceId=desk.id; closePopup(); render(); toast('請點選要交換的座位（再點一次同座位可取消）'); }
      else if(action==='cancel'){ closePopup(); }
    });
  });
}

function renderCanvas(){
  const el = document.getElementById('canvas');
  el.classList.toggle('paint-mode', state.groupPaintMode);
  el.innerHTML = '';

  // 依目前分組標記與學生身分，計算各組每種身分的人數，用於標示超出每組上限的座位
  const groupRoleCounts = {};
  state.desks.forEach(d=>{
    if(d.groupIndex!=null && d.studentId){
      const s = state.students.find(x=>x.id===d.studentId);
      if(s && s.role){
        groupRoleCounts[d.groupIndex] = groupRoleCounts[d.groupIndex] || {};
        groupRoleCounts[d.groupIndex][s.role] = (groupRoleCounts[d.groupIndex][s.role]||0)+1;
      }
    }
  });

  state.desks.forEach(desk=>{
    const student = desk.studentId ? state.students.find(x=>x.id===desk.studentId) : null;
    let overCap = false;
    if(desk.groupIndex!=null && student && student.role && isCappedRole(student.role)){
      const max = roleDef(student.role).max;
      const count = (groupRoleCounts[desk.groupIndex] && groupRoleCounts[desk.groupIndex][student.role]) || 0;
      if(count > max) overCap = true;
    }

    const div = document.createElement('div');
    div.className = 'desk' + (desk.studentId?'':' empty') + (desk.locked?' locked':'') + (state.swapSourceId===desk.id?' swap-source':'') + (overCap?' conflict':'');
    div.style.left = desk.x+'px'; div.style.top = desk.y+'px';
    if(desk.groupIndex!=null){
      const idx = desk.groupIndex % GROUP_BORDER.length;
      div.style.borderLeft = '4px solid ' + GROUP_BORDER[idx];
      if(!desk.studentId) div.style.background = GROUP_BG[idx];
    }

    const handle = document.createElement('div');
    handle.className = 'desk-handle';
    handle.textContent = '⠿⠿⠿';
    handle.title = '拖曳移動桌子';
    handle.addEventListener('pointerdown', e=> startDeskDrag(e, desk, div));

    const slot = document.createElement('div');
    slot.className = 'desk-slot';
    if(desk.studentId){
      const span = document.createElement('span');
      span.textContent = studentLabel(desk.studentId);
      slot.appendChild(span);
      if(student && student.role){
        const tag = document.createElement('div');
        tag.className = 'role-tag' + (overCap?' over-cap':'');
        tag.textContent = roleDisplay(student.role) + (overCap?' ⚠':'');
        tag.title = (state.showRoleIcons && roleDef(student.role) && roleDef(student.role).icon ? student.role+'　' : '') + (overCap ? `此組「${student.role}」人數已超出每組上限` : '');
        slot.appendChild(tag);
      }
      if(!desk.locked){
        slot.draggable = true;
        slot.addEventListener('dragstart', e=>{
          e.dataTransfer.setData('text/plain', JSON.stringify({studentId: desk.studentId, from:'desk', deskId: desk.id}));
        });
      }
      if(desk.locked){
        const badge = document.createElement('div');
        badge.className = 'desk-badges';
        badge.innerHTML = '<span title="已鎖定">🔒</span>';
        slot.appendChild(badge);
      }
    } else {
      const span = document.createElement('span');
      span.className = 'placeholder';
      span.textContent = '空位';
      slot.appendChild(span);
    }
    slot.addEventListener('click', ()=> onDeskSlotClick(desk, slot));
    slot.addEventListener('dragover', e=> e.preventDefault());
    slot.addEventListener('drop', e=>{
      e.preventDefault();
      let data;
      try{ data = JSON.parse(e.dataTransfer.getData('text/plain')||'{}'); } catch(err){ return; }
      if(!data.studentId) return;
      if(desk.locked){ toast('此座位已鎖定'); return; }
      if(data.from==='pool'){
        pushHistory();
        desk.studentId = data.studentId;
        render();
      } else if(data.from==='desk'){
        const srcDesk = state.desks.find(x=>x.id===data.deskId);
        if(srcDesk && !srcDesk.locked){
          pushHistory();
          const tmp = desk.studentId;
          desk.studentId = srcDesk.studentId;
          srcDesk.studentId = tmp;
          render();
        }
      }
    });

    const removeBtn = document.createElement('button');
    removeBtn.className = 'desk-remove';
    removeBtn.title = '刪除桌子';
    removeBtn.textContent = '×';
    removeBtn.addEventListener('click', e=>{ e.stopPropagation(); removeDesk(desk.id); });

    div.appendChild(handle);
    div.appendChild(slot);
    div.appendChild(removeBtn);
    el.appendChild(div);
  });
  fitCanvasSize();
}

function renderPool(){
  const poolEl = document.getElementById('pool');
  poolEl.innerHTML = '';
  const unassigned = state.students.filter(s=>!isAssigned(s.id)).sort((a,b)=>a.number-b.number);
  if(unassigned.length===0){
    poolEl.innerHTML = '<span class="hint">（全部學生皆已入座）</span>';
  }
  unassigned.forEach(s=>{
    const chip = document.createElement('div');
    chip.className = 'pool-chip';
    chip.draggable = true;
    chip.textContent = studentLabelObj(s) + (s.role ? `（${roleDisplay(s.role)}）` : '');
    chip.addEventListener('dragstart', e=>{
      e.dataTransfer.setData('text/plain', JSON.stringify({studentId: s.id, from:'pool'}));
    });
    poolEl.appendChild(chip);
  });
  poolEl.ondragover = e=>{ e.preventDefault(); poolEl.classList.add('dragover'); };
  poolEl.ondragleave = ()=> poolEl.classList.remove('dragover');
  poolEl.ondrop = e=>{
    e.preventDefault(); poolEl.classList.remove('dragover');
    let data;
    try{ data = JSON.parse(e.dataTransfer.getData('text/plain')||'{}'); } catch(err){ return; }
    if(data.from==='desk'){
      const d = state.desks.find(x=>x.id===data.deskId);
      if(d && !d.locked){
        pushHistory();
        d.studentId = null;
        render();
      }
    }
  };
}

function renderModeSettings(){
  document.querySelectorAll('.scenario').forEach(b=> b.classList.toggle('active', b.getAttribute('data-mode')===state.mode));
  const el = document.getElementById('modeSettings');
  let html = '';
  if(state.mode==='group') html = `<label>組數：<input type="number" id="setGroupCount" min="1" max="20" value="${state.settings.groupCount}"></label>`;
  else if(state.mode==='rows') html = `<label>每排座位數：<input type="number" id="setPerRow" min="1" max="20" value="${state.settings.perRow}"></label>`;
  else if(state.mode==='pairs') html = `<label>每排配對數：<input type="number" id="setPairsPerRow" min="1" max="10" value="${state.settings.pairsPerRow}"></label>`;
  else if(state.mode==='zonerows') html = `<label>每排配對數：<input type="number" id="setPairsPerRow" min="1" max="10" value="${state.settings.pairsPerRow}"></label><label>分色組數：<input type="number" id="setGroupCount" min="1" max="20" value="${state.settings.groupCount}"></label>`;
  else if(state.mode==='ushape') html = `<p class="hint">依學生人數自動排成 ㄇ 字型：左右各分上下兩區（各2欄），中間只在下方排一區，中間上方留空對著黑板，各區自動上色。</p>`;
  else if(state.mode==='xuesida') html = `<p class="hint">固定每組4人、兩兩併坐（2欄2列），並依 U 字型原則排列：左右各分上下疊放，中間只排下方、上方留空對著黑板，不需額外設定。</p>`;
  else if(state.mode==='free') html = `<p class="hint">用右側按鈕新增／清空桌子，並拖曳桌子上緣調整位置。</p>`;
  else if(state.mode.startsWith('custom:')){
    const sc = state.customScenarios.find(x=> 'custom:'+x.id===state.mode);
    html = `<p class="hint">已選擇自訂情境「${escapeHtml(sc?sc.name:'')}」，按「產生佈局」套用此範本。</p>`;
  }
  el.innerHTML = html;
  const gc = document.getElementById('setGroupCount');
  if(gc) gc.addEventListener('change', e=>{ state.settings.groupCount = Math.max(1, parseInt(e.target.value)||1); });
  const pr = document.getElementById('setPerRow');
  if(pr) pr.addEventListener('change', e=>{ state.settings.perRow = Math.max(1, parseInt(e.target.value)||1); });
  const ppr = document.getElementById('setPairsPerRow');
  if(ppr) ppr.addEventListener('change', e=>{ state.settings.pairsPerRow = Math.max(1, parseInt(e.target.value)||1); });
}

function renderConstraintSelectors(){
  const opts = state.students.slice().sort((a,b)=>a.number-b.number)
    .map(s=>`<option value="${s.id}">${escapeHtml(studentLabelObj(s))}</option>`).join('');
  ['apartA','apartB','frontSelect','groupApartA','groupApartB'].forEach(id=>{
    const el = document.getElementById(id);
    const prev = el.value;
    el.innerHTML = opts;
    if(prev) el.value = prev;
  });
}

function renderApartList(){
  const el = document.getElementById('apartList');
  el.innerHTML = '';
  state.constraints.apart.forEach((pair, idx)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `${escapeHtml(studentLabel(pair[0]))} ↔ ${escapeHtml(studentLabel(pair[1]))} <button>×</button>`;
    chip.querySelector('button').addEventListener('click', ()=>{ pushHistory(); state.constraints.apart.splice(idx,1); render(); });
    el.appendChild(chip);
  });
}

function renderFrontList(){
  const el = document.getElementById('frontList');
  el.innerHTML = '';
  state.constraints.front.forEach((sid, idx)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `${escapeHtml(studentLabel(sid))} <button>×</button>`;
    chip.querySelector('button').addEventListener('click', ()=>{ pushHistory(); state.constraints.front.splice(idx,1); render(); });
    el.appendChild(chip);
  });
}

function renderGroupApartList(){
  const el = document.getElementById('groupApartList');
  el.innerHTML = '';
  state.constraints.groupApart.forEach((pair, idx)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `${escapeHtml(studentLabel(pair[0]))} ⊘ ${escapeHtml(studentLabel(pair[1]))} <button>×</button>`;
    chip.querySelector('button').addEventListener('click', ()=>{ pushHistory(); state.constraints.groupApart.splice(idx,1); render(); });
    el.appendChild(chip);
  });
}

function renderRolesConfig(){
  const el = document.getElementById('rolesConfig');
  el.innerHTML = '';
  state.roles.forEach(r=>{
    const count = state.students.filter(s=>s.role===r.name).length;
    const groupCount = new Set(state.desks.filter(d=>d.groupIndex!=null).map(d=>d.groupIndex)).size;
    const overPopulated = r.max!=null && groupCount>0 && count > groupCount*r.max;
    const row = document.createElement('div');
    row.className = 'role-row';
    row.innerHTML = `
      <input type="text" class="role-name" value="${escapeHtml(r.name)}">
      <select class="role-icon-pick" title="快速選擇代號">
        <option value="">代號…</option>
        ${CIRCLED_NUMBERS.map(c=>`<option value="${c}"${r.icon===c?' selected':''}>${c}</option>`).join('')}
      </select>
      <input type="text" class="role-icon" maxlength="4" placeholder="圖示/代號" title="小圖示或代號，例如 🧢，也可以用左邊下拉選單挑圈圈數字" value="${escapeHtml(r.icon||'')}">
      <span class="role-max-label">每組上限</span>
      <input type="number" class="role-max" min="0" placeholder="不限" value="${r.max==null?'':r.max}">
      <span class="role-count${overPopulated?' role-count-warn':''}" title="${overPopulated?`目前 ${count} 人，但只標記了 ${groupCount} 組，每組上限 ${r.max}，至少會有座位超出上限`:''}">目前 ${count} 人${overPopulated?' ⚠':''}</span>
      <button class="role-remove" title="刪除身分">×</button>
    `;
    row.querySelector('.role-icon-pick').addEventListener('change', e=>{
      const val = e.target.value;
      if(!val || val===(r.icon||'')) return;
      pushHistory();
      r.icon = val;
      saveRolesToStorage();
      render();
    });
    row.querySelector('.role-icon').addEventListener('change', e=>{
      const newIcon = e.target.value.trim();
      if(newIcon===(r.icon||'')) return;
      pushHistory();
      r.icon = newIcon;
      saveRolesToStorage();
      render();
    });
    row.querySelector('.role-name').addEventListener('change', e=>{
      const newName = e.target.value.trim();
      if(!newName){ e.target.value = r.name; return; }
      if(newName===r.name) return;
      if(state.roles.some(x=>x.id!==r.id && x.name===newName)){
        toast('已有相同名稱的身分'); e.target.value = r.name; return;
      }
      pushHistory();
      const oldName = r.name;
      r.name = newName;
      state.students.forEach(s=>{ if(s.role===oldName) s.role = newName; });
      const fixed = reconcileStudentRoles();
      saveRolesToStorage();
      render();
      if(fixed>0) toast(`已重新命名身分，並自動修正 ${fixed} 位學生姓名中誤含的身分文字`);
    });
    row.querySelector('.role-max').addEventListener('change', e=>{
      const v = e.target.value.trim();
      const newMax = v==='' ? null : Math.max(0, parseInt(v)||0);
      if(newMax===r.max) return;
      pushHistory();
      r.max = newMax;
      saveRolesToStorage();
      render();
    });
    row.querySelector('.role-remove').addEventListener('click', ()=>{
      if(!confirm(`確定要刪除身分「${r.name}」嗎？該身分的學生會變為未分類。`)) return;
      pushHistory();
      state.students.forEach(s=>{ if(s.role===r.name) s.role = null; });
      state.roles = state.roles.filter(x=>x.id!==r.id);
      saveRolesToStorage();
      render();
    });
    el.appendChild(row);
  });
}

function addRole(){
  pushHistory();
  let base = '新身分', name = base, i = 1;
  while(state.roles.some(r=>r.name===name)){ i++; name = base+i; }
  state.roles.push({id: uid('r'), name, max: null, icon: ''});
  const fixed = reconcileStudentRoles();
  saveRolesToStorage();
  render();
  if(fixed>0) toast(`已新增身分，並自動修正 ${fixed} 位學生姓名中誤含的身分文字`);
}

function renderStudentRoleList(){
  const el = document.getElementById('studentRoleList');
  el.innerHTML = '';
  if(state.students.length===0){ el.innerHTML = '<span class="hint">尚無學生，請先匯入名單</span>'; return; }
  state.students.slice().sort((a,b)=>a.number-b.number).forEach(s=>{
    const row = document.createElement('div');
    row.className = 'student-role-row';
    const nameSpan = document.createElement('span');
    nameSpan.className = 'srl-name';
    nameSpan.textContent = studentLabelObj(s);
    const select = document.createElement('select');
    select.className = 'srl-select';
    const noneOpt = document.createElement('option');
    noneOpt.value = ''; noneOpt.textContent = '未分類';
    if(!s.role) noneOpt.selected = true;
    select.appendChild(noneOpt);
    state.roles.forEach(r=>{
      const opt = document.createElement('option');
      opt.value = r.name; opt.textContent = r.name;
      if(s.role===r.name) opt.selected = true;
      select.appendChild(opt);
    });
    select.addEventListener('change', ()=>{ pushHistory(); s.role = select.value || null; render(); });
    row.appendChild(nameSpan); row.appendChild(select);
    el.appendChild(row);
  });
}

function renderGroupPaintBar(){
  const bar = document.getElementById('groupPaintBar');
  const toggleBtn = document.getElementById('btnToggleGroupPaint');
  toggleBtn.classList.toggle('active', state.groupPaintMode);
  toggleBtn.textContent = state.groupPaintMode ? '結束分組設定' : '設定分組（點選座位）';
  bar.style.display = state.groupPaintMode ? 'flex' : 'none';
  bar.innerHTML = '';
  if(!state.groupPaintMode) return;

  const countLabel = document.createElement('label');
  countLabel.className = 'paint-count-label';
  countLabel.innerHTML = `組數：<input type="number" id="paintGroupCount" min="1" max="20" value="${state.settings.groupCount}">`;
  bar.appendChild(countLabel);

  for(let i=0;i<state.settings.groupCount;i++){
    const btn = document.createElement('button');
    btn.className = 'paint-swatch' + (state.activeGroupPaint===i ? ' active' : '');
    btn.style.background = GROUP_BORDER[i % GROUP_BORDER.length];
    btn.textContent = '組' + (i+1);
    btn.addEventListener('click', ()=>{ state.activeGroupPaint = i; renderGroupPaintBar(); });
    bar.appendChild(btn);
  }

  const eraseBtn = document.createElement('button');
  eraseBtn.className = 'paint-swatch erase' + (state.activeGroupPaint===null ? ' active' : '');
  eraseBtn.textContent = '清除標記';
  eraseBtn.addEventListener('click', ()=>{ state.activeGroupPaint = null; renderGroupPaintBar(); });
  bar.appendChild(eraseBtn);

  const clearAllBtn = document.createElement('button');
  clearAllBtn.className = 'btn small danger-outline';
  clearAllBtn.textContent = '清除全部分組標記';
  clearAllBtn.addEventListener('click', ()=>{
    if(!confirm('確定清除所有座位的分組標記？')) return;
    pushHistory();
    state.desks.forEach(d=> d.groupIndex = null);
    render();
  });
  bar.appendChild(clearAllBtn);

  document.getElementById('paintGroupCount').addEventListener('change', e=>{
    state.settings.groupCount = Math.max(1, Math.min(20, parseInt(e.target.value)||1));
    renderGroupPaintBar();
  });
}

function renderCustomScenarios(){
  const el = document.getElementById('customScenarioList');
  el.innerHTML = '';
  state.customScenarios.forEach(sc=>{
    const wrap = document.createElement('div');
    wrap.className = 'custom-scenario-chip';
    const btn = document.createElement('button');
    btn.className = 'btn custom-scenario-btn' + (state.mode==='custom:'+sc.id ? ' active' : '');
    btn.textContent = sc.name;
    btn.title = '點選後按「產生佈局」套用';
    btn.addEventListener('click', ()=>{
      state.mode = 'custom:'+sc.id;
      render();
    });
    const del = document.createElement('button');
    del.className = 'custom-scenario-del';
    del.textContent = '×';
    del.title = '刪除此情境';
    del.addEventListener('click', e=>{ e.stopPropagation(); deleteCustomScenario(sc.id); });
    wrap.appendChild(btn);
    wrap.appendChild(del);
    el.appendChild(wrap);
  });
}

function renderGroupCountInfo(){
  const groupCount = new Set(state.desks.filter(d=>d.groupIndex!=null).map(d=>d.groupIndex)).size;
  document.getElementById('groupCountInfo').textContent = `目前已標記 ${groupCount} 組`;
  updateArrangeCheckboxAvailability(groupCount>0);
}

function updateArrangeCheckboxAvailability(hasGroups){
  ['chkRoleCap','chkGroupApart'].forEach(id=>{
    const cb = document.getElementById(id);
    cb.disabled = !hasGroups;
    const label = cb.closest('label');
    if(label) label.classList.toggle('disabled', !hasGroups);
  });
}

function render(){
  reconcileStudentRoles();
  ensureRoleDefinitionsExist();
  applyDefaultRoleIcons();
  document.getElementById('studentCount').textContent = state.students.length + ' 位學生';
  document.getElementById('chkShowRoleIcons').checked = state.showRoleIcons;
  renderModeSettings();
  renderCustomScenarios();
  renderConstraintSelectors();
  renderApartList();
  renderFrontList();
  renderGroupApartList();
  renderRolesConfig();
  renderStudentRoleList();
  renderGroupCountInfo();
  renderGroupPaintBar();
  renderCanvas();
  renderPool();
  updateHistoryButtons();
}

// ---------- 保存 / 加載 / 預覽 / 重置 ----------
function dateStamp(){
  const d = new Date();
  const pad2 = n => String(n).padStart(2,'0');
  return pad2(d.getMonth()+1) + pad2(d.getDate());
}

function saveJSON(){
  const data = {
    version: 1,
    title: document.getElementById('roomTitle').value,
    students: state.students,
    desks: state.desks,
    mode: state.mode,
    settings: state.settings,
    constraints: state.constraints,
    roles: state.roles
  };
  const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (data.title ? data.title : '座位表') + dateStamp() + '.json';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  toast('已保存座位表');
}

function csvEscape(v){
  const s = String(v==null ? '' : v);
  if(/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

// 匯出組別名單：座號／姓名／組別（格式為「組別中文數字-身分代號」，例如「五-⑤」）
function exportGroupRoster(){
  if(state.students.length===0){ toast('目前沒有學生資料'); return; }
  const rows = [['座號','姓名','組別']];
  state.students.slice().sort((a,b)=>a.number-b.number).forEach(s=>{
    const desk = state.desks.find(d=>d.studentId===s.id);
    let groupLabel = '';
    if(desk && desk.groupIndex!=null){
      const r = s.role ? roleDef(s.role) : null;
      const code = r && r.icon ? r.icon : '';
      groupLabel = chineseNumeral(desk.groupIndex+1) + (code ? ('-'+code) : '');
    }
    rows.push([s.number, s.name, groupLabel]);
  });
  const csv = rows.map(r=> r.map(csvEscape).join(',')).join('\r\n');
  const bom = String.fromCharCode(0xFEFF);
  const blob = new Blob([bom+csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (document.getElementById('roomTitle').value || '座位表') + '_組別名單' + dateStamp() + '.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  toast('已匯出組別名單');
}

function loadJSONFile(file){
  const reader = new FileReader();
  reader.onload = e=>{
    try{
      const data = JSON.parse(e.target.result);
      if(!Array.isArray(data.students) || !Array.isArray(data.desks)) throw new Error('bad format');
      pushHistory();
      state.students = data.students;
      state.desks = data.desks;
      state.mode = data.mode || 'group';
      state.settings = Object.assign({groupCount:5, perRow:6, pairsPerRow:3}, data.settings||{});
      state.constraints = Object.assign({apart:[], front:[], groupApart:[]}, data.constraints||{});
      state.roles = (Array.isArray(data.roles) && data.roles.length) ? data.roles : defaultRoles();
      state.swapSourceId = null;
      state.groupPaintMode = false;
      state.activeGroupPaint = 0;
      document.getElementById('roomTitle').value = data.title || '';
      syncStudentTextarea();
      saveRolesToStorage();
      render();
      toast('已加載座位表');
    } catch(err){
      toast('加載失敗：檔案格式不正確');
    }
  };
  reader.readAsText(file, 'utf-8');
}

let previewMirrored = false; // 老師視角：整個佈局旋轉180°（黑板移到下方，左右也相反）

function renderPreviewDesks(){
  const pc = document.getElementById('previewCanvas');
  pc.innerHTML = '';
  let maxX = 0, maxY = 0;
  state.desks.forEach(d=>{
    maxX = Math.max(maxX, d.x+DESK_W);
    maxY = Math.max(maxY, d.y+DESK_H);
  });
  state.desks.forEach(d=>{
    const div = document.createElement('div');
    div.className = 'preview-desk';
    const x = previewMirrored ? (maxX - d.x - DESK_W) : d.x;
    const y = previewMirrored ? (maxY - d.y - DESK_H) : d.y;
    div.style.left = x+'px'; div.style.top = y+'px';
    if(d.groupIndex!=null){
      const idx = d.groupIndex % GROUP_BORDER.length;
      div.style.borderLeft = '4px solid ' + GROUP_BORDER[idx];
      if(!d.studentId) div.style.background = GROUP_BG[idx];
      const badge = document.createElement('div');
      badge.className = 'preview-group-badge';
      badge.textContent = chineseNumeral(d.groupIndex+1) + '組';
      div.appendChild(badge);
    }
    const nameEl = document.createElement('div');
    nameEl.textContent = d.studentId ? studentLabel(d.studentId) : '';
    div.appendChild(nameEl);
    pc.appendChild(div);
  });
  pc.style.width = maxX>0 ? maxX+'px' : '100%';
  pc.style.height = (maxY+30)+'px';

  document.getElementById('previewLayoutWrap').classList.toggle('mirrored', previewMirrored);
  document.getElementById('btnTeacherView').textContent = previewMirrored ? '↩ 切換為學生視角' : '🔄 切換為老師視角';
}

function openPreview(){
  document.getElementById('previewTitle').textContent = document.getElementById('roomTitle').value || '教室座位表';
  previewMirrored = false;
  renderPreviewDesks();
  document.getElementById('previewModal').style.display = 'block';
}

function resetAll(){
  if(!confirm('確定要重置所有資料嗎？此操作無法復原。')) return;
  pushHistory();
  state = {
    students:[], desks:[], mode:'group',
    settings:{groupCount:5,perRow:6,pairsPerRow:3},
    constraints:{apart:[],front:[],groupApart:[]},
    roles: defaultRoles(),
    customScenarios: state.customScenarios,
    showRoleIcons: state.showRoleIcons,
    swapSourceId:null,
    groupPaintMode:false, activeGroupPaint:0
  };
  document.getElementById('studentInput').value = '';
  document.getElementById('roomTitle').value = '';
  saveRolesToStorage();
  render();
  toast('已重置');
}

// ---------- 事件綁定 ----------
document.getElementById('btnParseStudents').addEventListener('click', parseStudents);

document.querySelectorAll('.scenario').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    state.mode = btn.getAttribute('data-mode');
    render();
  });
});

document.getElementById('btnGenerate').addEventListener('click', generateLayout);
document.getElementById('btnAddDesk').addEventListener('click', addDesk);
document.getElementById('btnClearDesks').addEventListener('click', ()=>{
  if(state.desks.length===0) return;
  if(confirm('確定清空所有桌子？學生會回到未分配名單。')){ pushHistory(); state.desks = []; render(); }
});
document.getElementById('btnSaveScenario').addEventListener('click', saveCurrentAsScenario);

document.getElementById('btnUndo').addEventListener('click', undo);
document.getElementById('btnRedo').addEventListener('click', redo);
document.getElementById('btnToggleGroupPaint').addEventListener('click', ()=>{
  state.groupPaintMode = !state.groupPaintMode;
  state.swapSourceId = null;
  render();
});

document.getElementById('btnAutoFill').addEventListener('click', autoFillInOrder);
document.getElementById('btnArrange').addEventListener('click', runArrange);
document.getElementById('btnAddRole').addEventListener('click', addRole);
document.getElementById('btnResetRoles').addEventListener('click', resetRolesFromStudents);
document.getElementById('chkShowRoleIcons').addEventListener('change', e=>{
  state.showRoleIcons = e.target.checked;
  saveShowRoleIcons();
  render();
});
document.getElementById('btnClearAssign').addEventListener('click', ()=>{
  if(confirm('確定要清空目前的座位指派嗎？（鎖定的座位不受影響）')){
    pushHistory();
    state.desks.forEach(d=>{ if(!d.locked) d.studentId = null; });
    render();
  }
});

document.getElementById('btnAddApart').addEventListener('click', ()=>{
  const a = document.getElementById('apartA').value, b = document.getElementById('apartB').value;
  if(!a || !b || a===b){ toast('請選擇兩位不同的學生'); return; }
  if(!isApartPair(a,b)){ pushHistory(); state.constraints.apart.push([a,b]); render(); }
});
document.getElementById('btnAddFront').addEventListener('click', ()=>{
  const id = document.getElementById('frontSelect').value;
  if(!id){ toast('請選擇學生'); return; }
  if(!state.constraints.front.includes(id)){ pushHistory(); state.constraints.front.push(id); render(); }
  else toast('已在前排名單中');
});
document.getElementById('btnAddGroupApart').addEventListener('click', ()=>{
  const a = document.getElementById('groupApartA').value, b = document.getElementById('groupApartB').value;
  if(!a || !b || a===b){ toast('請選擇兩位不同的學生'); return; }
  if(!isNotSameGroupPair(a,b)){ pushHistory(); state.constraints.groupApart.push([a,b]); render(); }
});

document.getElementById('btnSave').addEventListener('click', saveJSON);
document.getElementById('btnLoad').addEventListener('click', ()=> document.getElementById('fileLoad').click());
document.getElementById('fileLoad').addEventListener('change', e=>{
  if(e.target.files && e.target.files[0]) loadJSONFile(e.target.files[0]);
  e.target.value = '';
});
document.getElementById('btnPreview').addEventListener('click', openPreview);
document.getElementById('btnExportRoster').addEventListener('click', exportGroupRoster);
document.getElementById('btnClosePreview').addEventListener('click', ()=> document.getElementById('previewModal').style.display='none');
document.getElementById('btnTeacherView').addEventListener('click', ()=>{
  previewMirrored = !previewMirrored;
  renderPreviewDesks();
});
document.getElementById('btnPrint').addEventListener('click', ()=> window.print());
document.getElementById('btnReset').addEventListener('click', resetAll);

document.getElementById('overlay').addEventListener('click', closePopup);
window.addEventListener('resize', fitCanvasSize);

// init
document.querySelector('.scenario[data-mode="group"]').classList.add('active');
render();
