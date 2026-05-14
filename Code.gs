// ============================================================
//  Google Apps Script — 캘린더 + 할일 통합 백엔드
//  script.google.com 에서 붙여넣기 후 웹앱으로 배포
// ============================================================

// ── 시트 초기화 ─────────────────────────────────────────────
function getEventSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('events');
  if (!s) {
    s = ss.insertSheet('events');
    s.appendRow(['dateKey','id','title','time','note','color','category','createdAt']);
    s.setFrozenRows(1);
  }
  return s;
}

function getTodoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('todos');
  if (!s) {
    s = ss.insertSheet('todos');
    s.appendRow(['id','text','done','category','priority','createdAt','updatedAt']);
    s.setFrozenRows(1);
  }
  return s;
}

// ── CORS 헤더 포함 응답 ──────────────────────────────────────
function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET 라우터 ───────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getEvents') return resp(getAllEvents());
  if (action === 'getTodos')  return resp(getAllTodos());
  if (action === 'getAll')    return resp({ events: getAllEvents().events, todos: getAllTodos().todos });
  return resp({ error: 'Unknown action' });
}

// ── POST 라우터 ──────────────────────────────────────────────
function doPost(e) {
  const data   = JSON.parse(e.postData.contents);
  const action = data.action;

  // 이벤트
  if (action === 'addEvent')    return resp(addEvent(data.event));
  if (action === 'deleteEvent') return resp(deleteEvent(data.id));

  // 할일
  if (action === 'addTodo')     return resp(addTodo(data.todo));
  if (action === 'updateTodo')  return resp(updateTodo(data.todo));
  if (action === 'deleteTodo')  return resp(deleteTodo(data.id));
  if (action === 'reorderTodos')return resp(reorderTodos(data.ids));

  return resp({ error: 'Unknown action' });
}

// ── 이벤트 CRUD ──────────────────────────────────────────────
function getAllEvents() {
  const rows = getEventSheet().getDataRange().getValues();
  const events = {};
  for (let i = 1; i < rows.length; i++) {
    const [dateKey, id, title, time, note, color, category, createdAt] = rows[i];
    if (!dateKey || !id) continue;
    if (!events[dateKey]) events[dateKey] = [];
    events[dateKey].push({ id: String(id), title, time, note, color, category, createdAt });
  }
  for (const k in events)
    events[k].sort((a, b) => (a.time || '').localeCompare(b.time || ''));
  return { events };
}

function addEvent(ev) {
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  getEventSheet().appendRow([ev.dateKey, id, ev.title, ev.time||'', ev.note||'', ev.color||'blue', ev.category||'', now]);
  return { success: true, id };
}

function deleteEvent(id) {
  const s = getEventSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(id)) { s.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'Not found' };
}

// ── 할일 CRUD ────────────────────────────────────────────────
function getAllTodos() {
  const rows = getTodoSheet().getDataRange().getValues();
  const todos = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, text, done, category, priority, createdAt, updatedAt] = rows[i];
    if (!id) continue;
    todos.push({ id: String(id), text: String(text), done: done === true || done === 'TRUE', category: String(category||''), priority: String(priority||'보통'), createdAt: String(createdAt||''), updatedAt: String(updatedAt||'') });
  }
  return { todos };
}

function addTodo(todo) {
  const id  = Utilities.getUuid();
  const now = new Date().toISOString();
  getTodoSheet().appendRow([id, todo.text, false, todo.category||'', todo.priority||'보통', now, now]);
  return { success: true, id };
}

function updateTodo(todo) {
  const s    = getTodoSheet();
  const rows = s.getDataRange().getValues();
  const now  = new Date().toISOString();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(todo.id)) {
      // 체크 토글만 할 경우 text/category/priority 유지
      s.getRange(i+1, 2).setValue(todo.text     !== undefined ? todo.text     : rows[i][1]);
      s.getRange(i+1, 3).setValue(todo.done     !== undefined ? todo.done     : rows[i][2]);
      s.getRange(i+1, 4).setValue(todo.category !== undefined ? todo.category : rows[i][3]);
      s.getRange(i+1, 5).setValue(todo.priority !== undefined ? todo.priority : rows[i][4]);
      s.getRange(i+1, 7).setValue(now);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function deleteTodo(id) {
  const s    = getTodoSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { s.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'Not found' };
}

function reorderTodos(ids) {
  // ids 배열 순서대로 시트 행 재정렬
  const s    = getTodoSheet();
  const rows = s.getDataRange().getValues();
  const header = rows[0];
  const map  = {};
  for (let i = 1; i < rows.length; i++) map[String(rows[i][0])] = rows[i];
  s.clearContents();
  s.appendRow(header);
  ids.forEach(id => { if (map[id]) s.appendRow(map[id]); });
  return { success: true };
}
