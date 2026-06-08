// ============================================================
// 制作管理チームツール — Google Apps Script バックエンド
// ============================================================
// 【セットアップ手順】
// 1. Googleスプレッドシートを新規作成
// 2. 拡張機能 > Apps Script を開く
// 3. このコードを貼り付けて保存
// 4. 「デプロイ」>「新しいデプロイ」> 種類:ウェブアプリ
//    - 実行ユーザー: 自分
//    - アクセス: 全員
// 5. デプロイURLをコピーして team-manager.html に貼り付ける
// ============================================================

const SHEET_WORKSPACES = "workspaces";
const SHEET_MEMBERS    = "members";
const SHEET_PROJECTS   = "projects";
const SHEET_TASKS      = "tasks";

// ----------------------------------------------------------
// エントリポイント
// ----------------------------------------------------------

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const body   = parseBody(e);
  const action = params.action || body.action;

  const handlers = {
    // ワークスペース
    createWorkspace : createWorkspace,
    loginWorkspace  : loginWorkspace,
    // メンバー
    getMembers      : getMembers,
    addMember       : addMember,
    removeMember    : removeMember,
    // プロジェクト
    getProjects     : getProjects,
    createProject   : createProject,
    updateProject   : updateProject,
    deleteProject   : deleteProject,
    // タスク
    getProjectStats : getProjectStats,
    getTasks        : getTasks,
    createTask      : createTask,
    updateTask      : updateTask,
    deleteTask      : deleteTask,
  };

  try {
    if (!action || !handlers[action]) {
      return json({ ok: false, error: "unknown action: " + action });
    }
    const result = handlers[action](body.data || params);
    return json({ ok: true, data: result });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

// ----------------------------------------------------------
// ワークスペース
// ----------------------------------------------------------

function createWorkspace({ name, password }) {
  requireParams({ name, password });
  const sheet = getSheet(SHEET_WORKSPACES);
  const rows  = sheet.getDataRange().getValues();

  // 重複チェック（1行目はヘッダー）
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) throw new Error("同じ名前のワークスペースが既に存在します");
  }

  const hashed = hashPassword(password);
  sheet.appendRow([name, hashed, new Date().toISOString()]);
  return { name };
}

function loginWorkspace({ name, password }) {
  requireParams({ name, password });
  const sheet = getSheet(SHEET_WORKSPACES);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) {
      if (rows[i][1] !== hashPassword(password)) {
        throw new Error("パスワードが違います");
      }
      return { name };
    }
  }
  throw new Error("ワークスペースが見つかりません");
}

// ----------------------------------------------------------
// メンバー
// ----------------------------------------------------------

function getMembers({ workspace }) {
  requireParams({ workspace });
  const sheet = getSheet(SHEET_MEMBERS);
  const rows  = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(r => r[0] === workspace)
    .map(r => ({ id: r[2], name: r[1] }));
}

function addMember({ workspace, name }) {
  requireParams({ workspace, name });
  const sheet = getSheet(SHEET_MEMBERS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === workspace && rows[i][1] === name) {
      throw new Error("同名のメンバーが既に存在します");
    }
  }

  const id = generateId();
  sheet.appendRow([workspace, name, id]);
  return { id, name };
}

function removeMember({ workspace, memberId }) {
  requireParams({ workspace, memberId });
  const sheet = getSheet(SHEET_MEMBERS);
  deleteRowWhere(sheet, row => row[0] === workspace && row[2] === memberId);
}

// ----------------------------------------------------------
// プロジェクト
// ----------------------------------------------------------

function getProjects({ workspace }) {
  requireParams({ workspace });
  const sheet = getSheet(SHEET_PROJECTS);
  const rows  = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(r => r[1] === workspace)
    .map(r => ({ id: r[0], workspace: r[1], name: r[2], color: r[3], createdAt: r[4] }));
}

function createProject({ workspace, name, color }) {
  requireParams({ workspace, name });
  const sheet = getSheet(SHEET_PROJECTS);
  const id    = generateId();
  sheet.appendRow([id, workspace, name, color || "#4f8ef7", new Date().toISOString()]);
  return { id, workspace, name, color: color || "#4f8ef7" };
}

function updateProject({ workspace, projectId, name, color }) {
  requireParams({ workspace, projectId });
  const sheet = getSheet(SHEET_PROJECTS);
  updateRowWhere(
    sheet,
    row => row[0] === projectId && row[1] === workspace,
    row => {
      if (name)  row[2] = name;
      if (color) row[3] = color;
      return row;
    }
  );
}

function deleteProject({ workspace, projectId }) {
  requireParams({ workspace, projectId });
  // プロジェクトに紐づくタスクも削除
  const taskSheet = getSheet(SHEET_TASKS);
  deleteRowWhere(taskSheet, row => row[1] === projectId);

  const projSheet = getSheet(SHEET_PROJECTS);
  deleteRowWhere(projSheet, row => row[0] === projectId && row[1] === workspace);
}

// ----------------------------------------------------------
// タスク
// ----------------------------------------------------------

function getProjectStats({ workspace }) {
  requireParams({ workspace });
  // ワークスペース内の全プロジェクトIDを取得
  const projects = getSheet(SHEET_PROJECTS).getDataRange().getValues().slice(1)
    .filter(r => r[1] === workspace).map(r => r[0]);

  const tasks = getSheet(SHEET_TASKS).getDataRange().getValues().slice(1);
  const stats = {};
  projects.forEach(id => {
    const t = tasks.filter(r => r[1] === id);
    stats[id] = { total: t.length, done: t.filter(r => r[4] === true || r[4] === "TRUE").length };
  });
  return stats;
}

function getTasks({ workspace, projectId }) {
  requireParams({ workspace, projectId });
  const sheet = getSheet(SHEET_TASKS);
  const rows  = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(r => r[1] === projectId)
    .map(r => ({
      id         : r[0],
      projectId  : r[1],
      name       : r[2],
      assignee   : r[3],
      done       : r[4] === true || r[4] === "TRUE",
      order      : r[5] || 0,
      createdAt  : r[6],
    }));
}

function createTask({ workspace, projectId, name, assignee }) {
  requireParams({ workspace, projectId, name });
  const sheet = getSheet(SHEET_TASKS);
  const id    = generateId();
  sheet.appendRow([id, projectId, name, assignee || "", false, Date.now(), new Date().toISOString()]);
  return { id, projectId, name, assignee: assignee || "", done: false };
}

function updateTask({ workspace, projectId, taskId, name, assignee, done }) {
  requireParams({ workspace, taskId });
  const sheet = getSheet(SHEET_TASKS);
  updateRowWhere(
    sheet,
    row => row[0] === taskId && row[1] === projectId,
    row => {
      if (name     !== undefined) row[2] = name;
      if (assignee !== undefined) row[3] = assignee;
      if (done     !== undefined) row[4] = done;
      return row;
    }
  );
}

function deleteTask({ workspace, projectId, taskId }) {
  requireParams({ workspace, taskId });
  const sheet = getSheet(SHEET_TASKS);
  deleteRowWhere(sheet, row => row[0] === taskId && row[1] === projectId);
}

// ----------------------------------------------------------
// ユーティリティ
// ----------------------------------------------------------

function getSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = {
      workspaces : ["name", "passwordHash", "createdAt"],
      members    : ["workspace", "name", "id"],
      projects   : ["id", "workspace", "name", "color", "createdAt"],
      tasks      : ["id", "projectId", "name", "assignee", "done", "order", "createdAt"],
    };
    if (headers[name]) sheet.appendRow(headers[name]);
  }
  return sheet;
}

function deleteRowWhere(sheet, predicate) {
  const data = sheet.getDataRange().getValues();
  // 下から削除しないとインデックスがずれる
  for (let i = data.length - 1; i >= 1; i--) {
    if (predicate(data[i])) sheet.deleteRow(i + 1);
  }
}

function updateRowWhere(sheet, predicate, updater) {
  const range = sheet.getDataRange();
  const data  = range.getValues();
  for (let i = 1; i < data.length; i++) {
    if (predicate(data[i])) {
      const updated = updater([...data[i]]);
      sheet.getRange(i + 1, 1, 1, updated.length).setValues([updated]);
      return;
    }
  }
}

function generateId() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 16);
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function requireParams(params) {
  for (const [key, val] of Object.entries(params)) {
    if (val === undefined || val === null || val === "") {
      throw new Error("パラメーター '" + key + "' が必要です");
    }
  }
}

function parseBody(e) {
  try {
    return JSON.parse(e.postData.contents);
  } catch (_) {
    return {};
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
