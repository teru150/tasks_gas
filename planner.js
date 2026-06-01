/**
 * ============================================================
 * Conversational Unified AI Planner
 *
 * Flow:
 *   1) ダイアログを開くと「全タスク」と来週の既存「週予定」を読み込む
 *   2) Claude と自由にチャットで会話して計画を煮詰める
 *   3) Claude が "ready_for_blueprint": true で plan を返したら
 *      ダイアログ内に編集可能なブループリント (Google Calendar 風) を表示
 *   4) ユーザがブループリントを直接編集できる（クリックして書き換え）
 *   5) 「Write it down」ボタンを押した時のみシートに反映
 *
 * 何も書かれない限り、Claude のチャットや編集は一切シートに影響しない。
 * ============================================================
 */

const PLANNER_CONFIG = {
  MODEL: 'claude-sonnet-4-6',
  MAX_TOKENS: 8192,
  API_URL: 'https://api.anthropic.com/v1/messages',
  API_VERSION: '2023-06-01',
  API_KEY_PROP: 'ANTHROPIC_API_KEY',
  KB_PROP: 'AI_KNOWLEDGE_BASE',          // ナレッジベース (固有名詞・プロジェクト・人物etc.)
  CALENDARS_PROP: 'AI_PLANNER_CALENDARS', // 改行区切りのカレンダーID/メールアドレス
  COLOR_LABELS_PROP: 'WEEK_COLOR_LABELS', // 改行区切りで "#rrggbb=ラベル"

  TASKS_SHEET: '全タスク',
  WEEK_SHEET: '週予定',

  // 全タスク レイアウト
  TASKS_HEADER_ROW: 3,
  TASKS_START_ROW: 5,
  TASKS_SCAN_ROW_LIMIT: 300,

  // 週予定 レイアウト（calendarsync.js の CONFIG と一致）
  WEEK_DATE_ROW: 1,
  WEEK_CLASS_ROW: 2,
  WEEK_TIME_START_ROW: 3,
  WEEK_TIME_COL: 1,
  WEEK_DATA_START_COL: 2,
  WEEK_TIME_ROW_LIMIT: 25,
  WEEK_SCAN_COL_LIMIT: 200,

  MAX_TASKS_TO_AI: 50,
  MAX_HISTORY_TURNS: 24,                      // Claude に送るターン数 (12ペア)
  CHAT_LOG_SHEET: '_PlannerChatLog',          // 永続会話ログを置く隠しシート
  CHAT_HISTORY_RENDER: 40,                    // ダイアログ起動時に表示する直近ターン数
};


// ==================== メニュー入口 ====================

function openPlannerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('plannerDialog')
    .setWidth(900)
    .setHeight(740);
  SpreadsheetApp.getUi().showModalDialog(html, 'AI 週次プランナー（チャット）');
}

function plannerHasApiKey() {
  return !!PropertiesService.getScriptProperties().getProperty(PLANNER_CONFIG.API_KEY_PROP);
}


// ==================== 永続会話ログ ====================

/** 隠しシート _PlannerChatLog を取得（無ければ作成） */
function getChatLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PLANNER_CONFIG.CHAT_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PLANNER_CONFIG.CHAT_LOG_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'role', 'content']]);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 800);
    try { sheet.hideSheet(); } catch (e) { /* already hidden */ }
  }
  return sheet;
}

/** ダイアログ起動時に呼ばれる: 直近 N ターンを返す */
function plannerLoadHistory() {
  const sheet = getChatLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - PLANNER_CONFIG.CHAT_HISTORY_RENDER + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 3).getValues();
  return values.map(r => ({
    timestamp: (r[0] instanceof Date) ? r[0].toISOString() : String(r[0]),
    role: String(r[1] || ''),
    content: String(r[2] || ''),
  })).filter(m => m.role && m.content);
}

function appendChatTurn_(role, content) {
  const sheet = getChatLogSheet_();
  sheet.appendRow([new Date(), role, content]);
}

/** ダイアログから「会話履歴リセット」ボタンで呼ばれる */
function plannerClearHistory() {
  const sheet = getChatLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  return { ok: true };
}


// ==================== 色 → ラベル マッピング (週予定の色付け解釈) ====================

/** 改行区切り "#hex=ラベル" 文字列をパースして { hex: label } を返す */
function plannerGetColorLabelsRaw() {
  return PropertiesService.getScriptProperties().getProperty(PLANNER_CONFIG.COLOR_LABELS_PROP) || '';
}

function plannerParseColorLabels_() {
  const raw = plannerGetColorLabelsRaw();
  const map = {};
  String(raw).split('\n').forEach(line => {
    const m = line.trim().match(/^(#[0-9a-fA-F]{6})\s*=\s*(.+)$/);
    if (m) map[m[1].toLowerCase()] = m[2].trim();
  });
  return map;
}

function plannerGetColorLabels() {
  return { raw: plannerGetColorLabelsRaw(), map: plannerParseColorLabels_() };
}

function plannerSaveColorLabels(text) {
  PropertiesService.getScriptProperties().setProperty(PLANNER_CONFIG.COLOR_LABELS_PROP, String(text || ''));
  return { ok: true, map: plannerParseColorLabels_() };
}

/** 週予定シートで対象週内に出現する全ユニーク背景色を返す (色診断用) */
function plannerScanWeekColors(startYmd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekSheet = ss.getSheetByName(PLANNER_CONFIG.WEEK_SHEET);
  if (!weekSheet) return [];

  const weekDates = detectNextWeek_(startYmd);
  const timeSlots = readTimeSlots_(weekSheet);
  const weekColumns = findWeekColumns_(weekSheet, weekDates);

  const counts = {}; // hex -> { count, samples: [text] }

  weekDates.forEach(d => {
    const col = weekColumns[ymdKey_(d)];
    if (!col) return;
    if (!timeSlots.length) return;
    const minRow = Math.min(timeSlots[0].row, PLANNER_CONFIG.WEEK_CLASS_ROW);
    const maxRow = timeSlots[timeSlots.length - 1].row;
    const numRows = maxRow - minRow + 1;
    const bgs = weekSheet.getRange(minRow, col, numRows, 1).getBackgrounds();
    const vals = weekSheet.getRange(minRow, col, numRows, 1).getValues();
    for (let i = 0; i < bgs.length; i++) {
      const hex = String(bgs[i][0] || '').toLowerCase();
      if (!hex || hex === '#ffffff' || hex === '#000000') continue;
      if (!counts[hex]) counts[hex] = { count: 0, samples: [] };
      counts[hex].count++;
      const text = String(vals[i][0] || '').trim();
      if (text && counts[hex].samples.length < 3 && counts[hex].samples.indexOf(text) < 0) {
        counts[hex].samples.push(text);
      }
    }
  });

  return Object.keys(counts).map(hex => ({
    hex: hex,
    count: counts[hex].count,
    samples: counts[hex].samples,
  })).sort((a, b) => b.count - a.count);
}


// ==================== ナレッジベース ====================

/** ダイアログから読み出される: 現在の KB テキストを返す */
function plannerGetKnowledgeBase() {
  return PropertiesService.getScriptProperties().getProperty(PLANNER_CONFIG.KB_PROP) || '';
}

/** ダイアログから保存される: KB テキストを上書き保存 */
function plannerSaveKnowledgeBase(text) {
  const t = String(text || '');
  const MAX = 9000;
  const truncated = t.length > MAX ? t.slice(0, MAX) : t;
  PropertiesService.getScriptProperties().setProperty(PLANNER_CONFIG.KB_PROP, truncated);
  return { ok: true, length: truncated.length, truncated: t.length > MAX };
}


// ==================== カレンダー設定 ====================

/**
 * 監視対象のカレンダーID一覧。改行区切り。
 * "primary" は実行ユーザーのデフォルトカレンダー。
 * 他アカウントのカレンダーは、その所有アカウントから本アカウントへ「共有」設定をしておくこと。
 * 共有後、共有されたメールアドレス(例: family-account@gmail.com) を1行追加。
 */
function plannerGetCalendarIds() {
  const raw = PropertiesService.getScriptProperties().getProperty(PLANNER_CONFIG.CALENDARS_PROP);
  if (raw === null || raw === undefined) return ['primary']; // 未設定時のデフォルト
  return String(raw).split('\n').map(s => s.trim()).filter(Boolean);
}

function plannerSaveCalendarIds(text) {
  const lines = String(text || '').split('\n').map(s => s.trim()).filter(Boolean);
  PropertiesService.getScriptProperties().setProperty(PLANNER_CONFIG.CALENDARS_PROP, lines.join('\n'));

  // 各IDが実際にアクセス可能か検証
  const tz = Session.getScriptTimeZone();
  const checks = lines.map(id => {
    try {
      const cal = (id === 'primary') ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(id);
      if (!cal) return { id: id, ok: false, error: '見つかりません(共有されていない可能性)' };
      return { id: id, ok: true, name: cal.getName() };
    } catch (err) {
      return { id: id, ok: false, error: String(err && err.message ? err.message : err) };
    }
  });
  return { ok: true, count: lines.length, checks: checks };
}

/**
 * 指定週の全カレンダーイベントを { 'YYYY-MM-DD': [{time, title, source}, ...] } 形式で返す。
 * 終日イベントは time = '終日'。タイトル長は 60 文字でカット。
 */
function readCalendarsForWeek_(weekDates) {
  const ids = plannerGetCalendarIds();
  if (!ids.length) return {};

  const tz = Session.getScriptTimeZone();
  const start = new Date(weekDates[0].getFullYear(), weekDates[0].getMonth(), weekDates[0].getDate(), 0, 0, 0);
  const last = weekDates[weekDates.length - 1];
  const end = new Date(last.getFullYear(), last.getMonth(), last.getDate(), 23, 59, 59);

  const byDate = {};
  weekDates.forEach(d => { byDate[ymdKey_(d)] = []; });

  for (const id of ids) {
    let cal;
    try {
      cal = (id === 'primary') ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(id);
    } catch (e) { continue; }
    if (!cal) continue;

    const name = (cal.getName && cal.getName()) || id;

    let events;
    try { events = cal.getEvents(start, end); } catch (e) { continue; }

    for (const ev of events) {
      const s = ev.getStartTime();
      const e = ev.getEndTime();
      const dateKey = ymdKey_(s);
      if (!byDate[dateKey]) continue;

      const isAllDay = ev.isAllDayEvent();
      const time = isAllDay ? '終日' : Utilities.formatDate(s, tz, 'HH:mm') + '-' + Utilities.formatDate(e, tz, 'HH:mm');
      let title = String(ev.getTitle() || '').trim();
      if (title.length > 60) title = title.slice(0, 60) + '…';

      byDate[dateKey].push({
        time: time,
        startHour: isAllDay ? null : s.getHours(),
        startMin: isAllDay ? null : s.getMinutes(),
        endHour: isAllDay ? null : e.getHours(),
        endMin: isAllDay ? null : e.getMinutes(),
        title: title,
        source: name,
        isAllDay: isAllDay,
      });
    }
  }

  // 各日付内で開始時刻順
  Object.keys(byDate).forEach(k => {
    byDate[k].sort((a, b) => {
      if (a.isAllDay && !b.isAllDay) return -1;
      if (!a.isAllDay && b.isAllDay) return 1;
      return (a.startHour || 0) * 60 + (a.startMin || 0) - ((b.startHour || 0) * 60 + (b.startMin || 0));
    });
  });
  return byDate;
}


// ==================== コンテキスト収集 ====================

function getPlannerContext(startYmd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const weekDates = detectNextWeek_(startYmd);

  const tasksSheet = ss.getSheetByName(PLANNER_CONFIG.TASKS_SHEET);
  const tasks = tasksSheet ? readAllTasks_(tasksSheet) : [];

  const weekSheet = ss.getSheetByName(PLANNER_CONFIG.WEEK_SHEET);
  const timeSlots = weekSheet ? readTimeSlots_(weekSheet) : [];

  let existing = [];
  let weekColumns = {};
  if (weekSheet) {
    weekColumns = findWeekColumns_(weekSheet, weekDates);
    existing = readExistingSchedule_(weekSheet, weekColumns, weekDates, timeSlots);
  }

  // カレンダー (両アカウント分)
  const calendarIds = plannerGetCalendarIds();
  const calendarEvents = readCalendarsForWeek_(weekDates);
  let calendarEventCount = 0;
  Object.keys(calendarEvents).forEach(k => { calendarEventCount += calendarEvents[k].length; });

  return {
    targetWeek: {
      dates: weekDates.map(d => ymdKey_(d)),
      labels: weekDates.map(d => Utilities.formatDate(d, tz, 'M/d (EEE)')),
    },
    tasksCount: tasks.length,
    topTasks: tasks.slice(0, 16).map(formatTaskForPreview_),
    timeSlots: timeSlots.map(t => t.label),
    existingFixed: existing,
    weekColumnsFound: Object.keys(weekColumns).length,
    hasTasksSheet: !!tasksSheet,
    hasWeekSheet: !!weekSheet,
    calendarIds: calendarIds,
    calendarEvents: calendarEvents,
    calendarEventCount: calendarEventCount,
    colorMap: plannerParseColorLabels_(),
    colorMapRaw: plannerGetColorLabelsRaw(),
  };
}


// ==================== 会話エンドポイント ====================

/**
 * Claude と会話する。
 * @param {Array<{role:'user'|'assistant', content:string}>} history
 * @param {string} userMessage
 * @returns {{ ok:boolean, reply?:string, plan?:object|null, ready?:boolean, error?:string }}
 */
function plannerChat(history, userMessage, startYmd) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(PLANNER_CONFIG.API_KEY_PROP);
    if (!apiKey) throw new Error('ANTHROPIC_API_KEY が未設定です。');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const weekDates = detectNextWeek_(startYmd);

    const tasksSheet = ss.getSheetByName(PLANNER_CONFIG.TASKS_SHEET);
    const allTasks = tasksSheet ? readAllTasks_(tasksSheet) : [];

    const weekSheet = ss.getSheetByName(PLANNER_CONFIG.WEEK_SHEET);
    const timeSlots = weekSheet ? readTimeSlots_(weekSheet) : [];
    let weekColumns = {};
    let fixedSchedule = [];
    if (weekSheet) {
      weekColumns = findWeekColumns_(weekSheet, weekDates);
      fixedSchedule = readExistingSchedule_(weekSheet, weekColumns, weekDates, timeSlots);
    }

    const knowledgeBase = plannerGetKnowledgeBase();
    const calendarEvents = readCalendarsForWeek_(weekDates);
    const colorMap = plannerParseColorLabels_();

    const systemPrompt = buildSystemPrompt_({
      weekDates: weekDates,
      tasks: allTasks.slice(0, PLANNER_CONFIG.MAX_TASKS_TO_AI),
      timeSlots: timeSlots,
      fixedSchedule: fixedSchedule,
      tz: tz,
      knowledgeBase: knowledgeBase,
      calendarEvents: calendarEvents,
      colorMap: colorMap,
    });

    // 永続ログを真実の情報源にする (クライアント history は無視)
    const persisted = plannerLoadHistory();
    const messages = persisted
      .slice(-PLANNER_CONFIG.MAX_HISTORY_TURNS)
      .map(m => ({ role: m.role, content: m.content }));
    messages.push({ role: 'user', content: String(userMessage || '') });

    const body = {
      model: PLANNER_CONFIG.MODEL,
      max_tokens: PLANNER_CONFIG.MAX_TOKENS,
      system: systemPrompt,
      messages: messages,
    };

    const res = UrlFetchApp.fetch(PLANNER_CONFIG.API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': PLANNER_CONFIG.API_VERSION },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    const text = res.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('Anthropic API HTTP ' + code + ': ' + text.slice(0, 500));
    }

    const json = JSON.parse(text);
    const content = (json.content || []).map(c => c.text || '').join('\n');
    if (!content) throw new Error('Claude のレスポンスが空でした。');

    // 両ターンを永続ログに追記
    appendChatTurn_('user', String(userMessage || ''));
    appendChatTurn_('assistant', content);

    const parsed = parseAssistantTurn_(content);
    return {
      ok: true,
      reply: parsed.reply,
      plan: parsed.plan,
      ready: !!parsed.ready,
      raw: content,
    };
  } catch (err) {
    Logger.log('plannerChat error: ' + (err && err.stack ? err.stack : err));
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/**
 * Claude の出力を { reply, plan, ready } に分解。
 * 期待フォーマット:
 *   何かの自然言語テキスト
 *   ```plan
 *   { "ready_for_blueprint": true, "summary": "...", "days": [...] }
 *   ```
 */
function parseAssistantTurn_(content) {
  const m = content.match(/```plan\s*([\s\S]*?)```/);
  let plan = null;
  let ready = false;
  let reply = content;

  if (m) {
    try {
      const planJson = JSON.parse(m[1].trim());
      ready = !!planJson.ready_for_blueprint;
      plan = planJson;
      reply = content.replace(m[0], '').trim();
      if (!reply) reply = ready ? '計画案を作成しました。下のブループリントで確認・編集してください。' : '計画ドラフトを更新しました。';
    } catch (e) {
      // パース失敗時は plan なし、リプライはそのまま
      plan = null;
      ready = false;
    }
  }
  return { reply: reply, plan: plan, ready: ready };
}


// ==================== コミットエンドポイント ====================

/**
 * 編集後のブループリントをシートに反映。
 * @param {object} plan - { days: [{ date, daily_tasks, schedule }] }
 * @param {object} options - { writeToDaily: bool, writeToWeek: bool }
 */
function plannerCommit(plan, options) {
  try {
    if (!plan || !Array.isArray(plan.days)) throw new Error('plan が不正です。');
    options = options || {};
    const writeToDaily = options.writeToDaily !== false;
    const writeToWeek = options.writeToWeek !== false;
    // 承認された task_changes のみ適用 (ダイアログ側でフィルタ済み)
    const taskChanges = Array.isArray(options.taskChanges) ? options.taskChanges : [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weekDates = detectNextWeek_(options.startYmd);

    let dailyWritten = 0;
    let scheduleWritten = 0;
    let scheduleSkipped = 0;
    let taskChangesApplied = 0;
    let taskChangesFailed = 0;
    const taskChangeDetails = [];

    if (writeToDaily) {
      const dailyName = (typeof TASK_SYNC_CONFIG !== 'undefined' && TASK_SYNC_CONFIG.SHEET_NAME) || '毎日のタスク';
      const dailySheet = ss.getSheetByName(dailyName);
      if (!dailySheet) throw new Error('毎日のタスクシートが見つかりません: ' + dailyName);
      dailyWritten = writeToDailyTasks_(plan, dailySheet, weekDates);
    }
    if (writeToWeek) {
      const weekSheet = ss.getSheetByName(PLANNER_CONFIG.WEEK_SHEET);
      if (!weekSheet) throw new Error('週予定シートが見つかりません。');
      const timeSlots = readTimeSlots_(weekSheet);
      const weekColumns = findWeekColumns_(weekSheet, weekDates);
      const w = writeToWeekSchedule_(plan, weekSheet, weekColumns, weekDates, timeSlots);
      scheduleWritten = w.written;
      scheduleSkipped = w.skipped;
    }
    if (taskChanges.length) {
      const tasksSheet = ss.getSheetByName(PLANNER_CONFIG.TASKS_SHEET);
      if (!tasksSheet) throw new Error('全タスクシートが見つかりません。');
      const r = applyTaskChanges_(tasksSheet, taskChanges);
      taskChangesApplied = r.applied;
      taskChangesFailed = r.failed;
      r.details.forEach(d => taskChangeDetails.push(d));
    }

    return {
      ok: true,
      dailyWritten: dailyWritten,
      scheduleWritten: scheduleWritten,
      scheduleSkipped: scheduleSkipped,
      taskChangesApplied: taskChangesApplied,
      taskChangesFailed: taskChangesFailed,
      taskChangeDetails: taskChangeDetails,
    };
  } catch (err) {
    Logger.log('plannerCommit error: ' + (err && err.stack ? err.stack : err));
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

// ==================== 全タスク 編集 ====================

/**
 * field 名 → 全タスクシートの列番号 (1-indexed) のマッピング。
 * N(進捗%), O(時間経過%), R(期間中1日あたり) は通常 formula なので除外。
 */
// 全タスクシートの正しい列マッピング (1-indexed)。
// J(10) = "時間"ラベル / O(15) = 進捗% / P(16) = 時間経過% / S(19) = 1日あたり は formula なので書込対象外。
const TASKS_FIELD_TO_COL = {
  major: 1, mid: 2, sub: 3, task: 4, content: 5, priority: 6, order: 7,
  totalHours: 8,    // H かかる時間
  spentHours: 9,    // I 割いた時間
  hoursPerUnit: 11, // K 一単位当たり時間
  unit: 12,         // L 単位
  total: 13,        // M 全体分量
  current: 14,      // N 現在
  start: 17,        // Q 開始
  due: 18,          // R 締切
  notes: 20,        // T 備考
};

function applyTaskChanges_(sheet, changes) {
  const details = [];
  let applied = 0;
  let failed = 0;

  for (const c of changes) {
    try {
      if (!c || !c.action) throw new Error('action 欠落');

      if (c.action === 'update') {
        const row = Number(c.row);
        const field = String(c.field || '');
        const col = TASKS_FIELD_TO_COL[field];
        if (!row || row < PLANNER_CONFIG.TASKS_START_ROW) throw new Error('row 不正: ' + c.row);
        if (!col) throw new Error('field 不正: ' + field);
        const value = coerceTaskValue_(field, c.value);
        sheet.getRange(row, col).setValue(value);
        applied++;
        details.push({ ok: true, action: 'update', row: row, field: field, value: value });
      } else if (c.action === 'add') {
        const t = c.task || {};
        if (!String(t.task || '').trim()) throw new Error('add: task 名がありません');
        const newRow = findFirstEmptyRowInTaskSheet_(sheet);
        const rowValues = new Array(19).fill('');
        Object.keys(TASKS_FIELD_TO_COL).forEach(f => {
          if (t[f] !== undefined && t[f] !== null && t[f] !== '') {
            rowValues[TASKS_FIELD_TO_COL[f] - 1] = coerceTaskValue_(f, t[f]);
          }
        });
        sheet.getRange(newRow, 1, 1, 19).setValues([rowValues]);
        applied++;
        details.push({ ok: true, action: 'add', row: newRow, task: t.task });
      } else {
        throw new Error('未対応 action: ' + c.action);
      }
    } catch (err) {
      failed++;
      details.push({ ok: false, action: c && c.action, row: c && c.row, error: String(err && err.message || err) });
    }
  }
  return { applied: applied, failed: failed, details: details };
}

function coerceTaskValue_(field, value) {
  if (value === null || value === undefined) return '';
  if (field === 'start' || field === 'due') {
    if (value instanceof Date) return value;
    const s = String(value).trim();
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return s;
  }
  if (['order','totalHours','spentHours','hoursPerUnit','total','current'].indexOf(field) >= 0) {
    const n = Number(value);
    return isNaN(n) ? value : n;
  }
  return value;
}

function findFirstEmptyRowInTaskSheet_(sheet) {
  const startRow = PLANNER_CONFIG.TASKS_START_ROW;
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  // タスク列 (D = 4) で最初の空行を探す
  const numRows = lastRow - startRow + 1;
  if (numRows <= 0) return startRow;
  const values = sheet.getRange(startRow, 4, numRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) return startRow + i;
  }
  return lastRow + 1;
}


// ==================== System Prompt ====================

function buildSystemPrompt_(ctx) {
  const tz = ctx.tz;
  const dateLines = ctx.weekDates.map(d => Utilities.formatDate(d, tz, 'yyyy-MM-dd (EEE)'));
  const slotLines = ctx.timeSlots.map(s => '  ' + s.label).join('\n') || '  (時間スロット未取得)';

  const taskLines = ctx.tasks.map(t => {
    const path = [t.major, t.mid, t.sub, t.task].filter(Boolean).join(' / ');
    const meta = [];
    if (t.priority) meta.push('重要度=' + t.priority);
    if (t.remaining !== null && t.unit) meta.push('残=' + t.remaining + t.unit);
    if (t.remainingHours !== null) meta.push('残時間=' + t.remainingHours + 'h');
    if (t.hoursPerUnit !== null && t.unit) meta.push('単位時間=' + t.hoursPerUnit + 'h/' + t.unit);
    if (t.perDay !== null) meta.push('1日目安=' + t.perDay + 'h');
    if (t.start) meta.push('開始=' + Utilities.formatDate(t.start, tz, 'M/d'));
    if (t.due) meta.push('締切=' + Utilities.formatDate(t.due, tz, 'M/d'));
    if (t.progressText) meta.push('進捗=' + t.progressText);
    return `- [r${t.rowIndex}] ${path}  (${meta.join(' / ')})${t.content ? '  内容: ' + t.content.slice(0, 60) : ''}`;
  }).join('\n');

  const fixedLines = ctx.fixedSchedule.map(day => {
    if (day.status === 'no_column') return `■ ${day.label}: (週予定シートにこの日付の列がまだありません)`;
    if (!day.items.length) return `■ ${day.label}: (空き)`;
    return `■ ${day.label}\n` + day.items.map(it => {
      const lab = it.label ? `[${it.label}]` : (it.bg ? `[${it.bg}]` : '');
      const txt = it.text || '(空)';
      return `   - ${it.time} ${lab} ${txt}`.replace(/\s+/g, ' ');
    }).join('\n');
  }).join('\n');

  // 色ラベル一覧を AI にも明示
  const colorLegendLines = Object.keys(ctx.colorMap || {}).map(hex => `   - [${ctx.colorMap[hex]}] (${hex})`).join('\n');

  // Google カレンダーから読み込んだ実際の予定 (両アカウント)
  const calLines = (ctx.calendarEvents && Object.keys(ctx.calendarEvents).length)
    ? ctx.weekDates.map(d => {
        const key = ymdKey_(d);
        const events = (ctx.calendarEvents[key] || []);
        const dateLabel = Utilities.formatDate(d, ctx.tz, 'M/d (EEE)');
        if (!events.length) return `■ ${dateLabel}: (なし)`;
        return `■ ${dateLabel}\n` + events.map(e => `   - ${e.time} [${e.source}] ${e.title}`).join('\n');
      }).join('\n')
    : '(カレンダー設定なし)';

  const kbBlock = (ctx.knowledgeBase && ctx.knowledgeBase.trim())
    ? [
        '【ユーザに関する事前知識（ナレッジベース）】',
        '以下はユーザが事前に登録した自分自身・プロジェクト・固有名詞などの知識です。',
        'これを踏まえて会話してください。タスク名や略語の意味はここから推測してください。',
        '----',
        ctx.knowledgeBase.trim(),
        '----',
        '',
      ].join('\n')
    : '';

  return [
    'あなたはユーザの学習計画パートナーです。',
    'ユーザと対話しながら、来週の最適な計画を一緒に作っていきます。',
    '',
    kbBlock,
    '',
    '【あなたの動き方】',
    '1. 最初は質問・対話モードです。ユーザの目標、優先順位、避けるべき日や時間帯、',
    '   想定学習時間などを聞き出してください。',
    '2. 必要な情報が揃ったと判断したら、計画案を JSON で出力します。',
    '3. ユーザがブループリントを編集して再度メッセージを送ってきた場合、',
    '   その変更を尊重したうえで再提案できます（必要なら再度 plan を JSON で出す）。',
    '4. ユーザが新しいタスクを口頭で追加した、または条件を変えた場合も計画を再生成。',
    '',
    '【出力フォーマット】',
    '通常の応答は普通の文章で構いません。',
    '計画を提案するときは、文章のあとに ```plan ブロック``` を続けて以下のJSONを出力します:',
    '',
    '```plan',
    '{',
    '  "ready_for_blueprint": true,',
    '  "summary": "今週の方針を1〜2文で",',
    '  "days": [',
    '    {',
    '      "date": "YYYY-MM-DD",',
    '      "daily_tasks": ["毎日のタスクシートに書く具体タスク（短文）"],',
    '      "schedule": [{ "time": "8:00", "title": "タスク名", "task_row": 5 }]',
    '    }',
    '  ],',
    '  "task_changes": [',
    '    { "action": "update", "row": 5, "field": "current", "value": 12, "reason": "ユーザが12問完了と発言" },',
    '    { "action": "update", "row": 8, "field": "due", "value": "2026-05-15", "reason": "締切延長" },',
    '    { "action": "add", "task": { "major": "...", "mid": "...", "sub": "...", "task": "新タスク", "content": "...", "priority": "高", "totalHours": 5, "hoursPerUnit": 0.5, "unit": "問", "total": 10, "due": "2026-05-30" }, "reason": "ユーザが新規追加を依頼" }',
    '  ]',
    '}',
    '```',
    '',
    '【task_changes の使い方】',
    '- 全タスクシート (元データ) を編集したい時だけ task_changes に入れる',
    '- action は "update" / "add" のみ',
    '- update: row + field + value (field は task,content,priority,order,totalHours,spentHours,hoursPerUnit,unit,total,current,start,due,notes のいずれか)',
    '- add: task オブジェクトに必要なフィールドを入れる',
    '- reason は短い理由（ユーザに見せる）',
    '- 1ターンに最大 5 個まで。確信が無い変更は入れない',
    '- 変更がない場合は task_changes キー自体を出さない',
    '',
    '- 計画提案ターン以外で plan ブロックは付けない。',
    '- 文章で「○○はどうですか？」と聞き返したい場合、plan は出さなくてよい。',
    '- ユーザが「これでOK」「これでお願い」と言ったら、最終版の plan を再度出してから締めくくる。',
    '',
    '【対象期間】',
    dateLines.join(', '),
    '',
    '【週予定セルの色ラベル定義】',
    colorLegendLines || '   (色マップ未設定 — 色情報は無視してください)',
    '',
    '【週予定シートに既に入っている要素 (色ラベル付き)】',
    '※ 「[授業]」=授業中・絶対変更禁止。「[空きコマ]」=空き、ここに学習を入れて良い。',
    '   「[内職]」=授業中だが既に学習中、二重ブッキング禁止。',
    '   「[#hex]」と表示された色はマップ未登録、解釈不能なので避けて配置。',
    fixedLines || '(なし)',
    '',
    '【Google カレンダーから取得した実予定（変更禁止・避けて配置 / 自分用+家族用 など複数アカウント可能）】',
    calLines,
    '',
    '【利用可能な時間スロット (週予定シートの行)】',
    slotLines,
    '',
    '【全タスク（重要度 高>中>低 → 締切 → 順番 で並び替え済み）】',
    taskLines || '(タスクなし)',
    '',
    '【ルール】',
    '- date は対象期間の各日付を YYYY-MM-DD で必ず7件',
    '- daily_tasks は 1日 最大10件、各30文字以内推奨',
    '- schedule.time は上の利用可能時間スロットのいずれかと一致させる',
    '- schedule.title は全タスクの末端名（タスク列）をそのまま使うとよい',
    '- schedule.task_row は元タスクの行番号 (例 [r5] なら 5)',
    '- 既存固定要素と同じ (date, time) には書かない',
    '- 重要度=高 と 締切が近いタスク を優先',
  ].join('\n');
}


// ==================== 全タスク 読み取り ====================

function readAllTasks_(sheet) {
  const startRow = PLANNER_CONFIG.TASKS_START_ROW;
  const lastRow = Math.min(sheet.getLastRow(), startRow + PLANNER_CONFIG.TASKS_SCAN_ROW_LIMIT);
  if (lastRow < startRow) return [];

  // 全タスクシートの実レイアウト (A〜T, 20列):
  //  A 大項目  B 中項目  C 小項目  D タスク  E 内容
  //  F 重要度  G 順番  H かかる時間 (total)  I 割いた時間 (spent)
  //  J "時間" (H/I のユニット表示用ラベル — 文字列)
  //  K 一単位当たり時間  L 単位  M 全体分量  N 現在
  //  O 進捗% (formula text)  P 時間経過% (formula text)
  //  Q 開始日  R 締切日
  //  S 期間中1日あたり (formula)  T 備考
  // 以前は 19 列しか読まず、J列の存在を見落として全マッピングが 1 列ズレていた。
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 20).getValues();
  const out = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const taskName = String(row[3] || '').trim();
    if (!taskName) continue;

    const totalUnits = toNumberOrNull_(row[12]);    // M
    const currentUnits = toNumberOrNull_(row[13]);  // N
    const remainingUnits = (totalUnits !== null && currentUnits !== null)
      ? Math.max(totalUnits - currentUnits, 0) : null;

    const totalHours = toNumberOrNull_(row[7]);     // H
    const spentHours = toNumberOrNull_(row[8]);     // I
    const remainingHours = (totalHours !== null && spentHours !== null)
      ? Math.max(totalHours - spentHours, 0) : null;

    const start = (row[16] instanceof Date) ? row[16] : null;  // Q
    const due   = (row[17] instanceof Date) ? row[17] : null;  // R
    const perDay = toNumberOrNull_(row[18]);                    // S

    out.push({
      rowIndex: startRow + i,
      major: String(row[0] || '').trim(),
      mid:   String(row[1] || '').trim(),
      sub:   String(row[2] || '').trim(),
      task:  taskName,
      content: String(row[4] || '').trim(),
      priority: String(row[5] || '').trim(),
      order: toNumberOrNull_(row[6]),
      // 時間系
      totalHours: totalHours,
      spentHours: spentHours,
      remainingHours: remainingHours,
      hoursPerUnit: toNumberOrNull_(row[10]),       // K
      // 量系
      unit: String(row[11] || '').trim(),           // L
      total: totalUnits,
      current: currentUnits,
      remaining: remainingUnits,
      // ステータス系 (formula 出力 — 文字列)
      progressText: String(row[14] || '').trim(),   // O
      timePctText:  String(row[15] || '').trim(),   // P
      start: start,
      due: due,
      perDay: perDay,
      notes: String(row[19] || '').trim(),          // T
    });
  }

  const prioRank = { '高': 0, '中': 1, '低': 2 };
  out.sort((a, b) => {
    const pa = prioRank[a.priority] !== undefined ? prioRank[a.priority] : 3;
    const pb = prioRank[b.priority] !== undefined ? prioRank[b.priority] : 3;
    if (pa !== pb) return pa - pb;
    const da = a.due ? a.due.getTime() : Infinity;
    const db = b.due ? b.due.getTime() : Infinity;
    if (da !== db) return da - db;
    return (a.order || 99) - (b.order || 99);
  });

  return out;
}

function formatTaskForPreview_(t) {
  const parts = [];
  if (t.major) parts.push(t.major);
  if (t.mid)   parts.push(t.mid);
  if (t.sub)   parts.push(t.sub);
  parts.push(t.task);
  const head = parts.join(' / ');
  const meta = [];
  if (t.priority) meta.push('重要度' + t.priority);
  if (t.remaining !== null && t.unit) meta.push('残' + t.remaining + t.unit);
  if (t.due) meta.push('〆' + Utilities.formatDate(t.due, Session.getScriptTimeZone(), 'M/d'));
  return head + (meta.length ? '  [' + meta.join(' / ') + ']' : '');
}

function toNumberOrNull_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? null : n;
}


// ==================== 週予定 読み取り ====================

/**
 * 対象7日間の Date 配列を返す。
 * @param {string=} startYmd "YYYY-MM-DD" 指定があればその日から7日間。
 *                            無ければ今日から7日間 (今日を含む)。
 */
function detectNextWeek_(startYmd) {
  let start;
  if (startYmd) {
    const m = String(startYmd).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) start = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  if (!start) {
    const now = new Date();
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }
  const week = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    week.push(d);
  }
  return week;
}

function readTimeSlots_(weekSheet) {
  const startRow = PLANNER_CONFIG.WEEK_TIME_START_ROW;
  const limit = PLANNER_CONFIG.WEEK_TIME_ROW_LIMIT;
  const values = weekSheet.getRange(startRow, PLANNER_CONFIG.WEEK_TIME_COL, limit, 1).getValues();
  const slots = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v === '' || v === null || v === undefined) continue;
    const parsed = parseTimeSlotValue_(v);
    if (!parsed) continue;
    slots.push({
      row: startRow + i,
      label: parsed.label,
      hour: parsed.hour,
      minute: parsed.minute,
    });
  }
  return slots;
}

function parseTimeSlotValue_(v) {
  if (typeof v === 'number') {
    return { label: v + ':00', hour: Math.floor(v), minute: 0 };
  }
  if (v instanceof Date) {
    return { label: v.getHours() + ':' + ('0' + v.getMinutes()).slice(-2), hour: v.getHours(), minute: v.getMinutes() };
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return { label: s, hour: Number(m[1]), minute: Number(m[2]) };
  m = s.match(/^(\d{1,2})時(\d{1,2})?分?$/);
  if (m) return { label: s, hour: Number(m[1]), minute: m[2] ? Number(m[2]) : 0 };
  m = s.match(/^(\d{1,2})$/);
  if (m) return { label: s + ':00', hour: Number(m[1]), minute: 0 };
  return null;
}

function findWeekColumns_(weekSheet, weekDates) {
  const lastCol = Math.min(weekSheet.getLastColumn(), PLANNER_CONFIG.WEEK_SCAN_COL_LIMIT);
  if (lastCol < PLANNER_CONFIG.WEEK_DATA_START_COL) return {};
  const numCols = lastCol - PLANNER_CONFIG.WEEK_DATA_START_COL + 1;
  const headerRow = weekSheet.getRange(PLANNER_CONFIG.WEEK_DATE_ROW, PLANNER_CONFIG.WEEK_DATA_START_COL, 1, numCols).getValues()[0];

  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const d = parseAnyDate_(headerRow[i]);
    if (!d) continue;
    for (const target of weekDates) {
      if (sameYMD_(d, target)) {
        map[ymdKey_(target)] = PLANNER_CONFIG.WEEK_DATA_START_COL + i;
      }
    }
  }
  return map;
}

function readExistingSchedule_(weekSheet, weekColumns, weekDates, timeSlots) {
  const out = [];
  const colorMap = plannerParseColorLabels_();

  if (!timeSlots.length) {
    return weekDates.map(date => ({
      date: ymdKey_(date),
      label: Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d (EEE)'),
      status: 'no_slots',
      items: [],
    }));
  }

  for (const date of weekDates) {
    const key = ymdKey_(date);
    const col = weekColumns[key];
    const dateLabel = Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d (EEE)');
    if (!col) {
      out.push({ date: key, label: dateLabel, status: 'no_column', items: [] });
      continue;
    }

    const items = [];

    // 授業 行
    const classVal = String(weekSheet.getRange(PLANNER_CONFIG.WEEK_CLASS_ROW, col).getValue() || '').trim();
    const classBg = String(weekSheet.getRange(PLANNER_CONFIG.WEEK_CLASS_ROW, col).getBackground() || '').toLowerCase();
    if (classVal) {
      items.push({ time: '授業', text: classVal, bg: classBg, label: colorMap[classBg] || '' });
    }

    // 時間スロット (値 + 背景色 同時取得)
    const minRow = timeSlots[0].row;
    const maxRow = timeSlots[timeSlots.length - 1].row;
    const range = weekSheet.getRange(minRow, col, maxRow - minRow + 1, 1);
    const values = range.getValues();
    const bgs = range.getBackgrounds();
    for (const slot of timeSlots) {
      const idx = slot.row - minRow;
      const v = String(values[idx][0] || '').trim();
      const bg = String(bgs[idx][0] || '').toLowerCase();
      const label = colorMap[bg] || '';
      // テキストがある OR 色ラベルが付いてるなら item として記録
      // (例: 色だけ付いてる「空きコマ」も AI に伝えたい)
      if (v || label) {
        items.push({ time: slot.label, text: v, bg: bg, label: label });
      }
    }
    out.push({ date: key, label: dateLabel, status: 'ok', items: items });
  }
  return out;
}


// ==================== 書き込み: 毎日のタスク ====================

function writeToDailyTasks_(plan, dailySheet, weekDates) {
  if (typeof TASK_SYNC_CONFIG === 'undefined' ||
      typeof ensureDayPairColumnsThrough_ !== 'function' ||
      typeof findPairByDate_ !== 'function' ||
      typeof createDayPairColumn_ !== 'function' ||
      typeof findFirstEmptyRowInColumn_ !== 'function') {
    throw new Error('taskssync.js のヘルパが見つかりません。');
  }

  ensureDayPairColumnsThrough_(dailySheet, weekDates[weekDates.length - 1]);

  let count = 0;
  for (const day of plan.days) {
    const date = parseYmd_(day.date);
    if (!date) continue;
    const tasks = Array.isArray(day.daily_tasks) ? day.daily_tasks : [];
    if (!tasks.length) continue;

    let pair = findPairByDate_(dailySheet, date);
    if (!pair) pair = createDayPairColumn_(dailySheet, date);
    if (!pair) continue;

    for (const t of tasks) {
      const title = String(t || '').trim();
      if (!title) continue;

      // 2行ペア拡張ルール: 必要なら 日付マーカー行 + チェックボックス行 を追加し、後者を返す
      const row = (typeof findOrCreateEmptyTaskRow_ === 'function')
        ? findOrCreateEmptyTaskRow_(dailySheet, pair.titleCol, pair.doneCol)
        : findFirstEmptyRowInColumn_(dailySheet, pair.titleCol, TASK_SYNC_CONFIG.START_ROW);

      dailySheet.getRange(row, pair.titleCol).setValue(title);
      count++;
    }
  }
  return count;
}


// ==================== 書き込み: 週予定 ====================

function writeToWeekSchedule_(plan, weekSheet, weekColumns, weekDates, timeSlots) {
  if (!timeSlots.length) return { written: 0, skipped: 0 };

  const slotByLabel = {};
  for (const s of timeSlots) {
    slotByLabel[s.label] = s;
    slotByLabel[String(s.hour)] = s;
    slotByLabel[s.hour + ':00'] = s;
    slotByLabel[s.hour + ':' + ('0' + s.minute).slice(-2)] = s;
  }

  let written = 0;
  let skipped = 0;

  for (const day of plan.days) {
    const date = parseYmd_(day.date);
    if (!date) continue;
    const col = weekColumns[ymdKey_(date)];
    if (!col) { skipped += (day.schedule || []).length; continue; }

    const items = Array.isArray(day.schedule) ? day.schedule : [];
    for (const it of items) {
      const slot = slotByLabel[String(it.time).trim()];
      if (!slot) { skipped++; continue; }
      const cell = weekSheet.getRange(slot.row, col);
      const existing = String(cell.getValue() || '').trim();
      if (existing) { skipped++; continue; }
      cell.setValue(String(it.title || '').trim());
      written++;
    }
  }
  return { written: written, skipped: skipped };
}


// ==================== 日付ヘルパ ====================

function parseYmd_(s) {
  if (!s) return null;
  if (s instanceof Date) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  const m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function parseAnyDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = String(v).trim();
  let m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/(\d{1,2})[\/月](\d{1,2})/);
  if (m) {
    const y = new Date().getFullYear();
    return new Date(y, Number(m[1]) - 1, Number(m[2]));
  }
  return null;
}

function sameYMD_(a, b) {
  return a.getFullYear() === b.getFullYear()
      && a.getMonth() === b.getMonth()
      && a.getDate() === b.getDate();
}

function ymdKey_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
