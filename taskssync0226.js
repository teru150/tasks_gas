/************************************************************
 * TaskSync.gs（横持ちレイアウト版）
 * 「毎日のタスク」シート（2列1セット） <-> Google Tasks 双方向同期
 *
 * レイアウト想定：
 *   A1: 日付, A2:A: タスク名, B2:B: 完了チェック
 *   C1: 日付, C2:C: タスク名, D2:D: 完了チェック
 *   E1: 日付, E2:E: タスク名, F2:F: 完了チェック
 *   ...
 *
 * 機能:
 * - シート <-> Google Tasks 双方向同期
 * - Google Tasks側で作った新規タスクをシートへ取り込み
 * - サブタスク機能（双方向対応）:
 *    * シート側: (main)メインタスク名
 *    *           (sub1)サブタスク1
 *    *           (sub2)サブタスク2
 *    * Google Tasks側でサブタスクを作成すると、シートにもプレフィックス付きで反映
 * - 色付け:
 *    * 繰越未完了 = 赤
 *    * 繰越系列が完了済み = 元/繰越とも橙（赤優先）
 * - 繰り越しタスクの自動非表示:
 *    * 同じベースタイトルの繰り越しタスクが複数日付に存在する場合、
 *      Google Tasks側では最も古い日付（元タスク）のみを表示し、新しい日付（繰り越し先）は完了化して非表示
 *
 * 自動繰り越し機能は現在無効化（手動で繰り越しを行うため）
 * Task IDは見える列には置かず、隠しシート _TaskSyncMap に保存
 ************************************************************/

const TASK_SYNC_CONFIG = {
  SHEET_NAME: '毎日のタスク',
  MAP_SHEET_NAME: '_TaskSyncMap',
  TASKLIST_TITLE: 'マイタスク',
  SYNC_TAG: '[毎日のタスク同期]',
  TOMORROW_RELEASE_HOUR: 23,
  DAY_BOUNDARY_HOUR: 2,       // 2:00までは前日扱い（色判定など）

  HEADER_ROW: 1,
  START_ROW: 2,
  START_COL: 1,     // A列
  PAIR_WIDTH: 2,    // (タスク名列, チェック列) の2列セット
  MAX_SCAN_COLS: 80,

  // 色
  COLOR_NORMAL: null,         // 通常は塗りなし
  COLOR_LATE_DONE: '#f4b183', // 橙
  COLOR_LATE_OPEN: '#ff6666', // 赤
};

/* =========================
   公開関数（メニュー / トリガー）
   ========================= */

/**
 * メイン同期
 * - Google Tasks -> Sheet（完了状態反映 ※前日以前はチェック戻ししない）
 * - Google Tasks -> Sheet（Tasks側で新規作成されたタスク取り込み）
 * - Sheet -> Google Tasks（今日）
 * - 23時以降だけ明日も作成/更新
 * - 色更新
 */
function syncDailyTasksBidirectional() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート「${TASK_SYNC_CONFIG.SHEET_NAME}」が見つかりません。`);

  const taskListId = getGoogleTaskListId_();
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  // 1) Google Tasks -> Sheet（既存タスクの完了状態を反映）
  syncGoogleTasksToSheet_(sheet, taskListId);

  // 2) Google Tasks -> Sheet（Tasks側で新規作成されたタスクを取り込む）
  importNewGoogleTasksToSheet_(sheet, taskListId);

  // 3) Sheet -> Tasks（今日分）
  syncSheetDayToGoogleTasks_(sheet, taskListId, today);

  // 4) 23:00以降だけ明日分も作成/更新
  if (now.getHours() >= TASK_SYNC_CONFIG.TOMORROW_RELEASE_HOUR) {
    syncSheetDayToGoogleTasks_(sheet, taskListId, tomorrow);
  }

  // 5) 繰り越しタスクの古い日付をGoogle Tasks側で非表示化
  hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId);

  // 6) 色更新
  refreshTaskColors_();

  ss.toast(
    `Tasks同期完了（${Utilities.formatDate(now, tz, 'M/d HH:mm')}）`,
    'Google Tasks同期',
    5
  );
}

/**
 * シート編集時に即時同期（インストール型トリガー）
 * 対象: 「毎日のタスク」シート内のタスク列/チェック列
 */
function onEditDailyTasksSync(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== TASK_SYNC_CONFIG.SHEET_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < TASK_SYNC_CONFIG.START_ROW) return;

    // 2列セットのどちらか（タスク列 or チェック列）だけ反応
    if (!isTaskPairColumn_(col)) return;

    const taskListId = getGoogleTaskListId_();
    syncSingleCellTaskPairRow_(sheet, row, col, taskListId);

    refreshTaskColors_();
  } catch (err) {
    Logger.log('onEditDailyTasksSync error: ' + err);
  }
}

/**
 * Google Tasks側の完了変更をシートに戻す用（定期実行）
 * ついでに Tasks側の新規作成タスクも取り込む
 */
function pollGoogleTasksCompletionToSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート「${TASK_SYNC_CONFIG.SHEET_NAME}」が見つかりません。`);

  const taskListId = getGoogleTaskListId_();

  // 完了状態を反映（前日以前は戻さない）
  syncGoogleTasksToSheet_(sheet, taskListId);

  // Tasks側で新規作成されたタスクも取り込む
  importNewGoogleTasksToSheet_(sheet, taskListId);

  // 繰り越しタスクの古い日付をGoogle Tasks側で非表示化
  hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId);

  refreshTaskColors_();
}

/* ========================================
   自動繰り越し機能（現在無効化）
   手動で繰り越しを行うため、コメントアウト
   復活させる場合は以下のコメントを解除してください
   ======================================== */

/*
/**
 * 毎日2:00に実行する想定：
 * 今日の未完了タスクを翌日にコピー（元はシートに未完了で残す）
 * ただしGoogle Tasks側の「今日」の元タスクは完了化して、明日分だけ見えるようにする
 */
/*
function rolloverUncompletedTasksToTomorrow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート「${TASK_SYNC_CONFIG.SHEET_NAME}」が見つかりません。`);

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const todayPair = findPairByDate_(sheet, today);
  if (!todayPair) {
    Logger.log('今日の列が見つからないため、繰り越しスキップ');
    return;
  }

  // 明日の列がなければ作る
  let tomorrowPair = findPairByDate_(sheet, tomorrow);
  if (!tomorrowPair) {
    tomorrowPair = createDayPairColumn_(sheet, tomorrow);
  }

  const lastRow = Math.max(sheet.getLastRow(), TASK_SYNC_CONFIG.START_ROW);
  const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;

  // 一括でセルを読み込む（今日分と明日分）
  const todayTitles = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, todayPair.titleCol, numRows, 1).getValues() : [];
  const todayDones = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, todayPair.doneCol, numRows, 1).getValues() : [];
  const tomorrowTitles = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, tomorrowPair.titleCol, numRows, 1).getValues() : [];

  // 明日側の既存タスク（重複防止）
  const tomorrowTitlesSet = new Set();
  for (let i = 0; i < numRows; i++) {
    const t = String(tomorrowTitles[i][0] || '').trim();
    if (t) {
      tomorrowTitlesSet.add(normalizeTaskTitle_(t));
      tomorrowTitlesSet.add(normalizeTaskTitle_(normalizeTaskBaseTitle_(t)));
    }
  }

  const taskListId = getGoogleTaskListId_();
  const map = getTaskMapStore_();

  // 繰り越すタスクを収集
  const tasksToCarryover = [];
  for (let i = 0; i < numRows; i++) {
    const row = TASK_SYNC_CONFIG.START_ROW + i;
    const title = String(todayTitles[i][0] || '').trim();
    if (!title) continue;

    const done = !!todayDones[i][0];
    if (done) continue; // 完了済みは繰り越さない

    // 繰越N日目を付けて明日へ
    const carryTitle = makeCarryoverTitle_(title);

    // 重複防止（ベース名/繰越名どちらでもチェック）
    const keyTitleBase = normalizeTaskTitle_(normalizeTaskBaseTitle_(title));
    const keyTitleCarry = normalizeTaskTitle_(carryTitle);
    if (tomorrowTitlesSet.has(keyTitleBase) || tomorrowTitlesSet.has(keyTitleCarry)) {
      // 既に翌日に存在するなら、Tasks側の今日分だけ完了化しておく
      markTaskAsCompletedInGoogleTasksByCell_(sheet, row, todayPair.titleCol, todayPair.doneCol, taskListId, map);
      continue;
    }

    // 繰り越すタスクを記録
    tasksToCarryover.push({
      row: row,
      title: title,
      carryTitle: carryTitle,
      keyTitleBase: keyTitleBase,
      keyTitleCarry: keyTitleCarry
    });

    tomorrowTitlesSet.add(keyTitleBase);
    tomorrowTitlesSet.add(keyTitleCarry);

    // シート上の「今日」の元タスクは未完了のまま残すが、
    // Google Tasks側の「今日」のタスクは完了扱いにして一覧の混在を防ぐ
    markTaskAsCompletedInGoogleTasksByCell_(sheet, row, todayPair.titleCol, todayPair.doneCol, taskListId, map);
  }

  // 明日の列に一括で書き込む
  if (tasksToCarryover.length > 0) {
    // 最初の空行を見つける
    let destRow = findFirstEmptyRowInColumn_(sheet, tomorrowPair.titleCol, TASK_SYNC_CONFIG.START_ROW);

    const titleValues = [];
    const doneValues = [];

    for (const task of tasksToCarryover) {
      titleValues.push([task.carryTitle]);
      doneValues.push([false]);
    }

    // 一括書き込み
    sheet.getRange(destRow, tomorrowPair.titleCol, titleValues.length, 1).setValues(titleValues);
    sheet.getRange(destRow, tomorrowPair.doneCol, doneValues.length, 1).setValues(doneValues);
  }

  // map保存（上でupdated更新される）
  saveTaskMapStore_(map);

  // 繰り越したタスクだけをTasksへ反映（全列同期は時間がかかるため、新規タスクだけに限定）
  if (tasksToCarryover.length > 0) {
    let destRow = findFirstEmptyRowInColumn_(sheet, tomorrowPair.titleCol, TASK_SYNC_CONFIG.START_ROW);
    for (let i = 0; i < Math.min(tasksToCarryover.length, 50); i++) { // 最大50タスクまで即座に同期
      syncSingleTaskCellRow_(sheet, destRow + i, tomorrowPair.titleCol, tomorrowPair.doneCol, tomorrow, taskListId);
    }
    // 残りは次の定期同期（毎時）で自動的に同期される
  }

  refreshTaskColors_();

  ss.toast('未完了タスクを翌日に繰り越しました（Tasks側の前日分は完了化）', 'Tasks繰り越し', 5);
}
*/

/**
 * トリガー一括設定
 */
function setupDailyTaskSyncTriggers() {
  removeTaskSyncTriggers_();

  // 毎時：通常同期（今日＋23時以降は明日）
  ScriptApp.newTrigger('syncDailyTasksBidirectional')
    .timeBased()
    .everyHours(1)
    .create();

  // 15分ごと：Google Tasks側の完了状態をシートへ戻す＋新規取り込み
  ScriptApp.newTrigger('pollGoogleTasksCompletionToSheet')
    .timeBased()
    .everyMinutes(15)
    .create();

  // 毎日2:00前後：未完了を翌日に繰り越し（現在無効化）
  /*
  ScriptApp.newTrigger('rolloverUncompletedTasksToTomorrow')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .nearMinute(0)
    .create();
  */

  // onEdit：シート変更を即反映
  ScriptApp.newTrigger('onEditDailyTasksSync')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tasks同期トリガーを設定しました（毎時 / 15分 / onEdit）',
    'Tasks同期設定',
    5
  );
}

/**
 * TaskSync関連トリガーだけ削除
 */
function removeTaskSyncTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if ([
      'syncDailyTasksBidirectional',
      'pollGoogleTasksCompletionToSheet',
      'onEditDailyTasksSync',
      // 'rolloverUncompletedTasksToTomorrow'  // 自動繰り越し無効化のためコメントアウト
    ].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * 手動で色だけ再計算したい時用（任意メニュー用）
 */
function refreshTaskColors() {
  refreshTaskColors_();
  SpreadsheetApp.getActiveSpreadsheet().toast('タスク色を更新しました', 'Tasks色更新', 3);
}

/* =========================
   2列1セット（横持ち）処理
   ========================= */

/**
 * 指定日（today/tomorrow）の列セットを見つけて、その列のタスクをTasksへ同期
 */
function syncSheetDayToGoogleTasks_(sheet, taskListId, targetDate) {
  const pairs = getTaskColumnPairs_(sheet);

  for (const pair of pairs) {
    const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
    const headerDate = parseDateOnly_(headerVal);
    if (!headerDate) continue;
    if (!isSameDate_(headerDate, targetDate)) continue;

    syncTaskColumnPairToGoogleTasks_(sheet, pair.titleCol, pair.doneCol, headerDate, taskListId);
  }
}

/**
 * 編集された1セルに対応する「2列セットの1行」だけ同期
 */
function syncSingleCellTaskPairRow_(sheet, row, editedCol, taskListId) {
  const pair = normalizeToPair_(editedCol);
  if (!pair) return;

  const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
  const dueDate = parseDateOnly_(headerVal);
  if (!dueDate) return;

  syncSingleTaskCellRow_(sheet, row, pair.titleCol, pair.doneCol, dueDate, taskListId);
}

/**
 * ある日付列セット（例 A/B列）を全部同期
 * サブタスク機能に対応：メインタスクを先に処理してからサブタスクを処理
 */
function syncTaskColumnPairToGoogleTasks_(sheet, titleCol, doneCol, dueDate, taskListId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < TASK_SYNC_CONFIG.START_ROW) return;

  // 全タスクを読み込んで分類
  const tasks = [];
  for (let row = TASK_SYNC_CONFIG.START_ROW; row <= lastRow; row++) {
    const title = String(sheet.getRange(row, titleCol).getValue() || '').trim();
    if (!title) continue;

    tasks.push({
      row: row,
      title: title,
      isMain: isMainTask_(title),
      isSub: isSubTask_(title)
    });
  }

  // 1) メインタスクと通常タスクを先に処理
  for (const task of tasks) {
    if (!task.isSub) {
      syncSingleTaskCellRow_(sheet, task.row, titleCol, doneCol, dueDate, taskListId);
    }
  }

  // 2) サブタスクを処理（親タスクIDが必要なので後で処理）
  for (const task of tasks) {
    if (task.isSub) {
      // 直前のメインタスクを探す
      let parentRow = null;
      for (let i = task.row - 1; i >= TASK_SYNC_CONFIG.START_ROW; i--) {
        const prevTitle = String(sheet.getRange(i, titleCol).getValue() || '').trim();
        if (isMainTask_(prevTitle)) {
          parentRow = i;
          break;
        }
        // 別のサブタスクならスキップして探し続ける
        if (!isSubTask_(prevTitle) && prevTitle) {
          // 通常タスクに当たったら探索終了
          break;
        }
      }

      syncSingleTaskCellRow_(sheet, task.row, titleCol, doneCol, dueDate, taskListId, parentRow);
    }
  }
}

/**
 * 1タスク（1行）を同期
 * @param {number|null} parentRow - サブタスクの場合、親タスクの行番号
 */
function syncSingleTaskCellRow_(sheet, row, titleCol, doneCol, dueDate, taskListId, parentRow) {
  const title = String(sheet.getRange(row, titleCol).getValue() || '').trim();
  const done = !!sheet.getRange(row, doneCol).getValue();

  // 空欄なら何もしない（既存Task削除まではしない仕様）
  if (!title) return;

  // プレフィックスを削除したタイトル（Google Tasksに送信する）
  const cleanTitle = removePrefixFromTitle_(title);

  const map = getTaskMapStore_();
  const key = makeTaskMapKey_(sheet.getName(), row, titleCol, doneCol);
  const entry = map[key] || {};
  const taskId = (entry.taskId || '').trim();

  const dueIso = toTaskDueIso_(dueDate);

  // サブタスクの場合、親タスクIDを取得
  let parentTaskId = null;
  if (parentRow && isSubTask_(title)) {
    const parentKey = makeTaskMapKey_(sheet.getName(), parentRow, titleCol, doneCol);
    const parentEntry = map[parentKey];
    if (parentEntry && parentEntry.taskId) {
      parentTaskId = parentEntry.taskId;
    }
  }

  // 新規作成
  if (!taskId) {
    const resource = {
      title: cleanTitle,
      notes: buildTaskNotes_(sheet.getName(), row, titleCol),
      due: dueIso,
      status: done ? 'completed' : 'needsAction',
    };
    if (done) resource.completed = new Date().toISOString();
    if (parentTaskId) resource.parent = parentTaskId;

    const created = Tasks.Tasks.insert(resource, taskListId);

    map[key] = {
      taskId: created.id || '',
      updated: created.updated || '',
      sheetName: sheet.getName(),
      row: row,
      titleCol: titleCol,
      doneCol: doneCol,
    };
    saveTaskMapStore_(map);
    return;
  }

  // 既存更新
  try {
    const task = Tasks.Tasks.get(taskListId, taskId);

    task.title = cleanTitle;
    task.due = dueIso;
    task.notes = ensureSyncTag_(task.notes || '', sheet.getName(), row, titleCol);

    // 親タスクの更新（サブタスクの場合）
    if (parentTaskId && task.parent !== parentTaskId) {
      task.parent = parentTaskId;
    }

    if (done) {
      task.status = 'completed';
      if (!task.completed) task.completed = new Date().toISOString();
    } else {
      task.status = 'needsAction';
      task.completed = null;
    }

    const updated = Tasks.Tasks.update(task, taskListId, taskId);

    map[key].updated = updated.updated || '';
    saveTaskMapStore_(map);

  } catch (err) {
    // taskId無効なら再作成
    Logger.log(`Task update failed key=${key}, recreating: ${err}`);

    const recreateResource = {
      title: cleanTitle,
      notes: buildTaskNotes_(sheet.getName(), row, titleCol),
      due: dueIso,
      status: done ? 'completed' : 'needsAction',
      completed: done ? new Date().toISOString() : undefined,
    };
    if (parentTaskId) recreateResource.parent = parentTaskId;

    const recreated = Tasks.Tasks.insert(recreateResource, taskListId);

    map[key] = {
      taskId: recreated.id || '',
      updated: recreated.updated || '',
      sheetName: sheet.getName(),
      row: row,
      titleCol: titleCol,
      doneCol: doneCol,
    };
    saveTaskMapStore_(map);
  }
}

/* =========================
   Google Tasks -> Sheet
   ========================= */

/**
 * Google Tasks側の完了状態をシートに反映
 * （マップシートに保存してある taskId を使って逆引き）
 *
 * 仕様:
 * - 前日以前のタスクは、Google Tasks側の完了状態をシートへ戻さない
 *   （前日分をTasks側だけ完了化して一覧混在を防ぐため）
 */
function syncGoogleTasksToSheet_(sheet, taskListId) {
  const map = getTaskMapStore_();
  const keys = Object.keys(map);
  if (!keys.length) return;

  const taskMap = listAllGoogleTasksMap_(taskListId);

  // 0:00ではなく「運用上の今日」（例: 2:00までは前日扱い）
  const opToday = getOperationalToday_();

  keys.forEach(key => {
    const entry = map[key];
    if (!entry) return;
    if (entry.sheetName !== sheet.getName()) return;

    const taskId = entry.taskId;
    if (!taskId) return;

    const task = taskMap[taskId];
    if (!task) return; // 削除済みなど

    const row = Number(entry.row);
    const titleCol = Number(entry.titleCol);
    const doneCol = Number(entry.doneCol);

    if (row < TASK_SYNC_CONFIG.START_ROW || titleCol < 1 || doneCol < 1) return;

    const currentTitle = String(sheet.getRange(row, titleCol).getValue() || '').trim();
    if (!currentTitle) return;

    // そのセルが属する日付列
    const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).getValue();
    const cellDate = parseDateOnly_(headerVal);
    if (!cellDate) return;

    // 「運用上の今日」より前は、Tasks側の完了状態をシートに戻さない
    // （前日分をTasks側だけ完了化して一覧混在を防ぐため）
    const shouldSyncDoneToSheet = !isBeforeDate_(cellDate, opToday); // cellDate >= opToday

    if (shouldSyncDoneToSheet) {
      const done = task.status === 'completed';
      sheet.getRange(row, doneCol).setValue(done);
    }

    // 必要ならタイトルもTasks側→シートへ反映（通常はオフ推奨）
    // sheet.getRange(row, titleCol).setValue(task.title || '');

    map[key].updated = task.updated || '';
  });

  saveTaskMapStore_(map);
}

/**
 * Google Tasks側で新規作成されたタスクをシートに取り込む
 * 条件:
 * - まだ _TaskSyncMap に存在しない taskId
 * - due がある（期日付き）
 * サブタスク対応：親タスクを先に処理してから、サブタスクを処理
 */
function importNewGoogleTasksToSheet_(sheet, taskListId) {
  const allTasks = listAllGoogleTasksMap_(taskListId);
  const map = getTaskMapStore_();

  // 既存管理済み taskId 一覧
  const managedTaskIds = new Set(
    Object.keys(map)
      .map(k => (map[k] && map[k].taskId) ? String(map[k].taskId) : '')
      .filter(Boolean)
  );

  // 未管理タスクを親タスクとサブタスクに分類
  const newParentTasks = [];
  const newSubTasks = [];

  for (const taskId in allTasks) {
    const task = allTasks[taskId];
    if (!task) continue;

    // 管理済みはスキップ
    if (managedTaskIds.has(taskId)) continue;

    // dueなしは取り込み先日付がないのでスキップ
    if (!task.due) continue;

    if (task.parent) {
      newSubTasks.push(task);
    } else {
      newParentTasks.push(task);
    }
  }

  // 1) 親タスクを先に処理
  for (const task of newParentTasks) {
    const dueDate = parseTaskDueToLocalDate_(task.due);
    if (!dueDate) continue;

    let pair = findPairByDate_(sheet, dueDate);
    if (!pair) {
      pair = createDayPairColumn_(sheet, dueDate);
    }

    // メインタスクとしてプレフィックスを付与
    const titleWithPrefix = `(main)${task.title || ''}`;

    // 同じ日付列に同名タスクが既にあるならそこへ紐づけ（重複作成防止）
    let matchedRow = findRowByTitleInColumn_(sheet, pair.titleCol, titleWithPrefix, TASK_SYNC_CONFIG.START_ROW);

    if (!matchedRow) {
      matchedRow = findFirstEmptyRowInColumn_(sheet, pair.titleCol, TASK_SYNC_CONFIG.START_ROW);
      sheet.getRange(matchedRow, pair.titleCol).setValue(titleWithPrefix);
    }

    // 完了状態反映
    sheet.getRange(matchedRow, pair.doneCol).setValue(task.status === 'completed');

    // map登録
    const key = makeTaskMapKey_(sheet.getName(), matchedRow, pair.titleCol, pair.doneCol);
    map[key] = {
      taskId: task.id || '',
      updated: task.updated || '',
      sheetName: sheet.getName(),
      row: matchedRow,
      titleCol: pair.titleCol,
      doneCol: pair.doneCol,
    };

    managedTaskIds.add(task.id);
  }

  // 2) サブタスクを処理
  for (const task of newSubTasks) {
    const dueDate = parseTaskDueToLocalDate_(task.due);
    if (!dueDate) continue;

    let pair = findPairByDate_(sheet, dueDate);
    if (!pair) {
      pair = createDayPairColumn_(sheet, dueDate);
    }

    // 親タスクの行を探す
    let parentRow = null;
    for (const key in map) {
      const entry = map[key];
      if (entry.taskId === task.parent && entry.titleCol === pair.titleCol) {
        parentRow = entry.row;
        break;
      }
    }

    // 親タスクが見つかった場合、タイトルに(main)がなければ追加
    if (parentRow) {
      const parentTitle = String(sheet.getRange(parentRow, pair.titleCol).getValue() || '').trim();
      if (parentTitle && !isMainTask_(parentTitle) && !isSubTask_(parentTitle)) {
        const newParentTitle = `(main)${parentTitle}`;
        sheet.getRange(parentRow, pair.titleCol).setValue(newParentTitle);
      }
    }

    if (!parentRow) {
      // 親タスクが見つからない場合は通常タスクとして処理
      Logger.log(`親タスクが見つからないため、通常タスクとして処理: ${task.title}`);
      let matchedRow = findFirstEmptyRowInColumn_(sheet, pair.titleCol, TASK_SYNC_CONFIG.START_ROW);
      sheet.getRange(matchedRow, pair.titleCol).setValue(task.title || '');
      sheet.getRange(matchedRow, pair.doneCol).setValue(task.status === 'completed');

      const key = makeTaskMapKey_(sheet.getName(), matchedRow, pair.titleCol, pair.doneCol);
      map[key] = {
        taskId: task.id || '',
        updated: task.updated || '',
        sheetName: sheet.getName(),
        row: matchedRow,
        titleCol: pair.titleCol,
        doneCol: pair.doneCol,
      };
      managedTaskIds.add(task.id);
      continue;
    }

    // 親タスクの下に既存のサブタスク数を数える
    let subTaskCount = 0;
    for (let r = parentRow + 1; r <= sheet.getLastRow(); r++) {
      const title = String(sheet.getRange(r, pair.titleCol).getValue() || '').trim();
      if (!title) break;
      if (isSubTask_(title)) {
        subTaskCount++;
      } else if (isMainTask_(title)) {
        // 次のメインタスクに到達したら終了
        break;
      } else {
        // 通常タスクに到達したら終了
        break;
      }
    }

    // サブタスクとしてプレフィックスを付与
    const subIndex = subTaskCount + 1;
    const titleWithPrefix = `${makeSubTaskPrefix_(subIndex)}${task.title || ''}`;

    // 親タスクの直後（既存サブタスクの後）に挿入
    const insertRow = parentRow + subTaskCount + 1;

    // 行が足りない場合は挿入
    if (insertRow > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }

    // タイトルと完了状態を設定
    sheet.getRange(insertRow, pair.titleCol).setValue(titleWithPrefix);
    sheet.getRange(insertRow, pair.doneCol).setValue(task.status === 'completed');

    // map登録
    const key = makeTaskMapKey_(sheet.getName(), insertRow, pair.titleCol, pair.doneCol);
    map[key] = {
      taskId: task.id || '',
      updated: task.updated || '',
      sheetName: sheet.getName(),
      row: insertRow,
      titleCol: pair.titleCol,
      doneCol: pair.doneCol,
    };

    managedTaskIds.add(task.id);
  }

  saveTaskMapStore_(map);
}

/* =========================
   繰越 / タスク直接操作 helpers
   ========================= */

/**
 * 指定セルに対応するGoogle Taskだけを「完了」にする
 * （シートのチェック状態は変更しない）
 */
function markTaskAsCompletedInGoogleTasksByCell_(sheet, row, titleCol, doneCol, taskListId, mapOpt) {
  const map = mapOpt || getTaskMapStore_();
  const key = makeTaskMapKey_(sheet.getName(), row, titleCol, doneCol);
  const entry = map[key];
  if (!entry || !entry.taskId) return;

  try {
    const task = Tasks.Tasks.get(taskListId, entry.taskId);
    task.status = 'completed';
    if (!task.completed) task.completed = new Date().toISOString();

    const updated = Tasks.Tasks.update(task, taskListId, entry.taskId);
    entry.updated = updated.updated || '';
  } catch (err) {
    Logger.log(`markTaskAsCompletedInGoogleTasksByCell_ failed: ${err}`);
  }
}

/**
 * 繰り越しタスクの新しい日付をGoogle Tasks側で非表示化
 * 同じベースタイトルの繰り越しタスクが複数の日付にある場合、
 * 最も古い日付（元タスク）のみを残し、新しい日付（繰り越し先）をGoogle Tasks側で完了化（非表示化）
 */
function hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId) {
  const map = getTaskMapStore_();
  const pairs = getTaskColumnPairs_(sheet);
  if (!pairs.length) return;

  // 全繰り越しタスクを収集 {baseTitle -> [{row, titleCol, doneCol, date, taskId, done}]}
  const carryoverGroups = {}; // baseTitle -> array of task info

  for (const pair of pairs) {
    const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
    const date = parseDateOnly_(headerVal);
    if (!date) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < TASK_SYNC_CONFIG.START_ROW) continue;

    const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;
    const titleVals = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.titleCol, numRows, 1).getValues();
    const doneVals = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.doneCol, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const row = TASK_SYNC_CONFIG.START_ROW + i;
      const title = String(titleVals[i][0] || '').trim();
      if (!title) continue;

      // 繰り越しタスクのみ対象
      if (!isCarryoverTitle_(title)) continue;

      const baseTitle = normalizeTaskBaseTitle_(title);
      const done = !!doneVals[i][0];

      // taskIdを取得
      const key = makeTaskMapKey_(sheet.getName(), row, pair.titleCol, pair.doneCol);
      const entry = map[key];
      if (!entry || !entry.taskId) continue;

      if (!carryoverGroups[baseTitle]) {
        carryoverGroups[baseTitle] = [];
      }

      carryoverGroups[baseTitle].push({
        row: row,
        titleCol: pair.titleCol,
        doneCol: pair.doneCol,
        date: date,
        taskId: entry.taskId,
        done: done,
        title: title
      });
    }
  }

  // 各グループごとに処理
  for (const baseTitle in carryoverGroups) {
    const tasks = carryoverGroups[baseTitle];
    if (tasks.length <= 1) continue; // 1つしかない場合は処理不要

    // 日付で昇順ソート（最も古い日付が先頭）
    tasks.sort((a, b) => {
      const timeA = a.date.getTime();
      const timeB = b.date.getTime();
      return timeA - timeB; // 昇順
    });

    // 最も古いタスク（先頭）以外を完了化（新しい方を非表示）
    for (let i = 1; i < tasks.length; i++) {
      const task = tasks[i];
      try {
        const gtask = Tasks.Tasks.get(taskListId, task.taskId);
        if (gtask.status !== 'completed') {
          gtask.status = 'completed';
          if (!gtask.completed) gtask.completed = new Date().toISOString();
          Tasks.Tasks.update(gtask, taskListId, task.taskId);
          Logger.log(`新しい繰り越しタスクを完了化: ${task.title} (日付: ${task.date})`);
        }
      } catch (err) {
        Logger.log(`hideOlderCarryoverTasks error for taskId ${task.taskId}: ${err}`);
      }
    }
  }
}

function getOperationalToday_() {
  const now = new Date();
  const boundary = Number(TASK_SYNC_CONFIG.DAY_BOUNDARY_HOUR || 0);

  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (now.getHours() < boundary) {
    d.setDate(d.getDate() - 1); // 深夜帯は前日扱い
  }
  return d;
}

/* =========================
   色付けロジック
   ========================= */

/**
 * シート全体のタスクセル色を更新する
 * ルール:
 * - 繰越タスク（繰越N日目）で未完了 → 赤
 * - 繰越系列が完了済み（元/繰越どれか完了） → 系列全体を橙（赤優先）
 */
function refreshTaskColors_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
  if (!sheet) return;

  const pairs = getTaskColumnPairs_(sheet);
  if (!pairs.length) return;

  const lastRow = Math.max(sheet.getLastRow(), TASK_SYNC_CONFIG.START_ROW);

  // 全タスク読み取り
  const records = []; // {row,titleCol,doneCol,title,done,isCarry,baseTitle,date}
  for (const pair of pairs) {
    const header = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
    const date = parseDateOnly_(header);
    if (!date) continue;

    const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;
    if (numRows <= 0) continue;

    const titleVals = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.titleCol, numRows, 1).getValues();
    const doneVals  = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.doneCol,  numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const row = TASK_SYNC_CONFIG.START_ROW + i;
      const title = String(titleVals[i][0] || '').trim();
      if (!title) continue;

      const done = !!doneVals[i][0];
      const isCarry = isCarryoverTitle_(title);
      const baseTitle = normalizeTaskBaseTitle_(title);

      records.push({
        row: row,
        titleCol: pair.titleCol,
        doneCol: pair.doneCol,
        title: title,
        done: done,
        isCarry: isCarry,
        baseTitle: baseTitle,
        date: date
      });
    }
  }

  // ベースタイトル単位でグループ化
  const groups = {}; // baseTitle -> record[]
  records.forEach(r => {
    const k = r.baseTitle;
    if (!groups[k]) groups[k] = [];
    groups[k].push(r);
  });

  // セルごとの背景色マップ（キー row:col）
  const cellColorMap = {};

  // 一旦通常色
  records.forEach(r => {
    cellColorMap[`${r.row}:${r.titleCol}`] = TASK_SYNC_CONFIG.COLOR_NORMAL;
  });

  // 「運用上の今日」（例: 2:00までは前日扱い）
  const opToday = getOperationalToday_();

  // ルール適用
  Object.keys(groups).forEach(base => {
    const list = groups[base];
    const hasCarry = list.some(r => r.isCarry);
    if (!hasCarry) return;

    // 運用上の今日以前のレコードがある系列だけ色付け対象
    // （未来日の繰越タスクを先に赤/橙にしない）
    const hasPastOrToday = list.some(r => !isBeforeDate_(opToday, r.date)); // r.date <= opToday
    if (!hasPastOrToday) return;

    // 系列内のどれかが完了しているか（※日付問わず）
    const anyDone = list.some(r => r.done);

    // 繰越未完了（運用上の今日以前）
    const carryOpenExists = list.some(r => r.isCarry && !r.done && !isBeforeDate_(opToday, r.date));

    // 1) 繰越未完了を赤
    if (carryOpenExists) {
      list.forEach(r => {
        if (r.isCarry && !r.done && !isBeforeDate_(opToday, r.date)) {
          cellColorMap[`${r.row}:${r.titleCol}`] = TASK_SYNC_CONFIG.COLOR_LATE_OPEN;
        }
      });
    }

    // 2) 系列内のどれかが完了していれば系列全体を橙（ただし赤優先）
    if (anyDone) {
      list.forEach(r => {
        // 未来日（運用上の今日より後）はまだ塗らない
        if (isBeforeDate_(opToday, r.date)) return; // r.date > opToday

        const key = `${r.row}:${r.titleCol}`;
        if (cellColorMap[key] === TASK_SYNC_CONFIG.COLOR_LATE_OPEN) return; // 赤優先
        cellColorMap[key] = TASK_SYNC_CONFIG.COLOR_LATE_DONE;
      });
    }
  });

  // 列ごとに一括反映（タスク名列だけ）
  for (const pair of pairs) {
    const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;
    if (numRows <= 0) continue;

    const bg = [];
    for (let i = 0; i < numRows; i++) {
      const row = TASK_SYNC_CONFIG.START_ROW + i;
      const key = `${row}:${pair.titleCol}`;
      const color = Object.prototype.hasOwnProperty.call(cellColorMap, key)
        ? cellColorMap[key]
        : TASK_SYNC_CONFIG.COLOR_NORMAL;
      bg.push([color]);
    }

    sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.titleCol, numRows, 1).setBackgrounds(bg);
  }
}

/* =========================
   マップ保存（隠しシート）
   ========================= */

/**
 * _TaskSyncMap シート構造:
 * A:key
 * B:taskId
 * C:updated
 * D:sheetName
 * E:row
 * F:titleCol
 * G:doneCol
 */
function getTaskMapStore_() {
  const mapSheet = getOrCreateTaskMapSheet_();
  const lastRow = mapSheet.getLastRow();
  const out = {};

  if (lastRow < 2) return out;

  const values = mapSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  values.forEach(r => {
    const key = String(r[0] || '').trim();
    if (!key) return;
    out[key] = {
      taskId: String(r[1] || ''),
      updated: String(r[2] || ''),
      sheetName: String(r[3] || ''),
      row: Number(r[4] || 0),
      titleCol: Number(r[5] || 0),
      doneCol: Number(r[6] || 0),
    };
  });

  return out;
}

function saveTaskMapStore_(obj) {
  const mapSheet = getOrCreateTaskMapSheet_();
  const keys = Object.keys(obj);

  // 全消しして書き直し（件数少ない想定）
  mapSheet.clearContents();
  mapSheet.getRange(1, 1, 1, 7).setValues([[
    'key', 'taskId', 'updated', 'sheetName', 'row', 'titleCol', 'doneCol'
  ]]);

  if (!keys.length) return;

  const rows = keys.map(k => {
    const v = obj[k] || {};
    return [
      k,
      v.taskId || '',
      v.updated || '',
      v.sheetName || '',
      v.row || '',
      v.titleCol || '',
      v.doneCol || '',
    ];
  });

  mapSheet.getRange(2, 1, rows.length, 7).setValues(rows);
}

function getOrCreateTaskMapSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TASK_SYNC_CONFIG.MAP_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(TASK_SYNC_CONFIG.MAP_SHEET_NAME);
    sh.hideSheet();
    sh.getRange(1, 1, 1, 7).setValues([[
      'key', 'taskId', 'updated', 'sheetName', 'row', 'titleCol', 'doneCol'
    ]]);
  }

  return sh;
}

function makeTaskMapKey_(sheetName, row, titleCol, doneCol) {
  return `${sheetName}::r${row}::t${titleCol}::d${doneCol}`;
}

/* =========================
   列判定 / 日付列操作ヘルパー
   ========================= */

function getTaskColumnPairs_(sheet) {
  const pairs = [];
  const maxCol = Math.min(sheet.getLastColumn(), TASK_SYNC_CONFIG.MAX_SCAN_COLS);

  for (let titleCol = TASK_SYNC_CONFIG.START_COL; titleCol <= maxCol; titleCol += TASK_SYNC_CONFIG.PAIR_WIDTH) {
    const doneCol = titleCol + 1;
    if (doneCol > maxCol + 1) break;

    const header = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).getValue();
    const date = parseDateOnly_(header);
    if (!date) continue; // 日付がないペアは無視

    pairs.push({ titleCol, doneCol });
  }

  return pairs;
}

function isTaskPairColumn_(col) {
  if (col < TASK_SYNC_CONFIG.START_COL) return false;
  const offset = col - TASK_SYNC_CONFIG.START_COL;
  return (offset % TASK_SYNC_CONFIG.PAIR_WIDTH === 0) || (offset % TASK_SYNC_CONFIG.PAIR_WIDTH === 1);
}

function normalizeToPair_(col) {
  if (!isTaskPairColumn_(col)) return null;
  const offset = col - TASK_SYNC_CONFIG.START_COL;
  const pairIndex = Math.floor(offset / TASK_SYNC_CONFIG.PAIR_WIDTH);
  const titleCol = TASK_SYNC_CONFIG.START_COL + pairIndex * TASK_SYNC_CONFIG.PAIR_WIDTH;
  return { titleCol, doneCol: titleCol + 1 };
}

function findPairByDate_(sheet, targetDate) {
  const pairs = getTaskColumnPairs_(sheet);
  for (const pair of pairs) {
    const v = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
    const d = parseDateOnly_(v);
    if (d && isSameDate_(d, targetDate)) return pair;
  }
  return null;
}

function createDayPairColumn_(sheet, dateObj) {
  const pairs = getTaskColumnPairs_(sheet);
  let titleCol, doneCol;

  if (pairs.length === 0) {
    titleCol = TASK_SYNC_CONFIG.START_COL;
    doneCol = titleCol + 1;
  } else {
    const lastPair = pairs[pairs.length - 1];
    titleCol = lastPair.titleCol + TASK_SYNC_CONFIG.PAIR_WIDTH;
    doneCol = titleCol + 1;

    if (sheet.getMaxColumns() < doneCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), doneCol - sheet.getMaxColumns());
    }
  }

  // 見出し（日付）
  sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).setValue(formatDateHeader_(dateObj));

  // チェック列にチェックボックス設定（多め）
  const maxRows = Math.max(sheet.getMaxRows(), 200);
  sheet.getRange(TASK_SYNC_CONFIG.START_ROW, doneCol, maxRows - 1, 1).insertCheckboxes();

  return { titleCol, doneCol };
}

function formatDateHeader_(dateObj) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const y = dateObj.getFullYear();
  const m = dateObj.getMonth() + 1;
  const d = dateObj.getDate();
  const w = days[dateObj.getDay()];
  return `${y}/${m}/${d}/${w}`;
}

function findFirstEmptyRowInColumn_(sheet, col, startRow) {
  const maxRow = Math.max(sheet.getLastRow(), startRow);
  const scanRows = Math.min(maxRow - startRow + 100, 500); // 最大500行をスキャン

  if (scanRows <= 0) return startRow;

  // 一括でセルを読み込む
  const values = sheet.getRange(startRow, col, scanRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] || '').trim();
    if (!v) return startRow + i;
  }

  return startRow + scanRows;
}

function findRowByTitleInColumn_(sheet, col, title, startRow) {
  const target = normalizeTaskTitle_(title);
  const targetBase = normalizeTaskTitle_(normalizeTaskBaseTitle_(title));
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const numRows = lastRow - startRow + 1;

  if (numRows <= 0) return null;

  // 一括でセルを読み込む
  const values = sheet.getRange(startRow, col, numRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] || '').trim();
    if (!v) continue;

    const n = normalizeTaskTitle_(v);
    const b = normalizeTaskTitle_(normalizeTaskBaseTitle_(v));

    // 同名 or 同ベース名を一致とみなす
    if (n === target || b === targetBase) return startRow + i;
  }
  return null;
}

/* =========================
   Google Tasks API helpers
   ========================= */

function getGoogleTaskListId_() {
  const lists = Tasks.Tasklists.list();
  const items = (lists && lists.items) ? lists.items : [];
  if (!items.length) throw new Error('Google Tasks のリストが見つかりません。');

  const hit = items.find(x => x.title === TASK_SYNC_CONFIG.TASKLIST_TITLE);
  return (hit || items[0]).id;
}

function listAllGoogleTasksMap_(taskListId) {
  const map = {};
  let pageToken = null;

  do {
    const res = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken: pageToken,
    });

    const items = (res && res.items) ? res.items : [];
    items.forEach(t => { map[t.id] = t; });

    pageToken = res.nextPageToken || null;
  } while (pageToken);

  return map;
}

/* =========================
   日付/文字列 helpers
   ========================= */

function parseDateOnly_(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const str = String(value).trim();

  // 2026/2/22, 2026-2-22, 2026/2/22/日 など
  let m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  // 2/22, 2月22
  m = str.match(/(\d{1,2})[\/月](\d{1,2})/);
  if (m) {
    const y = new Date().getFullYear();
    return new Date(y, Number(m[1]) - 1, Number(m[2]));
  }

  return null;
}

function parseTaskDueToLocalDate_(dueStr) {
  if (!dueStr) return null;
  const d = new Date(dueStr);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function isSameDate_(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function isBeforeDate_(a, b) {
  const aa = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
  const bb = new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
  return aa < bb;
}

function toTaskDueIso_(dateObj) {
  // JST正午固定（UTC変換による前日ズレ防止）
  const y = dateObj.getFullYear();
  const m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  const d = ('0' + dateObj.getDate()).slice(-2);
  return `${y}-${m}-${d}T12:00:00+09:00`;
}

function normalizeTaskTitle_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

/* =========================
   サブタスク関連 helpers
   ========================= */

/**
 * タスクタイプを判定: (main), (sub1), (sub2), ...
 * @return {string|null} 'main', 'sub1', 'sub2', ... または null
 */
function getTaskType_(title) {
  const s = String(title || '').trim();
  const m = s.match(/^\((main|sub\d+)\)/);
  return m ? m[1] : null;
}

/**
 * メインタスクか判定
 */
function isMainTask_(title) {
  return getTaskType_(title) === 'main';
}

/**
 * サブタスクか判定
 */
function isSubTask_(title) {
  const type = getTaskType_(title);
  return type && type.startsWith('sub');
}

/**
 * プレフィックスを削除したタイトルを取得
 * 例: "(main)タスク名" -> "タスク名"
 */
function removePrefixFromTitle_(title) {
  const s = String(title || '').trim();
  return s.replace(/^\((main|sub\d+)\)\s*/, '');
}

/**
 * サブタスク番号を取得
 * 例: "(sub1)" -> 1, "(sub2)" -> 2
 */
function getSubTaskNumber_(title) {
  const type = getTaskType_(title);
  if (!type || !type.startsWith('sub')) return 0;
  const m = type.match(/sub(\d+)/);
  return m ? Number(m[1]) : 0;
}

/**
 * サブタスクプレフィックスを生成
 * 例: 1 -> "(sub1)", 2 -> "(sub2)"
 */
function makeSubTaskPrefix_(index) {
  return `(sub${index})`;
}

/**
 * 繰越タイトル判定: 〜（繰越N日目）
 */
function isCarryoverTitle_(title) {
  const s = String(title || '').trim();
  return /（繰越\d+日目）$/.test(s);
}

/**
 * 繰越日数取得（通常タスクは0）
 */
function getCarryoverDays_(title) {
  const s = String(title || '').trim();
  const m = s.match(/（繰越(\d+)日目）$/);
  return m ? Number(m[1]) : 0;
}

/**
 * 比較用ベースタイトル（繰越サフィックス除去）
 */
function normalizeTaskBaseTitle_(title) {
  let s = String(title || '').trim();
  s = s.replace(/（繰越\d+日目）$/g, '');
  s = s.replace(/\s+/g, ' ');
  return s;
}

/**
 * 次の繰越タイトルを生成
 * 例: 数学 -> 数学（繰越1日目）
 *     数学（繰越1日目） -> 数学（繰越2日目）
 */
function makeCarryoverTitle_(title) {
  const base = normalizeTaskBaseTitle_(title);
  const currentDays = getCarryoverDays_(title);
  const nextDays = currentDays + 1;
  return `${base}（繰越${nextDays}日目）`;
}

function buildTaskNotes_(sheetName, row, titleCol) {
  return `${TASK_SYNC_CONFIG.SYNC_TAG}\n${sheetName} R${row}C${titleCol}`;
}

function ensureSyncTag_(notes, sheetName, row, titleCol) {
  if (String(notes).indexOf(TASK_SYNC_CONFIG.SYNC_TAG) !== -1) return notes;
  return `${TASK_SYNC_CONFIG.SYNC_TAG}\n${sheetName} R${row}C${titleCol}\n${notes}`.trim();
}