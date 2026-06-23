/************************************************************
 * TaskSync.gs（横持ちレイアウト版 / サブタスク機能廃止版）
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
 * - 色付け:
 *    * 繰越未完了 = 赤
 *    * 繰越系列が完了済み = 元/繰越とも橙（赤優先）
 * - 繰り越しタスクの自動非表示:
 *    * 同じベースタイトルの繰り越しタスクが複数日付に存在する場合、
 *      Google Tasks側では最も古い日付（元タスク）のみを表示し、
 *      新しい日付（繰り越し先）は完了化して非表示
 *
 * 自動繰り越し機能は現在無効化（手動で繰り越しを行うため）
 * Task IDは見える列には置かず、隠しシート _TaskSyncMap に保存
 *
 * ※サブタスク機能は廃止済み
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
  CHECKBOX_ROW_COUNT: 15,
  TITLE_WIDTH_SOURCE_COL: 67, // BO列

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
  if (!sheet) throw new Error(`sheet not found: ${TASK_SYNC_CONFIG.SHEET_NAME}`);

  const taskListId = getGoogleTaskListId_();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  // 必要日付の列だけ確保
  ensureDayPairColumnsThrough_(sheet, today);
  if (now.getHours() >= TASK_SYNC_CONFIG.TOMORROW_RELEASE_HOUR) {
    ensureDayPairColumnsThrough_(sheet, tomorrow);
  }

  // チェックボックスは今後15行までだけ付与
  sanitizeTaskPairColumns_(sheet);

  // 旧(main)/(sub1)などが残っていたら除去
  cleanupLegacyTaskPrefixes_(sheet);

  // Google Tasks -> Sheets
  syncGoogleTasksToSheet_(sheet, taskListId);
  importNewGoogleTasksToSheet_(sheet, taskListId);

  // Sheets -> Google Tasks
  syncSheetDayToGoogleTasks_(sheet, taskListId, today);
  if (now.getHours() >= TASK_SYNC_CONFIG.TOMORROW_RELEASE_HOUR) {
    syncSheetDayToGoogleTasks_(sheet, taskListId, tomorrow);
  }

  hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId);
  refreshTaskColors_();

  ss.toast(
    `Tasks synced: ${Utilities.formatDate(now, tz, 'M/d HH:mm')}`,
    'Google Tasks',
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
  if (!sheet) throw new Error(`sheet not found: ${TASK_SYNC_CONFIG.SHEET_NAME}`);

  const taskListId = getGoogleTaskListId_();

  // 運用日の列だけは足りなければ作る
  ensureDayPairColumnsThrough_(sheet, getOperationalToday_());

  // チェックボックスは15行までだけ
  sanitizeTaskPairColumns_(sheet);

  // Google Tasks -> Sheets
  syncGoogleTasksToSheet_(sheet, taskListId);
  importNewGoogleTasksToSheet_(sheet, taskListId);

  hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId);
  refreshTaskColors_();
}

/* ========================================
   自動繰り越し機能（現在無効化）
   手動で繰り越しを行うため、コメントアウト
   復活させる場合は以下のコメントを解除してください
   ======================================== */

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

  let tomorrowPair = findPairByDate_(sheet, tomorrow);
  if (!tomorrowPair) {
    tomorrowPair = createDayPairColumn_(sheet, tomorrow);
  }

  const lastRow = Math.max(sheet.getLastRow(), TASK_SYNC_CONFIG.START_ROW);
  const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;

  const todayTitles = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, todayPair.titleCol, numRows, 1).getValues() : [];
  const todayDones = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, todayPair.doneCol, numRows, 1).getValues() : [];
  const tomorrowTitles = numRows > 0 ? sheet.getRange(TASK_SYNC_CONFIG.START_ROW, tomorrowPair.titleCol, numRows, 1).getValues() : [];

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

  const tasksToCarryover = [];
  for (let i = 0; i < numRows; i++) {
    const row = TASK_SYNC_CONFIG.START_ROW + i;
    const title = String(todayTitles[i][0] || '').trim();
    if (!title) continue;

    const done = !!todayDones[i][0];
    if (done) continue;

    const carryTitle = makeCarryoverTitle_(title);

    const keyTitleBase = normalizeTaskTitle_(normalizeTaskBaseTitle_(title));
    const keyTitleCarry = normalizeTaskTitle_(carryTitle);
    if (tomorrowTitlesSet.has(keyTitleBase) || tomorrowTitlesSet.has(keyTitleCarry)) {
      markTaskAsCompletedInGoogleTasksByCell_(sheet, row, todayPair.titleCol, todayPair.doneCol, taskListId, map);
      continue;
    }

    tasksToCarryover.push({
      row: row,
      title: title,
      carryTitle: carryTitle,
      keyTitleBase: keyTitleBase,
      keyTitleCarry: keyTitleCarry
    });

    tomorrowTitlesSet.add(keyTitleBase);
    tomorrowTitlesSet.add(keyTitleCarry);

    markTaskAsCompletedInGoogleTasksByCell_(sheet, row, todayPair.titleCol, todayPair.doneCol, taskListId, map);
  }

  if (tasksToCarryover.length > 0) {
    let destRow = findFirstEmptyRowInColumn_(sheet, tomorrowPair.titleCol, TASK_SYNC_CONFIG.START_ROW);

    const titleValues = [];
    const doneValues = [];

    for (const task of tasksToCarryover) {
      titleValues.push([task.carryTitle]);
      doneValues.push([false]);
    }

    sheet.getRange(destRow, tomorrowPair.titleCol, titleValues.length, 1).setValues(titleValues);
    sheet.getRange(destRow, tomorrowPair.doneCol, doneValues.length, 1).setValues(doneValues);
  }

  saveTaskMapStore_(map);

  if (tasksToCarryover.length > 0) {
    let destRow = findFirstEmptyRowInColumn_(sheet, tomorrowPair.titleCol, TASK_SYNC_CONFIG.START_ROW);
    for (let i = 0; i < Math.min(tasksToCarryover.length, 50); i++) {
      syncSingleTaskCellRow_(sheet, destRow + i, tomorrowPair.titleCol, tomorrowPair.doneCol, tomorrow, taskListId);
    }
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

  ScriptApp.newTrigger('syncDailyTasksBidirectional')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('pollGoogleTasksCompletionToSheet')
    .timeBased()
    .everyMinutes(15)
    .create();

  /*
  ScriptApp.newTrigger('rolloverUncompletedTasksToTomorrow')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .nearMinute(0)
    .create();
  */

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
    ].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * 手動で色だけ再計算したい時用
 */
function refreshTaskColors() {
  refreshTaskColors_();
  SpreadsheetApp.getActiveSpreadsheet().toast('タスク色を更新しました', 'Tasks色更新', 3);
}

/**
 * 毎日のタスク同期を完全リセット。
 *
 * 状態がぐちゃぐちゃになった場合の最終手段。
 * 以下の3つを同時にクリアして「片側だけ残る → 重複作成」を防ぐ。
 *   1. _TaskSyncMap シートを削除 (マップ全消去)
 *   2. 「毎日のタスク」シートのタスク文字とチェックを全クリア
 *      (日付ヘッダ・チェックボックスの列構造は残す)
 *   3. Google Tasks の対象リスト内のタスクを全削除
 *
 * メニュー → 「同期リセット (全削除・要確認)」 から実行。
 * 二段階の確認ダイアログあり。
 */
function resetDailyTaskSync() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const confirm1 = ui.alert(
    '⚠️ 毎日のタスク同期リセット',
    '以下を一括実行します:\n\n' +
    '  ① マップシート (_TaskSyncMap) を削除\n' +
    '  ② 毎日のタスクシートのタスク文字＋チェックを全クリア\n' +
    '       (日付ヘッダとチェックボックス列は残します)\n' +
    '  ③ Google Tasks の対象リスト内タスクを全削除\n\n' +
    '元に戻せません。続行しますか?',
    ui.ButtonSet.YES_NO
  );
  if (confirm1 !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  const confirm2 = ui.alert(
    '本当に実行しますか?',
    '最終確認です。Google Tasks 側のタスクも消えます。',
    ui.ButtonSet.YES_NO
  );
  if (confirm2 !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  let mapDeleted = false;
  let cellsCleared = 0;
  let tasksDeleted = 0;
  let tasksFailed = 0;
  const errors = [];

  // ① マップシート削除
  try {
    const mapSheet = ss.getSheetByName(TASK_SYNC_CONFIG.MAP_SHEET_NAME);
    if (mapSheet) {
      ss.deleteSheet(mapSheet);
      mapDeleted = true;
    }
  } catch (e) {
    errors.push('マップ削除失敗: ' + e);
  }

  // ② 毎日のタスクシート: タスク文字＋チェックを全クリア
  try {
    const dailySheet = ss.getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
    if (dailySheet) {
      const pairs = getTaskColumnPairs_(dailySheet);
      const lastRow = dailySheet.getLastRow();
      if (lastRow >= TASK_SYNC_CONFIG.START_ROW) {
        const numRows = lastRow - TASK_SYNC_CONFIG.START_ROW + 1;
        for (const pair of pairs) {
          // タスク文字 (titleCol) を空に
          dailySheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.titleCol, numRows, 1).clearContent();
          // チェック (doneCol) は false にリセット (チェックボックスは保持)
          // 一度値を消してから再度チェックボックスを貼る
          const doneRange = dailySheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.doneCol, numRows, 1);
          doneRange.clearContent();
          // 既存セルが checkbox なら false 維持。範囲内の全セルに改めてチェックボックス保証
          const checkboxRows = Math.min(numRows, TASK_SYNC_CONFIG.CHECKBOX_ROW_COUNT);
          if (checkboxRows > 0) {
            const cbRange = dailySheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.doneCol, checkboxRows, 1);
            cbRange.clearDataValidations();
            cbRange.insertCheckboxes();
          }
          cellsCleared += numRows;
        }
      }
    }
  } catch (e) {
    errors.push('シートクリア失敗: ' + e);
  }

  // ③ Google Tasks のタスク全削除
  try {
    const taskListId = getGoogleTaskListId_();
    const allTasks = listAllGoogleTasksMap_(taskListId);
    for (const taskId in allTasks) {
      try {
        Tasks.Tasks.remove(taskListId, taskId);
        tasksDeleted++;
      } catch (e) {
        tasksFailed++;
        Logger.log('Google Tasks delete failed: ' + taskId + ' / ' + e);
      }
    }
  } catch (e) {
    errors.push('Google Tasks 削除失敗: ' + e);
  }

  const msg =
    '✅ リセット完了\n\n' +
    '・マップシート: ' + (mapDeleted ? '削除' : 'なし(または失敗)') + '\n' +
    '・シートクリア: ' + cellsCleared + ' セル\n' +
    '・Google Tasks 削除: ' + tasksDeleted + ' 件' + (tasksFailed ? ' (失敗 ' + tasksFailed + ')' : '') + '\n' +
    (errors.length ? '\n⚠️ エラー:\n' + errors.join('\n') : '');

  ui.alert('リセット完了', msg, ui.ButtonSet.OK);
}

/* =========================
   2列1セット（横持ち）処理
   ========================= */

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

function syncSingleCellTaskPairRow_(sheet, row, editedCol, taskListId) {
  const pair = normalizeToPair_(editedCol);
  if (!pair) return;

  const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
  const dueDate = parseDateOnly_(headerVal);
  if (!dueDate) return;

  syncSingleTaskCellRow_(sheet, row, pair.titleCol, pair.doneCol, dueDate, taskListId);
}

function syncTaskColumnPairToGoogleTasks_(sheet, titleCol, doneCol, dueDate, taskListId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < TASK_SYNC_CONFIG.START_ROW) return;

  for (let row = TASK_SYNC_CONFIG.START_ROW; row <= lastRow; row++) {
    const title = String(sheet.getRange(row, titleCol).getValue() || '').trim();
    if (!title) continue;
    syncSingleTaskCellRow_(sheet, row, titleCol, doneCol, dueDate, taskListId);
  }
}

function syncSingleTaskCellRow_(sheet, row, titleCol, doneCol, dueDate, taskListId) {
  const title = String(sheet.getRange(row, titleCol).getValue() || '').trim();
  const done = !!sheet.getRange(row, doneCol).getValue();
  if (!title) return;

  const cleanTitle = removePrefixFromTitle_(title);
  if (cleanTitle !== title) {
    sheet.getRange(row, titleCol).setValue(cleanTitle);
  }

  const map = getTaskMapStore_();
  const key = makeTaskMapKey_(sheet.getName(), row, titleCol, doneCol);
  const entry = map[key] || {};
  const taskId = (entry.taskId || '').trim();
  const dueIso = toTaskDueIso_(dueDate);

  if (!taskId) {
    const resource = {
      title: cleanTitle,
      notes: buildTaskNotes_(sheet.getName(), row, titleCol),
      due: dueIso,
      status: done ? 'completed' : 'needsAction',
    };
    if (done) resource.completed = new Date().toISOString();

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

  try {
    const task = Tasks.Tasks.get(taskListId, taskId);
    task.title = cleanTitle;
    task.due = dueIso;
    task.notes = ensureSyncTag_(task.notes || '', sheet.getName(), row, titleCol);

    if (task.parent) delete task.parent;

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
    Logger.log(`Task update failed key=${key}, recreating: ${err}`);

    const resource = {
      title: cleanTitle,
      notes: buildTaskNotes_(sheet.getName(), row, titleCol),
      due: dueIso,
      status: done ? 'completed' : 'needsAction',
    };
    if (done) resource.completed = new Date().toISOString();

    const recreated = Tasks.Tasks.insert(resource, taskListId);
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

function syncGoogleTasksToSheet_(sheet, taskListId) {
  const map = getTaskMapStore_();
  const keys = Object.keys(map);
  if (!keys.length) return;

  const taskMap = listAllGoogleTasksMap_(taskListId);
  const opToday = getOperationalToday_();

  keys.forEach(key => {
    const entry = map[key];
    if (!entry) return;
    if (entry.sheetName !== sheet.getName()) return;

    const taskId = entry.taskId;
    if (!taskId) return;

    const task = taskMap[taskId];
    if (!task) return;

    const row = Number(entry.row);
    const titleCol = Number(entry.titleCol);
    const doneCol = Number(entry.doneCol);

    if (row < TASK_SYNC_CONFIG.START_ROW || titleCol < 1 || doneCol < 1) return;

    const currentTitle = String(sheet.getRange(row, titleCol).getValue() || '').trim();
    if (!currentTitle) return;

    const headerVal = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).getValue();
    const cellDate = parseDateOnly_(headerVal);
    if (!cellDate) return;

    const shouldSyncDoneToSheet = !isBeforeDate_(cellDate, opToday);

    if (shouldSyncDoneToSheet) {
      const done = task.status === 'completed';
      sheet.getRange(row, doneCol).setValue(done);
    }

    map[key].updated = task.updated || '';
  });

  saveTaskMapStore_(map);
}

function importNewGoogleTasksToSheet_(sheet, taskListId) {
  const allTasks = listAllGoogleTasksMap_(taskListId);
  const map = getTaskMapStore_();
  const managedTaskIds = new Set(
    Object.keys(map)
      .map(k => (map[k] && map[k].taskId) ? String(map[k].taskId) : '')
      .filter(Boolean)
  );

  for (const taskId in allTasks) {
    const task = allTasks[taskId];
    if (!task || managedTaskIds.has(taskId)) continue;

    const normalizedTitle = removePrefixFromTitle_(task.title || '');
    if (!normalizedTitle) continue;

    // 1) 期日があってその日付列がある(or作れる) → その列ペア
    // 2) 期日が無い / 列が見つからない → 一番左の表示中ペア(=実質的な「今日」列)
    let pair = null;
    let dueDate = null;
    if (task.due) {
      dueDate = parseTaskDueToLocalDate_(task.due);
      if (dueDate) {
        pair = findPairByDate_(sheet, dueDate) || createDayPairColumn_(sheet, dueDate, map);
      }
    }
    if (!pair) {
      // 期日なし / 期日と一致する列が作れない → 「今日」の列に入れる
      // (古い日付を非表示にしている運用なので「今日」は実質的に左端の表示中列になる)
      const _now = new Date();
      const _today = new Date(_now.getFullYear(), _now.getMonth(), _now.getDate());
      pair = findPairByDate_(sheet, _today) || createDayPairColumn_(sheet, _today, map);
      if (!pair) {
        // どうしても作れなければ最終手段で左端表示中列
        pair = findLeftmostVisiblePair_(sheet);
        if (!pair) continue;
      }
    }

    let matchedRow = findRowByTitleInColumn_(sheet, pair.titleCol, normalizedTitle, TASK_SYNC_CONFIG.START_ROW);
    if (!matchedRow) {
      // 「次に使える行」を取得 (必要なら 2 行ペアでシート拡張: 日付マーカー行 + チェックボックス行)
      matchedRow = findOrCreateEmptyTaskRow_(sheet, pair.titleCol, pair.doneCol);
    } else {
      // 既存行を使う場合もチェックボックスは保証
      ensureCheckbox_(sheet, matchedRow, pair.doneCol);
    }

    sheet.getRange(matchedRow, pair.titleCol).setValue(normalizedTitle);
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
  }

  saveTaskMapStore_(map);
}

/**
 * 一番左の「表示中」(非表示でない) ペア列を返す。
 * 古い日付の列はユーザが非表示にしている運用なので、これが実質「今日」のCC列。
 */
function findLeftmostVisiblePair_(sheet) {
  const pairs = getTaskColumnPairs_(sheet);
  for (const pair of pairs) {
    // isColumnHiddenByUser は 1-indexed
    if (!sheet.isColumnHiddenByUser(pair.titleCol)) return pair;
  }
  // 全部非表示なら最後のペア
  return pairs.length ? pairs[pairs.length - 1] : null;
}

/**
 * doneCol の指定行にチェックボックス DataValidation が無ければ追加。
 */
function ensureCheckbox_(sheet, row, doneCol) {
  const cell = sheet.getRange(row, doneCol);
  const dv = cell.getDataValidation();
  const isCheckbox = dv && dv.getCriteriaType && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
  if (!isCheckbox) {
    cell.clearDataValidations();
    cell.insertCheckboxes();
  }
}

/**
 * 指定 pair (titleCol, doneCol) で「次に使える空タスク行」を返す。
 *
 * 1) 既存範囲で titleCol が空の行を探す。あればその行を返す（チェックボックスは保証）。
 * 2) 無ければシートを **2 行ずつ** 拡張する:
 *      ・拡張で増えた 1 行目 → titleCol にその列の日付ヘッダ値を書く（マーカー行）
 *      ・拡張で増えた 2 行目 → doneCol にチェックボックスを挿入（こちらが使える行）
 *    2 行目の行番号を返す。
 *
 * これにより毎回の拡張で「日付マーカー行 + チェックボックス付きタスク行」のペアが
 * 揃う。スクロールで下に行っても何の日付の塊か判別できるようになる。
 */
function findOrCreateEmptyTaskRow_(sheet, titleCol, doneCol) {
  const startRow = TASK_SYNC_CONFIG.START_ROW;
  const physicalMax = sheet.getMaxRows();

  if (physicalMax >= startRow) {
    const numRows = physicalMax - startRow + 1;
    const values = sheet.getRange(startRow, titleCol, numRows, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (!String(values[i][0] || '').trim()) {
        const row = startRow + i;
        ensureCheckbox_(sheet, row, doneCol);
        return row;
      }
    }
  }

  // 既存に空き無し → 2 行ペアで拡張
  sheet.insertRowsAfter(physicalMax, 2);
  const markerRow = physicalMax + 1;
  const taskRow   = physicalMax + 2;

  const dateHeader = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).getValue();
  sheet.getRange(markerRow, titleCol).setValue(dateHeader);

  ensureCheckbox_(sheet, taskRow, doneCol);

  return taskRow;
}

/**
 * 互換用シム: 旧 findFirstEmptyRowInColumn_ + 旧 ensureRowHasStructure_ の組合せを
 * findOrCreateEmptyTaskRow_ に集約。古い呼び出し元向けに 1引数版も維持する。
 *
 * ※ 旧版は doneCol を知らずに行番号だけ返していたが、このシムは doneCol を
 *    自動推定して構造を整える。
 */
function ensureRowHasStructure_(sheet, row, titleCol, doneCol) {
  if (!row || row < TASK_SYNC_CONFIG.START_ROW) return;
  // 物理行が足りなければ 2 行追加 + マーカー(日付)
  const physicalMax = sheet.getMaxRows();
  if (row > physicalMax) {
    sheet.insertRowsAfter(physicalMax, Math.max(row - physicalMax, 2));
    // 拡張で初めて生まれた最初の行をマーカーにする
    const dateHeader = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).getValue();
    sheet.getRange(physicalMax + 1, titleCol).setValue(dateHeader);
  }
  ensureCheckbox_(sheet, row, doneCol);
}

/* =========================
   繰越 / タスク直接操作 helpers
   ========================= */

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

function hideOlderCarryoverTasksInGoogleTasks_(sheet, taskListId) {
  const map = getTaskMapStore_();
  const pairs = getTaskColumnPairs_(sheet);
  if (!pairs.length) return;

  const carryoverGroups = {};

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
      if (!isCarryoverTitle_(title)) continue;

      const baseTitle = normalizeTaskBaseTitle_(title);
      const done = !!doneVals[i][0];

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

  for (const baseTitle in carryoverGroups) {
    const tasks = carryoverGroups[baseTitle];
    if (tasks.length <= 1) continue;

    tasks.sort((a, b) => a.date.getTime() - b.date.getTime());

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
    d.setDate(d.getDate() - 1);
  }
  return d;
}

/* =========================
   色付けロジック
   ========================= */

function refreshTaskColors_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SYNC_CONFIG.SHEET_NAME);
  if (!sheet) return;

  const pairs = getTaskColumnPairs_(sheet);
  if (!pairs.length) return;

  const lastRow = Math.max(sheet.getLastRow(), TASK_SYNC_CONFIG.START_ROW);
  const records = [];

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

  const groups = {};
  records.forEach(r => {
    const k = r.baseTitle;
    if (!groups[k]) groups[k] = [];
    groups[k].push(r);
  });

  const cellColorMap = {};

  records.forEach(r => {
    cellColorMap[`${r.row}:${r.titleCol}`] = TASK_SYNC_CONFIG.COLOR_NORMAL;
  });

  const opToday = getOperationalToday_();

  Object.keys(groups).forEach(base => {
    const list = groups[base];
    const hasCarry = list.some(r => r.isCarry);
    if (!hasCarry) return;

    const hasPastOrToday = list.some(r => !isBeforeDate_(opToday, r.date));
    if (!hasPastOrToday) return;

    const anyDone = list.some(r => r.done);
    const carryOpenExists = list.some(r => r.isCarry && !r.done && !isBeforeDate_(opToday, r.date));

    if (carryOpenExists) {
      list.forEach(r => {
        if (r.isCarry && !r.done && !isBeforeDate_(opToday, r.date)) {
          cellColorMap[`${r.row}:${r.titleCol}`] = TASK_SYNC_CONFIG.COLOR_LATE_OPEN;
        }
      });
    }

    if (anyDone) {
      list.forEach(r => {
        if (isBeforeDate_(opToday, r.date)) return;

        const key = `${r.row}:${r.titleCol}`;
        if (cellColorMap[key] === TASK_SYNC_CONFIG.COLOR_LATE_OPEN) return;
        cellColorMap[key] = TASK_SYNC_CONFIG.COLOR_LATE_DONE;
      });
    }
  });

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
  const maxCol = sheet.getLastColumn();
  if (maxCol < TASK_SYNC_CONFIG.START_COL) return pairs;

  const headerValues = sheet.getRange(
    TASK_SYNC_CONFIG.HEADER_ROW,
    TASK_SYNC_CONFIG.START_COL,
    1,
    maxCol - TASK_SYNC_CONFIG.START_COL + 1
  ).getValues()[0];

  for (let titleCol = TASK_SYNC_CONFIG.START_COL; titleCol <= maxCol; titleCol += TASK_SYNC_CONFIG.PAIR_WIDTH) {
    const doneCol = titleCol + 1;
    const header = headerValues[titleCol - TASK_SYNC_CONFIG.START_COL];
    const date = parseDateOnly_(header);
    if (!date) continue;

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

function createDayPairColumn_(sheet, dateObj, mapOpt) {
  const pairs = getTaskColumnPairs_(sheet);
  let titleCol, doneCol;

  if (pairs.length === 0) {
    titleCol = TASK_SYNC_CONFIG.START_COL;
    doneCol = titleCol + 1;
  } else {
    const nextPair = pairs.find(pair => {
      const header = sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue();
      const pairDate = parseDateOnly_(header);
      return pairDate && isBeforeDate_(dateObj, pairDate);
    });

    if (nextPair) {
      titleCol = nextPair.titleCol;
      sheet.insertColumnsBefore(titleCol, TASK_SYNC_CONFIG.PAIR_WIDTH);
      shiftTaskMapColumnsAfterInsertion_(
        sheet.getName(),
        titleCol,
        TASK_SYNC_CONFIG.PAIR_WIDTH,
        mapOpt
      );
    } else {
      const lastPair = pairs[pairs.length - 1];
      titleCol = lastPair.titleCol + TASK_SYNC_CONFIG.PAIR_WIDTH;
      sheet.insertColumnsAfter(lastPair.doneCol, TASK_SYNC_CONFIG.PAIR_WIDTH);
      shiftTaskMapColumnsAfterInsertion_(
        sheet.getName(),
        titleCol,
        TASK_SYNC_CONFIG.PAIR_WIDTH,
        mapOpt
      );
    }
    doneCol = titleCol + 1;
  }

  if (sheet.getMaxColumns() < doneCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), doneCol - sheet.getMaxColumns());
  }

  // ヘッダだけ設定
  sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, titleCol).setValue(formatDateHeader_(dateObj));

  // チェックボックスは今後15行だけ付与
  const rowCount = TASK_SYNC_CONFIG.CHECKBOX_ROW_COUNT; // = 15
  const doneRange = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, doneCol, rowCount, 1);
  doneRange.clearDataValidations();
  doneRange.insertCheckboxes();

  return { titleCol, doneCol };
}

function ensureDayPairColumnsThrough_(sheet, targetDate) {
  let pairs = getTaskColumnPairs_(sheet);
  if (pairs.length === 0) {
    createDayPairColumn_(sheet, targetDate);
    return;
  }

  const targetPair = findPairByDate_(sheet, targetDate);
  if (targetPair) {
    const targetIndex = pairs.findIndex(pair => pair.titleCol === targetPair.titleCol);
    const previousPair = targetIndex > 0 ? pairs[targetIndex - 1] : null;
    const isTailPair = targetIndex === pairs.length - 1;

    if (!previousPair || !isTailPair) return;

    const previousDate = parseDateOnly_(
      sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, previousPair.titleCol).getValue()
    );
    const expectedDate = previousDate ? addDays_(previousDate, 1) : null;

    if (!expectedDate || !isBeforeDate_(expectedDate, targetDate)) return;

    // 旧版は80列より右を走査できず、末尾の日付を今日で上書きしていた。
    // その列を本来の翌日に戻し、不足日をこの後で新しい列として補完する。
    sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, targetPair.titleCol)
      .setValue(formatDateHeader_(expectedDate));
    pairs = getTaskColumnPairs_(sheet);
  }

  let latestDateBeforeTarget = null;
  for (const pair of pairs) {
    const pairDate = parseDateOnly_(
      sheet.getRange(TASK_SYNC_CONFIG.HEADER_ROW, pair.titleCol).getValue()
    );
    if (!pairDate || !isBeforeDate_(pairDate, targetDate)) continue;
    if (!latestDateBeforeTarget || isBeforeDate_(latestDateBeforeTarget, pairDate)) {
      latestDateBeforeTarget = pairDate;
    }
  }

  if (!latestDateBeforeTarget) {
    createDayPairColumn_(sheet, targetDate);
    return;
  }

  let nextDate = new Date(latestDateBeforeTarget);
  while (isBeforeDate_(nextDate, targetDate)) {
    nextDate.setDate(nextDate.getDate() + 1);
    if (!findPairByDate_(sheet, nextDate)) {
      createDayPairColumn_(sheet, nextDate);
    }
  }
}

/**
 * 日付ペアの途中へ列を挿入した際、隠しマップ内の列番号とキーを追従させる。
 *
 * @param {string} sheetName - 列を挿入したシート名
 * @param {number} startCol - 挿入位置（1-indexed）
 * @param {number} columnCount - 挿入列数
 * @param {Object=} mapOpt - 呼び出し元が保持しているマップ
 */
function shiftTaskMapColumnsAfterInsertion_(sheetName, startCol, columnCount, mapOpt) {
  const map = mapOpt || getTaskMapStore_();
  const shiftedMap = {};

  Object.keys(map).forEach(key => {
    const entry = map[key];
    if (!entry) return;

    if (entry.sheetName === sheetName) {
      if (Number(entry.titleCol) >= startCol) entry.titleCol = Number(entry.titleCol) + columnCount;
      if (Number(entry.doneCol) >= startCol) entry.doneCol = Number(entry.doneCol) + columnCount;
    }

    const shiftedKey = makeTaskMapKey_(
      entry.sheetName,
      entry.row,
      entry.titleCol,
      entry.doneCol
    );
    shiftedMap[shiftedKey] = entry;
  });

  Object.keys(map).forEach(key => { delete map[key]; });
  Object.keys(shiftedMap).forEach(key => { map[key] = shiftedMap[key]; });
  saveTaskMapStore_(map);
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
  const physicalMax = sheet.getMaxRows();

  // 物理的に存在する範囲だけ走査する（trim されたシートで getRange が範囲外エラーを起こさないように）
  let scanRows = Math.min(maxRow - startRow + 100, 500);
  scanRows = Math.min(scanRows, Math.max(physicalMax - startRow + 1, 0));

  if (scanRows <= 0) {
    // 物理的に startRow すら無い → 拡張対象として startRow を返す
    return startRow;
  }

  const values = sheet.getRange(startRow, col, scanRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] || '').trim();
    if (!v) return startRow + i;
  }

  // 全部埋まっていた → スキャン直後の行を返す（呼び出し側で ensureRowHasStructure_ が拡張する）
  return startRow + scanRows;
}

function findRowByTitleInColumn_(sheet, col, title, startRow) {
  const target = normalizeTaskTitle_(title);
  const targetBase = normalizeTaskTitle_(normalizeTaskBaseTitle_(title));
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const numRows = lastRow - startRow + 1;

  if (numRows <= 0) return null;

  const values = sheet.getRange(startRow, col, numRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] || '').trim();
    if (!v) continue;

    const n = normalizeTaskTitle_(v);
    const b = normalizeTaskTitle_(normalizeTaskBaseTitle_(v));

    if (n === target || b === targetBase) return startRow + i;
  }
  return null;
}

function cleanupLegacyTaskPrefixes_(sheet) {
  const pairs = getTaskColumnPairs_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < TASK_SYNC_CONFIG.START_ROW) return;

  for (const pair of pairs) {
    for (let row = TASK_SYNC_CONFIG.START_ROW; row <= lastRow; row++) {
      const value = String(sheet.getRange(row, pair.titleCol).getValue() || '').trim();
      if (!value) continue;
      const cleaned = removePrefixFromTitle_(value);
      if (cleaned !== value) {
        sheet.getRange(row, pair.titleCol).setValue(cleaned);
      }
    }
  }
}

function sanitizeTaskPairColumns_(sheet) {
  const pairs = getTaskColumnPairs_(sheet);
  if (!pairs.length) return;

  // 今後チェックボックスを追加するのは15行まで
  const ensureRows = TASK_SYNC_CONFIG.CHECKBOX_ROW_COUNT; // = 15

  for (const pair of pairs) {
    const doneRange = sheet.getRange(TASK_SYNC_CONFIG.START_ROW, pair.doneCol, ensureRows, 1);

    // 値は消さず、15行分だけチェックボックスを付与
    doneRange.clearDataValidations();
    doneRange.insertCheckboxes();
  }
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

  let m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

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

function addDays_(dateObj, days) {
  const result = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  result.setDate(result.getDate() + days);
  return result;
}

function toTaskDueIso_(dateObj) {
  const y = dateObj.getFullYear();
  const m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  const d = ('0' + dateObj.getDate()).slice(-2);
  return `${y}-${m}-${d}T12:00:00+09:00`;
}

function normalizeTaskTitle_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

/* =========================
   タイトル / 繰越 helpers
   ========================= */

/**
 * 旧サブタスク用プレフィックスを削除
 * 例: "(main)タスク名" -> "タスク名"
 *     "(sub1)タスク名" -> "タスク名"
 */
function removePrefixFromTitle_(title) {
  const s = String(title || '').trim();
  return s.replace(/^\((main|sub\d+)\)\s*/, '');
}

function isCarryoverTitle_(title) {
  const s = String(title || '').trim();
  return /（繰越\d+日目）$/.test(s);
}

function getCarryoverDays_(title) {
  const s = String(title || '').trim();
  const m = s.match(/（繰越(\d+)日目）$/);
  return m ? Number(m[1]) : 0;
}

function normalizeTaskBaseTitle_(title) {
  let s = String(title || '').trim();
  s = removePrefixFromTitle_(s);
  s = s.replace(/（繰越\d+日目）$/g, '');
  s = s.replace(/\s+/g, ' ');
  return s;
}

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