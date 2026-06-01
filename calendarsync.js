/**
 * ============================================================
 * Google スプレッドシート → Google カレンダー 自動同期スクリプト
 * ============================================================
 *
 * 【スプレッドシートの想定レイアウト（「週予定」シート）】
 *
 *        |   A    |   B        |     C      |     D      | ...
 *   -----+--------+------------+------------+------------+----
 *   1行目|        |試験まであと| 2026/1/12  | 2026/1/13  | ...  ← 日付（C列以降）
 *   2行目| 授業   |            | 数学α,物理 | 漢文       | ...  ← 授業一覧
 *   3行目|   8    |            | 青チャIIIC |            | ...  ← 8時の予定
 *   4行目|   9    |            |            |            | ...  ← 9時の予定
 *   ...  |  ...   |   ...      |   ...      |   ...      | ...
 *  19行目|  24    |            |(J PREP単語)|            | ...  ← 24時の予定
 *
 * 【機能】
 *   - ★ 表示中の列（非表示でない列）だけを同期対象にする
 *     → 過去の非表示週は無視し、今週分だけをカレンダーに反映
 *     → 過去にカレンダーに追加済みの予定は削除しない
 *   - 結合セルを自動検出し、結合範囲全体で1つのイベントとして扱う
 *   - 同じ予定名が連続する時間帯は1つのイベントにまとめる
 *   - 「@場所名」を含むセルは場所情報としてカレンダーに反映
 *   - セル内改行で複数予定が書かれている場合、それぞれ別イベントとして作成
 *   - 表示中の日付範囲のカレンダー予定を毎回クリア＆再作成（上書きモード）
 *   - トリガーで自動実行（毎日 or 毎時間）
 *   - 「授業」行は終日イベントとして登録（オプション）
 *
 * 【セットアップ】
 *   1. スプレッドシートの「拡張機能」→「Apps Script」を開く
 *   2. このコードを貼り付けて保存
 *   3. syncToCalendar() を一度手動実行して権限を承認
 *   4. setupDailyTrigger() を実行してトリガーを設定
 */

// ==================== 設定 ====================

const CONFIG = {
  // 同期先カレンダーID（デフォルトのカレンダーを使う場合は 'primary'）
  // 特定のカレンダーに同期したい場合はカレンダーIDを指定
  // 例: 'abc123@group.calendar.google.com'
  CALENDAR_ID: 'primary',

  // シート名（スプレッドシートのタブ名に合わせてください）
  SHEET_NAME: '週予定',

  // スクリプトが作成した予定を識別するためのタグ（説明欄に付与）
  SYNC_TAG: '[スプレッドシート同期]',

  // 1コマの時間（分）。時間が1時間刻みなら60、30分刻みなら30
  SLOT_DURATION_MINUTES: 60,

  // データの開始位置（0-indexed）
  DATE_ROW: 0,           // 1行目 = 日付
  CLASS_ROW: 1,          // 2行目 = 授業一覧
  TIME_START_ROW: 2,     // 3行目 = 最初の時間スロット（8時）
  TIME_COL: 0,           // A列 = 時間
  DATA_START_COL: 1,     // B列から走査開始（B列の「試験まであと」は日付として認識されずスキップ）

  // 「授業」行を終日イベントとして登録するか
  SYNC_CLASS_ROW: true,

  // カレンダーに登録しないキーワード（完全一致）
  // ※「休み」は授業が休みという意味なのでスキップしない
  SKIP_KEYWORDS: [],

  // カッコ付きの予定（例:「(J PREP単語)」）もカレンダーに登録するか
  INCLUDE_PARENTHESIZED: true,
};


// ==================== メイン同期関数 ====================

/**
 * 同期モードの定数
 *   'day'      : 今日1日だけ更新（デフォルト）
 *   'week_end' : 今日〜今週末（日曜）まで更新
 *   'week'     : 表示中の1週間すべてを更新（従来の動作）
 */
const SYNC_MODE = {
  DAY: 'day',
  WEEK_END: 'week_end',
  WEEK: 'week',
};

/** 1日更新（デフォルト）のエントリーポイント */
function syncToCalendarDay()     { syncToCalendar(SYNC_MODE.DAY); }
/** 週末まで更新のエントリーポイント */
function syncToCalendarWeekEnd() { syncToCalendar(SYNC_MODE.WEEK_END); }
/** 1週間更新のエントリーポイント */
function syncToCalendarWeek()    { syncToCalendar(SYNC_MODE.WEEK); }

/**
 * スプレッドシートの予定をGoogleカレンダーに同期する
 *
 * @param {string} mode - SYNC_MODE の値。省略時は 'day'（1日更新）
 */
function syncToCalendar(mode) {
  mode = mode || SYNC_MODE.DAY;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error(
      `シート「${CONFIG.SHEET_NAME}」が見つかりません。CONFIG.SHEET_NAME を確認してください。`
    );
  }

  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    throw new Error(
      `カレンダー（ID: ${CONFIG.CALENDAR_ID}）が見つかりません。CONFIG.CALENDAR_ID を確認してください。`
    );
  }

  // ★ 結合セルの値を展開する（結合範囲の全セルに同じ値を入れる）
  fillMergedCells(sheet, data);

  // ★ 表示中（非表示でない）の列だけを取得する
  const visibleCols = getVisibleColumns(sheet, CONFIG.DATA_START_COL, data[0].length);
  Logger.log(`表示中の列: ${visibleCols.length} 列 / 全 ${data[0].length - CONFIG.DATA_START_COL} 列`);

  // 日付一覧を取得（表示中の列のみ）
  let dates = [];
  for (const col of visibleCols) {
    const cell = data[CONFIG.DATE_ROW][col];
    dates.push({ col: col, date: cell ? parseDate(cell) : null });
  }

  const validEntries = dates.filter(d => d.date !== null);
  if (validEntries.length === 0) {
    Logger.log('表示中の列に日付が見つかりませんでした。');
    return;
  }

  // ★ モードに応じて対象日付を絞り込む
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (mode === SYNC_MODE.DAY) {
    // 今日のみ
    dates = dates.filter(d => {
      if (!d.date) return false;
      const dd = new Date(d.date);
      dd.setHours(0, 0, 0, 0);
      return dd.getTime() === today.getTime();
    });
    Logger.log('同期モード: 1日更新（今日のみ）');

  } else if (mode === SYNC_MODE.WEEK_END) {
    // 今日〜今週の日曜（dayOfWeek 0）まで
    const dayOfWeek = today.getDay(); // 0=日, 1=月, ..., 6=土
    const daysUntilSunday = dayOfWeek === 0 ? 0 : 7 - dayOfWeek;
    const sunday = new Date(today);
    sunday.setDate(today.getDate() + daysUntilSunday);
    sunday.setHours(23, 59, 59, 999);

    dates = dates.filter(d => {
      if (!d.date) return false;
      const dd = new Date(d.date);
      dd.setHours(0, 0, 0, 0);
      return dd.getTime() >= today.getTime() && d.date <= sunday;
    });
    Logger.log(`同期モード: 週末まで更新（今日〜${Utilities.formatDate(sunday, Session.getScriptTimeZone(), 'M/d')}）`);

  } else {
    // WEEK: 表示中の全日付（従来の動作）
    Logger.log('同期モード: 1週間更新（表示中の全日付）');
  }

  if (dates.filter(d => d.date).length === 0) {
    Logger.log('指定されたモードで同期対象の日付が見つかりませんでした。');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '同期対象の日付が見つかりませんでした（シートに今日以降の日付がありますか？）',
      '同期スキップ',
      5
    );
    return;
  }

  // 時間一覧を取得（A列の3行目以降）
  const times = parseTimesFromColumn(data, CONFIG.TIME_COL, CONFIG.TIME_START_ROW);
  if (times.length === 0) {
    Logger.log('時間が見つかりませんでした。A列に時間が入っているか確認してください。');
    return;
  }

  // ★ 絞り込み後の日付範囲のみ既存の同期イベントを削除
  const targetDates = dates.filter(d => d.date !== null).map(d => d.date);
  deleteExistingSyncedEvents(calendar, targetDates);

  let createdCount = 0;

  // 対象の各日付列について予定を作成
  for (const entry of dates) {
    if (!entry.date) continue;

    // 「授業」行を終日イベントとして登録
    if (CONFIG.SYNC_CLASS_ROW) {
      const classCell = data[CONFIG.CLASS_ROW][entry.col];
      if (classCell) {
        const className = String(classCell).trim();
        if (className && !shouldSkip(className)) {
          createAllDayEvent(calendar, entry.date, className);
          createdCount++;
        }
      }
    }

    // 時間スロットの予定を作成
    const events = buildEventsForColumn(data, entry.col, entry.date, times);
    for (const event of events) {
      createCalendarEvent(calendar, event);
      createdCount++;
    }
  }

  const modeLabel = mode === SYNC_MODE.DAY ? '1日' : mode === SYNC_MODE.WEEK_END ? '週末まで' : '1週間';
  Logger.log(`同期完了（${modeLabel}）: ${createdCount} 件のイベントを作成しました。`);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${createdCount} 件の予定をカレンダーに同期しました（${modeLabel}更新）`,
    '同期完了',
    5
  );
}


// ==================== 表示列の検出 ====================

/**
 * 表示中（非表示でない）の列インデックスを取得する
 * sheet.isColumnHiddenByUser() で非表示かどうか判定
 *
 * @param {Sheet} sheet - 対象シート
 * @param {number} startCol - 走査開始列（0-indexed）
 * @param {number} totalCols - 全列数
 * @returns {number[]} 表示中の列インデックス（0-indexed）の配列
 */
function getVisibleColumns(sheet, startCol, totalCols) {
  const visible = [];
  for (let col = startCol; col < totalCols; col++) {
    // isColumnHiddenByUser は 1-indexed
    if (!sheet.isColumnHiddenByUser(col + 1)) {
      visible.push(col);
    }
  }
  return visible;
}


// ==================== 結合セル処理 ====================

/**
 * 結合セルの値を結合範囲全体に展開する
 *
 * Google Sheets の getValues() は結合セルの左上のみ値を返し、
 * 残りは空文字になる。この関数で結合範囲すべてに同じ値を入れることで、
 * 後続の「連続する同じ予定名をまとめる」ロジックが正しく動作する。
 *
 * 例: 「鳥コン作業 @有楽町」が P5:P11 で結合されている場合
 *   → data[4][15]〜data[10][15] すべてに「鳥コン作業\n@有楽町」が入る
 *   → buildEventsForDate で 10:00〜17:00 の1イベントとして構築される
 */
function fillMergedCells(sheet, data) {
  const dataRange = sheet.getDataRange();
  const mergedRanges = dataRange.getMergedRanges();

  for (const range of mergedRanges) {
    const startRow = range.getRow() - 1;     // 1-indexed → 0-indexed
    const startCol = range.getColumn() - 1;
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    const value = data[startRow][startCol];   // 左上セルの値

    // 結合範囲の全セルに同じ値をセット
    for (let r = startRow; r < startRow + numRows; r++) {
      for (let c = startCol; c < startCol + numCols; c++) {
        if (r < data.length && c < data[r].length) {
          data[r][c] = value;
        }
      }
    }
  }

  Logger.log(`結合セル ${mergedRanges.length} 箇所を展開しました。`);
}


// ==================== 日付・時間のパース ====================

/**
 * 1行目から日付を読み取る
 */
function parseDatesFromRow(row, startCol) {
  const dates = [];
  for (let i = startCol; i < row.length; i++) {
    const cell = row[i];
    if (!cell) {
      dates.push(null);
      continue;
    }
    dates.push(parseDate(cell));
  }
  return dates;
}

/**
 * セルの値を Date オブジェクトに変換
 * 対応形式: Date オブジェクト / "2026/2/15/土" / "2026/2/15" / "2/15" / "2月15日"
 */
function parseDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  const str = String(value).trim();

  // 「2026/2/15」「2026/2/15/土」「2026-02-15」形式
  const fullMatch = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (fullMatch) {
    return new Date(Number(fullMatch[1]), Number(fullMatch[2]) - 1, Number(fullMatch[3]));
  }

  // 「2/15」「2月15日」形式（年は現在年を補完）
  const shortMatch = str.match(/(\d{1,2})[\/月](\d{1,2})/);
  if (shortMatch) {
    const year = new Date().getFullYear();
    return new Date(year, Number(shortMatch[1]) - 1, Number(shortMatch[2]));
  }

  return null;
}

/**
 * A列から時間を読み取る（3行目以降）
 */
function parseTimesFromColumn(data, timeCol, startRow) {
  const times = [];
  for (let i = startRow; i < data.length; i++) {
    const cell = data[i][timeCol];
    if (!cell && cell !== 0) {
      times.push(null);
      continue;
    }
    times.push(parseTime(cell));
  }
  return times;
}

/**
 * セルの値を { hours, minutes } に変換
 * 対応形式: 数値(8, 13) / "8:00" / "8時" / "8時30分" / Dateオブジェクト
 */
function parseTime(value) {
  if (typeof value === 'number') {
    return { hours: Math.floor(value), minutes: 0 };
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return { hours: value.getHours(), minutes: value.getMinutes() };
  }

  const str = String(value).trim();

  const colonMatch = str.match(/(\d{1,2})[：:](\d{2})/);
  if (colonMatch) {
    return { hours: Number(colonMatch[1]), minutes: Number(colonMatch[2]) };
  }

  const jpMatch = str.match(/(\d{1,2})時(?:(\d{1,2})分)?/);
  if (jpMatch) {
    return { hours: Number(jpMatch[1]), minutes: jpMatch[2] ? Number(jpMatch[2]) : 0 };
  }

  const numMatch = str.match(/^(\d{1,2})$/);
  if (numMatch) {
    return { hours: Number(numMatch[1]), minutes: 0 };
  }

  return null;
}


// ==================== イベント構築 ====================

/**
 * セルの値から予定情報を抽出する
 * セル内改行で複数の予定がある場合は配列で返す
 * 「@場所名」は場所情報として分離する
 *
 * 例: "鳥コン作業\n@有楽町" → [{ name: "鳥コン作業", location: "有楽町" }]
 * 例: "青チャIIIC\n鳥コン作業 @有楽町" → [{ name: "青チャIIIC" }, { name: "鳥コン作業", location: "有楽町" }]
 */
function parseCellContent(cellValue) {
  if (!cellValue) return [];

  const text = String(cellValue).trim();
  if (!text) return [];

  const lines = text.split(/\n/).map(l => l.trim()).filter(l => l);

  const results = [];
  let pendingEvent = null;

  for (const line of lines) {
    // 「@場所」だけの行 → 直前のイベントの場所情報
    if (line.match(/^[@＠]/)) {
      const location = line.replace(/^[@＠]\s*/, '');
      if (pendingEvent) {
        pendingEvent.location = location;
      }
      continue;
    }

    if (pendingEvent) {
      results.push(pendingEvent);
    }

    // 「イベント名 @場所」形式
    const locationMatch = line.match(/^(.+?)\s*[@＠]\s*(.+)$/);
    if (locationMatch) {
      pendingEvent = {
        name: locationMatch[1].trim(),
        location: locationMatch[2].trim(),
      };
    } else {
      pendingEvent = { name: line, location: '' };
    }
  }

  if (pendingEvent) {
    results.push(pendingEvent);
  }

  return results;
}

/**
 * 予定をスキップすべきか判定
 */
function shouldSkip(name) {
  if (CONFIG.SKIP_KEYWORDS.includes(name)) return true;
  if (!CONFIG.INCLUDE_PARENTHESIZED && /^[（(].+[）)]$/.test(name)) return true;
  return false;
}

/**
 * 1日分の予定一覧を構築する
 * - fillMergedCells() 済みなので、結合セルは全行に同じ値が入っている
 * - 連続する同じ予定名（結合セル含む）は1つのイベントにまとめる
 * - セル内改行で複数予定がある場合は、1時間内で均等分割する
 *
 * @param {Array[]} data - シートの全データ（2次元配列）
 * @param {number} col - 対象列の絶対インデックス（0-indexed）
 * @param {Date} date - この列の日付
 * @param {Object[]} times - 時間スロットの配列
 */
function buildEventsForColumn(data, col, date, times) {
  const events = [];
  const actualCol = col;
  let currentEvent = null;

  for (let i = 0; i < times.length; i++) {
    const rowIdx = CONFIG.TIME_START_ROW + i;
    if (rowIdx >= data.length) break;

    const time = times[i];
    if (!time) continue;

    const cellValue = data[rowIdx][actualCol];
    const parsedItems = parseCellContent(cellValue);
    const validItems = parsedItems.filter(item => !shouldSkip(item.name));

    if (validItems.length === 0) {
      if (currentEvent) {
        events.push(currentEvent);
        currentEvent = null;
      }
      continue;
    }

    if (validItems.length === 1) {
      const item = validItems[0];

      // 結合セルの場合、展開後は同じ name + location が連続するのでここでまとまる
      if (currentEvent && currentEvent.name === item.name && currentEvent.location === item.location) {
        currentEvent.endTime = addMinutes(date, time.hours, time.minutes, CONFIG.SLOT_DURATION_MINUTES);
      } else {
        if (currentEvent) {
          events.push(currentEvent);
        }
        currentEvent = {
          name: item.name,
          location: item.location || '',
          startTime: createDateTime(date, time.hours, time.minutes),
          endTime: addMinutes(date, time.hours, time.minutes, CONFIG.SLOT_DURATION_MINUTES),
        };
      }
    } else {
      // 複数の予定がセル内にある → 前の連続イベントを確定
      if (currentEvent) {
        events.push(currentEvent);
        currentEvent = null;
      }

      const slotMinutes = CONFIG.SLOT_DURATION_MINUTES;
      const perItem = Math.floor(slotMinutes / validItems.length);

      for (let j = 0; j < validItems.length; j++) {
        const item = validItems[j];
        const startMin = time.minutes + (perItem * j);
        const endMin = (j === validItems.length - 1)
          ? time.minutes + slotMinutes
          : time.minutes + (perItem * (j + 1));

        events.push({
          name: item.name,
          location: item.location || '',
          startTime: createDateTime(date, time.hours, startMin),
          endTime: createDateTime(date, time.hours, endMin),
        });
      }
    }
  }

  if (currentEvent) {
    events.push(currentEvent);
  }

  return events;
}

function createDateTime(date, hours, minutes) {
  const dt = new Date(date);
  dt.setHours(hours, minutes, 0, 0);
  return dt;
}

function addMinutes(date, hours, minutes, addMin) {
  const dt = new Date(date);
  dt.setHours(hours, minutes + addMin, 0, 0);
  return dt;
}


// ==================== カレンダー操作 ====================

/**
 * 同期タグが付いた既存イベントを削除する
 */
function deleteExistingSyncedEvents(calendar, dates) {
  const validDates = dates.filter(d => d !== null);
  if (validDates.length === 0) return;

  const minDate = new Date(Math.min(...validDates.map(d => d.getTime())));
  const maxDate = new Date(Math.max(...validDates.map(d => d.getTime())));

  minDate.setHours(0, 0, 0, 0);
  maxDate.setDate(maxDate.getDate() + 1);
  maxDate.setHours(23, 59, 59, 999);

  const existingEvents = calendar.getEvents(minDate, maxDate);
  let deletedCount = 0;

  for (const event of existingEvents) {
    const description = event.getDescription() || '';
    if (description.includes(CONFIG.SYNC_TAG)) {
      event.deleteEvent();
      deletedCount++;
    }
  }

  Logger.log(`既存の同期イベント ${deletedCount} 件を削除しました。`);
}

/**
 * カレンダーにイベントを作成する（時間指定）
 */
function createCalendarEvent(calendar, event) {
  const options = {
    description: CONFIG.SYNC_TAG + '\nスプレッドシートから自動同期された予定です。',
  };
  if (event.location) {
    options.location = event.location;
  }

  calendar.createEvent(event.name, event.startTime, event.endTime, options);

  Logger.log(
    `作成: ${event.name}${event.location ? ' @' + event.location : ''} ` +
    `(${formatDateTime(event.startTime)} - ${formatDateTime(event.endTime)})`
  );
}

/**
 * カレンダーに終日イベントを作成する（授業行用）
 */
function createAllDayEvent(calendar, date, title) {
  calendar.createAllDayEvent(title, date, {
    description: CONFIG.SYNC_TAG + '\nスプレッドシートの「授業」行から自動同期。',
  });
  Logger.log(`作成（終日）: ${title} (${Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d')})`);
}

function formatDateTime(dt) {
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'M/d HH:mm');
}


// ==================== トリガー設定 ====================

/**
 * 毎日自動同期するトリガーを設定する（毎朝6時に実行）
 */
function setupDailyTrigger() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncToCalendar')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  Logger.log('毎日6時に自動同期するトリガーを設定しました。');
  SpreadsheetApp.getActiveSpreadsheet().toast('毎日朝6時に自動同期するトリガーを設定しました', 'トリガー設定完了', 5);
}

/**
 * 毎時間自動同期するトリガーを設定する
 */
function setupHourlyTrigger() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncToCalendar')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('毎時間自動同期するトリガーを設定しました。');
  SpreadsheetApp.getActiveSpreadsheet().toast('毎時間自動同期するトリガーを設定しました', 'トリガー設定完了', 5);
}

/**
 * スプレッドシート編集時に自動同期するトリガーを設定する
 */
function setupOnEditTrigger() {
  removeExistingTriggers();
  ScriptApp.newTrigger('syncToCalendar')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  Logger.log('スプレッドシート編集時に自動同期するトリガーを設定しました。');
  SpreadsheetApp.getActiveSpreadsheet().toast('スプレッドシート編集時に自動同期するトリガーを設定しました', 'トリガー設定完了', 5);
}

function removeExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'syncToCalendar') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function stopAutoSync() {
  removeExistingTriggers();
  Logger.log('自動同期を停止しました。');
  SpreadsheetApp.getActiveSpreadsheet().toast('自動同期を停止しました', '停止完了', 5);
}