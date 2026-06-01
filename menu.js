function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('同期メニュー')
    // カレンダー同期 ─ 3モード（デフォルト: 1日更新）
    .addItem('カレンダー同期（今すぐ・1日更新）', 'syncToCalendarDay')
    .addItem('カレンダー同期（今すぐ・週末まで更新）', 'syncToCalendarWeekEnd')
    .addItem('カレンダー同期（今すぐ・1週間更新）', 'syncToCalendarWeek')
    .addSeparator()
    // Tasks同期
    .addItem('Tasks同期（今すぐ）', 'syncDailyTasksBidirectional')
    .addItem('Tasks→シート反映（今すぐ）', 'pollGoogleTasksCompletionToSheet')
    .addItem('Tasks同期トリガー設定', 'setupDailyTaskSyncTriggers')
    .addItem('Tasks色を更新', 'refreshTaskColors')
    .addItem('🚨 同期リセット (全削除・要確認)', 'resetDailyTaskSync')
    .addSeparator()
    // AI 週次プランナー
    .addItem('AI週次プランナーを開く', 'openPlannerDialog')
    .addToUi();
}