/**
 * ユニットテストコード
 *
 * 【使い方】
 * 1. GASのスクリプトエディタで「runAllTests」を実行
 * 2. ログを確認してテスト結果をチェック
 * 3. すべてのテストが通れば「すべてのテストが成功しました」と表示される
 *
 * @author YoyogiPinball
 * @version 1.0
 */

// ================================================================================
// テストランナー
// ================================================================================

/**
 * すべてのテストを実行
 *
 * この関数をGASのスクリプトエディタから実行してテストを開始します
 */
function runAllTests() {
  console.log('========================================');
  console.log('ユニットテスト開始');
  console.log('========================================\n');

  const tests = [
    testShouldExecute,
    testGetActionLabel,
    testShowDryRunSummary,
    testMockScheduleDataStructure,
    testMockConfigStructure,
    testParseScheduleDate,
    testFilterTodayTomorrowSchedules
  ];

  let passedCount = 0;
  let failedCount = 0;
  const failedTests = [];

  for (const test of tests) {
    const testName = test.name;
    try {
      console.log(`テスト実行: ${testName}`);
      test();
      console.log(`✓ ${testName} - 成功\n`);
      passedCount++;
    } catch (error) {
      console.error(`✗ ${testName} - 失敗`);
      console.error(`  エラー: ${error.message}\n`);
      failedCount++;
      failedTests.push({ name: testName, error: error.message });
    }
  }

  console.log('========================================');
  console.log('テスト結果サマリー');
  console.log('========================================');
  console.log(`成功: ${passedCount}/${tests.length}`);
  console.log(`失敗: ${failedCount}/${tests.length}`);

  if (failedCount > 0) {
    console.log('\n失敗したテスト:');
    failedTests.forEach(t => {
      console.log(`  - ${t.name}: ${t.error}`);
    });
    throw new Error(`${failedCount}件のテストが失敗しました`);
  } else {
    console.log('\n🎉 すべてのテストが成功しました！');
  }
}

// ================================================================================
// A案: 統一dry-run管理機能のテスト
// ================================================================================

/**
 * shouldExecute() 関数のテスト
 */
function testShouldExecute() {
  // テストケース1: DRY_RUN_SPREADSHEET=trueの場合、falseを返す
  const config1 = getMockConfig({ DRY_RUN_SPREADSHEET: true });
  const result1 = shouldExecute('SPREADSHEET', config1);
  assertFalse(result1, 'DRY_RUN_SPREADSHEET=trueの場合、falseを返すべき');

  // テストケース2: DRY_RUN_SPREADSHEET=falseの場合、trueを返す
  const config2 = getMockConfig({ DRY_RUN_SPREADSHEET: false });
  const result2 = shouldExecute('SPREADSHEET', config2);
  assertTrue(result2, 'DRY_RUN_SPREADSHEET=falseの場合、trueを返すべき');

  // テストケース3: すべてのアクションタイプで動作する
  const config3 = getMockConfig({
    DRY_RUN_CALENDAR: true,
    DRY_RUN_DISCORD: false,
    DRY_RUN_FILE_MOVE: true
  });
  assertFalse(shouldExecute('CALENDAR', config3), 'CALENDARがスキップされるべき');
  assertTrue(shouldExecute('DISCORD', config3), 'DISCORDが実行されるべき');
  assertFalse(shouldExecute('FILE_MOVE', config3), 'FILE_MOVEがスキップされるべき');
}

/**
 * getActionLabel() 関数のテスト
 */
function testGetActionLabel() {
  assertEqual(
    getActionLabel('SPREADSHEET'),
    'スプレッドシート書き込み',
    'SPREADSHEETのラベルが正しいこと'
  );
  assertEqual(
    getActionLabel('CALENDAR'),
    'カレンダー登録',
    'CALENDARのラベルが正しいこと'
  );
  assertEqual(
    getActionLabel('DISCORD'),
    'Discord通知',
    'DISCORDのラベルが正しいこと'
  );
  assertEqual(
    getActionLabel('FILE_MOVE'),
    'ファイル移動',
    'FILE_MOVEのラベルが正しいこと'
  );
  assertEqual(
    getActionLabel('UNKNOWN'),
    'UNKNOWN',
    '未知のアクションタイプはそのまま返すこと'
  );
}

/**
 * showDryRunSummary() 関数のテスト（実行時にエラーが出ないことを確認）
 */
function testShowDryRunSummary() {
  const config1 = getMockConfig({ DRY_RUN: true, TEST_MODE: false });
  showDryRunSummary(config1);  // エラーが出なければOK

  const config2 = getMockConfig({ DRY_RUN: false, TEST_MODE: true });
  showDryRunSummary(config2);  // エラーが出なければOK
}

// ================================================================================
// モックデータのテスト
// ================================================================================

/**
 * getMockScheduleData() の構造をテスト
 */
function testMockScheduleDataStructure() {
  const schedules = getMockScheduleData();

  // 配列であることを確認
  assertTrue(Array.isArray(schedules), 'スケジュールデータは配列であるべき');

  // 少なくとも1件のデータがあることを確認
  assertTrue(schedules.length > 0, 'スケジュールデータは1件以上あるべき');

  // 最初のデータの構造を確認
  const firstSchedule = schedules[0];
  assertHasProperty(firstSchedule, 'vtuber', 'vtuberプロパティが必要');
  assertHasProperty(firstSchedule, 'affiliation', 'affiliationプロパティが必要');
  assertHasProperty(firstSchedule, 'date', 'dateプロパティが必要');
  assertHasProperty(firstSchedule, 'day', 'dayプロパティが必要');
  assertHasProperty(firstSchedule, 'time', 'timeプロパティが必要');
  assertHasProperty(firstSchedule, 'content', 'contentプロパティが必要');
  assertHasProperty(firstSchedule, 'note', 'noteプロパティが必要');
}

/**
 * getMockConfig() の構造をテスト
 */
function testMockConfigStructure() {
  const config = getMockConfig();

  // 必須項目が含まれているか確認
  assertHasProperty(config, 'GEMINI_API_KEY', 'GEMINI_API_KEYが必要');
  assertHasProperty(config, 'INPUT_FOLDER_ID', 'INPUT_FOLDER_IDが必要');
  assertHasProperty(config, 'DONE_FOLDER_ID', 'DONE_FOLDER_IDが必要');
  assertHasProperty(config, 'SPREADSHEET_ID', 'SPREADSHEET_IDが必要');
  assertHasProperty(config, 'CALENDAR_ID', 'CALENDAR_IDが必要');
  assertHasProperty(config, 'DRY_RUN', 'DRY_RUNが必要');

  // オーバーライドが機能するか確認
  const customConfig = getMockConfig({ DRY_RUN: false });
  assertFalse(customConfig.DRY_RUN, 'オーバーライドが機能すべき');
}

// ================================================================================
// 日付処理のテスト
// ================================================================================

/**
 * 日付文字列のパースをテスト
 */
function testParseScheduleDate() {
  // 正常な日付文字列
  const validDate = '2025-12-15';
  const parsed = new Date(validDate);
  assertTrue(parsed instanceof Date, '日付オブジェクトが生成されるべき');
  assertTrue(!isNaN(parsed.getTime()), '有効な日付であるべき');

  // 無効な日付文字列
  const invalidDate = 'invalid-date';
  const parsedInvalid = new Date(invalidDate);
  assertTrue(isNaN(parsedInvalid.getTime()), '無効な日付はNaNになるべき');
}

/**
 * 今日・明日のスケジュールフィルタリングをテスト
 */
function testFilterTodayTomorrowSchedules() {
  const schedules = getMockScheduleData();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  const todaySchedules = [];
  const tomorrowSchedules = [];

  for (const schedule of schedules) {
    const scheduleDate = new Date(schedule.date);
    scheduleDate.setHours(0, 0, 0, 0);

    if (scheduleDate.getTime() === today.getTime()) {
      todaySchedules.push(schedule);
    } else if (scheduleDate.getTime() === tomorrow.getTime()) {
      tomorrowSchedules.push(schedule);
    }
  }

  // モックデータは今日と明日のスケジュールを含むはず
  assertTrue(
    todaySchedules.length > 0 || tomorrowSchedules.length > 0,
    '今日または明日のスケジュールが少なくとも1件あるべき'
  );
}

// ================================================================================
// 個別テスト実行用関数（デバッグ用）
// ================================================================================

/**
 * 個別のテストを実行（デバッグ用）
 *
 * @param {string} testName - 実行するテスト関数名
 */
function runSingleTest(testName) {
  console.log(`========================================`);
  console.log(`単体テスト: ${testName}`);
  console.log(`========================================\n`);

  try {
    const testFunction = this[testName];
    if (typeof testFunction === 'function') {
      testFunction();
      console.log(`✓ ${testName} - 成功`);
    } else {
      throw new Error(`テスト関数「${testName}」が見つかりません`);
    }
  } catch (error) {
    console.error(`✗ ${testName} - 失敗`);
    console.error(`エラー: ${error.message}`);
    throw error;
  }
}
