/**
 * テスト用のモックデータとヘルパー関数
 *
 * 【使い方】
 * - GASのスクリプトエディタでこのファイルも一緒にデプロイする
 * - test.gsのテスト関数から参照される
 *
 * @author YoyogiPinball
 * @version 1.0
 */

// ================================================================================
// モックスケジュールデータ
// ================================================================================

/**
 * テスト用のサンプルスケジュールデータを取得
 *
 * @returns {Array<Object>} スケジュールの配列
 */
function getMockScheduleData() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  const formatDate = (date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  return [
    {
      vtuber: '星野ひかり',
      affiliation: '個人勢',
      date: formatDate(today),
      day: '月',
      time: '20:00',
      content: 'ゲーム配信',
      note: ''
    },
    {
      vtuber: '星野ひかり',
      affiliation: '個人勢',
      date: formatDate(today),
      day: '月',
      time: '22:00',
      content: '雑談配信',
      note: 'メンバー限定'
    },
    {
      vtuber: '月城そら',
      affiliation: 'ホロライブ',
      date: formatDate(tomorrow),
      day: '火',
      time: '21:00',
      content: '歌枠',
      note: ''
    },
    {
      vtuber: '月城そら',
      affiliation: 'ホロライブ',
      date: formatDate(tomorrow),
      day: '火',
      time: '-',
      content: '未定',
      note: ''
    }
  ];
}

/**
 * テスト用の設定オブジェクトを取得
 *
 * @param {Object} overrides - 上書きする設定項目
 * @returns {Object} テスト用設定オブジェクト
 */
function getMockConfig(overrides = {}) {
  const defaultConfig = {
    GEMINI_API_KEY: 'test-api-key-12345',
    INPUT_FOLDER_ID: 'test-input-folder-id',
    DONE_FOLDER_ID: 'test-done-folder-id',
    SPREADSHEET_ID: 'test-spreadsheet-id',
    SHEET_NAME: 'スケジュール',
    CALENDAR_ID: 'test-calendar-id',
    DISCORD_WEBHOOK_URL: 'https://discord.com/api/webhooks/test',
    DRY_RUN: true,
    DRY_RUN_SPREADSHEET: true,
    DRY_RUN_CALENDAR: true,
    DRY_RUN_DISCORD: true,
    DRY_RUN_FILE_MOVE: true,
    TEST_MODE: false
  };

  return { ...defaultConfig, ...overrides };
}

/**
 * テスト用のGemini APIレスポンスをモック
 *
 * @returns {Object} Gemini APIのモックレスポンス
 */
function getMockGeminiResponse() {
  return {
    candidates: [
      {
        content: {
          parts: [
            {
              text: JSON.stringify([
                {
                  date: '2025-12-15',
                  day: '月',
                  time: '20:00',
                  content: 'ゲーム配信',
                  note: ''
                },
                {
                  date: '2025-12-15',
                  day: '月',
                  time: '22:00',
                  content: '雑談配信',
                  note: 'メンバー限定'
                }
              ])
            }
          ]
        }
      }
    ]
  };
}

// ================================================================================
// アサーション関数
// ================================================================================

/**
 * 値が等しいことをアサート
 *
 * @param {*} actual - 実際の値
 * @param {*} expected - 期待される値
 * @param {string} message - エラーメッセージ
 * @throws {Error} アサーション失敗時
 */
function assertEqual(actual, expected, message = '') {
  if (actual !== expected) {
    const errorMsg = message
      ? `${message}\n期待値: ${expected}\n実際の値: ${actual}`
      : `期待値: ${expected}\n実際の値: ${actual}`;
    throw new Error(errorMsg);
  }
}

/**
 * 値が真であることをアサート
 *
 * @param {*} value - チェックする値
 * @param {string} message - エラーメッセージ
 * @throws {Error} アサーション失敗時
 */
function assertTrue(value, message = '') {
  if (!value) {
    throw new Error(message || `期待: true\n実際: ${value}`);
  }
}

/**
 * 値が偽であることをアサート
 *
 * @param {*} value - チェックする値
 * @param {string} message - エラーメッセージ
 * @throws {Error} アサーション失敗時
 */
function assertFalse(value, message = '') {
  if (value) {
    throw new Error(message || `期待: false\n実際: ${value}`);
  }
}

/**
 * 配列の長さをアサート
 *
 * @param {Array} array - チェックする配列
 * @param {number} expectedLength - 期待される長さ
 * @param {string} message - エラーメッセージ
 * @throws {Error} アサーション失敗時
 */
function assertArrayLength(array, expectedLength, message = '') {
  if (!Array.isArray(array)) {
    throw new Error('配列ではありません');
  }
  if (array.length !== expectedLength) {
    const errorMsg = message
      ? `${message}\n期待される長さ: ${expectedLength}\n実際の長さ: ${array.length}`
      : `期待される長さ: ${expectedLength}\n実際の長さ: ${array.length}`;
    throw new Error(errorMsg);
  }
}

/**
 * オブジェクトが特定のプロパティを持つことをアサート
 *
 * @param {Object} obj - チェックするオブジェクト
 * @param {string} property - プロパティ名
 * @param {string} message - エラーメッセージ
 * @throws {Error} アサーション失敗時
 */
function assertHasProperty(obj, property, message = '') {
  if (!obj.hasOwnProperty(property)) {
    const errorMsg = message
      ? `${message}\nプロパティ「${property}」が見つかりません`
      : `プロパティ「${property}」が見つかりません`;
    throw new Error(errorMsg);
  }
}
