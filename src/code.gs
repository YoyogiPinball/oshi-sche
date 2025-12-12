// ================================================================================
// ⚠️ ドライランモード設定（ブレーカー）
// ================================================================================
//
// 初回テスト時は必ず true に設定してください
//
// true:  テストモード（画像解析のみ、書き込み・移動なし）
// false: 本番モード（実際に書き込み・移動を行う）
// null:  PropertiesServiceから読み取り（推奨）
//
const DRY_RUN_MODE = null;
//
// ================================================================================

/**
 * 推しスケ（Oshi-Sche）- VTuber配信スケジュール自動登録システム
 *
 * 【概要】
 * Googleドライブにアップロードされた配信スケジュール画像を、Gemini APIで解析し、
 * Googleスプレッドシート＆Googleカレンダーに自動登録するGoogle Apps Scriptです。
 *
 * 【主な機能】
 * 1. Driveフォルダの監視（サブフォルダ単位：所属：VTuber名）
 * 2. Gemini 2.5 Flashによる画像解析
 * 3. スプレッドシートへのスケジュール書き込み
 * 4. Googleカレンダーへの予定登録
 * 5. 処理済み画像の自動移動
 *
 * 【セットアップ】
 * GASの「プロジェクトの設定」→「スクリプトプロパティ」に以下を設定してください：
 * - GEMINI_API_KEY: Google AI StudioのAPIキー
 * - INPUT_FOLDER_ID: 入力フォルダのID（DriveのURLから取得）
 * - DONE_FOLDER_ID: 処理済みフォルダのID
 * - SPREADSHEET_ID: スプレッドシートのID（URLから取得）
 * - SHEET_NAME: シート名（デフォルト: "スケジュール"）
 * - CALENDAR_ID: カレンダーID（カレンダー設定から取得）
 * - DRY_RUN: ドライランモード（"true"で有効、"false"で無効、未設定時はtrue）
 *
 * @author YoyogiPinball
 * @version 2.0
 * @license MIT
 */

// ================================================================================
// 設定の読み込み
// ================================================================================

/**
 * スクリプトプロパティから設定値を取得する関数
 *
 * PropertiesServiceとは：
 * GASで設定値を安全に保存・読み込むための仕組みです。
 * コード内にAPIキーなどの機密情報を直接書くのではなく、
 * 設定画面から登録した値を読み取ることで、セキュリティを高めます。
 *
 * @returns {Object} 設定オブジェクト
 * @throws {Error} 必須の設定値が不足している場合
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();

  // 必須の設定項目一覧
  const requiredKeys = [
    'GEMINI_API_KEY',
    'INPUT_FOLDER_ID',
    'DONE_FOLDER_ID',
    'SPREADSHEET_ID',
    'CALENDAR_ID'
  ];

  // オプション項目（デフォルト値あり）
  const optionalKeys = {
    'SHEET_NAME': 'スケジュール',  // デフォルトのシート名
    'DRY_RUN': 'true'               // デフォルトはドライランモード
  };

  // 設定値を格納するオブジェクト
  const config = {};

  // 必須項目のチェック
  // 理由: 未設定の項目があると、後で予期しないエラーが発生するため、
  //       事前に分かりやすいエラーメッセージを表示します
  const missingKeys = [];
  for (const key of requiredKeys) {
    const value = props.getProperty(key);
    if (!value) {
      missingKeys.push(key);
    } else {
      config[key] = value;
    }
  }

  // 未設定の項目がある場合はエラー
  if (missingKeys.length > 0) {
    throw new Error(
      `【設定エラー】以下のスクリプトプロパティが未設定です:\n` +
      missingKeys.map(k => `  - ${k}`).join('\n') + '\n\n' +
      `設定方法:\n` +
      `1. GASエディタで「プロジェクトの設定（⚙アイコン）」をクリック\n` +
      `2. 「スクリプトプロパティ」セクションで「プロパティを追加」\n` +
      `3. 上記の項目を設定してください`
    );
  }

  // オプション項目の読み込み（未設定ならデフォルト値を使用）
  for (const [key, defaultValue] of Object.entries(optionalKeys)) {
    config[key] = props.getProperty(key) || defaultValue;
  }

  // DRY_RUNをブール値に変換
  // 理由: プロパティは文字列として保存されるため、true/falseに変換します
  // 最上部のDRY_RUN_MODE定数が設定されている場合はそちらを優先
  if (DRY_RUN_MODE !== null) {
    config.DRY_RUN = DRY_RUN_MODE;
  } else {
    config.DRY_RUN = config.DRY_RUN === 'true';
  }

  // 細分化されたドライランフラグ
  // 未設定の場合はDRY_RUNの値を使用
  const dryRunFlags = [
    'DRY_RUN_SPREADSHEET',  // スプレッドシート書き込み
    'DRY_RUN_CALENDAR',     // カレンダー登録
    'DRY_RUN_DISCORD',      // Discord通知
    'DRY_RUN_FILE_MOVE'     // ファイル移動
  ];

  for (const flag of dryRunFlags) {
    const value = props.getProperty(flag);
    if (value !== null && value !== undefined && value !== '') {
      // 個別フラグが設定されている場合はその値を使用
      config[flag] = value === 'true';
    } else {
      // 未設定の場合はDRY_RUNの値を使用
      config[flag] = config.DRY_RUN;
    }
  }

  return config;
}

// ================================================================================
// メイン処理
// ================================================================================

/**
 * メイン処理：フォルダを監視して新しい画像を処理
 *
 * 【処理フロー】
 * 1. 入力フォルダ内のサブフォルダを走査（所属：VTuber名 形式）
 * 2. 各サブフォルダ内の画像ファイルをチェック
 * 3. 未処理の画像をGeminiで解析
 * 4. スプレッドシート＆カレンダーに登録
 * 5. 処理済みフォルダに移動
 *
 * 【トリガー設定推奨】
 * GASのトリガー設定で、この関数を定期実行（例: 1時間ごと）に設定してください。
 *
 * @returns {number} 処理したファイルの件数
 */
function processNewImages() {
  // ドライランモードの判定
  // まず設定を読み込んで、ドライランかどうかを確認します
  let CONFIG;
  try {
    CONFIG = getConfig();
  } catch (error) {
    console.error(error.message);
    throw error;
  }

  if (CONFIG.DRY_RUN) {
    console.log('========================================');
    console.log('【ドライランモード】');
    console.log('これはテスト実行です。');
    console.log('実際の書き込み・移動は行われません。');
    console.log('========================================');
  }

  // Driveフォルダの取得
  // 理由: DriveApp.getFolderByIdは、存在しないIDを渡すとエラーになるため、
  //       try-catchでエラーを捕捉して分かりやすいメッセージを表示します
  let inputFolder, doneFolder;
  try {
    inputFolder = DriveApp.getFolderById(CONFIG.INPUT_FOLDER_ID);
  } catch (error) {
    throw new Error(
      `【エラー】入力フォルダが見つかりません。\n` +
      `INPUT_FOLDER_ID: ${CONFIG.INPUT_FOLDER_ID}\n` +
      `フォルダIDが正しいか確認してください。`
    );
  }

  try {
    doneFolder = DriveApp.getFolderById(CONFIG.DONE_FOLDER_ID);
  } catch (error) {
    throw new Error(
      `【エラー】処理済みフォルダが見つかりません。\n` +
      `DONE_FOLDER_ID: ${CONFIG.DONE_FOLDER_ID}\n` +
      `フォルダIDが正しいか確認してください。`
    );
  }

  // 処理件数のカウンター
  let processedCount = 0;  // 正常に処理できた件数
  let skippedCount = 0;    // スキップした件数（重複など）
  let errorCount = 0;      // エラーが発生した件数

  // サブフォルダを走査
  // 理由: 各VTuberごとにサブフォルダを分けることで、
  //       どのVTuberのスケジュールかを自動判別できます
  const subFolders = inputFolder.getFolders();

  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    const folderName = subFolder.getName(); // 例: 「個人勢：星野ひかり」

    // フォルダ名をパース（全角コロン「：」で分割）
    // 理由: 「所属：VTuber名」形式から、所属とVTuber名を抽出します
    const parts = folderName.split('：');
    const affiliation = parts[0] || '';      // 所属（例: 個人勢）
    const vtuberName = parts[1] || folderName; // VTuber名（例: 星野ひかり）

    console.log(`フォルダ確認: ${affiliation} - ${vtuberName}`);

    // 処理済みフォルダ内に同名のサブフォルダを取得or作成
    // 理由: 処理済み画像も同じフォルダ構造で管理することで、
    //       どのVTuberの画像が処理済みかを一目で確認できます
    const doneSubFolder = getDoneSubFolder(doneFolder, folderName);

    // 処理済みファイル名一覧を取得
    // 理由: 同じ画像を何度も処理しないように、重複チェックを行います
    const doneFileNames = getDoneFileNames(doneSubFolder);

    // サブフォルダ内のファイルを走査
    const files = subFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      // 画像ファイルのみ処理
      // 理由: PDFやテキストファイルなど、画像以外のファイルは
      //       Geminiの画像解析APIでは処理できないため、スキップします
      if (!mimeType.startsWith('image/')) {
        console.log(`スキップ（画像以外）: ${fileName}`);
        continue;
      }

      // 重複チェック
      // 理由: 同じファイルが入力フォルダと処理済みフォルダの両方に
      //       ある場合は、既に処理済みなので削除します
      if (doneFileNames.has(fileName)) {
        console.log(`削除（処理済み）: ${fileName}`);
        if (!CONFIG.DRY_RUN_FILE_MOVE) {
          file.setTrashed(true);  // ゴミ箱に移動
        } else {
          console.log('【ドライラン】ファイル削除はスキップされました');
        }
        skippedCount++;
        continue;
      }

      console.log(`処理開始: ${vtuberName} - ${fileName}`);

      try {
        // 画像をGeminiで解析
        // 引数として、VTuber名と所属を渡すことで、
        // 解析結果に自動的にこれらの情報を含めます
        const schedules = analyzeScheduleImage(file, vtuberName, affiliation);

        if (schedules && schedules.length > 0) {
          console.log(`${schedules.length}件のスケジュールを抽出しました`);

          // スプレッドシートへの書き込み
          if (!CONFIG.DRY_RUN_SPREADSHEET) {
            writeSchedulesToSheet(schedules, CONFIG);
            console.log(`スプレッドシートに書き込みました`);
          } else {
            console.log('【ドライラン】スプレッドシート書き込みはスキップされました');
            console.log('抽出されたスケジュール:');
            schedules.forEach((s, i) => {
              console.log(`  ${i + 1}. ${s.date} ${s.time} - ${s.content}`);
            });
          }

          // カレンダーへの登録
          if (!CONFIG.DRY_RUN_CALENDAR) {
            addSchedulesToCalendar(schedules, CONFIG);
            console.log(`カレンダーに登録しました`);
          } else {
            console.log('【ドライラン】カレンダー登録はスキップされました');
          }

          // 処理済みサブフォルダに移動
          // 理由: 同じファイルを何度も処理しないように、
          //       処理が完了したら別フォルダに移動します
          if (!CONFIG.DRY_RUN_FILE_MOVE) {
            file.moveTo(doneSubFolder);
          } else {
            console.log('【ドライラン】ファイル移動はスキップされました');
          }
          processedCount++;

        } else {
          console.log('スケジュールを抽出できませんでした');
          errorCount++;
        }

      } catch (error) {
        // エラーが発生した場合は、そのファイルだけスキップして
        // 他のファイルの処理は続行します
        console.error(`エラー: ${fileName} - ${error.message}`);
        errorCount++;
      }
    }
  }

  // 処理結果のサマリーを表示
  console.log('========================================');
  console.log(`処理完了サマリー:`);
  console.log(`  成功: ${processedCount}件`);
  console.log(`  スキップ: ${skippedCount}件`);
  console.log(`  エラー: ${errorCount}件`);
  if (CONFIG.DRY_RUN) {
    console.log('');
    console.log('【ドライランモード】実際の書き込みは行われていません。');
    console.log('本番実行する場合は、DRY_RUNプロパティを"false"に設定してください。');
  }
  console.log('========================================');

  return processedCount;
}

// ================================================================================
// サブ関数（フォルダ・ファイル操作）
// ================================================================================

/**
 * 処理済みフォルダ内のサブフォルダを取得（なければ作成）
 *
 * なぜこの関数が必要か：
 * 処理済みフォルダも入力フォルダと同じ構造（所属：VTuber名）で管理することで、
 * どのVTuberの画像が処理済みかを視覚的に分かりやすくします。
 * サブフォルダが存在しない場合は自動的に作成します。
 *
 * @param {GoogleAppsScript.Drive.Folder} doneFolder 処理済みフォルダ
 * @param {string} folderName サブフォルダ名（例: 個人勢：星野ひかり）
 * @returns {GoogleAppsScript.Drive.Folder} サブフォルダ
 */
function getDoneSubFolder(doneFolder, folderName) {
  // 同名のフォルダを検索
  const folders = doneFolder.getFoldersByName(folderName);

  // 既に存在する場合はそれを返す
  if (folders.hasNext()) {
    return folders.next();
  }

  // 存在しない場合は新規作成
  return doneFolder.createFolder(folderName);
}

/**
 * フォルダ内のファイル名一覧をSetで取得
 *
 * なぜSetを使うのか：
 * Set（集合）は、重複チェックが高速に行えるデータ構造です。
 * 配列でも実装できますが、ファイル数が多い場合はSetの方が効率的です。
 *
 * @param {GoogleAppsScript.Drive.Folder} folder フォルダ
 * @returns {Set<string>} ファイル名の集合
 */
function getDoneFileNames(folder) {
  const fileNames = new Set();
  const files = folder.getFiles();

  // フォルダ内の全ファイル名をSetに追加
  while (files.hasNext()) {
    fileNames.add(files.next().getName());
  }

  return fileNames;
}

// ================================================================================
// Gemini API連携
// ================================================================================

/**
 * Gemini APIで画像を解析してスケジュールを抽出
 *
 * 【処理の流れ】
 * 1. 画像をBase64エンコード（APIに送信できる形式に変換）
 * 2. Gemini APIにリクエスト送信
 * 3. レスポンスからJSON形式のスケジュールを抽出
 * 4. VTuber名・所属を付加して返す
 *
 * 【エラーハンドリング】
 * - APIキー未設定: 分かりやすいエラーメッセージ
 * - ネットワークエラー: リトライ機能付き
 * - JSON抽出失敗: Geminiの生の応答を表示
 *
 * @param {GoogleAppsScript.Drive.File} file 画像ファイル
 * @param {string} vtuberName VTuber名
 * @param {string} affiliation 所属
 * @returns {Array<Object>} スケジュール配列
 * @throws {Error} API呼び出し失敗時
 */
function analyzeScheduleImage(file, vtuberName, affiliation) {
  // APIキーの取得
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error(
      `【エラー】GEMINI_API_KEYが設定されていません。\n` +
      `Google AI Studio (https://aistudio.google.com/) でAPIキーを取得し、\n` +
      `スクリプトプロパティに設定してください。`
    );
  }

  // 画像をBase64エンコード
  // 理由: Gemini APIは、画像をBase64形式で受け取るため、
  //       バイナリデータを文字列に変換します
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  // 現在の年月を取得（年またぎ対応用）
  // 理由: スケジュール画像には年が書かれていないことが多いため、
  //       現在の年月から推測します
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  // Geminiへのプロンプト（指示文）
  // なぜ詳細な指示が必要か：
  // AIは指示が曖昧だと、期待通りの結果を返さないため、
  // 出力形式や注意事項を明確に伝えます
  const prompt = `この画像はVTuberの配信スケジュールです。以下の情報を抽出してJSON形式で出力してください。

本日は${currentYear}年${currentMonth}月です。

出力形式（厳密に守ってください）:
{
  "schedules": [
    {
      "date": "YYYY/MM/DD形式",
      "day": "曜日（月/火/水/木/金/土/日）",
      "time": "HH:MM形式（配信なし/休みの場合は'-'）",
      "content": "配信内容",
      "note": "備考（あれば）"
    }
  ]
}

注意事項:
- 年が書かれていない場合は以下のルールで判定してください：
  - 基本は${currentYear}年とする
  - ただし現在12月で、スケジュールに1月や2月がある場合は${currentYear + 1}年
  - ただし現在1月で、スケジュールに12月がある場合は${currentYear - 1}年
- 複数の配信が同じ日にある場合は別々のエントリとして出力
- 配信がない日（休み、OFFなど）も含めてください
- JSON以外の文字は出力しないでください`;

  // Gemini APIのエンドポイント（2.5-flashに変更）
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  // リクエストペイロード（送信するデータ）
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Data
          }
        }
      ]
    }]
  };

  // HTTPリクエストのオプション
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true  // エラー時もレスポンスを取得
  };

  // リトライ機能付きAPI呼び出し
  // 理由: ネットワークエラーや一時的なAPI障害に対応するため、
  //       最大3回までリトライします
  let response;
  let lastError;
  const maxRetries = 3;

  for (let retry = 0; retry < maxRetries; retry++) {
    try {
      response = UrlFetchApp.fetch(url, options);
      break;  // 成功したらループを抜ける
    } catch (error) {
      lastError = error;
      console.warn(`API呼び出し失敗 (${retry + 1}/${maxRetries}): ${error.message}`);

      // 最後のリトライでなければ、少し待ってから再試行
      if (retry < maxRetries - 1) {
        Utilities.sleep(1000 * (retry + 1));  // 1秒、2秒、3秒と待機時間を増やす
      }
    }
  }

  // 全てのリトライが失敗した場合
  if (!response) {
    throw new Error(
      `【ネットワークエラー】Gemini APIへの接続に失敗しました。\n` +
      `エラー: ${lastError.message}\n` +
      `ネットワーク接続を確認してください。`
    );
  }

  // レスポンスをパース
  const result = JSON.parse(response.getContentText());

  // APIエラーのチェック（エラーコード別の詳細メッセージ）
  if (result.error) {
    const errorCode = result.error.code;
    let errorDetail = '';

    switch (errorCode) {
      case 400:
        errorDetail = 'リクエストの形式が不正です。画像データが破損している可能性があります。';
        break;
      case 401:
        errorDetail = 'APIキーが無効です。GEMINI_API_KEYを確認してください。';
        break;
      case 403:
        errorDetail = 'APIへのアクセスが拒否されました。APIキーの権限を確認してください。';
        break;
      case 429:
        errorDetail = '無料枠の上限に達しました。次回のトリガー実行時に自動リトライされます。';
        break;
      case 500:
        errorDetail = 'Gemini APIの内部エラーです。しばらく待ってから再試行してください。';
        break;
      case 503:
        errorDetail = 'Gemini APIが一時的に利用できません。しばらく待ってから再試行してください。';
        break;
      default:
        errorDetail = 'APIキーや利用上限を確認してください。';
    }

    throw new Error(
      `【Gemini APIエラー】${result.error.message}\n` +
      `エラーコード: ${errorCode || '不明'}\n` +
      `対処方法: ${errorDetail}`
    );
  }

  // レスポンスからテキストを抽出
  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
    console.error('予期しないレスポンス形式:', JSON.stringify(result));
    throw new Error('Gemini APIから有効なレスポンスを取得できませんでした。');
  }

  const text = result.candidates[0].content.parts[0].text;

  // JSON部分を正規表現で抽出
  // 理由: Geminiは指示通りにJSONのみを返すはずですが、
  //       念のため前後の余分なテキストを除去します
  const jsonMatch = text.match(/\{[\s\S]*\}/);

  if (!jsonMatch) {
    console.log('Geminiの応答（JSON抽出失敗）:', text);
    throw new Error(
      `【解析エラー】画像からスケジュールをJSON形式で抽出できませんでした。\n` +
      `画像が不鮮明、または複雑すぎる可能性があります。`
    );
  }

  // JSONをパース
  let data;
  try {
    data = JSON.parse(jsonMatch[0]);
  } catch (error) {
    console.error('JSON解析エラー:', jsonMatch[0]);
    throw new Error(`【JSON解析エラー】Geminiの出力をパースできませんでした: ${error.message}`);
  }

  // スケジュールデータを整形
  // 理由: Geminiの出力にはVTuber名や所属が含まれないため、
  //       フォルダ名から取得した情報を追加します
  if (!data.schedules || !Array.isArray(data.schedules)) {
    console.error('予期しないJSON構造:', data);
    throw new Error('スケジュールデータが配列形式ではありません。');
  }

  return data.schedules.map(s => ({
    vtuber: vtuberName,
    affiliation: affiliation,
    date: s.date || '',
    day: s.day || '',
    time: s.time || '',
    content: s.content || '',
    note: s.note || ''
  }));
}

// ================================================================================
// Googleスプレッドシート連携
// ================================================================================

/**
 * スプレッドシートにスケジュールを書き込み
 *
 * 【処理の流れ】
 * 1. スプレッドシートを開く
 * 2. 全データ保管用シートに書き込み
 * 3. VTuber別シート（所属：VTuber名）に今週分のみ書き込み
 *
 * @param {Array<Object>} schedules スケジュール配列
 * @param {Object} config 設定オブジェクト
 * @throws {Error} スプレッドシート操作失敗時
 */
function writeSchedulesToSheet(schedules, config) {
  // スプレッドシートを開く
  let ss;
  try {
    ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  } catch (error) {
    throw new Error(
      `【エラー】スプレッドシートが見つかりません。\n` +
      `SPREADSHEET_ID: ${config.SPREADSHEET_ID}\n` +
      `IDが正しいか、権限があるか確認してください。`
    );
  }

  // 現在時刻（登録日時として記録）
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // ================================================================================
  // 1. 全データ保管用シート（既存の「スケジュール」シート）に書き込み
  // ================================================================================
  let mainSheet = ss.getSheetByName(config.SHEET_NAME);

  if (!mainSheet) {
    // シートが存在しない場合は新規作成
    mainSheet = ss.insertSheet(config.SHEET_NAME);

    // ヘッダー行を設定
    const headers = ['VTuber名', '所属', '日付', '曜日', '時間', '内容', '備考', '登録日時'];
    mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mainSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');  // 太字
    mainSheet.setFrozenRows(1);  // ヘッダー行を固定

    console.log(`新しいシート「${config.SHEET_NAME}」を作成しました`);
  }

  // スケジュールデータを2次元配列に変換
  const rows = schedules.map(s => [
    s.vtuber,
    s.affiliation,
    s.date,
    s.day,
    s.time,
    s.content,
    s.note,
    now
  ]);

  // 最終行の次の行から書き込み
  const lastRow = mainSheet.getLastRow();
  mainSheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);

  console.log(`全データシートに${rows.length}行を追加しました`);

  // ================================================================================
  // 2. VTuber別シート（所属：VTuber名）に今週分のみ書き込み
  // ================================================================================

  // VTuber別にグループ化
  const groupedByVtuber = {};
  for (const s of schedules) {
    const sheetName = `${s.affiliation}：${s.vtuber}`;
    if (!groupedByVtuber[sheetName]) {
      groupedByVtuber[sheetName] = [];
    }
    groupedByVtuber[sheetName].push(s);
  }

  // 今週の開始日（日曜日）を計算
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=日曜, 1=月曜, ..., 6=土曜
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - dayOfWeek); // 今週の日曜日
  weekStart.setHours(0, 0, 0, 0);

  // 各VTuber用のシートに書き込み
  for (const [sheetName, vtuberSchedules] of Object.entries(groupedByVtuber)) {
    let vtuberSheet = ss.getSheetByName(sheetName);

    // シートが存在しない場合は新規作成
    if (!vtuberSheet) {
      vtuberSheet = ss.insertSheet(sheetName);

      // ヘッダー行を設定
      const headers = ['日付', '曜日', '時間', '内容', '備考', '登録日時'];
      vtuberSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      vtuberSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');  // 太字
      vtuberSheet.setFrozenRows(1);  // ヘッダー行を固定

      console.log(`新しいVTuber別シート「${sheetName}」を作成しました`);
    }

    // 今週より前のデータを削除
    if (vtuberSheet.getLastRow() > 1) {
      const dataRange = vtuberSheet.getRange(2, 1, vtuberSheet.getLastRow() - 1, 6);
      const values = dataRange.getValues();

      // 削除する行を特定（下から削除）
      for (let i = values.length - 1; i >= 0; i--) {
        const dateStr = values[i][0]; // 日付列
        if (dateStr) {
          try {
            const scheduleDate = new Date(dateStr);
            if (scheduleDate < weekStart) {
              // 今週より前のデータは削除
              vtuberSheet.deleteRow(i + 2); // +2 はヘッダー行分
            }
          } catch (error) {
            console.warn(`日付のパースに失敗: ${dateStr}`);
          }
        }
      }
    }

    // 新しいスケジュールを追加
    const vtuberRows = vtuberSchedules.map(s => [
      s.date,
      s.day,
      s.time,
      s.content,
      s.note,
      now
    ]);

    const lastVtuberRow = vtuberSheet.getLastRow();
    vtuberSheet.getRange(lastVtuberRow + 1, 1, vtuberRows.length, 6).setValues(vtuberRows);

    console.log(`VTuber別シート「${sheetName}」に${vtuberRows.length}行を追加しました`);
  }
}

// ================================================================================
// Googleカレンダー連携
// ================================================================================

/**
 * Googleカレンダーにスケジュールを登録
 *
 * 【処理の流れ】
 * 1. カレンダーを取得
 * 2. 各スケジュールをループ
 * 3. 配信なし（時間が'-'）の場合はスキップ
 * 4. 日付・時間をパースしてDateオブジェクトに変換
 * 5. 重複チェック（同じ時間・同じタイトルがあればスキップ）
 * 6. イベントを作成
 *
 * @param {Array<Object>} schedules スケジュール配列
 * @param {Object} config 設定オブジェクト
 * @throws {Error} カレンダー操作失敗時
 */
function addSchedulesToCalendar(schedules, config) {
  // カレンダーを取得
  let calendar;
  try {
    calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
  } catch (error) {
    throw new Error(
      `【エラー】カレンダーにアクセスできません。\n` +
      `CALENDAR_ID: ${config.CALENDAR_ID}\n` +
      `IDが正しいか、権限があるか確認してください。`
    );
  }

  if (!calendar) {
    throw new Error(
      `【エラー】カレンダーが見つかりません。\n` +
      `CALENDAR_ID: ${config.CALENDAR_ID}`
    );
  }

  let addedCount = 0;  // 追加したイベント数

  for (const s of schedules) {
    // 配信なし/休みはスキップ
    // 理由: カレンダーには実際に配信がある予定のみを登録します
    if (!s.time || s.time === '-') {
      continue;
    }

    try {
      // 日付と時間をパース
      // 例: "2025/12/25" → [2025, 12, 25]
      //     "20:00" → [20, 00]
      const dateParts = s.date.split('/');
      const timeParts = s.time.split(':');

      if (dateParts.length !== 3 || timeParts.length < 1) {
        console.warn(`日付・時間の形式が不正です: ${s.date} ${s.time}`);
        continue;
      }

      const year = parseInt(dateParts[0]);
      const month = parseInt(dateParts[1]) - 1; // JavaScriptのDateは0始まり
      const day = parseInt(dateParts[2]);
      const hour = parseInt(timeParts[0]);
      const minute = parseInt(timeParts[1]) || 0;

      // 開始時刻と終了時刻を作成
      // 理由: Googleカレンダーのイベントには開始・終了時刻が必要なため、
      //       終了時刻は開始2時間後に設定します（配信の平均的な長さ）
      const startTime = new Date(year, month, day, hour, minute);
      const endTime = new Date(year, month, day, hour + 2, minute); // 2時間枠

      // イベントタイトル
      const title = `【${s.vtuber}】${s.content}`;

      // 重複チェック（同じ時間に同じタイトルがあればスキップ）
      // 理由: 同じスケジュールを複数回登録してしまうのを防ぎます
      const existingEvents = calendar.getEventsForDay(startTime);
      const isDuplicate = existingEvents.some(e =>
        e.getTitle() === title &&
        e.getStartTime().getTime() === startTime.getTime()
      );

      if (isDuplicate) {
        console.log(`カレンダー重複スキップ: ${title}`);
        continue;
      }

      // イベントの説明文
      const description = [
        `所属: ${s.affiliation}`,
        s.note ? `備考: ${s.note}` : ''
      ].filter(x => x).join('\n');  // 空文字列を除去

      // イベント作成
      calendar.createEvent(title, startTime, endTime, {
        description: description
      });

      console.log(`カレンダー追加: ${title} (${s.date} ${s.time})`);
      addedCount++;

    } catch (error) {
      // 個別のスケジュール登録に失敗しても、他のスケジュールは続行
      console.error(`カレンダー追加エラー: ${s.vtuber} ${s.date} - ${error.message}`);
    }
  }

  console.log(`カレンダー登録完了: ${addedCount}件追加しました`);
}

// ================================================================================
// Discord通知連携
// ================================================================================

/**
 * 毎朝6時に今日・明日のスケジュールをDiscordに通知
 *
 * 【処理の流れ】
 * 1. スプレッドシートから全データを取得
 * 2. 今日・明日のスケジュールのみフィルタリング
 * 3. VTuber別にグループ化
 * 4. Discord用のテキストに整形
 * 5. Webhook経由で送信
 *
 * 【セットアップ】
 * - DISCORD_WEBHOOK_URL をスクリプトプロパティに設定
 * - トリガー設定: 日タイマー 午前6時～7時
 *
 * @throws {Error} Discord送信失敗時
 */
function sendDailyScheduleToDiscord() {
  // 設定を取得
  const CONFIG = getConfig();

  // Discord Webhook URLを取得
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
  if (!webhookUrl) {
    console.warn('DISCORD_WEBHOOK_URLが設定されていません。Discord通知をスキップします。');
    return;
  }

  // スプレッドシートを開く
  let ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (error) {
    throw new Error(`スプレッドシートを開けませんでした: ${error.message}`);
  }

  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!mainSheet) {
    console.log('スケジュールシートが見つかりません。通知をスキップします。');
    return;
  }

  // 全データを取得
  const lastRow = mainSheet.getLastRow();
  if (lastRow <= 1) {
    console.log('スケジュールデータがありません。通知をスキップします。');
    return;
  }

  const dataRange = mainSheet.getRange(2, 1, lastRow - 1, 8);
  const values = dataRange.getValues();

  // 今日と明日の日付を計算
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const dayAfterTomorrow = new Date(today);
  dayAfterTomorrow.setDate(today.getDate() + 2);

  // 今日・明日のスケジュールをフィルタリング
  const todaySchedules = [];
  const tomorrowSchedules = [];

  for (const row of values) {
    const vtuber = row[0];
    const affiliation = row[1];
    const dateStr = row[2];
    const day = row[3];
    const time = row[4];
    const content = row[5];
    const note = row[6];

    if (!dateStr || !time || time === '-') continue;

    try {
      const scheduleDate = new Date(dateStr);
      scheduleDate.setHours(0, 0, 0, 0);

      const schedule = {
        vtuber: vtuber,
        affiliation: affiliation,
        date: dateStr,
        day: day,
        time: time,
        content: content,
        note: note
      };

      if (scheduleDate.getTime() === today.getTime()) {
        todaySchedules.push(schedule);
      } else if (scheduleDate.getTime() === tomorrow.getTime()) {
        tomorrowSchedules.push(schedule);
      }
    } catch (error) {
      console.warn(`日付のパースに失敗: ${dateStr}`);
    }
  }

  // スケジュールがない場合は通知しない
  if (todaySchedules.length === 0 && tomorrowSchedules.length === 0) {
    console.log('本日・明日のスケジュールがありません。通知をスキップします。');
    return;
  }

  // Discord用のメッセージを作成
  const formatDate = (date) => {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    const dayOfWeek = dayNames[date.getDay()];
    return `${month}/${day}（${dayOfWeek}）`;
  };

  const formatSchedulesByVtuber = (schedules) => {
    const grouped = {};
    for (const s of schedules) {
      const key = `${s.affiliation}：${s.vtuber}`;
      if (!grouped[key]) {
        grouped[key] = [];
      }
      grouped[key].push(s);
    }

    const lines = [];
    for (const [vtuberName, items] of Object.entries(grouped)) {
      lines.push(`\n**${vtuberName}**`);
      for (const item of items) {
        const noteText = item.note ? ` ※${item.note}` : '';
        lines.push(`・${item.time} - ${item.content}${noteText}`);
      }
    }

    return lines.join('\n');
  };

  let message = `📅 **本日・明日の配信予定**（${formatDate(today)}）\n`;
  message += `━━━━━━━━━━━━━━━━━━\n`;

  // 今日のスケジュール
  if (todaySchedules.length > 0) {
    message += `**【本日 ${formatDate(today)}】**`;
    message += formatSchedulesByVtuber(todaySchedules);
    message += `\n\n`;
  } else {
    message += `**【本日 ${formatDate(today)}】**\n`;
    message += `本日の配信予定はありません\n\n`;
  }

  message += `━━━━━━━━━━━━━━━━━━\n`;

  // 明日のスケジュール
  if (tomorrowSchedules.length > 0) {
    message += `**【明日 ${formatDate(tomorrow)}】**`;
    message += formatSchedulesByVtuber(tomorrowSchedules);
    message += `\n\n`;
  } else {
    message += `**【明日 ${formatDate(tomorrow)}】**\n`;
    message += `明日の配信予定はありません\n\n`;
  }

  message += `━━━━━━━━━━━━━━━━━━\n`;
  message += `※配信時間は変更される場合があります`;

  // Discordに送信
  if (CONFIG.DRY_RUN_DISCORD) {
    console.log('【ドライラン】Discord通知はスキップされました');
    console.log('送信予定のメッセージ:');
    console.log(message);
    return;
  }

  const payload = {
    content: message
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 204 || statusCode === 200) {
      console.log('Discordへの通知が成功しました');
    } else {
      console.error(`Discord通知失敗: ${statusCode} - ${response.getContentText()}`);
    }
  } catch (error) {
    console.error(`Discord通知エラー: ${error.message}`);
    throw error;
  }
}

// ================================================================================
// テスト用関数
// ================================================================================

/**
 * メイン実行関数
 *
 * 使い方:
 * 1. GASエディタで「実行」→「main」を選択
 * 2. 初回は権限の承認が求められるので、承認してください
 * 3. 実行ログで結果を確認
 *
 * 注意:
 * DRY_RUN="true"の間は、実際の書き込み・移動は行われません。
 * 本番実行する場合は、DRY_RUN="false"に変更してください。
 */
function main() {
  console.log('========================================');
  console.log('推しスケ実行開始');
  console.log('========================================');

  try {
    const count = processNewImages();
    console.log(`テスト完了: ${count}件処理しました`);
  } catch (error) {
    console.error('テスト実行中にエラーが発生しました:');
    console.error(error.message);
    console.error(error.stack);
  }
}
