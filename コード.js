// QRコードセミナー受付システム

// 列の定義
const COLUMNS = {
  TIMESTAMP: 1,    // A列: タイムスタンプ
  NAME: 2,         // B列: 氏名
  GRADUATION: 3,   // C列: 卒業年度
  EMAIL: 4,        // D列: メールアドレス
  TOKEN: 5,        // E列: トークン
  URL: 6,          // F列: URL
  QR: 7,           // G列: QR
  REPLY: 8,        // H列: 返信
  RECEPTION: 9     // I列: 受付
};

/**
 * スプレッドシート開く時に実行される関数
 */
function onOpen() {
  // ログシートを作成（存在しない場合のみ）
  createLogSheetIfNotExists();
  
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('QRコードセミナー受付システム')
    .addItem('📋 システム設定', 'showSettingsDialog')
    .addItem('🚀 デプロイエンドポイント設定', 'setActualDeployEndpoint')
    .addSeparator()
    .addItem('⚙️ トリガー設定手順', 'showTriggerSetupGuide')
    .addSeparator()
    .addItem('🔗 URL一括生成', 'generateUrls')
    .addSubMenu(ui.createMenu('📱 QRコード生成')
      .addItem('標準QRコード (200x200)', 'generateQRCodes')
      .addItem('大きいQRコード (300x300)', 'generateLargeQRCodes')
      .addItem('高品質QRコード (エラー訂正H)', 'generateHighQualityQRCodes')
      .addSeparator()
      .addItem('🔄 QRコード再生成 (G列クリア→再作成)', 'regenerateQRCodes'))
    .addSeparator()
    .addItem('📧 受付完了メール送信', 'sendReceptionEmails')
    .addSeparator()
    .addItem('📊 受付状況確認', 'showReceptionStatus')
    .addItem('🔍 システムチェック', 'showHealthCheck')
    .addItem('📋 受付ログ確認', 'showReceptionLogs')
    .addSeparator()
    .addItem('📖 使い方ガイド', 'showHelpDialog')
    .addToUi();
}

/**
 * 設定ダイアログを表示
 */
function showSettingsDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'システム設定');
}

/**
 * 現在の設定を取得
 */
function getCurrentSettings() {
  const properties = PropertiesService.getScriptProperties();
  return {
    sheetId: properties.getProperty('SHEET_ID') || '',
    sheetName: properties.getProperty('SHEET_NAME') || 'シート1',
    deployEndpoint: properties.getProperty('DEPLOY_ENDPOINT') || ''
  };
}

/**
 * 設定を保存
 */
function saveSettings(sheetId, sheetName, deployEndpoint) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'SHEET_ID': sheetId,
      'SHEET_NAME': sheetName || 'シート1',
      'DEPLOY_ENDPOINT': deployEndpoint || ''
    });
    
    // 設定をテスト
    const sheet = getSheet();
    
    return {
      success: true,
      message: '設定が保存されました。'
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

/**
 * メイン関数 - URL一括生成
 */
function generateUrls() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let generatedCount = 0;
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      const row = i + 1;
      const token = data[i][COLUMNS.TOKEN - 1];
      
      // トークンが存在し、URLが空の場合のみ生成
      if (token && !data[i][COLUMNS.URL - 1]) {
        const url = createUrl(token);
        sheet.getRange(row, COLUMNS.URL).setValue(url);
        generatedCount++;
      }
    }
    
    const message = `URL生成が完了しました。${generatedCount}件のURLを生成しました。`;
    console.log(message);
    
    // UIに結果を表示
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return message;
    
  } catch (error) {
    const errorMessage = 'URL生成に失敗しました: ' + error.message;
    console.error('URL生成エラー:', error);
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('エラー', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    throw new Error(errorMessage);
  }
}

/**
 * 受付処理 - Webアプリのエンドポイント
 */
function doGet(e) {
  try {
    const token = e.parameter.token;
    
    if (!token) {
      return createErrorResponse('無効なアクセスです。トークンが見つかりません。');
    }
    
    const result = processReception(token);
    
    if (result.success) {
      return createSuccessResponse(result.name, token);
    } else {
      return createErrorResponse(result.message);
    }
    
  } catch (error) {
    console.error('受付処理エラー:', error);
    return createErrorResponse('システムエラーが発生しました。');
  }
}

/**
 * 受付処理のロジック
 */
function processReception(token) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  // トークンに該当する行を検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.TOKEN - 1] === token) {
      const row = i + 1;
      const currentReception = data[i][COLUMNS.RECEPTION - 1];
      const name = data[i][COLUMNS.NAME - 1];
      
      // 既に受付済みかチェック
      if (currentReception === true || currentReception === 'TRUE') {
        return {
          success: false,
          message: 'この方は既に受付済みです。'
        };
      }
      
      // 受付状態をTRUEに更新
      sheet.getRange(row, COLUMNS.RECEPTION).setValue(true);
      
      // メールアドレスも取得してログ記録
      const email = data[i][COLUMNS.EMAIL - 1] || '';
      logReception(token, name, email, '成功');
      
      return {
        success: true,
        name: name
      };
    }
  }
  
  return {
    success: false,
    message: '無効なトークンです。'
  };
}

/**
 * URLを生成する
 */
function createUrl(token) {
  const properties = PropertiesService.getScriptProperties();
  const deployEndpoint = properties.getProperty('DEPLOY_ENDPOINT');
  
  if (deployEndpoint) {
    // 保存されたデプロイエンドポイントを使用
    return `${deployEndpoint}?token=${token}`;
  } else {
    // フォールバック: ScriptIDから生成（従来の方法）
    const scriptId = ScriptApp.getScriptId();
    return `https://script.google.com/macros/s/${scriptId}/exec?token=${token}`;
  }
}

/**
 * スプレッドシートを取得
 */
function getSheet() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const sheetId = properties.getProperty('SHEET_ID');
    const sheetName = properties.getProperty('SHEET_NAME') || 'シート1';
    
    let spreadsheet;
    
    if (sheetId) {
      // 設定されたIDを使用
      spreadsheet = SpreadsheetApp.openById(sheetId);
    } else {
      // バインドされたスプレッドシートを使用
      try {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      } catch (e) {
        throw new Error('スプレッドシートIDが設定されていません。メニューから「システム設定」を実行してください。');
      }
    }
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。シート名を確認してください。`);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('スプレッドシートにアクセスできません: ' + error.message);
  }
}

/**
 * 受付状況を表示
 */
function showReceptionStatus() {
  try {
    const status = checkReceptionStatus();
    const message = `受付状況:\n${status.received}/${status.total} 人が受付完了\n\n受付率: ${Math.round((status.received / status.total) * 100)}%`;
    SpreadsheetApp.getUi().alert('受付状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'データの取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システムチェック結果を表示
 */
function showHealthCheck() {
  try {
    const result = healthCheck();
    let message = `システム状態: ${result.status}\n`;
    message += `データ行数: ${result.totalRows}行\n`;
    
    if (result.duplicateTokens && result.duplicateTokens.length > 0) {
      message += `\n⚠️ 重複トークン: ${result.duplicateTokens.length}件\n`;
      message += `重複トークン: ${result.duplicateTokens.join(', ')}`;
    } else {
      message += '\n✅ 重複トークンはありません';
    }
    
    SpreadsheetApp.getUi().alert('システムチェック', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'チェックに失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ヘルプダイアログを表示
 */
function showHelpDialog() {
  const message = `QRコードセミナー受付システム 使い方ガイド

📋 初期設定:
1. Webアプリとしてデプロイ
2. 「システム設定」でスプレッドシートID、シート名、デプロイエンドポイントを設定
3. または「デプロイエンドポイント設定」で現在のエンドポイントを自動設定
4. 「URL一括生成」でF列にURLを生成

🔗 URL生成:
- E列にトークンが入力されている行のF列にURLを生成
- 既にURLがある行はスキップされます
- 設定されたデプロイエンドポイント + トークンでURLを構成

🚀 デプロイエンドポイント:
- WebアプリデプロイURL（/execで終わる）を設定
- 設定されていない場合はScriptIDから自動生成（推奨されません）
- 「デプロイエンドポイント設定」で現在のエンドポイントを自動設定可能

📱 QRコード生成 (api.qrserver.com使用):
- 標準QRコード: 200x200サイズ、エラー訂正レベルL
- 大きいQRコード: 300x300サイズ、エラー訂正レベルM
- 高品質QRコード: 250x250サイズ、エラー訂正レベルH
- G列に自動でQRコード画像が生成されます

📊 受付管理:
- 「受付状況確認」で現在の受付状況を確認
- 「システムチェック」で重複トークンなどをチェック

🚀 デプロイ:
1. 右上の「デプロイ」→「新しいデプロイ」
2. 種類「ウェブアプリ」、アクセス「全員」に設定
3. デプロイ後のURLでアクセス可能

💡 QRコードカスタマイズ:
手動でG列に数式を入力することで、色やサイズを自由にカスタマイズ可能です。詳細はセットアップ手順書をご参照ください。`;

  SpreadsheetApp.getUi().alert('使い方ガイド', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 成功時のHTMLレスポンス
 */
function createSuccessResponse(name, token) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>受付完了</title>
      <style>
        body {
          font-family: 'Helvetica Neue', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          display: flex;
          justify-content: center;
          align-items: center;
        }
        .container {
          background: white;
          padding: 40px;
          border-radius: 15px;
          box-shadow: 0 20px 40px rgba(0,0,0,0.1);
          text-align: center;
          max-width: 500px;
          width: 100%;
        }
        .success-icon {
          font-size: 4em;
          color: #28a745;
          margin-bottom: 20px;
        }
        h1 {
          color: #333;
          margin-bottom: 10px;
          font-size: 2em;
        }
        .name {
          font-size: 1.5em;
          color: #667eea;
          font-weight: bold;
          margin: 20px 0;
          padding: 15px;
          background: #f8f9fa;
          border-radius: 8px;
        }
        .message {
          color: #666;
          font-size: 1.1em;
          line-height: 1.6;
        }
        .timestamp {
          color: #999;
          font-size: 0.9em;
          margin-top: 20px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="success-icon">✅</div>
        <h1>受付完了</h1>
        <div class="name">${name || '参加者'} 様</div>
        <div class="message">
          QRコードセミナーの受付が完了いたしました。<br>
          ありがとうございました。
        </div>
        <div class="timestamp">
          受付日時: ${new Date().toLocaleString('ja-JP')}
        </div>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

/**
 * エラー時のHTMLレスポンス
 */
function createErrorResponse(message) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>エラー</title>
      <style>
        body {
          font-family: 'Helvetica Neue', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: linear-gradient(135deg, #ff6b6b 0%, #ee5a52 100%);
          min-height: 100vh;
          display: flex;
          justify-content: center;
          align-items: center;
        }
        .container {
          background: white;
          padding: 40px;
          border-radius: 15px;
          box-shadow: 0 20px 40px rgba(0,0,0,0.1);
          text-align: center;
          max-width: 500px;
          width: 100%;
        }
        .error-icon {
          font-size: 4em;
          color: #dc3545;
          margin-bottom: 20px;
        }
        h1 {
          color: #333;
          margin-bottom: 20px;
          font-size: 2em;
        }
        .message {
          color: #666;
          font-size: 1.1em;
          line-height: 1.6;
        }
        .contact {
          color: #999;
          font-size: 0.9em;
          margin-top: 20px;
          padding: 15px;
          background: #f8f9fa;
          border-radius: 8px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="error-icon">❌</div>
        <h1>エラー</h1>
        <div class="message">${message}</div>
        <div class="contact">
          問題が解決しない場合は、<br>
          セミナー運営事務局までお問い合わせください。
        </div>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

/**
 * ログシートを作成（存在しない場合のみ）
 */
function createLogSheetIfNotExists() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('受付ログ');
    
    if (!logSheet) {
      // ログシートを作成
      logSheet = spreadsheet.insertSheet('受付ログ');
      
      // ヘッダー行を設定
      const headers = [
        '受付日時',
        '氏名', 
        'メールアドレス',
        'トークン',
        'IPアドレス',
        'ユーザーエージェント',
        '受付状況'
      ];
      
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のスタイル設定
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4CAF50');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅を調整
      logSheet.setColumnWidth(1, 150); // 受付日時
      logSheet.setColumnWidth(2, 120); // 氏名
      logSheet.setColumnWidth(3, 200); // メールアドレス
      logSheet.setColumnWidth(4, 150); // トークン
      logSheet.setColumnWidth(5, 120); // IPアドレス
      logSheet.setColumnWidth(6, 250); // ユーザーエージェント
      logSheet.setColumnWidth(7, 100); // 受付状況
      
      console.log('受付ログシートを作成しました');
    }
  } catch (error) {
    console.error('ログシート作成エラー:', error);
  }
}

/**
 * 受付ログを記録
 */
function logReception(token, name, email = '', status = '成功') {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('受付ログ');
    
    // ログシートが存在しない場合は作成
    if (!logSheet) {
      createLogSheetIfNotExists();
      logSheet = spreadsheet.getSheetByName('受付ログ');
    }
    
    // 新しい行を追加
    const newRow = logSheet.getLastRow() + 1;
    const timestamp = new Date();
    
    // ログデータを設定
    const logData = [
      timestamp.toLocaleString('ja-JP'), // 受付日時
      name || '',                        // 氏名
      email || '',                       // メールアドレス
      token || '',                       // トークン
      '',                                // IPアドレス（Web経由でないため空）
      '',                                // ユーザーエージェント（Web経由でないため空）
      status                             // 受付状況
    ];
    
    logSheet.getRange(newRow, 1, 1, logData.length).setValues([logData]);
    
    // 受付成功の場合は背景色を緑に
    if (status === '成功') {
      logSheet.getRange(newRow, 1, 1, logData.length).setBackground('#e8f5e8');
    } else {
      logSheet.getRange(newRow, 1, 1, logData.length).setBackground('#ffeaea');
    }
    
    console.log(`[受付ログ] ${timestamp.toISOString()} - トークン: ${token}, 氏名: ${name}, 状況: ${status}`);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

/**
 * 受付状況を確認する管理関数
 */
function checkReceptionStatus() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  let received = 0;
  
  for (let i = 1; i < data.length; i++) {
    const token = data[i][COLUMNS.TOKEN - 1];
    if (token) {
      total++;
      if (data[i][COLUMNS.RECEPTION - 1] === true || data[i][COLUMNS.RECEPTION - 1] === 'TRUE') {
        received++;
      }
    }
  }
  
  console.log(`受付状況: ${received}/${total} 人が受付完了`);
  return { total, received };
}

/**
 * 受付ログを表示
 */
function showReceptionLogs() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('受付ログ');
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert(
        '受付ログ',
        '受付ログシートが見つかりません。まだ受付処理が実行されていない可能性があります。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const data = logSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert(
        '受付ログ',
        'まだ受付ログがありません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 最新10件を表示
    let message = '📋 受付ログ（最新10件）\n\n';
    const startRow = Math.max(1, data.length - 10);
    
    for (let i = data.length - 1; i >= startRow; i--) {
      const row = data[i];
      if (i === 0) continue; // ヘッダー行をスキップ
      
      const timestamp = row[0];
      const name = row[1];
      const email = row[2];
      const token = row[3];
      const status = row[6];
      
      message += `🕐 ${timestamp}\n`;
      message += `👤 ${name}\n`;
      if (email) {
        message += `📧 ${email}\n`;
      }
      message += `🔑 ${token}\n`;
      message += `✅ ${status}\n`;
      message += '─────────────\n';
    }
    
    message += `\n📊 総受付件数: ${data.length - 1}件`;
    message += '\n\n詳細は「受付ログ」シートでご確認ください。';
    
    SpreadsheetApp.getUi().alert(
      '受付ログ確認',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('受付ログ表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '受付ログの表示に失敗しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 単一のURLを生成する（テスト用）
 */
function generateSingleUrl(token) {
  if (!token) {
    throw new Error('トークンが必要です');
  }
  return createUrl(token);
}

/**
 * システムの健全性チェック
 */
function healthCheck() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    const duplicateTokens = checkDuplicateTokens(data);
    if (duplicateTokens.length > 0) {
      console.warn('重複トークンが見つかりました:', duplicateTokens);
    }
    
    console.log('健全性チェック完了');
    return {
      status: 'OK',
      totalRows: data.length - 1,
      duplicateTokens: duplicateTokens
    };
  } catch (error) {
    console.error('健全性チェックエラー:', error);
    return {
      status: 'ERROR',
      error: error.message
    };
  }
}

/**
 * 重複トークンをチェック
 */
function checkDuplicateTokens(data) {
  const tokens = [];
  const duplicates = [];
  
  for (let i = 1; i < data.length; i++) {
    const token = data[i][COLUMNS.TOKEN - 1];
    if (token) {
      if (tokens.includes(token)) {
        duplicates.push(token);
      } else {
        tokens.push(token);
      }
    }
  }
  
  return duplicates;
}

/**
 * QRコード画像URLを生成 (api.qrserver.com使用)
 */
function generateQRCodes() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let generatedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = i + 1;
      const url = data[i][COLUMNS.URL - 1];
      
      if (url && !data[i][COLUMNS.QR - 1]) {
        // api.qrserver.comを使用してQRコード生成
        const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(url)}`;
        // URLそのものを保存（IMAGE関数ではなく）
        sheet.getRange(row, COLUMNS.QR).setValue(qrCodeUrl);
        generatedCount++;
      }
    }
    
    const message = `QRコード生成が完了しました。${generatedCount}件のQRコードを生成しました。`;
    console.log(message);
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return message;
  } catch (error) {
    const errorMessage = 'QRコード生成に失敗しました: ' + error.message;
    console.error('QRコード生成エラー:', error);
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('エラー', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    throw new Error(errorMessage);
  }
}

/**
 * より詳細なQRコード生成オプション付き
 */
function generateQRCodesWithOptions(size = '200x200', format = 'png', errorCorrection = 'L') {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let generatedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = i + 1;
      const url = data[i][COLUMNS.URL - 1];
      
      if (url && !data[i][COLUMNS.QR - 1]) {
        // api.qrserver.comのパラメータ
        // size: QRコードのサイズ (例: 200x200, 300x300)
        // format: 画像形式 (png, gif, jpeg, svg)
        // ecc: エラー訂正レベル (L, M, Q, H)
        const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=${size}&format=${format}&ecc=${errorCorrection}&data=${encodeURIComponent(url)}`;
        // URLそのものを保存（IMAGE関数ではなく）
        sheet.getRange(row, COLUMNS.QR).setValue(qrCodeUrl);
        generatedCount++;
      }
    }
    
    const message = `QRコード生成が完了しました。${generatedCount}件のQRコード（${size}, ${format}形式）を生成しました。`;
    console.log(message);
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return message;
  } catch (error) {
    const errorMessage = 'QRコード生成に失敗しました: ' + error.message;
    console.error('QRコード生成エラー:', error);
    
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('エラー', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    throw new Error(errorMessage);
  }
}

/**
 * 大きいQRコード生成 (300x300)
 */
function generateLargeQRCodes() {
  return generateQRCodesWithOptions('300x300', 'png', 'M');
}

/**
 * 高品質QRコード生成 (エラー訂正レベル H)
 */
function generateHighQualityQRCodes() {
  return generateQRCodesWithOptions('250x250', 'png', 'H');
}

/**
 * 実際のデプロイエンドポイントを設定する（初期設定用）
 */
function setActualDeployEndpoint() {
  const actualEndpoint = 'https://script.google.com/macros/s/AKfycbxCCwMm-LYJRr-v4OseL0pscN5w3PbO727qTvwyJCvxu814X5ksWS6pXwbxuK5HQcEt/exec';
  
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('DEPLOY_ENDPOINT', actualEndpoint);
  
  console.log('デプロイエンドポイントを設定しました:', actualEndpoint);
  
  if (typeof SpreadsheetApp !== 'undefined') {
    SpreadsheetApp.getUi().alert('完了', 
      'デプロイエンドポイントを設定しました:\n' + actualEndpoint + 
      '\n\n既存のF列のURLを更新する場合は、一度F列をクリアしてから「URL一括生成」を実行してください。', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  return actualEndpoint;
}

/**
 * QRコード再生成（G列をクリアしてから再作成）
 */
function regenerateQRCodes() {
  try {
    const sheet = getSheet();
    if (!sheet) {
      SpreadsheetApp.getUi().alert('エラー', 'シートが見つかりません。設定を確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const data = sheet.getDataRange().getValues();
    let clearCount = 0;
    
    // G列をクリア（ヘッダー行以外）
    for (let i = 1; i < data.length; i++) {
      const row = i + 1;
      if (data[i][COLUMNS.QR - 1]) {
        sheet.getRange(row, COLUMNS.QR).setValue('');
        clearCount++;
      }
    }
    
    console.log(`G列クリア完了: ${clearCount}件`);
    
    // QRコード再生成
    const result = generateQRCodes();
    
    SpreadsheetApp.getUi().alert(
      'QRコード再生成完了',
      `G列をクリアしてQRコードを再生成しました。\n${clearCount}件をクリア後、新しいQRコードを生成しました。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('QRコード再生成エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `QRコード再生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 受付完了メールを一括送信
 */
function sendReceptionEmails() {
  try {
    const sheet = getSheet();
    if (!sheet) {
      SpreadsheetApp.getUi().alert('エラー', 'シートが見つかりません。設定を確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const data = sheet.getDataRange().getValues();
    let emailCount = 0;
    let errorCount = 0;
    const errors = [];

    // 2行目から最終行まで処理（1行目はヘッダー）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // A列（タイムスタンプ）があり、H列（返信）がFALSEの場合
      if (row[COLUMNS.TIMESTAMP - 1] && row[COLUMNS.REPLY - 1] === false) {
        try {
          // データの型安全性チェック
          const name = String(row[COLUMNS.NAME - 1] || '').trim();
          const email = String(row[COLUMNS.EMAIL - 1] || '').trim();
          const qrCode = row[COLUMNS.QR - 1]; // 型チェックはsendReceptionEmail内で実行
          const token = String(row[COLUMNS.TOKEN - 1] || '').trim();
          
          console.log(`メール送信データ確認 - 行${i + 1}: 氏名=${name}, QRコード=${qrCode}, Type=${typeof qrCode}`);
          
          const result = sendReceptionEmail(name, email, qrCode, token);
          
          if (result.success) {
            // H列（返信）をTRUEに更新
            sheet.getRange(i + 1, COLUMNS.REPLY).setValue(true);
            emailCount++;
          } else {
            errors.push(`${row[COLUMNS.NAME - 1]}: ${result.error}`);
            errorCount++;
          }
        } catch (error) {
          console.error(`メール送信エラー (${row[COLUMNS.NAME - 1]}):`, error);
          errors.push(`${row[COLUMNS.NAME - 1]}: ${error.message}`);
          errorCount++;
        }
      }
    }

    // 結果表示
    let message = `メール送信完了\n\n送信成功: ${emailCount}件`;
    if (errorCount > 0) {
      message += `\nエラー: ${errorCount}件\n\nエラー詳細:\n${errors.join('\n')}`;
    }

    SpreadsheetApp.getUi().alert(
      'メール送信結果',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    console.error('受付メール送信エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `メール送信でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * フォーム送信時に自動実行される関数
 */
function onFormSubmit(e) {
  try {
    console.log('フォーム送信イベント発生');
    
    const sheet = e.source.getActiveSheet();
    const row = e.range.getRow();
    
    console.log(`処理対象行: ${row}`);
    
    // E列にユニークトークンを生成
    const token = generateUniqueToken();
    sheet.getRange(row, COLUMNS.TOKEN).setValue(token);
    console.log(`トークン生成: ${token}`);
    
    // F列にURLを生成
    const url = createUrl(token);
    sheet.getRange(row, COLUMNS.URL).setValue(url);
    console.log(`URL生成: ${url}`);
    
    // G列にQRコードを生成
    const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(url)}`;
    sheet.getRange(row, COLUMNS.QR).setValue(qrCodeUrl);
    console.log(`QRコード生成: ${qrCodeUrl}`);
    
    // H列（返信フラグ）をFALSEに設定
    sheet.getRange(row, COLUMNS.REPLY).setValue(false);
    
    // I列（受付フラグ）をFALSEに設定
    sheet.getRange(row, COLUMNS.RECEPTION).setValue(false);
    
    console.log('フォーム送信時の自動処理完了');
    
  } catch (error) {
    console.error('フォーム送信時処理エラー:', error);
  }
}

/**
 * トリガー設定の手順ガイドを表示
 */
function showTriggerSetupGuide() {
  const message = `🔧 トリガー設定手順ガイド

フォームに新しい回答が来たときに自動的にトークン・URL・QRコードを生成するには、以下の手順でトリガーを設定してください：

📝 手順1: Google Apps Scriptエディタを開く
1. スプレッドシートのメニューから「拡張機能」→「Apps Script」をクリック

📝 手順2: トリガーを設定
1. 左メニューの「トリガー」（時計アイコン）をクリック
2. 右下の「+ トリガーを追加」をクリック
3. 以下のように設定：
   - 実行する関数：onFormSubmit または onEdit
   - イベントのソース：スプレッドシートから
   - イベントの種類：フォーム送信時 または 編集時
4. 「保存」をクリック

💡 どちらの関数を選ぶか：
• onFormSubmit：Googleフォームと連携している場合
• onEdit：手動でデータを入力する場合

⚠️ 注意事項：
• トリガー設定後は、新しい回答/データが追加されるたびに自動処理が実行されます
• 既存データにはメニューの「URL一括生成」を使用してください

設定完了後、フォームからテスト送信またはスプレッドシートに新しいデータを入力して動作確認してください。`;

  SpreadsheetApp.getUi().alert(
    'トリガー設定手順',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * スプレッドシート編集時に自動実行される関数（フォールバック）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();
    
    // ヘッダー行は除外
    if (row === 1) return;
    
    // A列（タイムスタンプ）またはB列（氏名）が編集された場合のみ処理
    if (col !== COLUMNS.TIMESTAMP && col !== COLUMNS.NAME) return;
    
    console.log(`スプレッドシート編集イベント発生: 行${row}, 列${col}`);
    
    // 既にトークンが設定されている場合はスキップ
    const existingToken = sheet.getRange(row, COLUMNS.TOKEN).getValue();
    if (existingToken && existingToken.toString().trim() !== '') {
      console.log('既にトークンが設定済みのためスキップ');
      return;
    }
    
    // A列とB列に値がある場合のみ処理
    const timestamp = sheet.getRange(row, COLUMNS.TIMESTAMP).getValue();
    const name = sheet.getRange(row, COLUMNS.NAME).getValue();
    
    if (!timestamp || !name) {
      console.log('タイムスタンプまたは氏名が未入力のためスキップ');
      return;
    }
    
    console.log(`自動処理開始: ${name} (行${row})`);
    
    // E列にユニークトークンを生成
    const token = generateUniqueToken();
    sheet.getRange(row, COLUMNS.TOKEN).setValue(token);
    console.log(`トークン生成: ${token}`);
    
    // F列にURLを生成
    const url = createUrl(token);
    sheet.getRange(row, COLUMNS.URL).setValue(url);
    console.log(`URL生成: ${url}`);
    
    // G列にQRコードを生成
    const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(url)}`;
    sheet.getRange(row, COLUMNS.QR).setValue(qrCodeUrl);
    console.log(`QRコード生成: ${qrCodeUrl}`);
    
    // H列（返信フラグ）をFALSEに設定
    sheet.getRange(row, COLUMNS.REPLY).setValue(false);
    
    // I列（受付フラグ）をFALSEに設定
    sheet.getRange(row, COLUMNS.RECEPTION).setValue(false);
    
    console.log('スプレッドシート編集時の自動処理完了');
    
  } catch (error) {
    console.error('スプレッドシート編集時処理エラー:', error);
  }
}

/**
 * ユニークトークンを生成する関数
 */
function generateUniqueToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  
  // 15文字のランダムな文字列を生成
  for (let i = 0; i < 15; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  
  // 既存のトークンと重複していないかチェック
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    // 重複チェック
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLUMNS.TOKEN - 1] === token) {
        // 重複した場合は再生成
        return generateUniqueToken();
      }
    }
  } catch (error) {
    console.warn('重複チェック中にエラー:', error);
  }
  
  return token;
}

/**
 * 個別の受付完了メールを送信（HTML形式+添付ファイル）
 */
function sendReceptionEmail(name, email, qrCodeUrl, token) {
  try {
    if (!name || !email) {
      return { success: false, error: '氏名またはメールアドレスが未入力です' };
    }

    // qrCodeUrlの型と値をチェック
    console.log('QRコードURL確認:', qrCodeUrl, 'Type:', typeof qrCodeUrl);
    
    if (!qrCodeUrl || typeof qrCodeUrl !== 'string' || qrCodeUrl === 'CellImage' || qrCodeUrl.trim() === '') {
      return { success: false, error: 'QRコードが正しく生成されていません。「🔄 QRコード再生成」を実行してください。' };
    }

    // メール件名
    const subject = '同窓会 受付完了のお知らせ';

    // QRコード画像を取得
    let qrAttachment = null;
    
    try {
      // qrCodeUrlを文字列として正規化
      const qrUrlString = String(qrCodeUrl).trim();
      console.log('正規化後QRコードURL:', qrUrlString);
      
      if (qrUrlString && qrUrlString.startsWith('http')) {
        const response = UrlFetchApp.fetch(qrUrlString);
        const blob = response.getBlob();
        blob.setName(`QRコード_${name}_${token}.png`);
        qrAttachment = blob;
      } else {
        return { success: false, error: `無効なQRコードURL: ${qrUrlString}` };
      }
    } catch (error) {
      console.error('QRコード取得エラー:', error);
      return { success: false, error: 'QRコード画像の取得に失敗しました: ' + error.message };
    }


    // テキストメール本文
    const textBody = `${name}様

この度は同窓会にお申込みいただき、ありがとうございます。
受付が完了いたしました。

当日は添付のQRコード画像をお持ちください。
受付でこちらのQRコードをご提示いただくとスムーズに入場できます。

※QRコードは当日まで大切に保管してください。
※万が一QRコードを紛失された場合は、受付でお名前をお伝えください。

ご不明な点がございましたら、お気軽にお問い合わせください。
当日お会いできることを楽しみにしております。

---
同窓会実行委員会`;

    // テキストメール+添付ファイルでメール送信
    if (qrAttachment) {
      GmailApp.sendEmail(email, subject, textBody, {
        attachments: [qrAttachment],
        name: '同窓会実行委員会'
      });
    } else {
      return { success: false, error: 'QRコード画像の取得に失敗しました' };
    }

    console.log(`受付完了メール送信完了: ${name} (${email})`);
    return { success: true };

  } catch (error) {
    console.error(`メール送信エラー (${name}):`, error);
    return { success: false, error: error.message };
  }
}
