/**
 * DXスキル進捗ダッシュボード - Google Apps Script
 * スプレッドシート名: DXスキルアップ記録簿
 * 
 * セットアップ手順:
 * 1. Googleスプレッドシートを作成（名前: DXスキルアップ記録簿）
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付け
 * 4. SPREADSHEET_ID を実際のIDに置き換え
 * 5. testInit() を実行して初期化
 * 6. デプロイ → 新しいデプロイ → ウェブアプリ
 */

// ============================================
// 設定 - スプレッドシートIDを設定してください
// ============================================
const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';

// シート名
const SHEET_DATA = 'DXスキルアップデータ';
const SHEET_LOG = '更新ログ';

// ============================================
// Web API エンドポイント
// ============================================

/**
 * GETリクエスト処理（データ読み込み）
 */
function doGet(e) {
  try {
    const action = e.parameter.action || 'load';
    
    if (action === 'load') {
      const data = loadData();
      return createJsonResponse({ success: true, data: data });
    }
    
    return createJsonResponse({ success: false, error: 'Unknown action' });
    
  } catch (error) {
    console.error('doGet error:', error);
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * POSTリクエスト処理（データ保存）
 */
function doPost(e) {
  try {
    let requestData;
    
    // リクエストデータの解析
    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      requestData = JSON.parse(e.parameter.data);
    } else {
      throw new Error('No data received');
    }
    
    const action = requestData.action || 'save';
    
    if (action === 'save' && requestData.data) {
      saveData(requestData.data);
      addLog('保存', '全データを更新');
      return createJsonResponse({ success: true, message: 'Data saved successfully' });
    }
    
    return createJsonResponse({ success: false, error: 'Unknown action or missing data' });
    
  } catch (error) {
    console.error('doPost error:', error);
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * JSONレスポンスを作成
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// データ操作
// ============================================

/**
 * スプレッドシートからデータを読み込む
 */
function loadData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_DATA);
  
  if (!sheet) {
    return null;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  
  // A2セルからJSONデータを取得
  const jsonStr = sheet.getRange('A2').getValue();
  
  if (!jsonStr) {
    return null;
  }
  
  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    console.error('JSON parse error:', e);
    return null;
  }
}

/**
 * スプレッドシートにデータを保存
 */
function saveData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_DATA);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA);
    sheet.getRange('A1').setValue('データ（JSON形式）');
    sheet.getRange('B1').setValue('最終更新日時');
    sheet.getRange('C1').setValue('タブ数');
    sheet.getRange('D1').setValue('総記録数');
    sheet.getRange('E1').setValue('総学習時間');
    sheet.setColumnWidth(1, 500);
    sheet.setFrozenRows(1);
  }
  
  // データをJSON形式で保存
  const jsonStr = JSON.stringify(data);
  const now = new Date();
  
  sheet.getRange('A2').setValue(jsonStr);
  sheet.getRange('B2').setValue(now);
  
  // サマリー情報を更新
  const tabCount = data.tabs ? data.tabs.length : 0;
  let totalEntries = 0;
  let totalHours = 0;
  
  if (data.tabs) {
    data.tabs.forEach(tab => {
      if (tab.entries) {
        totalEntries += tab.entries.length;
        tab.entries.forEach(entry => {
          totalHours += parseFloat(entry.hours) || 0;
        });
      }
    });
  }
  
  sheet.getRange('C2').setValue(tabCount);
  sheet.getRange('D2').setValue(totalEntries);
  sheet.getRange('E2').setValue(totalHours);
}

/**
 * ログを追加
 */
function addLog(operation, details) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_LOG);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.getRange('A1:D1').setValues([['タイムスタンプ', '操作', 'ユーザー', '詳細']]);
    sheet.setFrozenRows(1);
  }
  
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || '不明';
  
  sheet.insertRowAfter(1);
  sheet.getRange('A2:D2').setValues([[now, operation, user, details]]);
}

// ============================================
// 初期化・テスト
// ============================================

/**
 * 初期化テスト - 最初に一度実行してください
 */
function testInit() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // データシートの作成
  let dataSheet = ss.getSheetByName(SHEET_DATA);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(SHEET_DATA);
    dataSheet.getRange('A1').setValue('データ（JSON形式）');
    dataSheet.getRange('B1').setValue('最終更新日時');
    dataSheet.getRange('C1').setValue('タブ数');
    dataSheet.getRange('D1').setValue('総記録数');
    dataSheet.getRange('E1').setValue('総学習時間');
    dataSheet.setColumnWidth(1, 500);
    dataSheet.setFrozenRows(1);
    console.log('データシートを作成しました');
  }
  
  // ログシートの作成
  let logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_LOG);
    logSheet.getRange('A1:D1').setValues([['タイムスタンプ', '操作', 'ユーザー', '詳細']]);
    logSheet.setFrozenRows(1);
    console.log('ログシートを作成しました');
  }
  
  // ログ追加
  addLog('初期化', 'シート初期化完了');
  
  console.log('初期化が完了しました！');
  console.log('次のステップ: デプロイ → 新しいデプロイ → ウェブアプリ');
}

/**
 * テスト用: データ読み込みテスト
 */
function testLoad() {
  const data = loadData();
  console.log('読み込んだデータ:', JSON.stringify(data, null, 2));
}
