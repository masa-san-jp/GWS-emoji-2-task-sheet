/**
 * 定期的に実行され、未同期の行をまとめて転記するスクリプトです。
 * （データに一部空欄があっても、対象列が空欄であれば転記します）
 */
function syncDataPeriodic() {
  // ==========================================
  // 【設定部分】 ご自身の環境に合わせて書き換えてください
  // ==========================================
  
  // ① 転記先（共有）のスプレッドシートID
  const TARGET_SPREADSHEET_ID = "共有スプレッドシートのIDをここにコピペ";
  
  // ② 転記先のシート名
  const TARGET_SHEET_NAME = "シート1";
  
  // ③ 監視する元のシート名（自分だけが見ているシート）
  const SOURCE_SHEET_NAME = "シート1";
  
  // ④ スクリプトが自動で「済」の目印を入れる列番号（例: D列なら 4）
  const STATUS_COLUMN = 4;
  
  // ⑤ 転記したいデータの列数（例: A列からC列までの値だけを転記するなら 3）
  const DATA_COLUMNS = 3;

  // ⑥ データが入力されている開始行（1行目がタイトルの場合は 2 を設定）
  const START_ROW = 2;

  // ==========================================
  // 【処理部分】 ここから下は変更不要です
  // ==========================================

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return;
  
  const lastRow = sourceSheet.getLastRow();
  // 該当データがない場合は終了する
  if (lastRow < START_ROW) return; 

  // シートのデータを一括取得して処理を高速化
  const maxColumn = Math.max(DATA_COLUMNS, STATUS_COLUMN);
  const range = sourceSheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, maxColumn);
  const values = range.getValues();
  
  let rowsToAppend = [];
  let statusUpdates = []; 
  
  for (let i = 0; i < values.length; i++) {
    const rowData = values[i];
    
    // 【変更箇所】
    // 転記対象のデータ(A〜C列など)を切り出す
    const dataToSync = rowData.slice(0, DATA_COLUMNS);
    
    // A〜C列のどれか1つでも入力があるか（完全にすべて空っぽの無駄な行を弾くため）
    // 文字を連結して文字数が0でないかをチェック
    const hasAnyData = (dataToSync.join("").trim() !== "");
    
    // すでに同期されたマークがないか確認（D列が空欄かどうか）
    const isAlreadySynced = (rowData[STATUS_COLUMN - 1] === "済");
    
    // 「完全に空欄」のゴミデータではなく、かつ未同期のものだけを拾う
    if (hasAnyData && !isAlreadySynced) {
      rowsToAppend.push(dataToSync);
      // あとで「済」を書き込むための行番号を記録
      statusUpdates.push(START_ROW + i);
    }
  }
  
  // 新しく同期する対象がなければ何もしない
  if (rowsToAppend.length === 0) return;

  try {
    const targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME);
    
    // ====================================
    // ① 共有先シートへデータをまとめて追加
    // ====================================
    const appendStartRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(appendStartRow, 1, rowsToAppend.length, DATA_COLUMNS).setValues(rowsToAppend);
    
    // ====================================
    // ② 元のシートに「済」を記入し二重同期を防ぐ
    // ====================================
    for (let i = 0; i < statusUpdates.length; i++) {
        sourceSheet.getRange(statusUpdates[i], STATUS_COLUMN).setValue("済");
    }
    
  } catch (error) {
    console.error("転記時にエラーが発生しました: " + error.message);
  }
}
