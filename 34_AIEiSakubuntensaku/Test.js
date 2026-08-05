/**
 * ファイル名: Test.gs
 * 役割: APIキー設定機能のテスト実行用関数
 */

function testApiKeySetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.API_KEY_SHEET);
  if (!sheet) {
    console.error("「" + Config.API_KEY_SHEET + "」シートが見つかりません。");
    return;
  }
  
  // 1. テスト用の値を一時的にB1セルにセット
  const targetCell = sheet.getRange("B1");
  const originalVal = targetCell.getValue();
  
  const testValue = "AIzaTestKey_" + Math.random().toString(36).substring(7);
  console.log("B1セルにテスト用キーを設定します:", testValue);
  targetCell.setValue(testValue);
  
  // 2. onEdit イベントオブジェクトを模擬して呼び出し
  const mockEvent = {
    range: targetCell
  };
  
  console.log("onEdit シミュレーション実行...");
  onEdit(mockEvent);
  
  // 3. 結果の確認
  const updatedVal = targetCell.getValue();
  console.log("onEdit実行後のセルの値（「設定済み」であるべき）:", updatedVal);
  
  try {
    const keys = getApiKeys();
    console.log("getApiKeys() の取得結果（テストキーが含まれているべき）:", keys);
    
    if (updatedVal === "設定済み" && keys.indexOf(testValue) !== -1) {
      console.log("✅ テスト成功: APIキーが正常に退避され、表示が「設定済み」に切り替わりました！");
    } else {
      console.error("❌ テスト失敗: 値の変換またはプロパティへの退避に不整合があります。");
    }
  } catch (e) {
    console.error("❌ エラー発生:", e.message);
  } finally {
    // 4. クリーンアップ：元の状態（または空欄）に戻す
    console.log("B1セルを元の値に戻します。");
    targetCell.setValue(originalVal);
  }
}
