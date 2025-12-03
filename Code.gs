const SHEET_NAME = "工作表1"; // 🚨 請確認您的 Google Sheet 工作表名稱
const RECIPIENT_EMAIL = "ava85110@yahoo.com.tw"; // 🚨 請替換為接收訂單通知的 Email

/**
 * 處理從網頁 POST 過來的數據
 */
function doPost(e) {
  // 設置 CORS 標頭，允許任何來源訪問（重要）
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*', 
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  // 預檢請求處理 (CORS Preflight)
  if (e.parameter.method === "OPTIONS") {
    return ContentService.createTextOutput("")
      .setHeaders(corsHeaders);
  }

  // 確保接收到數據
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Error: No data received" }))
                         .setMimeType(ContentService.MimeType.JSON)
                         .setHeaders(corsHeaders);
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `Error: Sheet '${SHEET_NAME}' not found` }))
                         .setMimeType(ContentService.MimeType.JSON)
                         .setHeaders(corsHeaders);
  }
  
  try {
    const data = JSON.parse(e.postData.contents);
    
    // --------------------------------------------------------------------------------
    // 1. 寫入數據到試算表
    // --------------------------------------------------------------------------------
    // 確保欄位順序與您的試算表表頭（步驟1）完全一致
    sheet.appendRow([
      new Date(),             // Timestamp (自動)
      data.orderName,         // 訂購人聯繫姓名
      data.orderContact,      // 連絡電話
      data.orderEmail,        // Email
      data.coupletSign,       // 落款名單/公司行號
      data.relationship,      // 與亡者關係
      data.selectedStyle,     // 花籃樣式
      data.totalPrice,        // 總金額
      data.couplet,           // 輓聯上款
      data.note,              // 備註
      data.needReceipt,       // 是否需要收據
      data.receiptName,       // 收據收件人
      data.receiptAddress,    // 收據地址
      data.receiptPhone,      // 收據電話
      data.deceasedName,      // 亡者姓名
      data.funeralLocation,   // 會場地點
      data.funeralDate        // 出殯日期
    ]);
    
    // --------------------------------------------------------------------------------
    // 2. 發送 Email 通知 (通知禮儀師有新訂單)
    // --------------------------------------------------------------------------------
    const emailSubject = `【新花籃訂單】弔唁故 ${data.deceasedName} - 來自 ${data.orderName}`;
    const emailBody = `
新花籃訂單已登記：

弔唁對象：${data.deceasedName} (會場: ${data.funeralLocation})
訂購人：${data.orderName} (${data.orderContact})
落款名單：${data.coupletSign}
----------------------------
樣式代碼：${data.selectedStyle}
總金額：NT$${data.totalPrice.toLocaleString()}
輓聯上款：${data.couplet}
備註：${data.note || '無'}
----------------------------
需開收據：${data.needReceipt}
收據抬頭：${data.receiptName}
收據地址：${data.receiptAddress}
    
請盡快登入試算表確認訂單並聯繫客戶收款。
試算表連結：${ss.getUrl()}
    `;

    MailApp.sendEmail({
        to: RECIPIENT_EMAIL, 
        subject: emailSubject,
        body: emailBody
    });
    
    // --------------------------------------------------------------------------------
    // 3. 回傳成功訊息給前端
    // --------------------------------------------------------------------------------
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "訂單已成功登記" }))
                         .setMimeType(ContentService.MimeType.JSON)
                         .setHeaders(corsHeaders);
    
  } catch (error) {
    // 捕獲任何程式執行錯誤
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Apps Script 執行錯誤: " + error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON)
                         .setHeaders(corsHeaders);
  }
}
