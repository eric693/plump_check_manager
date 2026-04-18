// ==================== 💰 費用管理系統 API ====================
// ExpenseAPI.gs - 完整修正版

/**
 * 提交預支申請（修正版）
 */
function submitAdvanceApplication(e) {
  try {
    const userId = e.parameter.userId;
    const date = e.parameter.date;
    const amount = parseFloat(e.parameter.amount);
    const purpose = e.parameter.purpose;
    
    if (!userId || !date || !amount || !purpose) {
      return { ok: false, msg: '缺少必要參數' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('users');
    
    if (!userSheet) {
      return { ok: false, msg: '找不到用戶工作表' };
    }
    
    const userData = userSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === userId);
    
    if (!userRow) {
      return { ok: false, msg: '找不到用戶' };
    }
    
    const userName = userRow[2];  // C 欄：displayName
    
    // 寫入預支申請記錄
    const advanceSheet = getOrCreateExpenseSheet('AdvanceApplications');
    const timestamp = new Date().toISOString();
    const id = `ADV-${Date.now()}`;
    
    advanceSheet.appendRow([
      id,              // A: 申請ID
      userId,          // B: 用戶ID
      userName,        // C: 用戶名稱
      date,            // D: 申請日期
      amount,          // E: 申請金額
      purpose,         // F: 申請用途
      'PENDING',       // G: 狀態 (PENDING, APPROVED, REJECTED)
      timestamp,       // H: 申請時間
      '',              // I: 審核人
      '',              // J: 審核時間
      ''               // K: 審核意見
    ]);
    
    // 發送 LINE 通知給主管
    notifyAdminNewAdvanceApplication(userName, date, amount, purpose);
    
    return { 
      ok: true, 
      msg: '預支申請已送出',
      applicationId: id
    };
    
  } catch (error) {
    Logger.log('submitAdvanceApplication 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

/**
 * 提交報銷申請（修正版）
 */
function submitReimbursement(e) {
  try {
    const dataStr = e.parameter.data;
    const data = JSON.parse(dataStr);
    
    const userId = data.userId;
    const date = data.date;
    const summary = data.summary;
    const amount = parseFloat(data.amount);
    const invoiceNumber = data.invoiceNumber || '';
    const note = data.note || '';
    const invoiceImage = data.invoiceImage; // Base64
    const fileName = data.fileName;
    
    if (!userId || !date || !summary || !amount || !invoiceImage) {
      return { ok: false, msg: '缺少必要參數' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('users');
    
    if (!userSheet) {
      return { ok: false, msg: '找不到用戶工作表' };
    }
    
    const userData = userSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === userId);
    
    if (!userRow) {
      return { ok: false, msg: '找不到用戶' };
    }
    
    const userName = userRow[2];  // C 欄：displayName
    
    // 上傳發票照片到 Google Drive
    const folderId = getOrCreateExpenseFolder('Invoices');
    const folder = DriveApp.getFolderById(folderId);
    
    // 解碼 Base64 並儲存
    const blob = Utilities.newBlob(
      Utilities.base64Decode(invoiceImage), 
      'image/jpeg', 
      fileName
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const invoiceUrl = file.getUrl();
    
    // 寫入報銷申請記錄
    const reimbursementSheet = getOrCreateExpenseSheet('ReimbursementApplications');
    const timestamp = new Date().toISOString();
    const id = `REIMB-${Date.now()}`;
    
    reimbursementSheet.appendRow([
      id,              // A: 申請ID
      userId,          // B: 用戶ID
      userName,        // C: 用戶名稱
      date,            // D: 費用日期
      summary,         // E: 費用摘要
      amount,          // F: 報銷金額
      invoiceNumber,   // G: 發票號碼
      note,            // H: 備註
      invoiceUrl,      // I: 發票照片URL
      'PENDING',       // J: 狀態 (PENDING, APPROVED, REJECTED)
      timestamp,       // K: 申請時間
      '',              // L: 審核人
      '',              // M: 審核時間
      ''               // N: 審核意見
    ]);
    
    // 發送 LINE 通知給主管
    notifyAdminNewReimbursement(userName, date, summary, amount);
    
    return { 
      ok: true, 
      msg: '報銷申請已送出',
      applicationId: id,
      invoiceUrl: invoiceUrl
    };
    
  } catch (error) {
    Logger.log('submitReimbursement 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

/**
 * 取得預支申請記錄（修正版）
 */
function getAdvanceRecords(e) {
  try {
    const userId = e.parameter.userId;
    
    if (!userId) {
      return { ok: false, msg: '缺少用戶ID' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const advanceSheet = ss.getSheetByName('AdvanceApplications');
    
    if (!advanceSheet) {
      return { ok: true, records: [] };
    }
    
    const data = advanceSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { ok: true, records: [] };
    }
    
    // 過濾該用戶的記錄
    const records = data.slice(1)
      .filter(row => row[1] === userId)
      .map(row => ({
        id: row[0],
        userId: row[1],
        userName: row[2],
        date: row[3],
        amount: row[4],
        purpose: row[5],
        status: row[6],
        appliedAt: row[7],
        reviewer: row[8],
        reviewedAt: row[9],
        reviewComment: row[10]
      }))
      .reverse(); // 最新的在前面
    
    return { ok: true, records: records };
    
  } catch (error) {
    Logger.log('getAdvanceRecords 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

/**
 * 取得報銷申請記錄（修正版）
 */
function getReimbursementRecords(e) {
  try {
    const userId = e.parameter.userId;
    
    if (!userId) {
      return { ok: false, msg: '缺少用戶ID' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reimbursementSheet = ss.getSheetByName('ReimbursementApplications');
    
    if (!reimbursementSheet) {
      return { ok: true, records: [] };
    }
    
    const data = reimbursementSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { ok: true, records: [] };
    }
    
    // 過濾該用戶的記錄
    const records = data.slice(1)
      .filter(row => row[1] === userId)
      .map(row => ({
        id: row[0],
        userId: row[1],
        userName: row[2],
        date: row[3],
        summary: row[4],
        amount: row[5],
        invoiceNumber: row[6],
        note: row[7],
        invoiceUrl: row[8],
        status: row[9],
        appliedAt: row[10],
        reviewer: row[11],
        reviewedAt: row[12],
        reviewComment: row[13]
      }))
      .reverse(); // 最新的在前面
    
    return { ok: true, records: records };
    
  } catch (error) {
    Logger.log('getReimbursementRecords 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

/**
 * 審核預支申請（管理員）- 修正版
 */
function reviewAdvanceApplication(e) {
  try {
    const applicationId = e.parameter.id;
    const action = e.parameter.action; // 'approve' or 'reject'
    const comment = e.parameter.comment || '';
    const reviewerId = e.parameter.reviewerId;
    
    if (!applicationId || !action || !reviewerId) {
      return { ok: false, msg: '缺少必要參數' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const advanceSheet = ss.getSheetByName('AdvanceApplications');
    
    if (!advanceSheet) {
      return { ok: false, msg: '找不到預支申請工作表' };
    }
    
    const data = advanceSheet.getDataRange().getValues();
    
    // 找到該申請
    const rowIndex = data.findIndex(row => row[0] === applicationId);
    
    if (rowIndex === -1) {
      return { ok: false, msg: '找不到申請記錄' };
    }
    
    const status = action === 'approve' ? 'APPROVED' : 'REJECTED';
    const reviewTime = new Date().toISOString();
    
    // 取得審核人姓名
    const userSheet = ss.getSheetByName('users');
    const userData = userSheet.getDataRange().getValues();
    const reviewerRow = userData.find(row => row[0] === reviewerId);
    const reviewerName = reviewerRow ? reviewerRow[2] : reviewerId;  // C 欄：displayName
    
    // 更新狀態
    advanceSheet.getRange(rowIndex + 1, 7).setValue(status);           // G: 狀態
    advanceSheet.getRange(rowIndex + 1, 9).setValue(reviewerName);     // I: 審核人
    advanceSheet.getRange(rowIndex + 1, 10).setValue(reviewTime);      // J: 審核時間
    advanceSheet.getRange(rowIndex + 1, 11).setValue(comment);         // K: 審核意見
    
    // 發送通知給申請人
    const applicantId = data[rowIndex][1];
    const applicantName = data[rowIndex][2];
    const amount = data[rowIndex][4];
    
    notifyApplicantAdvanceResult(applicantId, applicantName, amount, status, comment);
    
    return { 
      ok: true, 
      msg: status === 'APPROVED' ? '已核准申請' : '已拒絕申請'
    };
    
  } catch (error) {
    Logger.log('reviewAdvanceApplication 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

/**
 * 審核報銷申請（管理員）- 修正版
 */
function reviewReimbursement(e) {
  try {
    const applicationId = e.parameter.id;
    const action = e.parameter.action; // 'approve' or 'reject'
    const comment = e.parameter.comment || '';
    const reviewerId = e.parameter.reviewerId;
    
    if (!applicationId || !action || !reviewerId) {
      return { ok: false, msg: '缺少必要參數' };
    }
    
    // ✅ 修正：使用 getActiveSpreadsheet()
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reimbursementSheet = ss.getSheetByName('ReimbursementApplications');
    
    if (!reimbursementSheet) {
      return { ok: false, msg: '找不到報銷申請工作表' };
    }
    
    const data = reimbursementSheet.getDataRange().getValues();
    
    // 找到該申請
    const rowIndex = data.findIndex(row => row[0] === applicationId);
    
    if (rowIndex === -1) {
      return { ok: false, msg: '找不到申請記錄' };
    }
    
    const status = action === 'approve' ? 'APPROVED' : 'REJECTED';
    const reviewTime = new Date().toISOString();
    
    // 取得審核人姓名
    const userSheet = ss.getSheetByName('users');
    const userData = userSheet.getDataRange().getValues();
    const reviewerRow = userData.find(row => row[0] === reviewerId);
    const reviewerName = reviewerRow ? reviewerRow[2] : reviewerId;  // C 欄：displayName
    
    // 更新狀態
    reimbursementSheet.getRange(rowIndex + 1, 10).setValue(status);         // J: 狀態
    reimbursementSheet.getRange(rowIndex + 1, 12).setValue(reviewerName);   // L: 審核人
    reimbursementSheet.getRange(rowIndex + 1, 13).setValue(reviewTime);     // M: 審核時間
    reimbursementSheet.getRange(rowIndex + 1, 14).setValue(comment);        // N: 審核意見
    
    // 發送通知給申請人
    const applicantId = data[rowIndex][1];
    const applicantName = data[rowIndex][2];
    const amount = data[rowIndex][5];
    
    notifyApplicantReimbursementResult(applicantId, applicantName, amount, status, comment);
    
    return { 
      ok: true, 
      msg: status === 'APPROVED' ? '已核准申請' : '已拒絕申請'
    };
    
  } catch (error) {
    Logger.log('reviewReimbursement 錯誤: ' + error);
    return { ok: false, msg: '系統錯誤：' + error.toString() };
  }
}

// ==================== 輔助函數 ====================

/**
 * 取得或建立費用工作表（專用版本）
 */
function getOrCreateExpenseSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // 根據不同的表格設定標題行
    if (sheetName === 'AdvanceApplications') {
      sheet.appendRow([
        '申請ID', '用戶ID', '用戶名稱', '申請日期', '申請金額', 
        '申請用途', '狀態', '申請時間', '審核人', '審核時間', '審核意見'
      ]);
      
      // 設定欄寬
      sheet.setColumnWidth(1, 150);  // 申請ID
      sheet.setColumnWidth(2, 200);  // 用戶ID
      sheet.setColumnWidth(3, 120);  // 用戶名稱
      sheet.setColumnWidth(6, 200);  // 申請用途
      
    } else if (sheetName === 'ReimbursementApplications') {
      sheet.appendRow([
        '申請ID', '用戶ID', '用戶名稱', '費用日期', '費用摘要', 
        '報銷金額', '發票號碼', '備註', '發票照片URL', '狀態', 
        '申請時間', '審核人', '審核時間', '審核意見'
      ]);
      
      // 設定欄寬
      sheet.setColumnWidth(1, 150);  // 申請ID
      sheet.setColumnWidth(2, 200);  // 用戶ID
      sheet.setColumnWidth(9, 300);  // 發票照片URL
    }
    
    Logger.log(`✅ 已建立工作表: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * 取得或建立 Drive 資料夾（費用專用）
 */
function getOrCreateExpenseFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    const folder = folders.next();
    Logger.log(`✅ 找到現有資料夾: ${folderName}`);
    return folder.getId();
  } else {
    const folder = DriveApp.createFolder(folderName);
    Logger.log(`✅ 已建立資料夾: ${folderName}`);
    return folder.getId();
  }
}

// ==================== LINE 通知函數 ====================

/**
 * 通知主管有新的預支申請
 */
function notifyAdminNewAdvanceApplication(userName, date, amount, purpose) {
  const message = `
📝 新的預支申請

👤 申請人：${userName}
📅 申請日期：${date}
💰 申請金額：NT$ ${amount.toLocaleString()}
📋 申請用途：${purpose}

請至系統審核此申請。
  `.trim();
  
  sendLineNotifyToAdmins(message);
}

/**
 * 通知主管有新的報銷申請
 */
function notifyAdminNewReimbursement(userName, date, summary, amount) {
  const message = `
📄 新的報銷申請

👤 申請人：${userName}
📅 費用日期：${date}
📋 費用摘要：${summary}
💰 報銷金額：NT$ ${amount.toLocaleString()}

請至系統審核此申請並查看發票。
  `.trim();
  
  sendLineNotifyToAdmins(message);
}

/**
 * 通知申請人預支申請結果
 */
function notifyApplicantAdvanceResult(userId, userName, amount, status, comment) {
  const statusText = status === 'APPROVED' ? '✅ 已核准' : '❌ 已拒絕';
  
  let message = `
預支申請審核結果

${statusText}
💰 申請金額：NT$ ${amount.toLocaleString()}
  `.trim();
  
  if (comment) {
    message += `\n\n📝 審核意見：${comment}`;
  }
  
  sendLineNotifyToUser(userId, message);
}

/**
 * 通知申請人報銷申請結果
 */
function notifyApplicantReimbursementResult(userId, userName, amount, status, comment) {
  const statusText = status === 'APPROVED' ? '✅ 已核准' : '❌ 已拒絕';
  
  let message = `
報銷申請審核結果

${statusText}
💰 報銷金額：NT$ ${amount.toLocaleString()}
  `.trim();
  
  if (comment) {
    message += `\n\n📝 審核意見：${comment}`;
  }
  
  sendLineNotifyToUser(userId, message);
}

/**
 * 發送 LINE 通知給所有管理員
 */
function sendLineNotifyToAdmins(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('users');
    
    if (!userSheet) {
      Logger.log('⚠️ 找不到用戶工作表，無法發送通知');
      return;
    }
    
    const userData = userSheet.getDataRange().getValues();
    
    // 找出所有管理員（F 欄 = '管理員'）
    const admins = userData.slice(1).filter(row => row[5] === '管理員');
    
    Logger.log(`📤 準備發送通知給 ${admins.length} 位管理員`);
    
    admins.forEach(admin => {
      const userId = admin[0];  // A 欄：userId
      sendLineNotifyToUser(userId, message);
    });
    
  } catch (error) {
    Logger.log('發送管理員通知失敗: ' + error);
  }
}

/**
 * 發送 LINE 通知給特定用戶
 */
function sendLineNotifyToUser(userId, message) {
  try {
    // 從 PropertiesService 取得 LINE Bot Token
    const LINE_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    
    if (!LINE_BOT_TOKEN) {
      Logger.log('⚠️ 未設定 LINE_CHANNEL_ACCESS_TOKEN，跳過通知');
      return;
    }
    
    const url = 'https://api.line.me/v2/bot/message/push';
    
    const payload = {
      to: userId,
      messages: [{
        type: 'text',
        text: message
      }]
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_BOT_TOKEN}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      Logger.log(`✅ 已發送 LINE 通知給: ${userId}`);
    } else {
      Logger.log(`⚠️ LINE 通知發送失敗 (${responseCode}): ${response.getContentText()}`);
    }
    
  } catch (error) {
    Logger.log('發送 LINE 通知失敗: ' + error);
  }
}

// ==================== 🧪 測試函數 ====================

/**
 * 測試費用管理系統完整流程
 */
function testExpenseSystemComplete() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試費用管理系統完整流程');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const testToken = '8a709f05-5124-4cce-a7fb-a0b98f0f8ea1';  // ⚠️ 替換成有效 token
  
  // 測試 1: 提交預支申請
  Logger.log('📝 測試 1: 提交預支申請');
  const advanceResult = handleSubmitAdvanceApplication({
    token: testToken,
    date: '2025-12-20',
    amount: '5000',
    purpose: '測試預支申請'
  });
  
  Logger.log('預支申請結果: ' + JSON.stringify(advanceResult, null, 2));
  Logger.log('');
  
  // 測試 2: 查詢預支記錄
  Logger.log('📋 測試 2: 查詢預支記錄');
  const recordsResult = handleGetAdvanceRecords({
    token: testToken
  });
  
  Logger.log('預支記錄: ' + JSON.stringify(recordsResult, null, 2));
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}


/**
 * 🏗️ 手動建立費用管理工作表
 */
function createExpenseTables() {
  Logger.log('🏗️ 開始建立費用管理工作表');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 建立預支申請表
  let advanceSheet = ss.getSheetByName('AdvanceApplications');
  
  if (!advanceSheet) {
    advanceSheet = ss.insertSheet('AdvanceApplications');
    advanceSheet.appendRow([
      '申請ID', '用戶ID', '用戶名稱', '申請日期', '申請金額', 
      '申請用途', '狀態', '申請時間', '審核人', '審核時間', '審核意見'
    ]);
    
    // 設定標題列格式
    const headerRange = advanceSheet.getRange(1, 1, 1, 11);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 設定欄寬
    advanceSheet.setColumnWidth(1, 150);  // 申請ID
    advanceSheet.setColumnWidth(2, 200);  // 用戶ID
    advanceSheet.setColumnWidth(3, 120);  // 用戶名稱
    advanceSheet.setColumnWidth(4, 110);  // 申請日期
    advanceSheet.setColumnWidth(5, 100);  // 申請金額
    advanceSheet.setColumnWidth(6, 250);  // 申請用途
    advanceSheet.setColumnWidth(7, 90);   // 狀態
    advanceSheet.setColumnWidth(8, 170);  // 申請時間
    advanceSheet.setColumnWidth(9, 120);  // 審核人
    advanceSheet.setColumnWidth(10, 170); // 審核時間
    advanceSheet.setColumnWidth(11, 200); // 審核意見
    
    // 凍結標題列
    advanceSheet.setFrozenRows(1);
    
    Logger.log('✅ 已建立「預支申請表」(AdvanceApplications)');
  } else {
    Logger.log('ℹ️ 「預支申請表」已存在');
  }
  
  Logger.log('');
  
  // 2. 建立報銷申請表
  let reimbSheet = ss.getSheetByName('ReimbursementApplications');
  
  if (!reimbSheet) {
    reimbSheet = ss.insertSheet('ReimbursementApplications');
    reimbSheet.appendRow([
      '申請ID', '用戶ID', '用戶名稱', '費用日期', '費用摘要', 
      '報銷金額', '發票號碼', '備註', '發票照片URL', '狀態', 
      '申請時間', '審核人', '審核時間', '審核意見'
    ]);
    
    // 設定標題列格式
    const headerRange = reimbSheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 設定欄寬
    reimbSheet.setColumnWidth(1, 150);  // 申請ID
    reimbSheet.setColumnWidth(2, 200);  // 用戶ID
    reimbSheet.setColumnWidth(3, 120);  // 用戶名稱
    reimbSheet.setColumnWidth(4, 110);  // 費用日期
    reimbSheet.setColumnWidth(5, 200);  // 費用摘要
    reimbSheet.setColumnWidth(6, 100);  // 報銷金額
    reimbSheet.setColumnWidth(7, 120);  // 發票號碼
    reimbSheet.setColumnWidth(8, 150);  // 備註
    reimbSheet.setColumnWidth(9, 350);  // 發票照片URL
    reimbSheet.setColumnWidth(10, 90);  // 狀態
    reimbSheet.setColumnWidth(11, 170); // 申請時間
    reimbSheet.setColumnWidth(12, 120); // 審核人
    reimbSheet.setColumnWidth(13, 170); // 審核時間
    reimbSheet.setColumnWidth(14, 200); // 審核意見
    
    // 凍結標題列
    reimbSheet.setFrozenRows(1);
    
    Logger.log('✅ 已建立「報銷申請表」(ReimbursementApplications)');
  } else {
    Logger.log('ℹ️ 「報銷申請表」已存在');
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('🎉 費用管理工作表建立完成！');
  Logger.log('');
  Logger.log('📋 工作表列表:');
  Logger.log('   1. AdvanceApplications（預支申請）');
  Logger.log('   2. ReimbursementApplications（報銷申請）');
  Logger.log('');
  Logger.log('💡 提示：工作表已設定好格式和欄寬');
}