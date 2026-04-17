// ====== 消防設備点検アプリ + 校正機器管理システム — 統合版 GAS ======
// ★ 既存のコードを全て削除して、このファイルの内容で上書きしてください
// ★ 上書き後「新しいデプロイ」をしてください

var SECRET_TOKEN = "KoudenSha_Secret_2026"; // ★ここを推測されにくい合言葉に変更してください

function verifyToken(token) {
  if (token !== SECRET_TOKEN) {
    throw new Error("Unauthorized: Invalid Token (合言葉が違います)");
  }
}

function doGet(e) {
  try {
    // 個別パラメータ方式（CORS回避: フロントエンドからGETで送信）
    var token = e.parameter.token;
    var action = e.parameter.action;
    var dataStr = e.parameter.data;
    var data = dataStr ? JSON.parse(dataStr) : undefined;

    return handleAction({ token: token, action: action, data: data });
  } catch (error) {
    return respondWithError(error.toString());
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    return handleAction(payload);
  } catch (error) {
    return respondWithError(error.toString());
  }
}

// 統一アクションルーター
function handleAction(payload) {
  verifyToken(payload.token);
  var action = payload.action;
  var data = payload.data;

  // 消防設備点検アプリ
  if (action === 'getTodaySites') return respondWithData(getTodaySites());
  if (action === 'saveSiteDetails') return respondWithData(saveSiteDetails(data));
  if (action === 'generatePDF') return respondWithData(generatePDF(data));
  if (action === 'addSiteMaster') return respondWithData(addSiteMaster(data));
  if (action === 'getAllSitesMaster') return respondWithData(getAllSitesMaster());
  if (action === 'updateSiteMaster') return respondWithData(updateSiteMaster(data));
  if (action === 'getPrepMaster') return respondWithData(getPrepMaster());
  if (action === 'addPrepItem') return respondWithData(addPrepItem(data));
  if (action === 'deletePrepItem') return respondWithData(deletePrepItem(data));
  if (action === 'addSiteHistory') return respondWithData(addSiteHistory(data));
  if (action === 'deleteSiteHistory') return respondWithData(deleteSiteHistory(data));

  // スタッフマスタ
  if (action === 'getStaffList') return respondWithData(getStaffList());
  if (action === 'addStaffMember') return respondWithData(addStaffMember(data));
  if (action === 'updateStaffMember') return respondWithData(updateStaffMember(data));

  // 校正機器管理
  if (action === 'getAllCalibrationData') return respondWithData(getAllCalibrationData());
  if (action === 'addCalibrationEquipment') return respondWithData(addCalibrationEquipment(data));
  if (action === 'updateCalibrationEquipment') return respondWithData(updateCalibrationEquipment(data));
  if (action === 'deleteCalibrationEquipment') return respondWithData(deleteCalibrationEquipment(data));
  if (action === 'sendForCalibration') return respondWithData(sendForCalibration(data));
  if (action === 'receiveFromCalibration') return respondWithData(receiveFromCalibration(data));
  if (action === 'getCalibrationCategories') return respondWithData(getCalibrationCategories());
  if (action === 'addCalibrationCategory') return respondWithData(addCalibrationCategory(data));
  if (action === 'updateCalibrationCategory') return respondWithData(updateCalibrationCategory(data));
  if (action === 'deleteCalibrationCategory') return respondWithData(deleteCalibrationCategory(data));
  if (action === 'uploadCalibrationCertificate') return respondWithData(uploadCalibrationCertificate(data));

  // Chat通知
  if (action === 'sendChatNotification') return respondWithData(sendChatNotification(data));

  return respondWithError('Invalid action: ' + action);
}

var TARGET_FOLDER_ID = "1s5SQcvKRb_-diRh1jZMoq5myHQ893SAz"; 

function respondWithData(data) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondWithError(errorMsg) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: errorMsg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetDataAsObjects(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("シートが見つかりません: " + sheetName);
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var rows = data.slice(1);
  return rows.map(function(row, rowIndex) {
    var obj = { _rowIndex: rowIndex + 2 };
    headers.forEach(function(h, i) {
      obj[String(h).trim()] = row[i];
    });
    return obj;
  });
}

function formatDate(dateObj) {
  if (!dateObj) return "";
  if (dateObj instanceof Date) {
    var mm = ("0" + (dateObj.getMonth() + 1)).slice(-2);
    var dd = ("0" + dateObj.getDate()).slice(-2);
    return dateObj.getFullYear() + "-" + mm + "-" + dd;
  }
  return String(dateObj);
}

function safeSetValue(sheet, rowIdx, headers, colName, value, missingHeaders) {
  var idx = headers.indexOf(colName);
  if (idx !== -1) {
    sheet.getRange(rowIdx, idx + 1).setValue(value);
  } else if (missingHeaders) {
    missingHeaders.push(colName);
  }
}


// =====================================================================
// ====== 消防設備点検アプリ API（既存機能） ======
// =====================================================================

function getTodaySites() {
  var masterData = getSheetDataAsObjects('案件マスタ');
  
  if (masterData.length === 0) {
     return [];
  }

  var targetList = masterData.map(function(m) {
    var historyRaw = m['現場履歴'] || '[]';
    var hists = [];
    try {
      var parsedHist = JSON.parse(historyRaw);
      hists = parsedHist.map(function(h, i) {
        return { id: i, date: h.date || '', author: h.author || '', text: h.text || '' };
      });
    } catch(e) {
      hists = [];
    }

    var prepArray = [];
    if (m['準備物リスト']) {
      var splits = String(m['準備物リスト']).split('\n');
      for (var i = 0; i < splits.length; i++) {
        if (splits[i].trim() !== '') {
          prepArray.push({ id: i, text: splits[i], checked: false });
        }
      }
    }

    return {
      id: m['案件ID'] || "ID未設定",
      type: "点検（スケ未定）",
      time: "--:--",
      name: m['物件名・建屋名'] || "名称未設定",
      groupName: m['契約グループ名(親案件)'] || "",
      address: m['住所'] || "",
      keyInfo: m['鍵対応'] || "",
      ky: m['KY事項'] || "",
      notes: m['打合せ事項 / 注意事項'] || "",
      prepItems: prepArray,
      sealColor: m['シール色'] || "",
      equipmentSealCount: m['機器シール数'] || 0,
      facilitySealCount: m['設備シール数'] || 0,
      frequency: m['点検頻度'] || "",
      generalMonth: m['総合点検予定月'] || "",
      equipmentMonth: m['機器点検予定月'] || "",
      history: hists,
      _masterRow: m._rowIndex,
      sealDetails: m['シール内訳・詳細'] || "",
      salesRep: m['営業担当者'] || "",
      clientContact: m['客先担当者'] || "",
      primeContractor: m['元請会社名'] || "",
      primeContact: m['元請担当者名・連絡先'] || "",
      fireManager: m['防火管理者'] || "",
      attendant: m['立会者'] || "",
      submitter: m['届出者'] || ""
    };
  });

  return targetList;
}

function saveSiteDetails(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件マスタ');
  if(!sheet) throw new Error("案件マスタシートがありません");

  var masterData = getSheetDataAsObjects('案件マスタ');
  var targetRowObj = null;
  for (var i = 0; i < masterData.length; i++) {
    if (masterData[i]['案件ID'] === data.id) {
      targetRowObj = masterData[i];
      break;
    }
  }
  
  if (targetRowObj) {
    var row = targetRowObj._rowIndex;
    var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var missingHeaders = [];
    
    var prepTextArray = [];
    for (var j = 0; j < data.prepItems.length; j++) {
      prepTextArray.push(data.prepItems[j].text);
    }
    var prepText = prepTextArray.join('\n');
    
    safeSetValue(sheet, row, headersRow, '打合せ事項 / 注意事項', data.notes, missingHeaders);
    safeSetValue(sheet, row, headersRow, 'KY事項', data.ky, missingHeaders);
    safeSetValue(sheet, row, headersRow, '準備物リスト', prepText, missingHeaders);
    
    safeSetValue(sheet, row, headersRow, 'シール色', data.sealColor || "", missingHeaders);
    safeSetValue(sheet, row, headersRow, '機器シール数', data.equipmentSealCount || "", missingHeaders);
    safeSetValue(sheet, row, headersRow, '設備シール数', data.facilitySealCount || "", missingHeaders);
    safeSetValue(sheet, row, headersRow, '点検頻度', data.frequency || "", missingHeaders);
    safeSetValue(sheet, row, headersRow, '総合点検予定月', data.generalMonth || "", missingHeaders);
    safeSetValue(sheet, row, headersRow, '機器点検予定月', data.equipmentMonth || "", missingHeaders);

    if (data.salesRep !== undefined) safeSetValue(sheet, row, headersRow, '営業担当者', data.salesRep || "", missingHeaders);
    if (data.clientContact !== undefined) safeSetValue(sheet, row, headersRow, '客先担当者', data.clientContact || "", missingHeaders);
    if (data.primeContractor !== undefined) safeSetValue(sheet, row, headersRow, '元請会社名', data.primeContractor || "", missingHeaders);
    if (data.primeContact !== undefined) safeSetValue(sheet, row, headersRow, '元請担当者名・連絡先', data.primeContact || "", missingHeaders);
    if (data.fireManager !== undefined) safeSetValue(sheet, row, headersRow, '防火管理者', data.fireManager || "", missingHeaders);
    if (data.attendant !== undefined) safeSetValue(sheet, row, headersRow, '立会者', data.attendant || "", missingHeaders);
    if (data.submitter !== undefined) safeSetValue(sheet, row, headersRow, '届出者', data.submitter || "", missingHeaders);
    
    if (missingHeaders.length > 0) {
      console.warn("見つからなかった見出し:", missingHeaders);
    }
  }

  return { success: true, updatedId: data.id, missingHeaders: missingHeaders || [] };
}

function addSiteMaster(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件マスタ');
  if(!sheet) throw new Error("案件マスタシートがありません");

  var masterData = getSheetDataAsObjects('案件マスタ');
  
  var newId = "B0001";
  var maxNum = 0;
  for (var i = 0; i < masterData.length; i++) {
    var existingId = String(masterData[i]['案件ID'] || "");
    if (existingId.match(/^B\d+$/)) {
      var num = parseInt(existingId.replace('B', ''), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }
  if (maxNum > 0) {
    var nextNum = maxNum + 1;
    var padded = ("000" + nextNum).slice(-4);
    newId = "B" + padded;
  }
  
  var prepTextArray = [];
  if (data.prepItems && data.prepItems.length > 0) {
    for (var j = 0; j < data.prepItems.length; j++) {
      prepTextArray.push(data.prepItems[j]);
    }
  }
  var prepText = prepTextArray.join('\n');
  
  var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = new Array(headersRow.length).fill("");

  newRow[headersRow.indexOf('案件ID')] = newId;
  newRow[headersRow.indexOf('物件名・建屋名')] = data.name || "名称未登録";
  var groupIdx = headersRow.indexOf('契約グループ名(親案件)');
  if (groupIdx !== -1) newRow[groupIdx] = data.groupName || "";
  newRow[headersRow.indexOf('住所')] = data.address || "";
  newRow[headersRow.indexOf('打合せ事項 / 注意事項')] = data.notes || "";
  newRow[headersRow.indexOf('鍵対応')] = data.keyInfo || "";
  newRow[headersRow.indexOf('KY事項')] = data.ky || "";
  newRow[headersRow.indexOf('準備物リスト')] = prepText;
  
  var fileIdx = headersRow.indexOf('関連ファイルURL');
  if (fileIdx !== -1) newRow[fileIdx] = "";
  
  newRow[headersRow.indexOf('シール色')] = data.sealColor || "";
  newRow[headersRow.indexOf('機器シール数')] = data.equipmentSealCount || "";
  newRow[headersRow.indexOf('設備シール数')] = data.facilitySealCount || "";
  newRow[headersRow.indexOf('点検頻度')] = data.frequency || "";
  newRow[headersRow.indexOf('総合点検予定月')] = data.generalMonth || "";
  newRow[headersRow.indexOf('機器点検予定月')] = data.equipmentMonth || "";
  
  var sealDetIdx = headersRow.indexOf('シール内訳・詳細');
  if (sealDetIdx !== -1) newRow[sealDetIdx] = data.sealDetails || "";
  
  var salesRepIdx = headersRow.indexOf('営業担当者');
  if (salesRepIdx !== -1) newRow[salesRepIdx] = data.salesRep || "";
  var clientContactIdx = headersRow.indexOf('客先担当者');
  if (clientContactIdx !== -1) newRow[clientContactIdx] = data.clientContact || "";
  var primeContractorIdx = headersRow.indexOf('元請会社名');
  if (primeContractorIdx !== -1) newRow[primeContractorIdx] = data.primeContractor || "";
  var primeContactIdx = headersRow.indexOf('元請担当者名・連絡先');
  if (primeContactIdx !== -1) newRow[primeContactIdx] = data.primeContact || "";
  var fireManagerIdx = headersRow.indexOf('防火管理者');
  if (fireManagerIdx !== -1) newRow[fireManagerIdx] = data.fireManager || "";
  var attendantIdx = headersRow.indexOf('立会者');
  if (attendantIdx !== -1) newRow[attendantIdx] = data.attendant || "";
  var submitterIdx = headersRow.indexOf('届出者');
  if (submitterIdx !== -1) newRow[submitterIdx] = data.submitter || "";

  sheet.appendRow(newRow);
  return { success: true, newId: newId };
}

function getAllSitesMaster() {
  var masterData = getSheetDataAsObjects('案件マスタ');
  var mappedData = masterData.map(function(row) {
    var prepRaw = String(row['準備物リスト'] || "");
    var prepArray = prepRaw.split('\n').filter(function(item) { return item.trim() !== ""; });
    
    return {
      id: String(row['案件ID'] || ""),
      name: String(row['物件名・建屋名'] || ""),
      groupName: String(row['契約グループ名(親案件)'] || ""),
      address: String(row['住所'] || ""),
      notes: String(row['打合せ事項 / 注意事項'] || ""),
      keyInfo: String(row['鍵対応'] || ""),
      ky: String(row['KY事項'] || ""),
      prepItems: prepArray,
      sealColor: String(row['シール色'] || ""),
      equipmentSealCount: row['機器シール数'] === "" ? 0 : Number(row['機器シール数']),
      facilitySealCount: row['設備シール数'] === "" ? 0 : Number(row['設備シール数']),
      frequency: String(row['点検頻度'] || ""),
      generalMonth: String(row['総合点検予定月'] || ""),
      equipmentMonth: String(row['機器点検予定月'] || ""),
      sealDetails: String(row['シール内訳・詳細'] || ""),
      salesRep: String(row['営業担当者'] || ""),
      clientContact: String(row['客先担当者'] || ""),
      primeContractor: String(row['元請会社名'] || ""),
      primeContact: String(row['元請担当者名・連絡先'] || ""),
      fireManager: String(row['防火管理者'] || ""),
      attendant: String(row['立会者'] || ""),
      submitter: String(row['届出者'] || ""),
      history: (function() {
        var raw = String(row['現場履歴'] || '[]');
        try {
          var arr = JSON.parse(raw);
          return arr.map(function(h, i) {
            return { id: i, date: h.date || '', author: h.author || '', text: h.text || '' };
          });
        } catch(e) { return []; }
      })()
    };
  });
  return mappedData;
}

function updateSiteMaster(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件マスタ');
  if(!sheet) throw new Error("案件マスタシートがありません");

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  
  var idIndex = headers.indexOf('案件ID');
  if (idIndex === -1) throw new Error("案件IDの列が見つかりません");

  var targetRowIdx = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(data.id)) {
      targetRowIdx = i + 1;
      break;
    }
  }

  if (targetRowIdx === -1) {
    throw new Error("指定された案件IDが見つかりません: " + data.id);
  }

  var missingHeaders = [];
  var prepTextArray = [];
  if (data.prepItems && data.prepItems.length > 0) {
    for (var j = 0; j < data.prepItems.length; j++) {
      prepTextArray.push(data.prepItems[j]);
    }
  }
  var prepText = prepTextArray.join('\n');
  
  safeSetValue(sheet, targetRowIdx, headers, '物件名・建屋名', data.name || "名称未登録", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '契約グループ名(親案件)', data.groupName || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '住所', data.address || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '打合せ事項 / 注意事項', data.notes || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '鍵対応', data.keyInfo || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, 'KY事項', data.ky || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '準備物リスト', prepText, missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, 'シール色', data.sealColor || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '機器シール数', data.equipmentSealCount || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '設備シール数', data.facilitySealCount || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '点検頻度', data.frequency || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '総合点検予定月', data.generalMonth || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '機器点検予定月', data.equipmentMonth || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, 'シール内訳・詳細', data.sealDetails || "", missingHeaders);

  safeSetValue(sheet, targetRowIdx, headers, '営業担当者', data.salesRep || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '客先担当者', data.clientContact || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '元請会社名', data.primeContractor || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '元請担当者名・連絡先', data.primeContact || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '防火管理者', data.fireManager || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '立会者', data.attendant || "", missingHeaders);
  safeSetValue(sheet, targetRowIdx, headers, '届出者', data.submitter || "", missingHeaders);

  if (missingHeaders.length > 0) {
    console.warn("見つからなかった見出し:", missingHeaders);
  }

  return { success: true, updatedId: data.id, missingHeaders: missingHeaders };
}

function generatePDF(payloadData) {
  var data = payloadData.siteData || payloadData;
  var pdfType = payloadData.pdfType || 'notes';
  
  var todayStr = formatDate(new Date());
  
  var prefix = (pdfType === 'prep') ? "準備物_" : "打合せ事項_";
  var docName = todayStr + "_" + prefix + data.name;
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();
  
  var title = (pdfType === 'prep') ? "現場準備物リスト" : "現場 打合せ・注意事項シート";
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("現場名: " + data.name + " (" + todayStr + ")").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendHorizontalRule();
  body.appendParagraph("住所: " + (data.address || "未登録"));
  body.appendParagraph("");
  
  if (pdfType === 'notes') {
    body.appendParagraph("【KY事項 (危険予知)】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    body.appendParagraph(data.ky || "特になし");
    body.appendParagraph("");
    
    body.appendParagraph("【鍵対応】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    body.appendParagraph(data.keyInfo || "特になし");
    body.appendParagraph("");

    body.appendParagraph("【打合せ事項 / 注意事項】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    body.appendParagraph(data.notes || "特記事項なし");
    body.appendParagraph("");
    
    body.appendParagraph("【過去の履歴・トラブル報告】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    if (data.history && data.history.length > 0) {
      data.history.forEach(function(h) {
        body.appendParagraph("■ " + h.date + " (" + h.author + ")");
        body.appendParagraph(h.text);
        body.appendParagraph("");
      });
    } else {
      body.appendParagraph("履歴なし");
    }

  } else if (pdfType === 'prep') {
    body.appendParagraph("【点検シール】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    body.appendParagraph("色: " + (data.sealColor || "未定") + " / 機器: " + (data.equipmentSealCount || 0) + "枚 / 設備: " + (data.facilitySealCount || 0) + "枚");
    body.appendParagraph("");
    
    body.appendParagraph("【準備物リスト】").setHeading(DocumentApp.ParagraphHeading.HEADING4);
    if (data.prepItems && data.prepItems.length > 0) {
      data.prepItems.forEach(function(item) {
        body.appendParagraph("□ " + item.text);
      });
    } else {
      body.appendParagraph("登録なし");
    }
  }

  doc.saveAndClose();
  
  var pdfBlob = doc.getAs('application/pdf');
  pdfBlob.setName(docName + ".pdf");
  
  var parentFolder = DriveApp.getRootFolder();
  if (TARGET_FOLDER_ID && TARGET_FOLDER_ID.trim() !== "") {
    try {
      parentFolder = DriveApp.getFolderById(TARGET_FOLDER_ID.trim());
    } catch(e) {
      console.error("指定されたフォルダIDが見つからないためルートに保存します。", e);
    }
  }
  
  var pdfFile = parentFolder.createFile(pdfBlob);
  
  try {
    DriveApp.getFileById(doc.getId()).setTrashed(true);
  } catch(e) {
  }
  
  return { success: true, pdfUrl: pdfFile.getUrl() };
}

// ====== 準備物マスタ管理 API ======

function getPrepMaster() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("準備物マスタ");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  
  var categories = data[0];
  var result = [];
  
  for (var col = 0; col < categories.length; col++) {
    var catName = String(categories[col]).trim();
    if (catName === "") continue;
    
    var items = [];
    for (var row = 1; row < data.length; row++) {
      if (data[row] && data[row][col] !== undefined) {
        var itemVal = String(data[row][col]).trim();
        if (itemVal !== "") {
          items.push(itemVal);
        }
      }
    }
    result.push({
      category: catName,
      items: items
    });
  }
  
  return result;
}

function addPrepItem(dataObj) {
  var catName = dataObj.category;
  var itemName = dataObj.item;
  if (!catName || !itemName) throw new Error("カテゴリまたはアイテム名が空です。");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("準備物マスタ");
  if (!sheet) throw new Error("準備物マスタシートが見つかりません。");
  
  var data = sheet.getDataRange().getValues();
  var categories = data.length > 0 ? data[0] : [];
  
  var targetColIndex = -1;
  for (var c = 0; c < categories.length; c++) {
    if (String(categories[c]).trim() === catName) {
      targetColIndex = c;
      break;
    }
  }
  
  if (targetColIndex === -1) {
    targetColIndex = categories.length;
    sheet.getRange(1, targetColIndex + 1).setValue(catName);
  }
  
  var insertRow = 2; 
  var maxRow = sheet.getLastRow();
  if (maxRow > 0) {
    var colValues = sheet.getRange(1, targetColIndex + 1, maxRow, 1).getValues();
    for (var r = 1; r < colValues.length; r++) {
      if (String(colValues[r][0]).trim() !== "") {
        insertRow = r + 2; 
      }
    }
  }
  
  sheet.getRange(insertRow, targetColIndex + 1).setValue(itemName);
  return getPrepMaster();
}

function deletePrepItem(dataObj) {
  var catName = dataObj.category;
  var itemName = dataObj.item;
  if (!catName || !itemName) throw new Error("カテゴリまたはアイテム名が空です。");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("準備物マスタ");
  if (!sheet) throw new Error("準備物マスタが見つかりません。");
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getPrepMaster();
  
  var categories = data[0];
  var targetColIndex = -1;
  for (var c = 0; c < categories.length; c++) {
    if (String(categories[c]).trim() === catName) {
      targetColIndex = c;
      break;
    }
  }
  
  if (targetColIndex !== -1) {
    for (var r = 1; r < data.length; r++) {
      if (data[r] && data[r][targetColIndex] !== undefined) {
        if (String(data[r][targetColIndex]).trim() === itemName) {
          sheet.getRange(r + 1, targetColIndex + 1).clearContent();
        }
      }
    }
  }
  
  return getPrepMaster();
}

// ====== 現場履歴管理 API ======

function addSiteHistory(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件マスタ');
  if(!sheet) throw new Error("案件マスタシートがありません");

  var masterData = getSheetDataAsObjects('案件マスタ');
  var targetRow = null;
  for (var i = 0; i < masterData.length; i++) {
    if (masterData[i]['案件ID'] === data.caseId) {
      targetRow = masterData[i];
      break;
    }
  }
  if (!targetRow) throw new Error("案件IDが見つかりません: " + data.caseId);

  var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = headersRow.indexOf('現場履歴');

  // カラムが存在しない場合は自動追加
  if (colIdx === -1) {
    var lastCol = sheet.getLastColumn();
    colIdx = lastCol; // 0-indexed for new column
    sheet.getRange(1, lastCol + 1).setValue('現場履歴');
  }

  var existingRaw = sheet.getRange(targetRow._rowIndex, colIdx + 1).getValue() || '[]';
  var histArray = [];
  try { histArray = JSON.parse(existingRaw); } catch(e) { histArray = []; }

  var dateStr = data.date || formatDate(new Date());
  histArray.unshift({
    date: dateStr,
    author: data.author || '',
    text: data.text || ''
  });

  sheet.getRange(targetRow._rowIndex, colIdx + 1).setValue(JSON.stringify(histArray));

  // 更新後の履歴を返す
  var returnHists = histArray.map(function(h, i) {
    return { id: i, date: h.date || '', author: h.author || '', text: h.text || '' };
  });
  return { success: true, history: returnHists };
}

function deleteSiteHistory(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件マスタ');
  if(!sheet) throw new Error("案件マスタシートがありません");

  var masterData = getSheetDataAsObjects('案件マスタ');
  var targetRow = null;
  for (var i = 0; i < masterData.length; i++) {
    if (masterData[i]['案件ID'] === data.caseId) {
      targetRow = masterData[i];
      break;
    }
  }
  if (!targetRow) throw new Error("案件IDが見つかりません: " + data.caseId);

  var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = headersRow.indexOf('現場履歴');
  if (colIdx === -1) throw new Error("現場履歴列が見つかりません");

  var existingRaw = sheet.getRange(targetRow._rowIndex, colIdx + 1).getValue() || '[]';
  var histArray = [];
  try { histArray = JSON.parse(existingRaw); } catch(e) { histArray = []; }

  // 指定されたインデックスの履歴を削除
  var deleteIndex = data.historyIndex;
  if (deleteIndex < 0 || deleteIndex >= histArray.length) {
    throw new Error("無効な履歴インデックスです: " + deleteIndex);
  }
  histArray.splice(deleteIndex, 1);

  sheet.getRange(targetRow._rowIndex, colIdx + 1).setValue(JSON.stringify(histArray));

  var returnHists = histArray.map(function(h, i) {
    return { id: i, date: h.date || '', author: h.author || '', text: h.text || '' };
  });
  return { success: true, history: returnHists };
}

// =====================================================================
// ====== 校正機器管理システム API（新規追加） ======
// =====================================================================

// ── 初期セットアップ（1回だけ実行） ──
function setupCalibrationSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 校正機器マスタ
  if (!ss.getSheetByName('校正機器マスタ')) {
    var s1 = ss.insertSheet('校正機器マスタ');
    s1.getRange(1, 1, 1, 16).setValues([[
      '管理番号', '機器種別', '機器名', '通称', 'メーカー', '型番',
      'シリアル番号', '購入日', '校正周期(月)', '校正業者', '保管場所',
      '使用者', 'ステータス', '備考', '登録日', '更新日'
    ]]);
    s1.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
    s1.setFrozenRows(1);
  }
  
  // 校正履歴
  if (!ss.getSheetByName('校正履歴')) {
    var s2 = ss.insertSheet('校正履歴');
    s2.getRange(1, 1, 1, 14).setValues([[
      '履歴ID', '管理番号', '送出日', '返却日', '所要日数',
      '校正実施日', '次回校正予定日', '結果', '証明書番号', '証明書URL',
      '費用', '校正業者', '備考', '記録日'
    ]]);
    s2.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
    s2.setFrozenRows(1);
  }
  
  // 校正機器種別マスタ
  if (!ss.getSheetByName('校正機器種別マスタ')) {
    var s3 = ss.insertSheet('校正機器種別マスタ');
    s3.getRange(1, 1, 1, 3).setValues([['種別ID', '種別名', 'デフォルト校正周期(月)']]);
    s3.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#bf9000').setFontColor('#ffffff');
    s3.setFrozenRows(1);
    
    var defaults = [
      ['cat-001', '絶縁抵抗計（メガー）', 12],
      ['cat-002', 'デジタルマルチメーター', 12],
      ['cat-003', '接地抵抗計', 12],
      ['cat-004', 'クランプメーター', 12],
      ['cat-005', '騒音計', 24],
      ['cat-006', '照度計', 24],
      ['cat-007', '風速計', 24],
      ['cat-008', '圧力計（連成計）', 12],
      ['cat-009', '差圧計（マノメーター）', 12],
      ['cat-010', 'ピトーゲージ', 24],
      ['cat-011', '加煙試験器', 0],
      ['cat-012', '加熱試験器', 0],
      ['cat-099', 'その他', 12]
    ];
    s3.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
  
  Logger.log('校正管理用の3シートを作成しました！');
}

// ====== 校正機器マスタ CRUD ======

function getCalibrationEquipment() {
  var data = getSheetDataAsObjects('校正機器マスタ');
  return data.map(function(row) {
    return {
      managementNumber: String(row['管理番号'] || ''),
      categoryId: String(row['機器種別'] || ''),
      name: String(row['機器名'] || ''),
      nickname: String(row['通称'] || ''),
      manufacturer: String(row['メーカー'] || ''),
      model: String(row['型番'] || ''),
      serialNumber: String(row['シリアル番号'] || ''),
      purchaseDate: formatDate(row['購入日']),
      calibrationCycleMonths: Number(row['校正周期(月)']) || 0,
      calibrationVendor: String(row['校正業者'] || ''),
      storageLocation: String(row['保管場所'] || ''),
      assignedTo: String(row['使用者'] || ''),
      status: String(row['ステータス'] || 'active'),
      notes: String(row['備考'] || ''),
      createdAt: formatDate(row['登録日']),
      updatedAt: formatDate(row['更新日'])
    };
  });
}

function addCalibrationEquipment(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器マスタ');
  if (!sheet) throw new Error('校正機器マスタシートがありません');
  
  var existingData = getSheetDataAsObjects('校正機器マスタ');
  var maxNum = 0;
  for (var i = 0; i < existingData.length; i++) {
    var mn = String(existingData[i]['管理番号'] || '');
    var match = mn.match(/^CAL-(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  var newNumber = 'CAL-' + ('000' + (maxNum + 1)).slice(-3);
  
  var now = formatDate(new Date());
  
  sheet.appendRow([
    newNumber,
    data.categoryId || '',
    data.name || '',
    data.nickname || '',
    data.manufacturer || '',
    data.model || '',
    data.serialNumber || '',
    data.purchaseDate || '',
    data.calibrationCycleMonths || 0,
    data.calibrationVendor || '',
    data.storageLocation || '',
    data.assignedTo || '',
    data.status || 'active',
    data.notes || '',
    now,
    now
  ]);
  
  return { success: true, managementNumber: newNumber };
}

function updateCalibrationEquipment(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器マスタ');
  if (!sheet) throw new Error('校正機器マスタシートがありません');
  
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var mnIdx = headers.indexOf('管理番号');
  
  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][mnIdx]) === String(data.managementNumber)) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) throw new Error('管理番号が見つかりません: ' + data.managementNumber);
  
  var missingHeaders = [];
  if (data.categoryId !== undefined) safeSetValue(sheet, targetRow, headers, '機器種別', data.categoryId, missingHeaders);
  if (data.name !== undefined) safeSetValue(sheet, targetRow, headers, '機器名', data.name, missingHeaders);
  if (data.nickname !== undefined) safeSetValue(sheet, targetRow, headers, '通称', data.nickname, missingHeaders);
  if (data.manufacturer !== undefined) safeSetValue(sheet, targetRow, headers, 'メーカー', data.manufacturer, missingHeaders);
  if (data.model !== undefined) safeSetValue(sheet, targetRow, headers, '型番', data.model, missingHeaders);
  if (data.serialNumber !== undefined) safeSetValue(sheet, targetRow, headers, 'シリアル番号', data.serialNumber, missingHeaders);
  if (data.purchaseDate !== undefined) safeSetValue(sheet, targetRow, headers, '購入日', data.purchaseDate, missingHeaders);
  if (data.calibrationCycleMonths !== undefined) safeSetValue(sheet, targetRow, headers, '校正周期(月)', data.calibrationCycleMonths, missingHeaders);
  if (data.calibrationVendor !== undefined) safeSetValue(sheet, targetRow, headers, '校正業者', data.calibrationVendor, missingHeaders);
  if (data.storageLocation !== undefined) safeSetValue(sheet, targetRow, headers, '保管場所', data.storageLocation, missingHeaders);
  if (data.assignedTo !== undefined) safeSetValue(sheet, targetRow, headers, '使用者', data.assignedTo, missingHeaders);
  if (data.status !== undefined) safeSetValue(sheet, targetRow, headers, 'ステータス', data.status, missingHeaders);
  if (data.notes !== undefined) safeSetValue(sheet, targetRow, headers, '備考', data.notes, missingHeaders);
  safeSetValue(sheet, targetRow, headers, '更新日', formatDate(new Date()), missingHeaders);
  
  return { success: true, updatedId: data.managementNumber };
}

function deleteCalibrationEquipment(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器マスタ');
  if (!sheet) throw new Error('校正機器マスタシートがありません');
  
  var values = sheet.getDataRange().getValues();
  var mnIdx = values[0].indexOf('管理番号');
  
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][mnIdx]) === String(data.managementNumber)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  
  var histSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正履歴');
  if (histSheet) {
    var histValues = histSheet.getDataRange().getValues();
    var histMnIdx = histValues[0].indexOf('管理番号');
    for (var j = histValues.length - 1; j >= 1; j--) {
      if (String(histValues[j][histMnIdx]) === String(data.managementNumber)) {
        histSheet.deleteRow(j + 1);
      }
    }
  }
  
  return { success: true, deletedId: data.managementNumber };
}

// ====== 校正ワークフロー ======

function sendForCalibration(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正履歴');
  if (!sheet) throw new Error('校正履歴シートがありません');
  
  var existingData = getSheetDataAsObjects('校正履歴');
  var maxNum = 0;
  for (var i = 0; i < existingData.length; i++) {
    var hid = String(existingData[i]['履歴ID'] || '');
    var match = hid.match(/^CH-(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  var newId = 'CH-' + ('0000' + (maxNum + 1)).slice(-4);
  
  sheet.appendRow([
    newId,
    data.managementNumber,
    data.sentDate || formatDate(new Date()),
    '', '', '', '', '', '', '', '',
    data.vendor || '',
    '',
    formatDate(new Date())
  ]);
  
  updateCalibrationEquipment({
    managementNumber: data.managementNumber,
    status: 'calibrating'
  });
  
  return { success: true, recordId: newId };
}

function receiveFromCalibration(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正履歴');
  if (!sheet) throw new Error('校正履歴シートがありません');
  
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idIdx = headers.indexOf('履歴ID');
  
  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(data.recordId)) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) throw new Error('履歴IDが見つかりません: ' + data.recordId);
  
  var sentDateStr = String(values[targetRow - 1][headers.indexOf('送出日')]);
  var turnaroundDays = '';
  if (sentDateStr && data.returnedDate) {
    var sent = new Date(sentDateStr);
    var returned = new Date(data.returnedDate);
    turnaroundDays = Math.round((returned - sent) / (1000 * 60 * 60 * 24));
  }
  
  var missingHeaders = [];
  safeSetValue(sheet, targetRow, headers, '返却日', data.returnedDate || '', missingHeaders);
  safeSetValue(sheet, targetRow, headers, '所要日数', turnaroundDays, missingHeaders);
  safeSetValue(sheet, targetRow, headers, '校正実施日', data.calibrationDate || '', missingHeaders);
  safeSetValue(sheet, targetRow, headers, '次回校正予定日', data.nextCalibrationDate || '', missingHeaders);
  safeSetValue(sheet, targetRow, headers, '結果', data.result || 'pass', missingHeaders);
  safeSetValue(sheet, targetRow, headers, '証明書番号', data.certificateNumber || '', missingHeaders);
  safeSetValue(sheet, targetRow, headers, '費用', data.cost || 0, missingHeaders);
  safeSetValue(sheet, targetRow, headers, '備考', data.notes || '', missingHeaders);
  
  if (data.certificatePdfBase64 && data.certificateFileName) {
    try {
      var pdfBlob = Utilities.newBlob(
        Utilities.base64Decode(data.certificatePdfBase64),
        'application/pdf',
        data.certificateFileName
      );
      
      var parentFolder = DriveApp.getRootFolder();
      if (TARGET_FOLDER_ID && TARGET_FOLDER_ID.trim() !== '') {
        try {
          parentFolder = DriveApp.getFolderById(TARGET_FOLDER_ID.trim());
        } catch(e) {}
      }
      
      var certFolders = parentFolder.getFoldersByName('校正証明書');
      var certFolder;
      if (certFolders.hasNext()) {
        certFolder = certFolders.next();
      } else {
        certFolder = parentFolder.createFolder('校正証明書');
      }
      
      var pdfFile = certFolder.createFile(pdfBlob);
      safeSetValue(sheet, targetRow, headers, '証明書URL', pdfFile.getUrl(), []);
    } catch(e) {
      console.error('PDF保存エラー:', e);
    }
  }
  
  var mnIdx = headers.indexOf('管理番号');
  var managementNumber = String(values[targetRow - 1][mnIdx]);
  updateCalibrationEquipment({
    managementNumber: managementNumber,
    status: 'active'
  });
  
  return { success: true, updatedId: data.recordId, turnaroundDays: turnaroundDays };
}

function getCalibrationRecords() {
  var data = getSheetDataAsObjects('校正履歴');
  return data.map(function(row) {
    return {
      id: String(row['履歴ID'] || ''),
      managementNumber: String(row['管理番号'] || ''),
      sentDate: formatDate(row['送出日']),
      returnedDate: row['返却日'] ? formatDate(row['返却日']) : null,
      turnaroundDays: row['所要日数'] !== '' ? Number(row['所要日数']) : null,
      calibrationDate: formatDate(row['校正実施日']),
      nextCalibrationDate: formatDate(row['次回校正予定日']),
      result: String(row['結果'] || ''),
      certificateNumber: String(row['証明書番号'] || ''),
      certificateUrl: String(row['証明書URL'] || ''),
      cost: Number(row['費用']) || 0,
      vendor: String(row['校正業者'] || ''),
      notes: String(row['備考'] || ''),
      createdAt: formatDate(row['記録日'])
    };
  });
}

// ====== 校正機器種別マスタ CRUD ======

function getCalibrationCategories() {
  var data = getSheetDataAsObjects('校正機器種別マスタ');
  return data.map(function(row) {
    return {
      id: String(row['種別ID'] || ''),
      name: String(row['種別名'] || ''),
      defaultCycleMonths: Number(row['デフォルト校正周期(月)']) || 0
    };
  });
}

function addCalibrationCategory(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器種別マスタ');
  if (!sheet) throw new Error('校正機器種別マスタシートがありません');
  
  var existingData = getSheetDataAsObjects('校正機器種別マスタ');
  var maxNum = 0;
  for (var i = 0; i < existingData.length; i++) {
    var cid = String(existingData[i]['種別ID'] || '');
    var match = cid.match(/^cat-(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  var newId = 'cat-' + ('000' + (maxNum + 1)).slice(-3);
  
  sheet.appendRow([newId, data.name || '', data.defaultCycleMonths || 12]);
  return { success: true, id: newId };
}

function updateCalibrationCategory(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器種別マスタ');
  if (!sheet) throw new Error('校正機器種別マスタシートがありません');
  
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idIdx = headers.indexOf('種別ID');
  
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(data.id)) {
      var row = i + 1;
      var missingHeaders = [];
      safeSetValue(sheet, row, headers, '種別名', data.name || '', missingHeaders);
      safeSetValue(sheet, row, headers, 'デフォルト校正周期(月)', data.defaultCycleMonths || 0, missingHeaders);
      break;
    }
  }
  return { success: true, id: data.id };
}

function deleteCalibrationCategory(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正機器種別マスタ');
  if (!sheet) throw new Error('校正機器種別マスタシートがありません');
  
  var values = sheet.getDataRange().getValues();
  var idIdx = values[0].indexOf('種別ID');
  
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idIdx]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true, deletedId: data.id };
}

// ====== 全データ一括取得（初回読み込み用） ======

function getAllCalibrationData() {
  return {
    equipment: getCalibrationEquipment(),
    records: getCalibrationRecords(),
    categories: getCalibrationCategories()
  };
}

// ====== Google Chat Webhook 通知 ======

function sendChatNotification(data) {
  if (!data.webhookUrl) throw new Error('Webhook URLが指定されていません');
  if (!data.text) throw new Error('送信メッセージが空です');
  
  var options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ text: data.text }),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(data.webhookUrl, options);
  var code = response.getResponseCode();
  
  if (code === 200) {
    return { success: true, message: '通知を送信しました' };
  } else {
    throw new Error('Chat送信失敗: HTTP ' + code + ' - ' + response.getContentText());
  }
}
// ====== 校正証明書PDF単体アップロード ======

function uploadCalibrationCertificate(data) {
  if (!data.base64Data || !data.fileName) {
    throw new Error('PDF データまたはファイル名が指定されていません');
  }
  
  var pdfBlob = Utilities.newBlob(
    Utilities.base64Decode(data.base64Data),
    'application/pdf',
    data.fileName
  );
  
  // 校正証明書専用フォルダに直接保存
  var CERT_FOLDER_ID = '18iVqC9njVEZ4wbUwByA7yMcoBmZnuE9x';
  var certFolder;
  try {
    certFolder = DriveApp.getFolderById(CERT_FOLDER_ID);
  } catch(e) {
    // フォルダが見つからない場合はルートに保存
    certFolder = DriveApp.getRootFolder();
    console.warn('校正証明書フォルダが見つかりません。ルートに保存します。');
  }
  
  var pdfFile = certFolder.createFile(pdfBlob);
  
  // 校正履歴シートに証明書URLを記録（履歴IDがある場合）
  if (data.recordId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('校正履歴');
    if (sheet) {
      var values = sheet.getDataRange().getValues();
      var headers = values[0];
      var idIdx = headers.indexOf('履歴ID');
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][idIdx]) === String(data.recordId)) {
          safeSetValue(sheet, i + 1, headers, '証明書URL', pdfFile.getUrl(), []);
          break;
        }
      }
    }
  }
  
  return {
    success: true,
    fileUrl: pdfFile.getUrl(),
    fileId: pdfFile.getId(),
    fileName: data.fileName
  };
}

// =====================================================================
// ====== スタッフマスタ API ======
// =====================================================================

var STAFF_HEADERS = ['社員ID', '氏名', '所属課', '電話番号', '保有資格', '夜勤可否', '休日出勤可否', '備考'];

function ensureStaffHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スタッフマスタ');
  if (!sheet) throw new Error('スタッフマスタシートがありません');
  
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]);
    sheet.getRange(1, 1, 1, STAFF_HEADERS.length).setFontWeight('bold').setBackground('#1a56db').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getStaffList() {
  ensureStaffHeaders();
  var data = getSheetDataAsObjects('スタッフマスタ');
  return data.map(function(row) {
    return {
      id: String(row['社員ID'] || ''),
      name: String(row['氏名'] || ''),
      department: String(row['所属課'] || ''),
      phone: String(row['電話番号'] || ''),
      qualification: String(row['保有資格'] || ''),
      notes: String(row['備考'] || '')
    };
  });
}

function addStaffMember(data) {
  var sheet = ensureStaffHeaders();
  
  var existing = getSheetDataAsObjects('スタッフマスタ');
  
  // 課別の採番: 2課→S2xx, 3課→S3xx, 4課→S4xx, 大阪→S5xx
  var prefix = 100;
  if (data.department === '2課') prefix = 200;
  else if (data.department === '3課') prefix = 300;
  else if (data.department === '4課') prefix = 400;
  else if (data.department === '大阪') prefix = 500;
  
  var deptExisting = existing.filter(function(r) { return String(r['所属課']) === data.department; });
  var deptMax = 0;
  for (var j = 0; j < deptExisting.length; j++) {
    var sid = String(deptExisting[j]['社員ID'] || '');
    var m = sid.match(/^S(\d+)$/);
    if (m) {
      var n = parseInt(m[1], 10) % 100;
      if (n > deptMax) deptMax = n;
    }
  }
  var newId = 'S' + (prefix + deptMax + 1);
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = new Array(headers.length).fill('');
  
  var idIdx = headers.indexOf('社員ID'); if (idIdx !== -1) newRow[idIdx] = newId;
  var nameIdx = headers.indexOf('氏名'); if (nameIdx !== -1) newRow[nameIdx] = data.name || '';
  var deptIdx = headers.indexOf('所属課'); if (deptIdx !== -1) newRow[deptIdx] = data.department || '';
  var phoneIdx = headers.indexOf('電話番号'); if (phoneIdx !== -1) newRow[phoneIdx] = data.phone || '';
  var qualIdx = headers.indexOf('保有資格'); if (qualIdx !== -1) newRow[qualIdx] = data.qualification || '';
  var notesIdx = headers.indexOf('備考'); if (notesIdx !== -1) newRow[notesIdx] = data.notes || '';
  
  sheet.appendRow(newRow);
  return { success: true, newId: newId };
}

function updateStaffMember(data) {
  var sheet = ensureStaffHeaders();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idIdx = headers.indexOf('社員ID');
  
  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIdx]) === String(data.id)) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) throw new Error('社員IDが見つかりません: ' + data.id);
  
  var missingHeaders = [];
  if (data.name !== undefined) safeSetValue(sheet, targetRow, headers, '氏名', data.name, missingHeaders);
  if (data.department !== undefined) safeSetValue(sheet, targetRow, headers, '所属課', data.department, missingHeaders);
  if (data.phone !== undefined) safeSetValue(sheet, targetRow, headers, '電話番号', data.phone, missingHeaders);
  if (data.qualification !== undefined) safeSetValue(sheet, targetRow, headers, '保有資格', data.qualification, missingHeaders);
  if (data.notes !== undefined) safeSetValue(sheet, targetRow, headers, '備考', data.notes, missingHeaders);
  
  return { success: true, updatedId: data.id };
}
