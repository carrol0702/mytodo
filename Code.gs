function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('我的手繪便條紙')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 取得或建立工作表
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Tasks');
  if (!sheet) {
    sheet = ss.insertSheet('Tasks');
    sheet.appendRow(['ID', 'Content', 'Date', 'Type', 'Color']); // 建立標題欄
  }
  return sheet;
}

function getTasks() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // 如果只有標題，回傳空陣列
    
    const headers = data.shift();
    return data.map(row => ({
      id: row[0],
      content: row[1],
      date: row[2] instanceof Date ? row[2].toLocaleDateString() : row[2],
      type: row[3],
      color: row[4]
    }));
  } catch (e) {
    return [];
  }
}

function addTask(task) {
  try {
    const sheet = getSheet();
    const id = new Date().getTime();
    sheet.appendRow([id, task.content, task.date, task.type, task.color]);
    return getTasks();
  } catch (e) {
    throw new Error("新增失敗：" + e.toString());
  }
}

function deleteTask(id) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    // 從最後一行往回找（這樣刪除行時才不會影響到前面的索引）
    for (let i = data.length - 1; i >= 1; i--) {
      // 強制將兩者轉為字串並去除空白，確保匹配成功
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        break; // 找到後刪除並跳出迴圈
      }
    }
    return getTasks(); // 回傳更新後的清單
  } catch (e) {
    throw new Error("刪除失敗：" + e.toString());
  }
}