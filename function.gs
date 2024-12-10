function sendLineMessage(newsRows, keyword) {
  const channelAccessToken = 'RYawmNOqXelUkfCQaic+h1FWwE/SPH/2cRzwuJiadzoISTW2cB7uI/kpOAJ/siYr7pTFdKVTEu/1xfdfj7EAGUpqiR58sB2yy+Kdqn99EUve4UzfXDJqMkz95vrFexl1BoQkOkaP0wttCwEvJOnF+wdB04t89/1O/w1cDnyilFU=';
  //const channelSecret = 'YOUR_CHANNEL_SECRET';

    const messages = [
      {
        type: 'text',
        text: '以下是今日更新標案'
      }
    ];
    const options_title = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + channelAccessToken
      },
      'payload': JSON.stringify({
        messages
      })
   };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', options_title);

  if (newsRows.length > 0) {
    const messages = [];
    newsRows.forEach(row => {
      const [date, type, title, category, unitName, url, filename, imageUrl] = row;

      const message = {
        "type": "bubble",
        "body": {
          "type": "box",
          "layout": "vertical",
          "contents": [
            {
              "type": "text",
              "text": `${type}`,
              "weight": "bold",
              "size": "xl"
            },
            {
              "type": "text",
              "text": `${title}`,
              "wrap": true
            },
            {
              "type": "text",
              "text": `單位：${unitName}`,
              "wrap": true
            },
            {
              "type": "text",
              "text": `發布日期：${date}`,
              "wrap": true
            }
          ]
        },
        "footer": {
          "type": "box",
          "layout": "vertical",
          "spacing": "sm",
          "contents": [
            {
              "type": "button",
              "style": "link",
              "height": "sm",
              "action": {
                "type": "uri",
                "label": "查看詳細資訊",
                "uri": url
              }
            }
          ]
        }
      }
      messages.push(message);
    })

    const flexMessage = {
      "type": "flex",
      "altText": `最新 ${keyword} 標案資訊`,
      "contents": {
        "type": "carousel",
        "contents": messages
      }
    };

    const options = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + channelAccessToken
      },
      'payload': JSON.stringify({
        messages: [flexMessage]
      })
    };

      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', options);
    Logger.log("LINE訊息已發送");
  } else {
    Logger.log("沒有新資料需要發送至 LINE");
  }
}

function sendEmailNotification(newsRows) {
  if (newsRows.length > 0) {
    // 建構 HTML 內容
    let htmlBody = "<p>以下是最新的標案資料：</p><table border='1'>";
    htmlBody += "<tr><th>類型</th><th>標題</th><th>分類</th><th>單位名稱</th><th>網址</th><th>關鍵字</th></tr>";
    newsRows.forEach(row => {
      htmlBody += `<tr><td>${row[1]}</td><td>${row[2]}</td><td>${row[3]}</td><td>${row[4]}</td><td><a href="${row[5]}">${row[5]}</a></td><td>${row[7]}</td></tr>`;
    });
    htmlBody += "</table>";

    // 設定收件人、標題和內容
    const recipient = PropertiesService.getScriptProperties().getProperty('EMAIL');
    const subject = "今日新標案通知";
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });

    Logger.log("Email 已發送!");
  } else {
    Logger.log("沒有新資料需要發送");
  }
}

function fetchAndFillData() {
  const spreadsheetId = "1rw8LnJxTzrVmNh84KTWNDqEJlVSgzlaRoYeqJRyljug"; // 試算表 ID
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const newsSheetName = "news";
  let newsSheet = spreadsheet.getSheetByName(newsSheetName);
  if (!newsSheet) {
    newsSheet = spreadsheet.insertSheet(newsSheetName);
    const headers = ["日期", "類型", "標題", "分類", "單位名稱", "URL", "檔案名稱", "關鍵字"];
    newsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 設置資料驗證（示例：日期欄位）
    const range = newsSheet.getRange(2, 1, newsSheet.getLastRow(), 1); // 假設日期欄位為第一列
    const rule = SpreadsheetApp.newDataValidation().requireDate().build();
    range.setDataValidation(rule);
  } else {
    // 清空資料，但保留表頭
    if (newsSheet.getLastRow() > 1) {
      newsSheet.getRange(2, 1, newsSheet.getLastRow() - 1, newsSheet.getLastColumn()).clearContent();
    }
  }

  // 收集所有新資料以添加到 news 工作表
  const newsRows = [];

  // 處理關鍵字的資料
  const keywordSheet = spreadsheet.getSheetByName("keyword");
  if (!keywordSheet) {
    throw new Error("找不到名為 'keyword' 的工作表。");
  }

  const keywords = keywordSheet.getRange(1, 1, keywordSheet.getLastRow(), 1).getValues().flat();

  keywords.forEach(keyword => {
    const sheetName = `data_${keyword}`;
    const apiUrl = `https://pcc.g0v.ronny.tw/api/searchbytitle?query=${keyword}`;

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      const headers = ["日期", "類型", "標題", "分類", "單位名稱", "URL", "檔案名稱"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const response = UrlFetchApp.fetch(apiUrl);
    const data = JSON.parse(response.getContentText());
    const records = data.records || [];
    const existingFilename = sheet.getLastRow() > 1
      ? sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues().flat()
      : [];

    const newRows = [];

    records.forEach(record => {
      const date = record.date || "";
      const type = record.brief.type || "";
      const title = record.brief.title || "";
      const category = record.brief.category || "";
      const unitName = record.unit_name || "";
      const tender_api_url = record.tender_api_url;
      const filename = record.filename || "";

      const urlResponse = UrlFetchApp.fetch(tender_api_url);
      const urlData = JSON.parse(urlResponse.getContentText());
      const urlRecords = urlData.records || [];

      const url = urlRecords.length > 0 && urlRecords[0]?.detail?.url
        ? urlRecords[0].detail.url
        : null;

      if (!existingFilename.includes(filename)) {
        const newRow = [date, type, title, category, unitName, url, filename];
        newRows.push(newRow);
        newsRows.push([...newRow, keyword]); // 將新資料加上關鍵字並添加到 news 列表
        existingFilename.push(filename);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    } else {
      Logger.log(`關鍵字 "${keyword}" 無新增資料`);
    }
  });

  // 更新 news 工作表
  if (newsRows.length > 0) {
    newsSheet.getRange(2, 1, newsRows.length, newsRows[0].length).setValues(newsRows);
    // 發送電子郵件通知
    sendEmailNotification(newsRows);

    // 發送 LINE 訊息通知
    sendLineMessage(newsRows);
  } else {
    Logger.log("此次執行無新增資料到 news 工作表");
  }
}
