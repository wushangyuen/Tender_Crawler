const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SHEETID');
const recipient = PropertiesService.getScriptProperties().getProperty('EMAIL');
const channelAccessToken = PropertiesService.getScriptProperties().getProperty('TOKEN');

function fetchTableData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the 'keyword' sheet and get the list of keywords in column A
  const keywordSheet = spreadsheet.getSheetByName('keyword');
  if (!keywordSheet) {
    Logger.log('Keyword sheet not found.');
    return;
  }
  
  const keywords = keywordSheet.getRange('A1:A').getValues().flat().filter(String);  // Get keywords, ignoring empty cells
  
  // Access or create the 'news' sheet
  let newsSheet = spreadsheet.getSheetByName('news');
  if (!newsSheet) {
    newsSheet = spreadsheet.insertSheet('news');
  } else {
    newsSheet.clear();  // Clear the 'news' sheet each time the script runs
  }
  
  // Add header row to the 'news' sheet
  newsSheet.appendRow(['Type', 'Organization', 'Tender Name', 'Tender Notice Date', 'Award Notice', 'Bid Submission Deadline', 'Public Viewing Date', 'Pre-announcement Date', 'View Link', 'Keyword']);
  
  // Get current ROC year
  const currentYear = new Date().getFullYear();
  const rocYear = currentYear - 1911;  // Calculate ROC year
  
  keywords.forEach(keyword => {
    // Construct the URL using the keyword and dynamic ROC year
    const url = `https://web.pcc.gov.tw/prkms/tender/common/bulletion/readBulletion?querySentence=${encodeURIComponent(keyword)}&tenderStatusType=%E6%8B%9B%E6%A8%99&tenderStatusType=%E6%B1%BA%E6%A8%99&tenderStatusType=%E5%85%AC%E9%96%8B%E9%96%B1%E8%A6%BD%E5%8F%8A%E5%85%AC%E9%96%8B%E5%BE%B5%E6%B1%82&tenderStatusType=%E6%94%BF%E5%BA%9C%E6%8E%A1%E8%B3%BC%E9%A0%90%E5%91%8A&sortCol=TENDER_NOTICE_DATE&timeRange=${rocYear}&pageSize=100`;
    
    // Fetch the content for the current keyword
    const response = UrlFetchApp.fetch(url);
    const htmlContent = response.getContentText();
    
    // Use regular expressions to extract the content between <table class="tb_01" id="bulletion"> and </table>
    const regex = /<table class="tb_01" id="bulletion">([\s\S]*?)<\/table>/gi;
    const match = regex.exec(htmlContent);
    
    if (match) {
      const tableContent = match[1];  // Get the content inside <table>
      const rows = [];
      const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
      let rowMatch;
      const existingLinks = new Set();  // To track existing links and avoid duplicates

      // Fetch existing links from the 'data_${keyword}' sheet if it exists
      let keywordSheet = spreadsheet.getSheetByName(`data_${keyword}`);
      let existingDataLinks = new Set();
      if (keywordSheet) {
        const existingData = keywordSheet.getDataRange().getValues();
        for (let i = 1; i < existingData.length; i++) {  // Start at index 1 to skip the header row
          const viewLink = existingData[i][8];  // Assuming the view link is in the 9th column
          if (viewLink) {
            existingDataLinks.add(viewLink);
          }
        }
      }

      // Loop through the rows and parse the data
      while ((rowMatch = rowRegex.exec(tableContent)) !== null) {
        const row = [];
        const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
        let cellMatch;
        let cellCount = 0;
        let rowData = {};
        let viewLink = null;

        while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
          switch (cellCount) {
            case 1:  // Type (Type)
              rowData.type = cellMatch[1].trim();
              break;
            case 2:  // Organization
              rowData.organization = cellMatch[1].trim();
              break;
            case 3:  // Tender Name
              const tenderNameMatch = /var hw = Geps3.CNS.pageCode2Img\("([^"]+)"\)/gi.exec(cellMatch[1]);
              if (tenderNameMatch) {
                rowData.tenderName = tenderNameMatch[1].trim();  // Extract Tender Name
              }
              break;
            case 4:  // 招標公告日期 (Tender Notice Date)
              rowData.tenderNoticeDate = cellMatch[1].trim(); // Already in ROC format
              break;
            case 5:  // 決標或無法決標公告 (Award or Non-Award Notice)
              let awardStatus = cellMatch[1].trim();
              if (awardStatus.includes('無法決標')) {
                awardStatus = awardStatus.replace('<span style="color:red"><br>(無法決標)</span>', '').trim() + "\n(non award)";
              }
              rowData.awardNotice = awardStatus;
              break;
            case 6:  // 截止投標日期 (Bid Submission Deadline)
              rowData.bidDeadline = cellMatch[1].trim();  // Already in ROC format
              break;
            case 7:  // 公開閱覽/徵求日期 (Public Viewing Date)
              const publicViewingDate = cellMatch[1].trim();
              // Handle concatenating the dates for Public Viewing Date
              const dateRangeMatch = /([\d/]+)\s*<br>\s*~\s*<br>\s*([\d/]+)/.exec(publicViewingDate);
              if (dateRangeMatch) {
                rowData.publicViewingDate = dateRangeMatch[1] + '~' + dateRangeMatch[2];  // Concatenate date range
              } else {
                rowData.publicViewingDate = publicViewingDate;
              }
              break;
            case 8:  // 預告公告日期 (Pre-announcement Date)
              rowData.preAnnouncementDate = cellMatch[1].trim();  // Already in ROC format
              break;
            case 9:  // View Link
              const viewLinkMatch = /<a href="([^"]+)">([\s\S]*?)<\/a>/gi.exec(cellMatch[1]);
              if (viewLinkMatch) {
                const tenderLinkId = viewLinkMatch[1].match(/pk=([^&]+)/)[1];
                viewLink = `https://web.pcc.gov.tw/tps/QueryTender/query/searchTenderDetail?pkPmsMain=${tenderLinkId}`;
              }
              break;
          }
          cellCount++;
        }

        if (viewLink && !existingLinks.has(viewLink) && !existingDataLinks.has(viewLink)) {
          existingLinks.add(viewLink);  // Add to the set to avoid duplicates
          rowData.viewLink = viewLink;
          rowData.keyword = keyword;  // Add the keyword to the row data for the 'news' sheet
          
          // Add the row data to the 'news' sheet
          newsSheet.appendRow([
            rowData.type || '', 
            rowData.organization || '', 
            rowData.tenderName || '', 
            rowData.tenderNoticeDate || '', 
            rowData.awardNotice || '', 
            rowData.bidDeadline || '', 
            rowData.publicViewingDate || '', 
            rowData.preAnnouncementDate || '', 
            rowData.viewLink || '', 
            rowData.keyword || ''
          ]);

          // Add the row data to the 'data_${keyword}' sheet
          let keywordSheet = spreadsheet.getSheetByName(`data_${keyword}`);
          if (!keywordSheet) {
            keywordSheet = spreadsheet.insertSheet(`data_${keyword}`);
            keywordSheet.appendRow(['Type', 'Organization', 'Tender Name', 'Tender Notice Date', 'Award Notice', 'Bid Submission Deadline', 'Public Viewing Date', 'Pre-announcement Date', 'View Link', 'Keyword']);
          }

          const rowValues = [
            rowData.type || '', 
            rowData.organization || '', 
            rowData.tenderName || '', 
            rowData.tenderNoticeDate || '', 
            rowData.awardNotice || '', 
            rowData.bidDeadline || '', 
            rowData.publicViewingDate || '', 
            rowData.preAnnouncementDate || '', 
            rowData.viewLink || '', 
            rowData.keyword || ''
          ];
          keywordSheet.appendRow(rowValues);
        }
      }

      Logger.log(`Data for keyword "${keyword}" successfully written to sheet "data_${keyword}"`);
      
    } else {
      Logger.log(`No data found for keyword "${keyword}"`);
    }
  });
  sendNotifications();
}

function getNewsData() {
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('news');
  const data = sheet.getDataRange().getValues();  // Get all the rows in the sheet
  let newsRows = [];  // Change const to let, as this will be modified

  // Skip the header row (assuming it's the first row)
  for (let i = 1; i < data.length; i++) {
    let row = data[i];  // Change const to let, as this will be modified
    let type = row[0];        // Type
    let organization = row[1]; // Organization
    let tenderName = row[2];   // Tender Name
    let tenderNoticeDate = row[3];  // Tender Notice Date
    let awardNotice = row[4];   // Award Notice
    let bidSubmissionDeadline = row[5]; // Bid Submission Deadline
    let publicViewingDate = row[6];  // Public Viewing Date
    let preAnnouncementDate = row[7]; // Pre-announcement Date
    let viewLink = row[8];         // View Link
    let keyword = row[9];         // Keyword

    // Convert dates from the format "民國年" to "yyyy/MM/dd" format
    tenderNoticeDate = formatDate(tenderNoticeDate);
    bidSubmissionDeadline = formatDate(bidSubmissionDeadline);
    publicViewingDate = formatDate(publicViewingDate);  // Handle range if necessary

    // Combine date ranges for Public Viewing Date
    if (publicViewingDate.includes("~")) {
      publicViewingDate = publicViewingDate.replace(/\s*~\s*/, "~");
    }

    // Add the row to the newsRows array if it contains valid data
    if (viewLink) {
      newsRows.push([
        tenderNoticeDate, 
        type, 
        tenderName, 
        organization, 
        bidSubmissionDeadline, 
        publicViewingDate, 
        viewLink, 
        keyword
      ]);
    }
  }

  return newsRows; // Return the array of news rows
}

function formatDate(date) {
  if (!date) return ""; // If the date is empty, return empty string
  if (date instanceof Date) {
    // If it's a JavaScript Date object, format it
    const year = date.getFullYear();
    const month = date.getMonth() + 1; // Months are 0-indexed
    const day = date.getDate();
    return `${year}/${pad(month)}/${pad(day)}`;
  }
  
  // If it's a string date, convert it into Date first and format
  const dateObj = new Date(date);
  if (!isNaN(dateObj.getTime())) {
    const year = dateObj.getFullYear() - 1911;
    const month = dateObj.getMonth() + 1;
    const day = dateObj.getDate();
    return `${year}/${pad(month)}/${pad(day)}`;
  }

  return date; // If it's not a valid date, return it as is
}

function pad(n) {
  return n < 10 ? '0' + n : n; // Ensure two-digit format for month and day
}


// Call the functions to send email and LINE message
function sendNotifications() {
  const newsRows = getNewsData(); // Get the news data from the "news" sheet
  
  // Send email notification if there are new rows
  sendEmailNotification(newsRows);
  
  // Send LINE message if there are new rows
  sendLineMessage(newsRows);
}

// Send email notification function
function sendEmailNotification(newsRows) {
  if (newsRows.length > 0) {
    // Fetch all email addresses from the 'Email' sheet
    const emailSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Email');
    const emailAddresses = emailSheet.getRange('A1:A').getValues().flat().filter(String);  // Get emails, ignoring empty cells
    
    if (emailAddresses.length > 0) {
      // Build HTML content for the email
      var date = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy");
      let htmlBody = `<p>以下是${date}標案資料：</p><table border='1'>`;
      htmlBody += "<tr><th>類型</th><th>標題</th><th>單位名稱</th><th>標案公告日期</th><th>截止投標日期</th><th>公開閱覽日期</th><th>網址</th><th>關鍵字</th></tr>";
      newsRows.forEach(row => {
        htmlBody += `<tr><td>${row[1]}</td><td>${row[2]}</td><td>${row[3]}</td><td>${row[0]}</td><td>${row[4]}</td><td>${row[5]}</td><td><a href="${row[6]}">${row[6]}</a></td><td>${row[7]}</td></tr>`;
      });
      htmlBody += "</table>";

      // Email settings
      const subject = `${date}新標案通知`;

      // Loop through each email address and send the email
      emailAddresses.forEach(function(email) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody
        });
        Logger.log(`Email sent to: ${email}`);
      });

      Logger.log("Email 已發送給所有人!");
    } else {
      Logger.log("沒有在Email表格中找到任何電子郵件地址");
    }
  } else {
    Logger.log("沒有新資料需要發送");
  }
}


// Send LINE message function
function sendLineMessage(newsRows) {
  const messages = [
    {
      type: 'text',
      text: Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy") + '新增標案'
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
    // Split data into chunks of 12 items each for better formatting in LINE
    const chunks = chunkArray(newsRows, 12);
    
    chunks.forEach(chunk => {
      const messages = chunk.map(row => {
        const [date, type, title, organization, bidSubmissionDeadline, publicViewingDate, url, keyword] = row;

        return {
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
                "text": `單位：${organization}`,
                "wrap": true
              },
              {
                "type": "text",
                "text": `關鍵字：${keyword}`,
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
        };
      });

      const flexMessage = {
        "type": "flex",
        "altText": Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy") + `標案資訊`,
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
    });

    Logger.log("LINE訊息已發送");
  } else {
    const messages = [
    {
      type: 'text',
      text: Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy") + '新增標案'
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
    Logger.log("沒有新資料需要發送至 LINE");
  }
}

// Helper function to split an array into smaller chunks
function chunkArray(array, chunkSize) {
  const result = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    result.push(array.slice(i, i + chunkSize));
  }
  return result;
}

// const GROUP_ID_LIST_SHEET_NAME = 'GroupIDs'; // Name of the sheet where group IDs are stored

// function doPost(e) {
//   // Log the entire request for debugging purposes
//   Logger.log('Request received: ' + JSON.stringify(e));

//   // Check if the event is a join or leave event
//   var data = JSON.parse(e.postData.contents);
//   if (data.events) {
//     data.events.forEach(function(event) {
//       Logger.log('Event received: ' + JSON.stringify(event));

//       if (event.type === 'join') {
//         handleJoin(event);
//       } else if (event.type === 'leave') {
//         handleLeave(event);
//       }
//     });
//   }
//   return ContentService.createTextOutput('OK');
// }


// function handleJoin(event) {
//   const groupId = event.source.groupId;
//   if (groupId) {
//     addGroupIdToSheet(groupId);  // Add the group ID to the list
//     Logger.log('Bot joined group: ' + groupId);
//   }
// }

// function handleLeave(event) {
//   const groupId = event.source.groupId;
//   if (groupId) {
//     removeGroupIdFromSheet(groupId);  // Remove the group ID from the list
//     Logger.log('Bot left group: ' + groupId);
//   }
// }

// function addGroupIdToSheet(groupId) {
//   const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(GROUP_ID_LIST_SHEET_NAME);
//   const existingGroupIds = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat(); // Get all existing group IDs
  
//   if (!existingGroupIds.includes(groupId)) {
//     sheet.appendRow([groupId]);  // Add the new group ID if it doesn't exist
//     Logger.log('Group ID added to the list: ' + groupId);
//   }
// }

// function removeGroupIdFromSheet(groupId) {
//   const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(GROUP_ID_LIST_SHEET_NAME);
//   const existingGroupIds = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat(); // Get all existing group IDs
  
//   const rowIndex = existingGroupIds.indexOf(groupId);
//   if (rowIndex !== -1) {
//     sheet.deleteRow(rowIndex + 1);  // Delete the row containing the group ID (1-indexed)
//     Logger.log('Group ID removed from the list: ' + groupId);
//   }
// }
