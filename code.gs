// ==========================================
// ส่วนตั้งค่า (Config)
// ==========================================
const CHANNEL_ACCESS_TOKEN = 'ISu56Cj+AJdBJA8HSK/svoiT0gkvJ9Hbv2f6dGrgwu1eO3PFh+01WHGIwjvSj5A0sARxvCKb3wqT/PS74ro3G4MkJXSx63SMHy5WBPgl7hxrJggJwGG/WhRVBM90Yo/btA5P2nUrpuX9G6SBQvT1yQdB04t89/1O/w1cDnyilFU='; 
const SHEET_NAME = 'Data'; 
const GROUP_ID_TARGET = 'Cc882e6046c69af3830ae71a0582e377b'; 

// ==========================================
// ส่วนที่ 1: Web App & Data Management
// ==========================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- ฟังก์ชันดึงข้อมูล ---
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0]; 

  const rows = sheet.getDataRange().getDisplayValues();
  if (rows.length <= 1) return "[]"; 
  rows.shift(); // เอาหัวตารางออก
  return JSON.stringify(rows);
}

// --- ฟังก์ชันบันทึกข้อมูล ---
function saveData(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  let birthDateVal = "";
  if (formObject.birthdate) {
    const parts = formObject.birthdate.split('-'); 
    if (parts.length === 3) {
       birthDateVal = `${parts[2]}/${parts[1]}/${parts[0]}`; 
    }
  }

  const dataRow = [
      formObject.name,
      formObject.nickname,
      formObject.position,
      formObject.dept,
      "'" + formObject.phone,
      formObject.lineId,
      formObject.email,
      birthDateVal
  ];

  if (formObject.recId && formObject.recId !== "") {
    const data = sheet.getDataRange().getValues(); 
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() == formObject.recId.toString()) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2, 1, 8).setValues([dataRow]);
        break;
      }
    }
  } else {
    const newId = new Date().getTime().toString();
    const newRow = [newId, ...dataRow]; 
    sheet.appendRow(newRow);
  }
  
  try { CacheService.getScriptCache().remove("memberData"); } catch(e){}
  return "บันทึกข้อมูลเรียบร้อย";
}

// ฟังก์ชันลบข้อมูล
function deleteData(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() == id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  try { CacheService.getScriptCache().remove("memberData"); } catch(e){}
  return "ลบข้อมูลเรียบร้อย";
}

// ==========================================
// ส่วนที่ 2: ระบบแจ้งเตือนวันเกิด (Birthday Alert)
// ==========================================

function checkAndSendBirthday() {
  console.log("เริ่มตรวจสอบวันเกิด...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  const data = sheet.getDataRange().getValues(); 
  
  const now = new Date();
  const timeZone = "Asia/Bangkok";
  const currentDay = parseInt(Utilities.formatDate(now, timeZone, "d")); 
  const currentMonth = parseInt(Utilities.formatDate(now, timeZone, "M"));
  
  let birthdayList = [];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let name = row[1];       
    let nickname = row[2];   
    let birthDateRaw = row[8]; 

    if (!birthDateRaw || birthDateRaw === "") continue;

    let bDay = 0;
    let bMonth = 0;

    if (birthDateRaw instanceof Date) {
      let dateString = Utilities.formatDate(birthDateRaw, "Asia/Bangkok", "d/M");
      let parts = dateString.split('/');
      bDay = parseInt(parts[0]);
      bMonth = parseInt(parts[1]);
    } else {
      let strDate = String(birthDateRaw).split('T')[0];
      if (strDate.includes('/')) {
         let parts = strDate.split('/');
         bDay = parseInt(parts[0]);
         bMonth = parseInt(parts[1]);
      } else if (strDate.includes('-')) {
         let parts = strDate.split('-');
         bDay = parseInt(parts[2]);
         bMonth = parseInt(parts[1]);
      }
    }

    if (bDay === currentDay && bMonth === currentMonth) {
      let displayName = nickname ? `${name} (${nickname})` : name;
      birthdayList.push(displayName);
    }
  }

  if (birthdayList.length > 0) {
    sendGroupPushMessage(birthdayList);
  } else {
    console.log("วันนี้ไม่มีสมาชิกเกิด");
  }
}

function sendGroupPushMessage(names) {
  if (!names || names.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetConfig = ss.getSheetByName("Config"); 
  let targetGroupsRaw = [];

  // ดึง ID กลุ่มทั้งหมดจาก Sheet "Config"
  if (sheetConfig) {
    const rows = sheetConfig.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      let gid = rows[i][0];
      if (gid && gid.toString().trim() !== "") {
        targetGroupsRaw.push(gid.toString().trim());
      }
    }
  }

  if (targetGroupsRaw.length === 0 && typeof GROUP_ID_TARGET !== 'undefined' && GROUP_ID_TARGET !== "") {
     targetGroupsRaw.push(GROUP_ID_TARGET);
  }

  if (targetGroupsRaw.length === 0) {
    console.log("ไม่พบ Group ID สำหรับส่งข้อความ");
    return;
  }

  // กรอง Group ID ที่ซ้ำกันออก เผื่อมีการบันทึกซ้ำ
  const targetGroups = [...new Set(targetGroupsRaw)];

  const today = new Date().toLocaleDateString('th-TH', {
    day: 'numeric', month: 'long', year: 'numeric'
  });

  const nameList = names.map(name => `✨ ${name}`).join('\n');
  const messageText = `🎂 Happy Birthday! 🎂 ท่านสมาชิก พสพ.13 ประจำวันที่ ${today}\n\n${nameList}\n\n🙏 ขออาราธนาคุณพระศรีรัตนตรัย โปรดดลบันดาลให้ท่านมีความสุข 💖 สุขภาพแข็งแรง 💪 คิดสิ่งใดสมปรารถนา 🌟 การงานก้าวหน้า 📈 ร่ำรวยเงินทอง 💰 และประสบความสำเร็จยิ่งๆ ขึ้นไป 🎉`;

  let successCount = 0;

  // วนลูปส่งทีละกลุ่ม พร้อมตั้งเวลาหน่วง (Delay) ป้องกัน LINE API บล็อก
  for (let i = 0; i < targetGroups.length; i++) {
    let targetId = targetGroups[i];

    const payload = {
      'to': targetId, 
      'messages': [{ 'type': 'text', 'text': messageText }]
    };

    const options = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true 
    };

    try {
      let response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
      let responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
         successCount++;
         console.log(`ส่งวันเกิดสำเร็จ -> ID กลุ่ม: ${targetId}`);
      } else {
         console.log(`ส่งพลาด ID: ${targetId} | โค้ด: ${responseCode} | สาเหตุ: ${response.getContentText()}`);
      }
    } catch (e) {
      console.log(`เกิดข้อผิดพลาดในการส่ง LINE (Push) หา ${targetId}: ${e}`);
    }

    // *** เพิ่มระบบหน่วงเวลา 1 วินาที (1000 ms) ก่อนส่งกลุ่มต่อไป ป้องกัน Rate Limit ***
    Utilities.sleep(1000); 
  }

  console.log(`กระบวนการเสร็จสิ้น: ส่งสำเร็จ ${successCount}/${targetGroups.length} ปลายทาง`);
}

// ==========================================
// ส่วนที่ 3: Chatbot (Line OA)
// ==========================================

function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var events = json.events;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      
      // === เช็คว่าเป็นกลุ่มไหม ถ้าใช่ให้บันทึก (แก้ไขให้ทำงานชัวร์ขึ้น) ===
      if (event.source.type === "group" || event.source.type === "room") {
         var groupId = event.source.groupId || event.source.roomId;
         if (groupId) {
             saveGroupIdToSheet(groupId);
         }
      }
      // ===============================================

      if (event.type === 'message' && event.message.type === 'text') {
        handleMessage(event);
      }
    }
  } catch (error) {
    console.log("Error: " + error);
  }
}

function handleMessage(event) {
  const userMsg = event.message.text.trim();
  const replyToken = event.replyToken;

  if (userMsg === 'วิธีค้นหาเพื่อน') {
    const helpText = `📌 วิธีค้นหาข้อมูลทำเนียบรุ่น\n\nกรุณาพิมพ์คำว่า "พสพ " (มีเว้นวรรค 1 ครั้ง) แล้วตามด้วยชื่อ, ชื่อเล่น, หรือสังกัดที่ต้องการค้นหา \n\n💡 ตัวอย่างการค้นหา:\nพสพ สมชาย\nพสพ เทคโน\nพสพ นก`;
    replyText(replyToken, helpText);
    return;
  }

  if (userMsg === 'สุ่ม' || userMsg === 'สุ่มชื่อ') {
    randomLuckyDraw(replyToken);
    return;
  }

  if (userMsg.startsWith('พสพ ')) {
    const keyword = userMsg.substring(4).trim(); 
    if (keyword.length === 0) return;

    const result = searchMemberForBot(keyword);
    if (result) {
      replyFlexMessage(replyToken, result);
    } else {
      replyText(replyToken, 'ไม่พบข้อมูลของ "' + keyword + '"');
    }
    return;
  }
}

function replyText(replyToken, text) {
  const payload = {
    "replyToken": replyToken,
    "messages": [{ "type": "text", "text": text }]
  };
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    'payload': JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

// ฟังก์ชันค้นหาสำหรับ Bot
function searchMemberForBot(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  
  const data = sheet.getDataRange().getDisplayValues(); 
  const searchKey = keyword.toString().toLowerCase().trim(); 
  const searchKeyDigits = searchKey.replace(/\D/g, ''); 

  let foundMembers = []; 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const fullName = (row[1] || "").toString().trim();
    const nickName = (row[2] || "").toString().trim();
    const position = (row[3] || "").toString().trim();
    const dept     = (row[4] || "").toString().trim();
    const phone    = (row[5] || "").toString().trim();
    const lineId   = (row[6] || "").toString().trim();
    const email    = (row[7] || "").toString().trim();

    const phoneDigits = phone.replace(/\D/g, '');
    const lowerName = fullName.toLowerCase();
    const lowerNick = nickName.toLowerCase();
    const lowerDept = dept.toLowerCase();

    let matchPriority = 99;

    if (lowerNick.includes(searchKey)) matchPriority = 1;
    else if (lowerName.startsWith(searchKey)) matchPriority = 2;
    else if (lowerName.includes(searchKey)) matchPriority = 3;
    else if (lowerDept.includes(searchKey)) matchPriority = 4;
    else if (
      lineId.toLowerCase().includes(searchKey) || 
      email.toLowerCase().includes(searchKey) || 
      (searchKeyDigits.length >= 3 && phoneDigits.includes(searchKeyDigits))
    ) {
      matchPriority = 5;
    }

    if (matchPriority !== 99) {
      foundMembers.push({
        name: fullName, nickname: nickName, position: position,
        dept: dept, phone: phone, line: lineId, email: email, _priority: matchPriority 
      });
    }
  }

  foundMembers.sort((a, b) => a._priority - b._priority);
  return foundMembers.length > 0 ? foundMembers : null;
}

// ฟังก์ชัน Flex Message
function replyFlexMessage(replyToken, members) {
  let webAppUrl = "https://script.google.com";
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) {}

  const displayMembers = members.slice(0, 10); 

  const bubbles = displayMembers.map(data => {
    const safeName = (data.name || "ไม่ระบุชื่อ").toString();
    const safeNick = (data.nickname || "-").toString();
    const safePos = (data.position || "-").toString();
    const safeDept = (data.dept || "-").toString();
    const safePhone = (data.phone || "-").toString();
    
    const cleanPhone = safePhone.replace(/\D/g, ''); 
    const lineId = (data.line || "").toString().trim();

    let footerContents = [];

    if (cleanPhone.length >= 9) {
       footerContents.push({
           "type": "button", "style": "primary", "height": "sm",
           "action": { "type": "uri", "label": "📞 โทร", "uri": "tel:" + cleanPhone },
           "color": "#2a5298"
       });
    }
    if (lineId !== "" && lineId !== "-") {
       footerContents.push({
             "type": "button", "style": "secondary", "height": "sm",
             "action": { "type": "uri", "label": "💬 แชท LINE", "uri": "https://line.me/ti/p/~" + lineId }
       });
    }
    footerContents.push({
         "type": "button", "style": "link", "height": "sm",
         "action": { "type": "uri", "label": "🌐 ดู/แก้ไข ทำเนียบ (Web)", "uri": webAppUrl }
    });

    return {
      "type": "bubble",
      "header": {
        "type": "box", "layout": "vertical", "backgroundColor": "#1e3c72",
        "contents": [ { "type": "text", "text": "ผลการค้นหา", "color": "#ffffff", "weight": "bold" } ]
      },
      "body": {
        "type": "box", "layout": "vertical",
        "contents": [
          { "type": "text", "text": safeName, "weight": "bold", "size": "lg", "color": "#1e3c72", "wrap": true },
          { "type": "text", "text": "ชื่อเล่น: " + safeNick, "size": "sm", "color": "#555555", "margin": "xs" },
          { "type": "separator", "margin": "md" },
          { "type": "box", "layout": "vertical", "margin": "md", "spacing": "sm", "contents": [
              { "type": "box", "layout": "baseline", "contents": [
                  { "type": "text", "text": "ตำแหน่ง", "color": "#aaaaaa", "size": "xs", "flex": 2 },
                  { "type": "text", "text": safePos, "wrap": true, "color": "#666666", "size": "xs", "flex": 4 }
              ]},
              { "type": "box", "layout": "baseline", "contents": [
                  { "type": "text", "text": "สังกัด", "color": "#aaaaaa", "size": "xs", "flex": 2 },
                  { "type": "text", "text": safeDept, "wrap": true, "color": "#666666", "size": "xs", "flex": 4 }
              ]},
              { "type": "box", "layout": "baseline", "contents": [
                  { "type": "text", "text": "เบอร์โทร", "color": "#aaaaaa", "size": "xs", "flex": 2 },
                  { "type": "text", "text": safePhone, "color": "#666666", "size": "xs", "flex": 4 }
              ]}
          ]}
        ]
      },
      "footer": { "type": "box", "layout": "vertical", "spacing": "sm", "contents": footerContents }
    };
  });

  const payload = {
    "replyToken": replyToken,
    "messages": [{ "type": "flex", "altText": "ผลการค้นหาข้อมูลสมาชิก พสพ.13", "contents": { "type": "carousel", "contents": bubbles } }]
  };
  
  const options = {
    'method': 'post', 'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    'payload': JSON.stringify(payload), 'muteHttpExceptions': true
  };

  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', options); } catch (e) { console.log(e); }
}

// --- ฟังก์ชันบันทึก Group ID อัตโนมัติ ---
function saveGroupIdToSheet(groupId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Config"); 
  if (!sheet) {
    sheet = ss.insertSheet("Config");
    sheet.appendRow(["GroupID", "GroupName"]);
  }

  const data = sheet.getDataRange().getValues();
  let exists = false;
  
  for (let i = 1; i < data.length; i++) {
    // ใช้ trim() เผื่อมีช่องว่างติดมา
    if (data[i][0] && data[i][0].toString().trim() === groupId.toString().trim()) {
      exists = true;
      break;
    }
  }

  if (!exists) {
    sheet.appendRow([groupId, "กลุ่มใหม่ (รอแก้ไขชื่อ)"]);
    console.log("บันทึกกลุ่มใหม่เรียบร้อย: " + groupId);
  }
}

// ==========================================
// ส่วนฟังก์ชัน: ระบบสุ่มรายชื่อ (Lucky Draw)
// ==========================================

function randomLuckyDraw(replyToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  
  const data = sheet.getDataRange().getDisplayValues(); 
  
  let validMembers = [];
  for (let i = 1; i < data.length; i++) {
    let name = data[i][1];
    if (name && name.toString().trim() !== "") {
      validMembers.push(data[i]);
    }
  }

  if (validMembers.length === 0) {
    replyText(replyToken, "ยังไม่มีรายชื่อสมาชิกในระบบครับ");
    return;
  }

  const randomIndex = Math.floor(Math.random() * validMembers.length);
  const winner = validMembers[randomIndex];

  const winnerData = {
    name: winner[1],
    nickname: winner[2] ? winner[2] : '-',
    position: winner[3] ? winner[3] : '-',
    dept: winner[4] ? winner[4] : 'ไม่ระบุสังกัด'
  };

  replyLuckyDrawFlex(replyToken, winnerData);
}

// ฟังก์ชันสร้างการ์ดประกาศรางวัล (Flex Message)
function replyLuckyDrawFlex(replyToken, data) {
  const payload = {
    "replyToken": replyToken,
    "messages": [{
      "type": "flex", "altText": "🎉 ประกาศรายชื่อผู้โชคดี!",
      "contents": {
        "type": "bubble", "size": "mega",
        "header": {
          "type": "box", "layout": "vertical", "backgroundColor": "#FFC107",
          "contents": [ { "type": "text", "text": "🎉 สมาชิกผู้โชคดี คือ 🎉", "color": "#ffffff", "weight": "bold", "size": "xl", "align": "center" } ]
        },
        "body": {
          "type": "box", "layout": "vertical",
          "contents": [
            { "type": "text", "text": "🎯", "size": "3xl", "align": "center", "margin": "md" },
            { "type": "text", "text": data.name, "weight": "bold", "size": "xxl", "align": "center", "color": "#1e3c72", "wrap": true, "margin": "md" },
            { "type": "text", "text": "(" + data.nickname + ")", "size": "md", "align": "center", "color": "#555555", "margin": "sm" },
            { "type": "separator", "margin": "xl" },
            {
              "type": "box", "layout": "vertical", "margin": "lg", "spacing": "sm",
              "contents": [
                {
                  "type": "box", "layout": "baseline",
                  "contents": [
                    { "type": "text", "text": "สังกัด", "color": "#aaaaaa", "size": "sm", "flex": 2 },
                    { "type": "text", "text": data.dept, "wrap": true, "color": "#666666", "size": "sm", "flex": 5 }
                  ]
                }
              ]
            }
          ]
        },
        "footer": {
          "type": "box", "layout": "vertical",
          "contents": [ { "type": "text", "text": "ยินดีด้วยครับ! 🥳", "align": "center", "color": "#aaaaaa", "size": "sm" } ]
        },
        "styles": { "header": { "backgroundColor": "#ffb300" } }
      }
    }]
  };
  
  const options = {
    'method': 'post', 'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    'payload': JSON.stringify(payload)
  };

  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', options); } catch (e) { console.log(e); }
}
