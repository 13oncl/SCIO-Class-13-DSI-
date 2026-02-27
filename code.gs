function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('ทำเนียบรุ่น พสพ.13')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Sheet1"); // **ต้องมั่นใจว่าชื่อ Sheet คือ "Sheet1" เป๊ะๆ**
  
  // เช็คว่ามีข้อมูลหรือไม่ (ป้องกัน Error กรณีมีแค่หัวตาราง)
  if (ws.getLastRow() < 2) return [];

  // ดึงข้อมูลถึงคอลัมน์ I (9)
  const data = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues(); 
  
  // กรองแถวว่าง และแปลงวันที่เป็น Text เพื่อส่งไปหน้าเว็บได้ชัวร์ๆ
  const cleanData = data.filter(r => r[0] && r[0] !== "").map(row => {
    // ถ้าช่อง I (index 8) เป็นวันที่ ให้แปลงเป็น ISO String
    if (row[8] && Object.prototype.toString.call(row[8]) === '[object Date]') {
      row[8] = row[8].toISOString();
    }
    return row;
  });

  return cleanData;
}

function saveData(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Sheet1");
  
  // จัดการวันเกิด
  let bd = form.birthdate;
  if (bd) {
    bd = new Date(bd); // แปลงกลับเป็น Date Object เพื่อลง Sheet
  } else {
    bd = "";
  }

  if (form.recId) {
    // แก้ไข
    const data = ws.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == form.recId) {
        ws.getRange(i + 1, 2, 1, 8).setValues([[
          form.name,
          form.nickname,
          form.position,
          form.dept,
          "'" + form.phone,
          form.lineId,
          form.email,
          bd // Column I
        ]]);
        return "อัปเดตข้อมูลสำเร็จ";
      }
    }
  } else {
    // เพิ่มใหม่
    const newId = new Date().getTime();
    ws.appendRow([
      newId,
      form.name,
      form.nickname,
      form.position,
      form.dept,
      "'" + form.phone,
      form.lineId,
      form.email,
      bd // Column I
    ]);
    return "บันทึกข้อมูลใหม่สำเร็จ";
  }
}

function deleteData(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Sheet1");
  const data = ws.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      ws.deleteRow(i + 1);
      return "ลบข้อมูลเรียบร้อย";
    }
  }
  return "ไม่พบข้อมูล";
}
