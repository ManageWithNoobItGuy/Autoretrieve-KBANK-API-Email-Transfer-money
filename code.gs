/**
 * @OnlyCurrentDoc
 *
 * สคริปต์นี้จะค้นหาอีเมลสรุปยอดจาก KBank QR API ใน Gmail,
 * ดึงข้อมูลยอดเงินรวม, และบันทึกลงใน Google Sheet ที่ใช้งานอยู่
 * พร้อมกับตั้งค่าทริกเกอร์ให้ทำงานอัตโนมัติทุกวัน
 */

// =================================================================
//                          การตั้งค่าหลัก
// =================================================================

// ชื่อร้านค้าของคุณ (ตามที่ปรากฏในหัวข้ออีเมล)
const SHOP_NAME = "ชื่อร้านค่าตามหัวข้ออีเมล";

// อีเมลผู้ส่ง
const SENDER_EMAIL = "KPLUSSHOP@kasikornbank.com";

// ชื่อชีตที่จะบันทึกข้อมูล
const SHEET_NAME = "ยอดขายรายวัน";


// =================================================================
//                      ฟังก์ชันหลักในการประมวลผล
// =================================================================

/**
 * ฟังก์ชันหลักที่จะทำงานเมื่อเรียกใช้สคริปต์
 * ทำหน้าที่ค้นหาอีเมลและบันทึกข้อมูล
 */
function extractAndLogSales() {
  try {
    // สร้างเงื่อนไขการค้นหาอีเมลที่ยังไม่ได้อ่าน จากผู้ส่งและหัวข้อที่กำหนด
    const searchQuery = `from:(${SENDER_EMAIL}) subject:("เรียน ร้านค้า - ${SHOP_NAME}") is:unread`;
    
    // ค้นหาอีเมลใน Gmail
    const threads = GmailApp.search(searchQuery);
    
    if (threads.length === 0) {
      Logger.log("ไม่พบอีเมลใหม่ที่ตรงตามเงื่อนไข");
      return;
    }

    // รับชีตที่จะบันทึกข้อมูล
    const sheet = getOrCreateSheet(SHEET_NAME);
    
    // ตรวจสอบและเพิ่มหัวตารางถ้ายังไม่มี
    addHeaderRowIfNeeded(sheet);

    // วนลูปเพื่อประมวลผลทุกอีเมลที่พบ
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        // ตรวจสอบอีกครั้งว่าข้อความยังไม่ได้ถูกประมวลผล
        if (!message.isUnread()) {
          return;
        }
        
        const emailBody = message.getPlainBody();
        const emailDate = message.getDate();
        
        // ดึงข้อมูลยอดเงินจากเนื้อหาอีเมล
        const amount = extractAmount(emailBody);

        if (amount !== null) {
          // เพิ่มแถวใหม่ใน Google Sheet พร้อมข้อมูลวันที่และยอดเงิน
          sheet.appendRow([emailDate, amount]);
          Logger.log(`บันทึกข้อมูลสำเร็จ: วันที่ ${emailDate}, ยอดเงิน ${amount}`);
        } else {
          Logger.log(`ไม่พบข้อมูลยอดเงินในอีเมล: ${message.getSubject()}`);
        }
        
        // ทำเครื่องหมายว่าอีเมลนี้ถูกอ่านแล้ว เพื่อไม่ให้ประมวลผลซ้ำ
        message.markRead();
      });
    });
  } catch (error) {
    // บันทึกข้อผิดพลาดที่เกิดขึ้น
    Logger.log(`เกิดข้อผิดพลาด: ${error.toString()}`);
    // สามารถตั้งค่าการแจ้งเตือนทางอีเมลได้หากต้องการ
    // MailApp.sendEmail("your-email@example.com", "Apps Script Error", error.toString());
  }
}


// =================================================================
//                         ฟังก์ชันเสริม
// =================================================================

/**
 * ดึงข้อมูลยอดเงินจากเนื้อหาอีเมลโดยใช้ Regular Expression
 * @param {string} body - เนื้อหาอีเมลแบบข้อความธรรมดา
 * @returns {number|null} - คืนค่าเป็นตัวเลขยอดเงิน หรือ null หากไม่พบ
 */
function extractAmount(body) {
  // รูปแบบการค้นหา: "ยอดเงินจำนวน(บาท) : 40,542.00"
  const regex = /ยอดเงินจำนวน\(บาท\)\s*:\s*([\d,]+\.\d{2})/;
  const match = body.match(regex);

  if (match && match[1]) {
    // แปลงข้อความตัวเลข (เช่น "40,542.00") ให้เป็นตัวเลขทศนิยม
    const amountString = match[1].replace(/,/g, ''); // ลบเครื่องหมายจุลภาค
    return parseFloat(amountString);
  }
  
  return null;
}

/**
 * รับชีตตามชื่อที่ระบุ หรือสร้างใหม่หากยังไม่มี
 * @param {string} name - ชื่อของชีต
 * @returns {Sheet} - อ็อบเจกต์ชีต
 */
function getOrCreateSheet(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  return sheet;
}

/**
 * เพิ่มหัวตาราง (Date, Amount) หากยังไม่มีในแถวแรก
 * @param {Sheet} sheet - อ็อบเจกต์ชีตที่ต้องการตรวจสอบ
 */
function addHeaderRowIfNeeded(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["วันที่", "ยอดเงิน (บาท)"]);
    // จัดรูปแบบหัวตาราง
    const headerRange = sheet.getRange("A1:B1");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f0f0f0");
    sheet.setColumnWidth(1, 200); // ตั้งค่าความกว้างคอลัมน์วันที่
    sheet.setColumnWidth(2, 150); // ตั้งค่าความกว้างคอลัมน์ยอดเงิน
  }
}


// =================================================================
//                      ฟังก์ชันสำหรับตั้งค่าทริกเกอร์
// =================================================================

/**
 * สร้างทริกเกอร์ (Trigger) เพื่อให้สคริปต์ทำงานอัตโนมัติทุกวัน
 * ให้เรียกใช้ฟังก์ชันนี้เพียงครั้งเดียวเพื่อตั้งค่า
 */
function createDailyTrigger() {
  // ลบทริกเกอร์เก่า (ถ้ามี) เพื่อป้องกันการทำงานซ้ำซ้อน
  deleteTriggers();
  
  // สร้างทริกเกอร์ใหม่ให้ทำงานทุกวัน ในช่วงเวลาที่กำหนด
  ScriptApp.newTrigger('extractAndLogSales')
      .timeBased()
      .everyDays(1)
      .atHour(7) // ตั้งเวลาทำงานตอน 7 โมงเช้า (สามารถเปลี่ยนได้ 0-23)
      .create();
      
  Logger.log("สร้างทริกเกอร์สำหรับทำงานทุกวันตอน 7 โมงเช้าสำเร็จแล้ว");
}

/**
 * ลบทริกเกอร์ทั้งหมดของโปรเจกต์นี้
 */
function deleteTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log("ลบทริกเกอร์เก่าทั้งหมดแล้ว");
}
