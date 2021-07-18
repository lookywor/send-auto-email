var Send_email_Success = "Send_email_Success";
function sendEmails() { 
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 4;  

  var numRows = 200 ;   
  var dataRange = sheet.getRange(startRow, 1, numRows, 13)
  var data = dataRange.getValues();
  var emailcc = "ithelpdesk.dsl@ktbcs.co.th";
 
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; 
    var message = "เรี่ยนคุณ " + row[1] + "," + "\n\n";  

    message = message + "จาก Case Incident : " + row[2] + " " + row[3] + " " + row[4] + " " + row[5]+ "\n\n";
   // message = message + "Description : " + row[6] + " " + row[3] + " " + row[4] + " " + row[5]+ "\n\n";

    message = message + "ผู้รับเรื่อง : " + row[1] + "\n";
    message = message + "วันที่เปิด : " + row[7] + "\n\n";

    message = message + "Resolution : ทีมพัฒนาระบบตรวจสอบแล้วแจ้งข้อมูลดังนี้ " + "\n";
    message = message + row[8] + "\n\n";
    
    message = message + "    ** รบกวนช่วยยืนยันกลับผ่านระบบ Service Now หรือ reply email นี้ หากปัญหาดังกล่าวได้รับการแก้ไขหรือยังไม่ได้รับการแก้ไข แจ้งกลับภายใน 5 วัน หากไม่แจ้งกลับระบบจะดำเนินการปิดงานอัตโนมัติ ** " + "\n\n" ;

    message = message + "ขอบคุณที่ใช้บริการ" + "\n";
    message = message + "-------------------------------------" + "\n";
    message = message + "IT Helpdesk DSL" + "\n";
    message = message + "รับแจ้งปัญหาด้าน IT กองทุนเงินให้กู้ยืมเพื่อการศึกษา (กยศ.)" + "\n";
    message = message + "หมายเลขโทรศัพท์ : 02-248-4006" + "\n";
    message = message + "อีเมล์ : ithelpdesk.dsl@ktbcs.co.th" + "\n";
    



    var emailSent = row[1];  
    if (emailSent != "") {  
      var subject = "ปัญหาที่แจ้งได้รับการตรวจสอบแล้ว Incident :  " + row[2] + " " + row[3] + " " + row[4] + " " + row[5];  
      MailApp.sendEmail({
        to: emailAddress, 
        cc: emailcc,
        subject: subject,
        body: message,
        
      });
      sheet.getRange(startRow + i, 13).setValue(Send_email_Success);
      SpreadsheetApp.flush();
    }
  }
}
