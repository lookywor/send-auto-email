var Send_email_Success = "Send_email_Success";
function sendEmails() { 
  var sheet = SpreadsheetApp.getActive().getSheetByName('sendstatus');
  var startRow = 4;  

  var numRows = 350 ;   
  var dataRange = sheet.getRange(startRow, 1, numRows, 16)
  var data = dataRange.getValues();
  // ithelpdesk.dsl@ktbcs.co.th
   var emailcc = "ithelpdesk.dsl@ktbcs.co.th";
  

// ฟังชั่นเปลี่ยนฟอร์แมตวันที่
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; 
    var dt = row[7];

      function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) 
        month = '0' + month;
    if (day.length < 2) 
        day = '0' + day;

    return [year, month, day].join('-');
}


var dateissue = dt.toLocaleString("en-EN", { timeZone: "Asia/Bangkok", hour12: false });

    var dateissuedate =  formatDate(dateissue);

    var dataissuetime = new Date(dt).toLocaleTimeString("en-EN", { timeZone: "Asia/Bangkok", hour12: false });


    var message = "";

    message = message + "<div style='white-space: pre-line'>";
    message = message + "เรียนคุณ " + row[1] + "\n\n"; 
    message = message + "จาก Case Incident : " + row[2] + " " + row[3] + " " + row[4] + "\n";
    message = message + "<b>Description : </b>" + row[6] + "\n\n";

  if (row[12] != "") {
  message = message + row[12] + "\n"; //ผู้รับเรื่อง
  message = message + row[13] + dateissuedate + " " + dataissuetime + "\n\n"; //วันที่เปิด
  }

    message = message + "<b>Resolution : </b>" +"ทีมพัฒนาระบบตรวจสอบแล้วแจ้งข้อมูลดังนี้ " + "\n";
    message = message + row[8] + "\n\n";

    message = message + "<p style='color:red;'><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;** รบกวนช่วยยืนยันกลับผ่านระบบ Service Now หรือ reply email นี้ หากปัญหาดังกล่าวได้รับการแก้ไขหรือยังไม่ได้รับการแก้ไข แจ้งกลับภายใน 5 วัน \n หากไม่แจ้งกลับระบบจะดำเนินการปิดงานอัตโนมัติ **</b></p> " + "\n" ;

    message = message + "ขอบคุณที่ใช้บริการ" + "\n";
    message = message + "-------------------------------------" + "\n";
    message = message + "IT Helpdesk DSL" + "\n";
    message = message + "รับแจ้งปัญหาด้าน IT กองทุนเงินให้กู้ยืมเพื่อการศึกษา (กยศ.)" + "\n";
    message = message + "หมายเลขโทรศัพท์ : 02-248-4006" + "\n";
    message = message + "อีเมล์ : ithelpdesk.dsl@ktbcs.co.th" + "\n";
    message = message + "</div>";


    var emailSent = row[1];  
    if (emailSent != "") {  
      var subject1 = "ปัญหาที่แจ้งได้รับการตรวจสอบแล้ว Incident :  " + row[2] + " " + row[3] + " " + row[4]  ;  
      MailApp.sendEmail({
        to: emailAddress, 
        cc: emailcc,
        subject: subject1,
        htmlBody: message,

      });
      sheet.getRange(startRow + i, 16).setValue(Send_email_Success);
      SpreadsheetApp.flush();
    }
  }
}
