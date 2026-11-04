const REPORT_TITLE = "RRAS Examination Center";


function sendMidDayAttendanceReport() {


 const ss = SpreadsheetApp.openById(LOGIN_SPREADSHEET_ID);


 const attendance = ss.getSheetByName("ATTENDANCE_LOG");
 const examSheet = ss.getSheetByName("Exam_Details");
 const settings = ss.getSheetByName("SETTINGS");
 const leaveSheet = ss.getSheetByName("Leave Application");
 const staffSheet = ss.getSheetByName("STAFF_MASTER");


 const ownerEmail = settings.getRange("D2").getValue();


 const today = Utilities.formatDate(
   new Date(),
   Session.getScriptTimeZone(),
   "dd-MM-yyyy"
 );


 const attendanceData = attendance.getDataRange().getValues();
 const examData = examSheet.getDataRange().getValues();
 const leaveData = leaveSheet.getDataRange().getValues();
 const staffData = staffSheet.getDataRange().getValues();


 let examName = "";
 let totalShift = "";


 for (let i = 1; i < examData.length; i++) {


   const d = Utilities.formatDate(
     new Date(examData[i][1]),
     Session.getScriptTimeZone(),
     "dd-MM-yyyy"
   );


   if (d == today) {


     examName = examData[i][2];
     totalShift = examData[i][3];
     break;


   }


 }


 const doc = DocumentApp.create(
   "RRAS_MidDay_Report_" + today
 );


 const body = doc.getBody();


body.appendParagraph("RADHA RAMAN ASSESSMENT SERVICES PVT. LTD.")
.setFontSize(16)
.setBold(true);


body.appendParagraph("RRAS EXAMINATION CENTER")
.setFontSize(13)
.setBold(true);


body.appendParagraph("MID-DAY ATTENDANCE REPORT")
.setBold(true);


body.appendHorizontalRule();


body.appendParagraph(
"Exam Date      : " + today +
"\nExam Name      : " + examName +
"\nTotal Shift    : " + totalShift +
"\nReport Type    : Mid-Day Attendance"
);


body.appendHorizontalRule();
let permanent = 0;
let temporary = 0;


const table = body.appendTable();
table.setBorderWidth(1);


table.appendTableRow();


const header = table.getRow(0);


header.appendTableCell("Sl");
header.appendTableCell("Staff Name");
header.appendTableCell("ID");
header.appendTableCell("Role");
header.appendTableCell("Type");
header.appendTableCell("Check-In");
header.appendTableCell("Dist.");
header.appendTableCell("Status");


// Make header bold
for (let i = 0; i < header.getNumCells(); i++) {
header.getCell(i)
     .editAsText()
     .setBold(true)
     .setFontSize(8);
}


 let sl = 1;


 let totalPresent = 0;
for (let i = 1; i < attendanceData.length; i++) {


 const reportDate = Utilities.formatDate(
   new Date(attendanceData[i][0]),
   Session.getScriptTimeZone(),
   "dd-MM-yyyy"
 );


 if (reportDate != today) continue;


   totalPresent++;
   if(attendanceData[i][7]=="Permanent Staff")
permanent++;


if(attendanceData[i][7]=="Temporary Staff")
temporary++;


   const row = table.appendTableRow();


   row.appendTableCell(String(sl++));


let staffName = attendanceData[i][1].toString();


if(staffName.length > 18){
 staffName = staffName.substring(0,18) + "..";
}


row.appendTableCell(staffName);


   row.appendTableCell(String(attendanceData[i][9]));


let role = attendanceData[i][8];


if(role=="Center Head")
 role="Center Hd";


if(role=="Invigilator")
 role="Invigil.";


row.appendTableCell(role);


row.appendTableCell(
 attendanceData[i][7] == "Permanent Staff"
   ? "Perm."
   : "Temp."
);


   row.appendTableCell(
     attendanceData[i][10]
       ? Utilities.formatDate(
           new Date(attendanceData[i][10]),
           Session.getScriptTimeZone(),
           "hh:mm a"
         )
       : "-"
   );


   row.appendTableCell(String(attendanceData[i][17]) + " m");


let status = attendanceData[i][20];


if(status=="Checked-In"){
 status = "Present";
}


if(status=="Checked-Out"){
 status = "Done";
}


row.appendTableCell(status);


 }



//========================================
// ACTIVE STAFF ABSENT LIST
//========================================

let totalActiveStaff = 0;
let absentStaff = [];

for(let i=1;i<staffData.length;i++){

  if(String(staffData[i][10]).toUpperCase() == "YES"){

    totalActiveStaff++;

    let found = false;

    for(let j=1;j<attendanceData.length;j++){

      const reportDate = Utilities.formatDate(
        new Date(attendanceData[j][0]),
        Session.getScriptTimeZone(),
        "dd-MM-yyyy"
      );

      if(
        reportDate == today &&
        String(attendanceData[j][9]).trim() ==
        String(staffData[i][9]).trim()
      ){
        found = true;
        break;
      }

    }

    if(!found){

      absentStaff.push({
        name:staffData[i][1],
        id:staffData[i][9],
        role:staffData[i][8]
      });

    }

  }

}
body.appendParagraph("");
body.appendParagraph("");

body.appendHorizontalRule();

body.appendParagraph("ACTIVE STAFF ABSENT LIST")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const absentTable = body.appendTable();

absentTable.appendTableRow();

absentTable.getRow(0).appendTableCell("Sl");
absentTable.getRow(0).appendTableCell("Name");
absentTable.getRow(0).appendTableCell("Resource ID");
absentTable.getRow(0).appendTableCell("Role");

if(absentStaff.length==0){

  const r = absentTable.appendTableRow();

  r.appendTableCell("-");
  r.appendTableCell("No Active Staff Absent");
  r.appendTableCell("-");
  r.appendTableCell("-");

}else{

  for(let i=0;i<absentStaff.length;i++){

    const r = absentTable.appendTableRow();

    r.appendTableCell(String(i+1));
    r.appendTableCell(absentStaff[i].name);
    r.appendTableCell(String(absentStaff[i].id));
    r.appendTableCell(absentStaff[i].role);

  }

}

body.appendParagraph("");
body.appendHorizontalRule();


body.appendParagraph("PERMANENT STAFF ON LEAVE")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);


const leaveTable = body.appendTable();


leaveTable.appendTableRow();


leaveTable.getRow(0).appendTableCell("Name");
leaveTable.getRow(0).appendTableCell("Resource ID");
leaveTable.getRow(0).appendTableCell("From");
leaveTable.getRow(0).appendTableCell("To");
leaveTable.getRow(0).appendTableCell("Reason");


let leaveFound = false;
let leaveCount = 0;


for (let i = 1; i < leaveData.length; i++) {


 const fromDate =
   Utilities.formatDate(
     new Date(leaveData[i][3]),
     Session.getScriptTimeZone(),
     "dd-MM-yyyy"
   );


 const toDate =
   Utilities.formatDate(
     new Date(leaveData[i][4]),
     Session.getScriptTimeZone(),
     "dd-MM-yyyy"
   );


if (today >= fromDate && today <= toDate) {


 leaveFound = true;
 leaveCount++;


   const r = leaveTable.appendTableRow();


   r.appendTableCell(leaveData[i][1]);


   r.appendTableCell(String(leaveData[i][2]));


   r.appendTableCell(fromDate);


   r.appendTableCell(toDate);


   r.appendTableCell(String(leaveData[i][5]));


 }


}


if (!leaveFound) {


 const r = leaveTable.appendTableRow();


 r.appendTableCell("No Staff On Leave");


 r.appendTableCell("-");


 r.appendTableCell("-");


 r.appendTableCell("-");


 r.appendTableCell("-");


}
// Reduce font size of attendance table
for (let r = 0; r < table.getNumRows(); r++) {
 const tr = table.getRow(r);


 for (let c = 0; c < tr.getNumCells(); c++) {
   tr.getCell(c)
     .editAsText()
     .setFontSize(8);
 }
}
body.appendParagraph("");


body.appendHorizontalRule();


body.appendParagraph("DEVICE SHARING ANALYSIS")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);


const dupTable = body.appendTable();


dupTable.appendTableRow();


dupTable.getRow(0).appendTableCell("Device Group");
dupTable.getRow(0).appendTableCell("Login Type");
dupTable.getRow(0).appendTableCell("Name");
dupTable.getRow(0).appendTableCell("Resource ID");
dupTable.getRow(0).appendTableCell("Mobile");
dupTable.getRow(0).appendTableCell("Check-In");


let deviceMap = {};
let duplicateFound = false;


// Group today's records
for (let i = 1; i < attendanceData.length; i++) {


 const reportDate = Utilities.formatDate(
   new Date(attendanceData[i][0]),
   Session.getScriptTimeZone(),
   "dd-MM-yyyy"
 );


 if(reportDate != today) continue;


 const fp = String(attendanceData[i][23]);


 if(!deviceMap[fp]){
   deviceMap[fp] = [];
 }


 deviceMap[fp].push(attendanceData[i]);


}


let deviceNo = 1;


for(const fp in deviceMap){


 const users = deviceMap[fp];


 if(users.length <= 1) continue;


 duplicateFound = true;


 users.sort(function(a,b){
   return new Date(a[10]) - new Date(b[10]);
 });


 for(let j=0;j<users.length;j++){


   const r = users[j];


   const row = dupTable.appendTableRow();


   row.appendTableCell("Device #" + deviceNo);


   row.appendTableCell(
     j==0 ? "Original" : "Duplicate"
   );


   row.appendTableCell(r[1]);


   row.appendTableCell(String(r[9]));


   row.appendTableCell(String(r[2]));


   row.appendTableCell(
     Utilities.formatDate(
       new Date(r[10]),
       Session.getScriptTimeZone(),
       "hh:mm a"
     )
   );


 }


 deviceNo++;


}


if(!duplicateFound){


 const row = dupTable.appendTableRow();


 row.appendTableCell("-");


 row.appendTableCell("-");


 row.appendTableCell("No Duplicate Mobile Login Found");


 row.appendTableCell("-");


 row.appendTableCell("-");


 row.appendTableCell("-");


}
body.appendParagraph("");
body.appendHorizontalRule();


body.appendParagraph("SUMMARY")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);


body.appendParagraph(
"Total Active Staff : " + totalActiveStaff
).setBold(true);

body.appendParagraph(
"Total Present Staff : " + totalPresent
);

body.appendParagraph(
"Total Active Staff Absent : " + absentStaff.length
);

body.appendParagraph(
"Permanent Staff : " + permanent
);

body.appendParagraph(
"Temporary Staff : " + temporary
);

body.appendParagraph(
"Duplicate Device Groups : " + (deviceNo - 1)
);

body.appendParagraph(
"Staff On Leave : " + leaveCount
);


body.appendParagraph("");


body.appendHorizontalRule();


body.appendHorizontalRule();


body.appendParagraph(
"CONFIDENTIAL REPORT"
).setBold(true);


body.appendParagraph(
"Generated Automatically by RRAS Attendance Management System"
);


body.appendParagraph(
"Radha Raman Assessment Services Pvt. Ltd."
);


body.appendParagraph(
"Generated On : " +
Utilities.formatDate(
new Date(),
Session.getScriptTimeZone(),
"dd-MM-yyyy hh:mm a"
)
);


 doc.saveAndClose();


 const pdf = DriveApp
   .getFileById(doc.getId())
   .getAs(MimeType.PDF)
   .setName("RRAS_MidDay_Report_" + today + ".pdf");


const html =


"<h2>RRAS Examination Center</h2>"+


"<b>MID-DAY ATTENDANCE REPORT</b><br><br>"+


"Date : "+today+"<br>"+


"Exam : "+examName+"<br>"+


"Total Shift : "+totalShift+"<br><br>"+


"<table border='1' cellpadding='6' cellspacing='0'>"+

"<tr><td>Total Active Staff</td><td>"+totalActiveStaff+"</td></tr>"+

"<tr><td>Total Present Staff</td><td>"+totalPresent+"</td></tr>"+

"<tr><td>Total Active Staff Absent</td><td>"+absentStaff.length+"</td></tr>"+

"<tr><td>Permanent Staff</td><td>"+permanent+"</td></tr>"+

"<tr><td>Temporary Staff</td><td>"+temporary+"</td></tr>"+

"</table><br>";


"This report contains the current attendance status, leave records, and duplicate mobile login analysis for today's examination. Please find the attached PDF report."


 MailApp.sendEmail({
   to: ownerEmail,
   subject: "RRAS Mid-Day Attendance Report - " + today,
   htmlBody: html,
   attachments: [pdf]
 });


 DriveApp.getFileById(doc.getId()).setTrashed(true);


}

