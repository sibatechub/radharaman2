

function sendFinalReport() {

  const ss = SpreadsheetApp.openById(LOGIN_SPREADSHEET_ID);

  const attendanceSheet = ss.getSheetByName("ATTENDANCE_LOG");
  const examSheet = ss.getSheetByName("Exam_Details");
  const settingsSheet = ss.getSheetByName("SETTINGS");
  const paymentSheet = ss.getSheetByName("Payment Details");
  const expenseSheet = ss.getSheetByName("Daily_Expenses");
  const leaveSheet = ss.getSheetByName("Leave Application");
  const staffSheet = ss.getSheetByName("STAFF_MASTER");

  const attendanceData = attendanceSheet.getDataRange().getValues();
  const examData = examSheet.getDataRange().getValues();
  const paymentData = paymentSheet.getDataRange().getValues();
  const expenseData = expenseSheet.getDataRange().getValues();
  const leaveData = leaveSheet.getDataRange().getValues();
  const staffData = staffSheet.getDataRange().getValues();

  const ownerEmail = settingsSheet.getRange("D2").getValue();

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd-MM-yyyy"
  );

  let examName = "";
  let totalShift = "";

  for (let i = 1; i < examData.length; i++) {

    const examDate = Utilities.formatDate(
      new Date(examData[i][1]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
    );

    if (examDate == today) {

      examName = examData[i][2];
      totalShift = examData[i][3];
      break;

    }

  }

  const doc = DocumentApp.create(
    "RRAS_Final_Report_" + today
  );

  const body = doc.getBody();

  body.appendParagraph(
    "RADHA RAMAN ASSESSMENT SERVICES PVT. LTD."
  )
  .setFontSize(16)
  .setBold(true);

  body.appendParagraph(
    "RRAS EXAMINATION CENTER"
  )
  .setFontSize(13)
  .setBold(true);

  body.appendParagraph(
    "FINAL ATTENDANCE & PAYMENT REPORT"
  )
  .setFontSize(12)
  .setBold(true);

  body.appendHorizontalRule();

  body.appendParagraph(
      "Exam Date      : " + today +
      "\nExam Name      : " + examName +
      "\nTotal Shift    : " + totalShift +
      "\nReport Type    : Final Report"
  );

  body.appendHorizontalRule();

  /***************************************************
      PART-2 STARTS FROM HERE
      Attendance Table
  ***************************************************/
  //=============================
// ATTENDANCE TABLE
//=============================

let permanent = 0;
let temporary = 0;
let totalPresent = 0;

body.appendParagraph("FINAL ATTENDANCE DETAILS")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const attendanceTable = body.appendTable();
attendanceTable.setBorderWidth(1);

attendanceTable.appendTableRow();

const head = attendanceTable.getRow(0);

head.appendTableCell("Sl");
head.appendTableCell("Staff Name");
head.appendTableCell("ID");
head.appendTableCell("Role");
head.appendTableCell("Type");
head.appendTableCell("Check-In");
head.appendTableCell("Check-Out");
head.appendTableCell("Duration");
head.appendTableCell("Shift");
head.appendTableCell("Status");

// Header Style
for (let i = 0; i < head.getNumCells(); i++) {

  head.getCell(i)
      .editAsText()
      .setBold(true)
      .setFontSize(8);

}

let sl = 1;

//=============================
// Attendance Data
//=============================

for (let i = 1; i < attendanceData.length; i++) {

  const reportDate = Utilities.formatDate(
      new Date(attendanceData[i][0]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
  );

  if (reportDate != today) continue;

  totalPresent++;

  if (attendanceData[i][7] == "Permanent Staff")
      permanent++;

  if (attendanceData[i][7] == "Temporary Staff")
      temporary++;

  const row = attendanceTable.appendTableRow();

  // Sl
  row.appendTableCell(String(sl++));

  // Staff Name
  let staff = String(attendanceData[i][1]);

  if (staff.length > 20)
      staff = staff.substring(0,20) + "..";

  row.appendTableCell(staff);

  // Resource ID
  row.appendTableCell(String(attendanceData[i][9]));

  // Role
  let role = attendanceData[i][8];

  if(role=="Center Head")
      role="Center Hd";

  if(role=="Invigilator")
      role="Invigil.";

  row.appendTableCell(role);

  // Employment Type
  row.appendTableCell(
      attendanceData[i][7]=="Permanent Staff"
      ? "Perm."
      : "Temp."
  );

  // Check In
  row.appendTableCell(
      attendanceData[i][10]
      ? Utilities.formatDate(
            new Date(attendanceData[i][10]),
            Session.getScriptTimeZone(),
            "hh:mm a"
        )
      : "-"
  );

  // Check Out
  row.appendTableCell(
      attendanceData[i][11]
      ? Utilities.formatDate(
            new Date(attendanceData[i][11]),
            Session.getScriptTimeZone(),
            "hh:mm a"
        )
      : "-"
  );

  // Duration
  row.appendTableCell(
      attendanceData[i][12]
      ? String(attendanceData[i][12])
      : "-"
  );

  // Shift Count
  row.appendTableCell(
      attendanceData[i][22]
      ? String(attendanceData[i][22])
      : "0"
  );

  // Status

  let status = attendanceData[i][20];

  if(status=="Checked-Out")
      status="Completed";

  if(status=="Checked-In")
      status="Present";

  row.appendTableCell(status);

}

// Reduce font size

for (let r = 0; r < attendanceTable.getNumRows(); r++) {

  const tr = attendanceTable.getRow(r);

  for (let c = 0; c < tr.getNumCells(); c++) {

      tr.getCell(c)
        .editAsText()
        .setFontSize(8);

  }

}

body.appendParagraph("");
body.appendHorizontalRule();
body.appendParagraph("");
body.appendHorizontalRule();

//========================================
// ACTIVE STAFF ABSENT CALCULATION
//========================================

let totalActiveStaff = 0;
let absentStaff = [];

for(let i=1;i<staffData.length;i++){

  if(String(staffData[i][10]).toUpperCase()=="YES"){

    totalActiveStaff++;

    let found = false;

    for(let j=1;j<attendanceData.length;j++){

      const reportDate = Utilities.formatDate(
        new Date(attendanceData[j][0]),
        Session.getScriptTimeZone(),
        "dd-MM-yyyy"
      );

      if(
        reportDate==today &&
        String(attendanceData[j][9]).trim()==
        String(staffData[i][9]).trim()
      ){
        found=true;
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

  const r=absentTable.appendTableRow();

  r.appendTableCell("-");
  r.appendTableCell("No Active Staff Absent");
  r.appendTableCell("-");
  r.appendTableCell("-");

}else{

  for(let i=0;i<absentStaff.length;i++){

    const r=absentTable.appendTableRow();

    r.appendTableCell(String(i+1));
    r.appendTableCell(absentStaff[i].name);
    r.appendTableCell(String(absentStaff[i].id));
    r.appendTableCell(absentStaff[i].role);

  }

}

// Reduce Font Size

for(let r=0;r<absentTable.getNumRows();r++){

  const tr=absentTable.getRow(r);

  for(let c=0;c<tr.getNumCells();c++){

    tr.getCell(c)
      .editAsText()
      .setFontSize(8);

  }

}
//==================================================
// PAYMENT CALCULATION
//==================================================

body.appendParagraph("PAYMENT CALCULATION")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const paymentTable = body.appendTable();
paymentTable.setBorderWidth(1);

paymentTable.appendTableRow();

const pHead = paymentTable.getRow(0);

pHead.appendTableCell("Sl");
pHead.appendTableCell("Staff Name");
pHead.appendTableCell("Role");
pHead.appendTableCell("Shift");
pHead.appendTableCell("Rate");
pHead.appendTableCell("Amount");

// Header Style
for (let i = 0; i < pHead.getNumCells(); i++) {

  pHead.getCell(i)
       .editAsText()
       .setBold(true)
       .setFontSize(8);

}

let paymentSl = 1;

let grandTotalPayment = 0;

let totalShiftWorked = 0;

//------------------------------------
// Read today's attendance
//------------------------------------

for (let i = 1; i < attendanceData.length; i++) {

  const reportDate = Utilities.formatDate(
      new Date(attendanceData[i][0]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
  );

  if(reportDate != today) continue;

  const role = attendanceData[i][8];
  // Only calculate payment for Invigilators
if (role != "Invigilator") {
  continue;
}

  const shiftWorked = Number(attendanceData[i][22]) || 0;

  let rate = 0;

  //------------------------------------
  // Find Rate From Payment Details
  //------------------------------------

for (let j = 1; j < paymentData.length; j++) {

    if (
        String(paymentData[j][1]).trim() == String(role).trim() &&
        Number(paymentData[j][2]) == shiftWorked
    ) {

        rate = Number(paymentData[j][3]);
        break;

    }

}

  const amount = rate;

  grandTotalPayment += amount;

  totalShiftWorked += shiftWorked;

  const row = paymentTable.appendTableRow();

  row.appendTableCell(String(paymentSl++));

  row.appendTableCell(String(attendanceData[i][1]));

  row.appendTableCell(role);

  row.appendTableCell(String(shiftWorked));

  row.appendTableCell("₹ " + rate);

  row.appendTableCell("₹ " + amount);

}

// Reduce Font Size

for(let r=0;r<paymentTable.getNumRows();r++){

    const tr = paymentTable.getRow(r);

    for(let c=0;c<tr.getNumCells();c++){

        tr.getCell(c)
          .editAsText()
          .setFontSize(8);

    }

}

body.appendParagraph("");

body.appendParagraph(
"Grand Total Payment : ₹ " + grandTotalPayment
).setBold(true);

body.appendParagraph("");

body.appendHorizontalRule();
//==================================================
// DAILY EXPENSES
//==================================================

body.appendParagraph("DAILY EXPENSES")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const expenseTable = body.appendTable();
expenseTable.setBorderWidth(1);

expenseTable.appendTableRow();

const eHead = expenseTable.getRow(0);

eHead.appendTableCell("Expense");
eHead.appendTableCell("Quantity");
eHead.appendTableCell("Rate");
eHead.appendTableCell("Total");

// Header Style
for (let i = 0; i < eHead.getNumCells(); i++) {

  eHead.getCell(i)
       .editAsText()
       .setBold(true)
       .setFontSize(9);

}

let totalExpenses = 0;
let expenseFound = false;

for (let i = 1; i < expenseData.length; i++) {

  const reportDate = Utilities.formatDate(
      new Date(expenseData[i][1]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
  );

  if(reportDate != today) continue;

  expenseFound = true;

  totalExpenses = Number(expenseData[i][15]) || 0;

  //------------------------------------------------

let row;

// Tiffin
row = expenseTable.appendTableRow();
row.appendTableCell("Tiffin");
row.appendTableCell(String(expenseData[i][3]));
row.appendTableCell("₹ " + expenseData[i][4]);
row.appendTableCell("₹ " + expenseData[i][5]);

// Lunch
row = expenseTable.appendTableRow();
row.appendTableCell("Lunch");
row.appendTableCell(String(expenseData[i][6]));
row.appendTableCell("₹ " + expenseData[i][7]);
row.appendTableCell("₹ " + expenseData[i][8]);

// Snacks
row = expenseTable.appendTableRow();
row.appendTableCell("Snacks");
row.appendTableCell("-");
row.appendTableCell("-");
row.appendTableCell("₹ " + expenseData[i][9]);

// Fuel
row = expenseTable.appendTableRow();
row.appendTableCell("Fuel");
row.appendTableCell(expenseData[i][10] + " L");
row.appendTableCell("-");
row.appendTableCell("₹ " + expenseData[i][11]);

// Sanitary Items
row = expenseTable.appendTableRow();
row.appendTableCell("Sanitary Items");
row.appendTableCell("-");
row.appendTableCell("-");
row.appendTableCell("₹ " + expenseData[i][12]);

// Other
row = expenseTable.appendTableRow();
row.appendTableCell("Other (" + expenseData[i][13] + ")");
row.appendTableCell("-");
row.appendTableCell("-");
row.appendTableCell("₹ " + expenseData[i][14]);

  break;

}

// No Expense Found

if(!expenseFound){

  let row = expenseTable.appendTableRow();

  row.appendTableCell("No Expense Record Found");

  row.appendTableCell("-");

}

// Reduce Font Size

for(let r=0;r<expenseTable.getNumRows();r++){

  const tr = expenseTable.getRow(r);

  for(let c=0;c<tr.getNumCells();c++){

      tr.getCell(c)
        .editAsText()
        .setFontSize(8);

  }

}

body.appendParagraph("");

body.appendParagraph(
"Total Expenses : ₹ " + totalExpenses
).setBold(true);

body.appendParagraph("");

body.appendHorizontalRule();
//====================================================
// LEAVE DETAILS
//====================================================

body.appendParagraph("");
body.appendHorizontalRule();

body.appendParagraph("STAFF ON LEAVE")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const leaveTable = body.appendTable();
leaveTable.setBorderWidth(1);

leaveTable.appendTableRow();

const leaveHead = leaveTable.getRow(0);

leaveHead.appendTableCell("Name");
leaveHead.appendTableCell("Resource ID");
leaveHead.appendTableCell("From");
leaveHead.appendTableCell("To");
leaveHead.appendTableCell("Reason");

// Header Style
for (let i = 0; i < leaveHead.getNumCells(); i++) {

  leaveHead.getCell(i)
      .editAsText()
      .setBold(true)
      .setFontSize(8);

}

let leaveCount = 0;

for (let i = 1; i < leaveData.length; i++) {

  const fromDate = Utilities.formatDate(
      new Date(leaveData[i][3]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
  );

  const toDate = Utilities.formatDate(
      new Date(leaveData[i][4]),
      Session.getScriptTimeZone(),
      "dd-MM-yyyy"
  );

  if (today >= fromDate && today <= toDate) {

      leaveCount++;

      const row = leaveTable.appendTableRow();

      row.appendTableCell(String(leaveData[i][1]));
      row.appendTableCell(String(leaveData[i][2]));
      row.appendTableCell(fromDate);
      row.appendTableCell(toDate);
      row.appendTableCell(String(leaveData[i][5]));

  }

}

if (leaveCount == 0) {

  const row = leaveTable.appendTableRow();

  row.appendTableCell("No Staff On Leave");
  row.appendTableCell("-");
  row.appendTableCell("-");
  row.appendTableCell("-");
  row.appendTableCell("-");

}

// Reduce font size

for (let r = 0; r < leaveTable.getNumRows(); r++) {

  const tr = leaveTable.getRow(r);

  for (let c = 0; c < tr.getNumCells(); c++) {

      tr.getCell(c)
        .editAsText()
        .setFontSize(8);

  }

}
//====================================================
// CHECK-IN DUPLICATE DEVICE ANALYSIS
//====================================================

body.appendParagraph("");
body.appendHorizontalRule();

body.appendParagraph("DEVICE SHARING ANALYSIS (CHECK-IN)")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const dupTable = body.appendTable();
dupTable.setBorderWidth(1);

dupTable.appendTableRow();

const dupHead = dupTable.getRow(0);

dupHead.appendTableCell("Device");
dupHead.appendTableCell("Login");
dupHead.appendTableCell("Staff");
dupHead.appendTableCell("ID");
dupHead.appendTableCell("Mobile");
dupHead.appendTableCell("Check-In");

// Header Style

for (let i = 0; i < dupHead.getNumCells(); i++) {

  dupHead.getCell(i)
      .editAsText()
      .setBold(true)
      .setFontSize(8);

}

let deviceMap = {};
let duplicateGroups = 0;

// Group today's attendance

for (let i = 1; i < attendanceData.length; i++) {

    const reportDate = Utilities.formatDate(
        new Date(attendanceData[i][0]),
        Session.getScriptTimeZone(),
        "dd-MM-yyyy"
    );

    if (reportDate != today) continue;

    const fingerprint = String(attendanceData[i][23]);

    if (!deviceMap[fingerprint]) {

        deviceMap[fingerprint] = [];

    }

    deviceMap[fingerprint].push(attendanceData[i]);

}

let deviceNo = 1;

for (const fp in deviceMap) {

    const users = deviceMap[fp];

    if (users.length <= 1) continue;

    duplicateGroups++;

    users.sort(function(a,b){

        return new Date(a[10]) - new Date(b[10]);

    });

    for (let j = 0; j < users.length; j++) {

        const r = users[j];

        const row = dupTable.appendTableRow();

        row.appendTableCell("Device #" + deviceNo);

        row.appendTableCell(
            j==0 ? "Original" : "Duplicate"
        );

        row.appendTableCell(String(r[1]));

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

if (duplicateGroups == 0) {

    const row = dupTable.appendTableRow();

    row.appendTableCell("-");
    row.appendTableCell("-");
    row.appendTableCell("No Duplicate Device Found");
    row.appendTableCell("-");
    row.appendTableCell("-");
    row.appendTableCell("-");

}

// Reduce font size

for (let r = 0; r < dupTable.getNumRows(); r++) {

    const tr = dupTable.getRow(r);

    for (let c = 0; c < tr.getNumCells(); c++) {

        tr.getCell(c)
          .editAsText()
          .setFontSize(8);

    }

}
//====================================================
// CHECK-OUT DUPLICATE DEVICE ANALYSIS
//====================================================

body.appendParagraph("");
body.appendHorizontalRule();

body.appendParagraph("DEVICE SHARING ANALYSIS (CHECK-OUT)")
.setHeading(DocumentApp.ParagraphHeading.HEADING2);

const checkoutTable = body.appendTable();
checkoutTable.setBorderWidth(1);

checkoutTable.appendTableRow();

const checkoutHead = checkoutTable.getRow(0);

checkoutHead.appendTableCell("Device");
checkoutHead.appendTableCell("Login");
checkoutHead.appendTableCell("Staff");
checkoutHead.appendTableCell("ID");
checkoutHead.appendTableCell("Mobile");
checkoutHead.appendTableCell("Check-Out");

// Header Style

for (let i = 0; i < checkoutHead.getNumCells(); i++) {

  checkoutHead.getCell(i)
      .editAsText()
      .setBold(true)
      .setFontSize(8);

}

let checkoutDeviceMap = {};
let checkoutDuplicateGroups = 0;

// Group today's Check-Out devices

for (let i = 1; i < attendanceData.length; i++) {

    const reportDate = Utilities.formatDate(
        new Date(attendanceData[i][0]),
        Session.getScriptTimeZone(),
        "dd-MM-yyyy"
    );

    if (reportDate != today) continue;

    const fingerprint = String(attendanceData[i][24]);

    if (!fingerprint) continue;

    if (!checkoutDeviceMap[fingerprint]) {

        checkoutDeviceMap[fingerprint] = [];

    }

    checkoutDeviceMap[fingerprint].push(attendanceData[i]);

}

let checkoutDeviceNo = 1;

for (const fp in checkoutDeviceMap) {

    const users = checkoutDeviceMap[fp];

    if (users.length <= 1) continue;

    checkoutDuplicateGroups++;

    users.sort(function(a,b){

        return new Date(a[11]) - new Date(b[11]);

    });

    for (let j = 0; j < users.length; j++) {

        const r = users[j];

        const row = checkoutTable.appendTableRow();

        row.appendTableCell("Device #" + checkoutDeviceNo);

        row.appendTableCell(
            j==0 ? "Original" : "Duplicate"
        );

        row.appendTableCell(String(r[1]));

        row.appendTableCell(String(r[9]));

        row.appendTableCell(String(r[2]));

        row.appendTableCell(

            r[11]

            ?

            Utilities.formatDate(

                new Date(r[11]),

                Session.getScriptTimeZone(),

                "hh:mm a"

            )

            :

            "-"

        );

    }

    checkoutDeviceNo++;

}

if (checkoutDuplicateGroups == 0) {

    const row = checkoutTable.appendTableRow();

    row.appendTableCell("-");
    row.appendTableCell("-");
    row.appendTableCell("No Duplicate Check-Out Device Found");
    row.appendTableCell("-");
    row.appendTableCell("-");
    row.appendTableCell("-");

}

// Reduce font size

for (let r = 0; r < checkoutTable.getNumRows(); r++) {

    const tr = checkoutTable.getRow(r);

    for (let c = 0; c < tr.getNumCells(); c++) {

        tr.getCell(c)
          .editAsText()
          .setFontSize(8);

    }

}
//====================================================
// FINAL SUMMARY
//====================================================

body.appendParagraph("");
body.appendHorizontalRule();

body.appendParagraph("FINAL SUMMARY")
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
"Staff On Leave : " + leaveCount
);

body.appendParagraph(
"Duplicate Check-In Device Groups : " + duplicateGroups
);

body.appendParagraph(
"Duplicate Check-Out Device Groups : " + checkoutDuplicateGroups
);

body.appendParagraph("");

body.appendHorizontalRule();

body.appendParagraph("FINANCIAL SUMMARY")
.setBold(true);

body.appendParagraph(
"Total Staff Payment : ₹" + grandTotalPayment
);

body.appendParagraph(
"Total Daily Expenses : ₹" + totalExpenses
);

body.appendParagraph(
"Overall Cost : ₹" +
(grandTotalPayment + totalExpenses)
).setBold(true);

body.appendParagraph("");

body.appendHorizontalRule();

body.appendParagraph("CONFIDENTIAL REPORT")
.setBold(true);

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


//====================================================
// PDF
//====================================================

const pdf = DriveApp
.getFileById(doc.getId())
.getAs(MimeType.PDF)
.setName(
"RRAS_Final_Report_" + today + ".pdf"
);


//====================================================
// EMAIL BODY
//====================================================

const html =

"<h2>RRAS Examination Center</h2>"+

"<b>FINAL EXAM REPORT</b><br><br>"+

"Exam Date : "+today+"<br>"+

"Exam Name : "+examName+"<br>"+

"Total Shift : "+totalShift+"<br><br>"+

"<table border='1' cellpadding='6' cellspacing='0'>"+

"<tr><td><b>Total Active Staff</b></td><td>"+totalActiveStaff+"</td></tr>"+

"<tr><td><b>Total Present Staff</b></td><td>"+totalPresent+"</td></tr>"+

"<tr><td><b>Total Active Staff Absent</b></td><td>"+absentStaff.length+"</td></tr>"+

"<tr><td><b>Permanent Staff</b></td><td>"+permanent+"</td></tr>"+

"<tr><td><b>Temporary Staff</b></td><td>"+temporary+"</td></tr>"+

"<tr><td><b>Staff On Leave</b></td><td>"+leaveCount+"</td></tr>"+

"<tr><td><b>Duplicate Check-In Devices</b></td><td>"+duplicateGroups+"</td></tr>"+

"<tr><td><b>Duplicate Check-Out Devices</b></td><td>"+checkoutDuplicateGroups+"</td></tr>"+

"<tr><td><b>Staff Payment</b></td><td>₹"+grandTotalPayment+"</td></tr>"+

"<tr><td><b>Daily Expenses</b></td><td>₹"+totalExpenses+"</td></tr>"+

"<tr><td><b>Overall Cost</b></td><td><b>₹"+(grandTotalPayment+totalExpenses)+"</b></td></tr>"+

"</table>"+

"<br><br>"+

"Please find the attached Final Examination Report PDF.";


//====================================================
// SEND EMAIL
//====================================================

MailApp.sendEmail({

to: ownerEmail,

subject:
"RRAS Final Examination Report - " + today,

htmlBody: html,

attachments: [pdf]

});


//====================================================
// DELETE TEMP DOC
//====================================================

DriveApp
.getFileById(doc.getId())
.setTrashed(true);

}
