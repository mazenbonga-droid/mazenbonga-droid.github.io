/**
 * =========================================================================
 * AUC FACULTY LAB - EQUIPMENT CHECKOUT SYSTEM (Google Apps Script Backend)
 * 100% Hosted inside AUC Google Workspace (@aucegypt.edu) - FERPA/Privacy Compliant
 * =========================================================================
 */

// 1. Serves the Web Application
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AUC Faculty Lab - Equipment Checkout')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. Automatically identify the logged-in AUC student or allow fallback ID login
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userEmail = "";
  try {
    userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (err) {
    userEmail = "";
  }

  const studentsSheet = ss.getSheetByName("Students");
  const equipmentSheet = ss.getSheetByName("Equipment");
  const bookingsSheet = ss.getSheetByName("Bookings");

  const studentsData = studentsSheet.getDataRange().getValues();
  const equipmentData = equipmentSheet.getDataRange().getValues();
  const bookingsData = bookingsSheet.getDataRange().getValues();

  // Parse equipment
  const equipmentList = [];
  for (let i = 1; i < equipmentData.length; i++) {
    const row = equipmentData[i];
    if (row[0]) {
      equipmentList.push({
        id: row[0].toString(),
        name: row[1].toString(),
        barcode: row[2].toString(),
        course: row[3].toString().toLowerCase().trim(),
        image: row[4].toString() || "https://images.unsplash.com/photo-1516035069371-29a1b244cc32?w=600",
        maxHours: row[5].toString() || "24 Hours",
        specs: row[6].toString() || "Standard Lab Equipment"
      });
    }
  }

  // Parse bookings
  const bookingsList = [];
  for (let i = 1; i < bookingsData.length; i++) {
    const row = bookingsData[i];
    if (row[0]) {
      bookingsList.push({
        id: row[0].toString(),
        equipmentId: row[1].toString(),
        equipmentName: row[2].toString(),
        studentName: row[3].toString(),
        studentEmail: row[4].toString(),
        pickup: row[5].toString(),
        return: row[6].toString(),
        status: row[7].toString()
      });
    }
  }

  // Check if current user is an authorized student
  let matchedStudent = null;
  if (userEmail) {
    for (let i = 1; i < studentsData.length; i++) {
      const sEmail = studentsData[i][2].toString().toLowerCase().trim();
      if (sEmail === userEmail) {
        matchedStudent = {
          id: studentsData[i][0].toString(),
          name: studentsData[i][1].toString(),
          email: sEmail,
          course: studentsData[i][3].toString().toLowerCase().trim(),
          status: studentsData[i][4].toString().toLowerCase().trim()
        };
        break;
      }
    }
  }

  return {
    currentUserEmail: userEmail,
    student: matchedStudent,
    equipment: equipmentList,
    bookings: bookingsList
  };
}

// 3. Fallback: Authenticate student by ID (if SSO email not auto-passed)
function verifyStudentById(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  const data = sheet.getDataRange().getValues();

  studentId = studentId.toString().trim();

  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][0].toString().trim();
    if (rowId === studentId) {
      return {
        success: true,
        student: {
          id: rowId,
          name: data[i][1].toString(),
          email: data[i][2].toString(),
          course: data[i][3].toString().toLowerCase().trim(),
          status: data[i][4].toString().toLowerCase().trim()
        }
      };
    }
  }

  return { success: false, message: "Student ID not found in authorized course roster." };
}

// 4. Save new Reservation directly into Google Sheet and send confirmation email
function createReservation(bookingPayload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");

  const bookingId = "AUC-BK-" + Math.floor(1000 + Math.random() * 9000);
  const timestamp = new Date();

  sheet.appendRow([
    bookingId,
    bookingPayload.equipmentId,
    bookingPayload.equipmentName,
    bookingPayload.studentName,
    bookingPayload.studentEmail,
    bookingPayload.pickup,
    bookingPayload.return,
    "Confirmed",
    bookingPayload.phone || "",
    timestamp
  ]);

  // Optional: Send official AUC confirmation email to student
  try {
    if (bookingPayload.studentEmail && bookingPayload.studentEmail.includes("@")) {
      MailApp.sendEmail({
        to: bookingPayload.studentEmail,
        subject: `[AUC Media Lab] Reservation Confirmed: ${bookingPayload.equipmentName} (#${bookingId})`,
        body: `Dear ${bookingPayload.studentName},\n\nYour equipment reservation at the AUC Media Lab has been confirmed.\n\nBooking ID: ${bookingId}\nEquipment: ${bookingPayload.equipmentName}\nPickup: ${bookingPayload.pickup}\nReturn: ${bookingPayload.return}\n\nPlease present your AUC Student ID card at the lab upon pickup.\n\nBest regards,\nAUC Media Lab Faculty Staff`
      });
    }
  } catch (e) {
    Logger.log("Email notification error: " + e.message);
  }

  return { success: true, bookingId: bookingId };
}

// 5. ONE-CLICK INITIAL SETUP: Run this ONCE to automatically create all sheets & sample data!
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Create Equipment Sheet
  let eqSheet = ss.getSheetByName("Equipment");
  if (!eqSheet) eqSheet = ss.insertSheet("Equipment");
  eqSheet.clear();
  eqSheet.appendRow(["ID", "Name", "Barcode", "Course", "Image_URL", "Max_Duration", "Specs"]);
  eqSheet.getRange(1, 1, 1, 7).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");

  eqSheet.appendRow(["1", "Sony FX3 Full-Frame Cinema Camera", "AUC-CAM-001", "cinematography", "https://images.unsplash.com/photo-1516035069371-29a1b244cc32?w=600", "24 Hours", "24-70mm f/2.8 Lens + 2x Batteries + 128GB Card"]);
  eqSheet.appendRow(["2", "Blackmagic Pocket Cinema 6K Pro", "AUC-CAM-002", "cinematography", "https://images.unsplash.com/photo-1502920917128-1aa500764cbd?w=600", "24 Hours", "Touch Screen + Cage + Pelican Case + Charger"]);
  eqSheet.appendRow(["3", "DJI Ronin RS3 Pro Gimbal", "AUC-GIM-001", "cinematography", "https://images.unsplash.com/photo-1584905066893-7d5c142ba4e1?w=600", "12 Hours", "Wireless Focus Motor + Dual Handgrip + Quick Plate"]);
  eqSheet.appendRow(["4", "Zoom H6 Audio Recorder", "AUC-AUD-001", "sound", "https://images.unsplash.com/photo-1590602847861-f357a9332bbc?w=600", "48 Hours", "4x XLR Combo Inputs + XY Mic Capsule + SD Card"]);
  eqSheet.appendRow(["5", "Sennheiser EW 100 G4 Wireless Lav Kit", "AUC-AUD-002", "sound", "https://images.unsplash.com/photo-1520523839898-507127054976?w=600", "24 Hours", "Transmitter + Bodypack Receiver + ME 2-II Microphone"]);
  eqSheet.appendRow(["6", "Aputure 300d II Light Kit", "AUC-LGT-001", "documentary", "https://images.unsplash.com/photo-1579783902614-a3fb3927b675?w=600", "24 Hours", "Heavy Duty Stand + Light Dome Softbox + Remote"]);

  // 2. Create Students Sheet
  let stSheet = ss.getSheetByName("Students");
  if (!stSheet) stSheet = ss.insertSheet("Students");
  stSheet.clear();
  stSheet.appendRow(["Student_ID", "Full_Name", "AUC_Email", "Course", "Status"]);
  stSheet.getRange(1, 1, 1, 5).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");

  stSheet.appendRow(["202401", "Alex Johnson", "alex.johnson@aucegypt.edu", "cinematography", "active"]);
  stSheet.appendRow(["202402", "Sarah Williams", "sarah.williams@aucegypt.edu", "sound", "active"]);
  stSheet.appendRow(["202403", "Michael Brown", "michael.brown@aucegypt.edu", "documentary", "active"]);
  stSheet.appendRow(["202404", "David Miller", "david.miller@aucegypt.edu", "cinematography", "suspended"]);

  // 3. Create Bookings Sheet
  let bkSheet = ss.getSheetByName("Bookings");
  if (!bkSheet) bkSheet = ss.insertSheet("Bookings");
  bkSheet.clear();
  bkSheet.appendRow(["Booking_ID", "Equipment_ID", "Equipment_Name", "Student_Name", "Student_Email", "Pickup_Time", "Return_Time", "Status", "Phone", "Timestamp"]);
  bkSheet.getRange(1, 1, 1, 10).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");

  // Remove default Sheet1 if exists
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Logger.log("✅ Initial setup complete! 3 sheets created with headers and sample data.");
}
