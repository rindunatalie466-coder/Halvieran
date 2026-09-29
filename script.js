/**
 * StorySlay 8th Grade - Google Apps Script Backend (Kode.gs)
 * Disesuaikan untuk Vercel / External Web Hosting via REST API (JSON)
 */

// Helper function to get the spreadsheet database
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("Running in standalone script mode or without bound spreadsheet.");
    return null;
  }
}

/**
 * REST API Endpoint GET (Menerima request dari Vercel)
 */
function doGet(e) {
  setupDatabaseSheets();
  
  const action = e && e.parameter ? e.parameter.action : '';
  let result = {};

  if (action === 'getTasks') {
    result = { success: true, tasks: getTasksData() };
  } else {
    result = { success: true, message: "StorySlay Apps Script REST API is active!" };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * REST API Endpoint POST (Menerima input penambahan/pembaharuan data dari Vercel)
 */
function doPost(e) {
  setupDatabaseSheets();
  
  let data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter || {};
  }

  let result = { success: false, message: "Action not recognized" };

  if (data.action === 'addTask') {
    result = addNewTask(data.task);
  } else if (data.action === 'updateStatus') {
    result = updateTaskStatus(data.taskId, data.newStatus);
  } else if (data.action === 'saveProfile') {
    result = saveProfileData(data.profile);
  } else if (data.action === 'submitContact') {
    result = submitContactMessage(data.contact);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Serves the HTML frontend application
 */
function doGet(e) {
  // Ensure default database structure exists in Google Sheets
  setupDatabaseSheets();

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('StorySlay 8th Grade - English Storytelling')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Utility helper to include partial HTML files if needed
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Automatically creates required sheets (Tasks, Profile, Contacts) with headers and initial data
 */
function setupDatabaseSheets() {
  const ss = getSpreadsheet();
  if (!ss) return;

  // 1. Setup 'Tasks' Sheet
  let tasksSheet = ss.getSheetByName('Tasks');
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet('Tasks');
    tasksSheet.appendRow(['ID', 'Subject', 'Title', 'Status', 'Deadline', 'Notes', 'CreatedAt']);
    
    tasksSheet.appendRow(['1', 'English Grade 8', 'Video Storytelling Practice: Malin Kundang', 'Proses', '2026-10-05', 'Rekam video durasi 2-3 menit. Boleh pakai kostum atau gambar pendukung!', new Date().toISOString()]);
    tasksSheet.appendRow(['2', 'English Grade 8', 'Draft Script Storytelling: Fable pilihan', 'Selesai', '2026-09-28', 'Pastikan memuat Orientation, Complication, dan Resolution yang jelas.', new Date().toISOString()]);
    tasksSheet.appendRow(['3', 'English Grade 8', 'Vocabulary Quiz: Action Verbs in Legend', 'Belum Mulai', '2026-10-10', 'Pelajari 20 kata kerja past tense dari PDF Modul Storytelling.', new Date().toISOString()]);
    tasksSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#E6E6FA');
  }

  // 2. Setup 'Profile' Sheet
  let profileSheet = ss.getSheetByName('Profile');
  if (!profileSheet) {
    profileSheet = ss.insertSheet('Profile');
    profileSheet.appendRow(['Key', 'Value']);
    profileSheet.appendRow(['Name', 'Aisyah Nur Ramadhani']);
    profileSheet.appendRow(['Class', 'Kelas 8-A • SMP Nusantara']);
    profileSheet.appendRow(['Bio', 'Suka nonton anime, dengerin K-Pop, sama lagi belajar pd ngomong English story tanpa text! 🎧📚']);
    profileSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#FFD1DC');
  }

  // 3. Setup 'Contacts' Sheet
  let contactsSheet = ss.getSheetByName('Contacts');
  if (!contactsSheet) {
    contactsSheet = ss.insertSheet('Contacts');
    contactsSheet.appendRow(['Timestamp', 'SenderName', 'Category', 'Message']);
    contactsSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#FEF08A');
  }
}

function getTasksData() {
  const ss = getSpreadsheet();
  if (!ss) return getSampleTasks();

  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return getSampleTasks();

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      tasks.push({
        id: row[0].toString(),
        subject: row[1] || 'English Grade 8',
        title: row[2] || '',
        status: row[3] || 'Belum Mulai',
        deadline: row[4] ? formatDate(row[4]) : '',
        notes: row[5] || ''
      });
    }
  }
  return tasks;
}

function addNewTask(taskObj) {
  const ss = getSpreadsheet();
  if (!ss) return { success: false, message: "Spreadsheet not connected" };

  let sheet = ss.getSheetByName('Tasks');
  if (!sheet) {
    setupDatabaseSheets();
    sheet = ss.getSheetByName('Tasks');
  }

  const newId = taskObj.id ? taskObj.id.toString() : Date.now().toString();
  sheet.appendRow([
    newId,
    taskObj.subject || 'English Grade 8',
    taskObj.title,
    taskObj.status || 'Belum Mulai',
    taskObj.deadline,
    taskObj.notes || '',
    new Date().toISOString()
  ]);

  return { success: true, id: newId };
}

function updateTaskStatus(taskId, newStatus) {
  const ss = getSpreadsheet();
  if (!ss) return { success: false, message: "Spreadsheet not connected" };

  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return { success: false, message: "Tasks sheet missing" };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === taskId.toString()) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      return { success: true, id: taskId, newStatus: newStatus };
    }
  }
  return { success: false, message: "Task ID not found" };
}

function saveProfileData(profileObj) {
  const ss = getSpreadsheet();
  if (!ss) return { success: false };

  let sheet = ss.getSheetByName('Profile');
  if (!sheet) {
    setupDatabaseSheets();
    sheet = ss.getSheetByName('Profile');
  }

  sheet.clearContents();
  sheet.appendRow(['Key', 'Value']);
  sheet.appendRow(['Name', profileObj.name]);
  sheet.appendRow(['Class', profileObj.userClass]);
  sheet.appendRow(['Bio', profileObj.bio]);

  return { success: true };
}

function submitContactMessage(contactObj) {
  const ss = getSpreadsheet();
  if (!ss) return { success: false };

  let sheet = ss.getSheetByName('Contacts');
  if (!sheet) {
    setupDatabaseSheets();
    sheet = ss.getSheetByName('Contacts');
  }

  sheet.appendRow([
    new Date().toLocaleString('id-ID'),
    contactObj.name,
    contactObj.category,
    contactObj.message
  ]);

  return { success: true };
}

function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    const yyyy = dateVal.getFullYear();
    const mm = String(dateVal.getMonth() + 1).padStart(2, '0');
    const dd = String(dateVal.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  return dateVal.toString();
}

function getSampleTasks() {
  return [
    { id: "1", subject: "English Grade 8", title: "Video Storytelling Practice: Malin Kundang", status: "Proses", deadline: "2026-10-05", notes: "Rekam video durasi 2-3 menit. Boleh pakai kostum atau gambar pendukung!" },
    { id: "2", subject: "English Grade 8", title: "Draft Script Storytelling: Fable pilihan", status: "Selesai", deadline: "2026-09-28", notes: "Pastikan memuat Orientation, Complication, dan Resolution yang jelas." },
    { id: "3", subject: "English Grade 8", title: "Vocabulary Quiz: Action Verbs in Legend", status: "Belum Mulai", deadline: "2026-10-10", notes: "Pelajari 20 kata kerja past tense dari PDF Modul Storytelling." }
  ];
}