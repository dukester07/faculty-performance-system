// =============================================================================
// VMRF-DU Faculty Weekly Performance System — Code.gs
// =============================================================================
// SETUP:
//  1. Open Google Sheet → Extensions → Apps Script
//  2. Paste this into Code.gs (replace all existing content)
//  3. Click + → New HTML file → name it exactly "Index" → paste Index.html
//  4. Run initializeSystem() once → approve permissions
//  5. Deploy → New Deployment → Web App
//     Execute as: Me | Access: Anyone (or your org)
//  6. Copy & share the Web App URL
// =============================================================================

var SH = {
  STAFF:       'Staff_Master',
  FACULTY:     'Faculty_Master',
  SUBMISSION:  'Weekly_Submission',
  TIMESHEET:   'Timesheet_Entries',
  SELF_ASSESS: 'Self_Assessment',
  HOD:         'HOD_Remarks',
  HOI:         'HOI_Remarks',
  IMO:         'IMO_Monitoring',
  NOTIF:       'Notifications'
};

var SCHEMA = {
  Staff_Master:      ['StaffID','StaffName','Email','Role','Department','PasswordHash','GoogleEmail','Status'],
  Faculty_Master:    ['FacultyID','FacultyName','Email','Department','Campus','Institution','Designation','PasswordHash','GoogleEmail','Status'],
  Weekly_Submission: ['SubmissionID','FacultyID','AcademicYearSemester','ReportingFrom','ReportingTo','Declaration','SubmittedDateTime'],
  Timesheet_Entries: ['SubmissionID','Day','TimeSlot','ActivityType','ActivityDetails'],
  Self_Assessment:   ['SubmissionID','OutcomeOfWeek','TargetPlanNextWeek'],
  HOD_Remarks:       ['SubmissionID','HOD_Remark','HOD_Status','HOD_DateTime'],
  HOI_Remarks:       ['SubmissionID','HOI_Remark','HOI_Status','HOI_DateTime'],
  IMO_Monitoring:    ['SubmissionID','IMO_Remark','IMO_Status','IMO_DateTime'],
  // NotifID | ForRole | Type | Title | Body | SubmissionID | FacultyName | IsRead | CreatedAt
  Notifications:     ['NotifID','ForRole','Type','Title','Body','SubmissionID','FacultyName','IsRead','CreatedAt']
};

var ACTIVITY_TYPES = [
  'Lecture Delivery (Theory/Practical)',
  'Course Material/PPT Preparation',
  'Question Paper Setting / Answer Script Evaluation',
  'LMS Content Uploads',
  'Research & Development Activities',
  'Project Proposal Development',
  'FDP Participation',
  'ERP Updates',
  'Student Mentoring',
  'Event Support & Committee Meetings',
  'FDP / SWAYAM / NPTEL Course Participation',
  'Career Guidance or Competitive Exam Support',
  'Alumni/Parent Interaction',
  'Department Coordination',
  'Other'
];

var TIME_SLOTS = [
  '8:30 AM – 9:30 AM',
  '9:30 AM – 10:30 AM',
  '10:30 AM – 11:30 AM',
  '11:30 AM – 12:30 PM',
  '1:00 PM – 2:00 PM',
  '2:00 PM – 3:00 PM',
  '3:00 PM – 3:30 PM'
];

var DEPARTMENTS = [
  // Medical & Clinical
  'Anatomy','Physiology','Biochemistry','Pharmacology','Pathology','Microbiology',
  'Forensic Medicine','Community Medicine','General Medicine','General Surgery',
  'Obstetrics & Gynaecology','Paediatrics','Orthopaedics','Ophthalmology','ENT',
  'Anaesthesia','Radiology','Dermatology','Psychiatry','Physiotherapy',
  'Nursing','Allied Health Sciences','Rehabilitation Sciences',
  // Engineering
  'Computer Science & Engineering',
  'Information Technology',
  'Electronics & Communication Engineering',
  'Electrical & Electronics Engineering',
  'Mechanical Engineering',
  'Civil Engineering',
  'Biomedical Engineering',
  'Artificial Intelligence & Data Science',
  'Computer Science & Business Systems',
  'Robotics & Automation',
  // Sciences & Humanities
  'Mathematics','Physics','Chemistry','Biology',
  'English','Tamil','Management Studies',
  'Commerce','Economics','Business Administration',
  'Library & Information Science','Physical Education',
  'Other'
];

var DESIGNATIONS   = ['Professor','Associate Professor','Assistant Professor','Lecturer','Other'];
var CAMPUSES       = ["Vinayaka Mission's Chennai Campus","Vinayaka Mission's Puducherry Campus"];
var ACADEMIC_YEARS = ['2025–2026 (Odd Semester)','2025–2026 (Even Semester)','2024–2025 (Odd Semester)','2024–2025 (Even Semester)'];
var DAYS           = ['Day 1 (Mon)','Day 2 (Tue)','Day 3 (Wed)','Day 4 (Thu)','Day 5 (Fri)'];
var INSTITUTIONS   = [
  'Aarupadai Veedu Medical College and Hospital (AVMC)',
  'Vinayaka Mission College of Nursing (VMCN)',
  'School of Physiotherapy (SOP)',
  'School of Allied Health Sciences (SAHS)',
  'School of Rehabilitation and Behavioral Sciences (SRBS)'
];

// ─── WEB APP ENTRY ────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VMRF-DU Faculty Performance System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VMRF IMO')
    .addItem('Setup / Re-initialise System', 'initializeSystem')
    .addItem('Open Web App', 'openWebApp')
    .addToUi();
}

function openWebApp() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<script>window.open("'+url+'","_blank");google.script.host.close();<\/script>').setWidth(1).setHeight(1),
    'Opening Web App...'
  );
}

// ─── INITIALIZE ───────────────────────────────────────────────────────────────
function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = Object.keys(SH).map(function(k){ return SH[k]; });
  names.forEach(function(n){ if(!ss.getSheetByName(n)) ss.insertSheet(n); });
  names.forEach(function(n){
    var sheet = ss.getSheetByName(n), hdrs = SCHEMA[n];
    if(sheet.getRange(1,1).getValue() !== hdrs[0]){
      var r = sheet.getRange(1,1,1,hdrs.length);
      r.setValues([hdrs]);
      r.setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });
  _applyValidations(ss);
  _applyFormatting(ss);
  _setColumnWidths(ss);
  _resetTriggers();
  // Set default registration codes if not already set
  var p=PropertiesService.getScriptProperties();
  if(!p.getProperty('REGCODE_HOD')) p.setProperty('REGCODE_HOD','HOD@VMRF');
  if(!p.getProperty('REGCODE_HOI')) p.setProperty('REGCODE_HOI','HOI@VMRF');
  if(!p.getProperty('REGCODE_IMO')) p.setProperty('REGCODE_IMO','IMO@VMRF');
  try { SpreadsheetApp.getUi().alert('✅ VMRF System Ready!\n\nDefault registration codes:\n  HOD: HOD@VMRF\n  HOI: HOI@VMRF\n  IMO: IMO@VMRF\n\nChange via Script Properties (REGCODE_HOD, REGCODE_HOI, REGCODE_IMO)\n\nDeploy as Web App when ready.'); } catch(e) {}
}

function _applyValidations(ss) {
  var R = 500;
  var fm = ss.getSheetByName(SH.FACULTY);
  _dv(fm,2,4,R,DEPARTMENTS); _dv(fm,2,5,R,CAMPUSES);
  _dv(fm,2,6,R,INSTITUTIONS); _dv(fm,2,7,R,DESIGNATIONS);
  var ws = ss.getSheetByName(SH.SUBMISSION);
  _dv(ws,2,3,R,ACADEMIC_YEARS); _dv(ws,2,6,R,['YES','NO']);
  ws.getRange(2,4,R,2).setNumberFormat('dd-MMM-yyyy');
  // Timesheet_Entries is written programmatically — no cell validation needed
  var ts = ss.getSheetByName(SH.TIMESHEET);
  ts.getRange(2,1,R,5).clearDataValidations();
  _dv(ss.getSheetByName(SH.HOD),2,3,R,['Approved','Needs Revision','Rejected']);
  _dv(ss.getSheetByName(SH.HOI),2,3,R,['Approved','Needs Revision','Rejected']);
  _dv(ss.getSheetByName(SH.IMO),2,3,R,['Under Review','Finalised','Escalated']);
}
function _dv(sheet,sr,col,nr,list){
  sheet.getRange(sr,col,nr,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(false).build());
}
function _applyFormatting(ss) {
  [{n:SH.HOD,c:3,m:[['Approved','#b7e1cd'],['Needs Revision','#fce8b2'],['Rejected','#f4c7c3']]},
   {n:SH.HOI,c:3,m:[['Approved','#b7e1cd'],['Needs Revision','#fce8b2'],['Rejected','#f4c7c3']]},
   {n:SH.IMO,c:3,m:[['Finalised','#b7e1cd'],['Under Review','#fce8b2'],['Escalated','#f4c7c3']]}
  ].forEach(function(cfg){
    var sheet=ss.getSheetByName(cfg.n),range=sheet.getRange(2,cfg.c,500,1);
    sheet.setConditionalFormatRules(cfg.m.map(function(r){
      return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setRanges([range]).build();
    }));
  });
}
function _setColumnWidths(ss) {
  var map={};
  map[SH.STAFF]=[120,180,220,80,180,160,200,80];
  map[SH.FACULTY]=[110,180,200,150,220,260,160,160,200,80];
  map[SH.SUBMISSION]=[200,120,170,120,120,80,160];
  map[SH.TIMESHEET]=[200,100,170,280,200];
  map[SH.SELF_ASSESS]=[200,360,360];
  map[SH.HOD]=map[SH.HOI]=map[SH.IMO]=[200,320,160,160];
  map[SH.NOTIF]=[180,80,120,300,500,200,160,60,160];
  Object.keys(map).forEach(function(n){
    var s=ss.getSheetByName(n);
    if(s) map[n].forEach(function(w,i){s.setColumnWidth(i+1,w);});
  });
}
function _resetTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('sendFridayReminders').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(12).create();
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
// Four separate role logins — each stored separately:
//   Faculty  → Faculty_Master sheet (Faculty ID + password or Google)
//   HOD/HOI/IMO → Staff_Master sheet (email + password or Google)
//
// Registration:
//   Faculty:  activateAccount(facultyID, password, confirmPassword)
//             — activates an IMO-pre-enrolled Pending row
//   Staff:    staffRegister(f)
//             — self-registers with a role-specific reg code
//             — codes stored in ScriptProperties: REGCODE_HOD, REGCODE_HOI, REGCODE_IMO

// ── Helper: hash password (Djb2) ─────────────────────────────────────────────
function _hashPwd(pwd) {
  var hash = 5381, s = String(pwd);
  for (var i = 0; i < s.length; i++) { hash = ((hash << 5) + hash) + s.charCodeAt(i); hash = hash & hash; }
  return 'H' + (hash >>> 0).toString(16).toUpperCase();
}

// ── Helper: generate random ID ───────────────────────────────────────────────
function _makeID(prefix) {
  var ch = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', r = prefix + '-';
  for (var k = 0; k < 6; k++) r += ch[Math.floor(Math.random() * ch.length)];
  return r;
}

// ── 1. Faculty login (Faculty ID + password) ──────────────────────────────────
function facultyLogin(facultyID, password) {
  if (!facultyID) throw new Error('Please enter your Faculty ID.');
  if (!password)  throw new Error('Please enter your password.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var data = sheet.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('FacultyID'), nmI = h.indexOf('FacultyName');
  var pwI = h.indexOf('PasswordHash'), stI = h.indexOf('Status');
  var fid = String(facultyID).trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idI]||'').trim().toUpperCase() !== fid) continue;
    if (String(data[i][stI]||'').trim() === 'Pending')
      throw new Error('Account not yet activated. Please activate your account first.');
    var stored = String(data[i][pwI]||'').trim();
    if (!stored) throw new Error('No password set. Please re-activate your account.');
    if (stored !== _hashPwd(password)) throw new Error('Incorrect password. Please try again.');
    return { success:true, role:'FACULTY', facultyID:String(data[i][idI]).trim(), facultyName:String(data[i][nmI]||'') };
  }
  throw new Error('Faculty ID not found. Check the ID given to you by the IMO.');
}

// ── 2. Staff login (HOD / HOI / IMO — by email + password) ───────────────────
function staffLogin(role, staffID, password) {
  if (!role)     throw new Error('Role is required.');
  if (!staffID)  throw new Error('Please enter your Staff ID.');
  if (!password) throw new Error('Please enter your password.');
  role = role.toUpperCase();
  if (['HOD','HOI','IMO'].indexOf(role) < 0) throw new Error('Invalid role.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
  if (!sheet) throw new Error('Staff sheet not found. Please run initializeSystem() first.');
  var data = sheet.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('StaffID'), nmI = h.indexOf('StaffName');
  var pwI = h.indexOf('PasswordHash'), stI = h.indexOf('Status'), rlI = h.indexOf('Role');
  var sid = String(staffID).trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idI]||'').trim().toUpperCase() !== sid) continue;
    if (String(data[i][rlI]||'').toUpperCase() !== role) continue;
    if (String(data[i][stI]||'').trim() !== 'Active')
      throw new Error('Account not active. Please contact the administrator.');
    var stored = String(data[i][pwI]||'').trim();
    if (!stored) throw new Error('No password set for this account.');
    if (stored !== _hashPwd(password)) throw new Error('Incorrect password. Please try again.');
    return { success:true, role:role, staffID:String(data[i][idI]).trim(), staffName:String(data[i][nmI]||'') };
  }
  throw new Error(role + ' ID not found. Check the ID shown to you after registration.');
}

// ── 3. Faculty Google login ──────────────────────────────────────────────────
function facultyGoogleLogin() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return { success:false, reason:'no_email' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var data = sheet.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('FacultyID'), nmI = h.indexOf('FacultyName');
  var gmI = h.indexOf('GoogleEmail'), stI = h.indexOf('Status');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][gmI]||'').toLowerCase() !== email.toLowerCase()) continue;
    if (String(data[i][stI]||'') !== 'Active') return { success:false, reason:'pending', email:email };
    return { success:true, role:'FACULTY', facultyID:String(data[i][idI]).trim(), facultyName:String(data[i][nmI]||''), email:email };
  }
  return { success:false, reason:'not_found', email:email };
}

// ── 4. Staff Google login ─────────────────────────────────────────────────────
function staffGoogleLogin(role) {
  if (!role) return { success:false, reason:'no_role' };
  role = role.toUpperCase();
  var email = Session.getActiveUser().getEmail();
  if (!email) return { success:false, reason:'no_email' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
  if (!sheet) return { success:false, reason:'no_sheet' };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), nmI = h.indexOf('StaffName');
  var gmI = h.indexOf('GoogleEmail'), stI = h.indexOf('Status'), rlI = h.indexOf('Role');
  var idI = h.indexOf('StaffID');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][gmI]||'').toLowerCase() !== email.toLowerCase()) continue;
    if (String(data[i][rlI]||'').toUpperCase() !== role) continue;
    if (String(data[i][stI]||'') !== 'Active') return { success:false, reason:'pending', email:email };
    return { success:true, role:role, staffID:String(data[i][idI]).trim(), staffName:String(data[i][nmI]||''), email:email };
  }
  return { success:false, reason:'not_found', email:email };
}

// ── 5. Faculty self-registration — auto-generates Faculty ID ─────────────────
// Faculty fill in their details and set a password. No pre-enrollment needed.
// Faculty ID (VMRF-XXXXXX) is generated automatically and returned to the user.
function facultyRegister(f) {
  if (!f.name)     throw new Error('Full name is required.');
  if (!f.email)    throw new Error('Email address is required.');
  if (!f.department)  throw new Error('Please select a department.');
  if (!f.designation) throw new Error('Please select a designation.');
  if (!f.campus)      throw new Error('Please select a campus.');
  if (!f.institution) throw new Error('Please select an institution.');
  if (!f.password)    throw new Error('Please set a password.');
  if (f.password.length < 6) throw new Error('Password must be at least 6 characters.');
  if (f.password !== f.confirmPassword) throw new Error('Passwords do not match.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var h    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();

  // Email uniqueness check
  var emI = h.indexOf('Email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').toLowerCase() === f.email.toLowerCase())
      throw new Error('An account with this email already exists. Please sign in.');
  }

  // Generate unique Faculty ID
  var existing = data.slice(1).map(function(r){ return String(r[0]); });
  var id = _makeID('VMRF');
  while (existing.indexOf(id) !== -1) id = _makeID('VMRF');

  // Write row by column name
  var vals = {
    'FacultyID': id, 'FacultyName': f.name, 'Email': f.email,
    'Department': f.department, 'Campus': f.campus, 'Institution': f.institution,
    'Designation': f.designation, 'PasswordHash': _hashPwd(f.password),
    'GoogleEmail': f.googleEmail || '', 'Status': 'Active'
  };
  var newRow = h.map(function(col){ return vals[col] !== undefined ? vals[col] : ''; });
  sheet.appendRow(newRow);
  return { success:true, facultyID:id, facultyName:f.name };
}

// ── 6. Staff self-registration (HOD / HOI / IMO) — auto-generates Staff ID ───
function staffRegister(f) {
  if (!f.role)     throw new Error('Role is required.');
  if (!f.name)     throw new Error('Full name is required.');
  if (!f.email)    throw new Error('Email address is required.');
  if (!f.password) throw new Error('Please set a password.');
  if (f.password.length < 6) throw new Error('Password must be at least 6 characters.');
  if (f.password !== f.confirmPassword) throw new Error('Passwords do not match.');

  var role = f.role.toUpperCase();
  if (['HOD','HOI','IMO'].indexOf(role) < 0) throw new Error('Invalid role.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
  if (!sheet) throw new Error('Staff sheet not found. Please run initializeSystem() first.');
  var data = sheet.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), rlI = h.indexOf('Role');

  // Check email+role uniqueness
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').toLowerCase() === f.email.toLowerCase() &&
        String(data[i][rlI]||'').toUpperCase() === role)
      throw new Error('An account with this email already exists for ' + role + '. Please sign in.');
  }

  // Generate unique Staff ID
  var existing = data.slice(1).map(function(r){ return String(r[0]); });
  var id = _makeID(role);
  while (existing.indexOf(id) !== -1) id = _makeID(role);

  // Build row by column name
  var vals = {
    'StaffID': id, 'StaffName': f.name, 'Email': f.email,
    'Role': role, 'Department': f.department||'',
    'PasswordHash': _hashPwd(f.password), 'GoogleEmail': '', 'Status': 'Active'
  };
  var newRow = h.map(function(col){ return vals[col] !== undefined ? vals[col] : ''; });
  sheet.appendRow(newRow);
  return { success:true, staffID:id, staffName:f.name, role:role };
}

// ── 7. Change password ───────────────────────────────────────────────────────
function changePassword(role, identifier, oldPwd, newPwd) {
  if (!newPwd || newPwd.length < 6) throw new Error('New password must be at least 6 characters.');
  role = String(role).toUpperCase();
  if (role === 'FACULTY') {
    facultyLogin(identifier, oldPwd); // throws if wrong
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
    _ensureFacultyColumns(sheet);
    var data = sheet.getDataRange().getValues(), h = data[0];
    var idI = h.indexOf('FacultyID'), pwI = h.indexOf('PasswordHash');
    var fid = String(identifier).trim().toUpperCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]||'').trim().toUpperCase() === fid) {
        sheet.getRange(i+1, pwI+1).setValue(_hashPwd(newPwd));
        return { ok:true };
      }
    }
    throw new Error('Faculty record not found.');
  } else {
    staffLogin(role, identifier, oldPwd); // throws if wrong
    var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
    var d2 = sh2.getDataRange().getValues(), h2 = d2[0];
    var emI2 = h2.indexOf('Email'), pwI2 = h2.indexOf('PasswordHash'), rlI2 = h2.indexOf('Role');
    for (var j = 1; j < d2.length; j++) {
      if (String(d2[j][emI2]||'').toLowerCase() === identifier.toLowerCase() &&
          String(d2[j][rlI2]||'').toUpperCase() === role) {
        sh2.getRange(j+1, pwI2+1).setValue(_hashPwd(newPwd));
        return { ok:true };
      }
    }
    throw new Error('Staff record not found.');
  }
}

// ── 8. IMO enrolls faculty (unchanged) ───────────────────────────────────────
function preEnrollFaculty(f) {
  if (!f.name||!f.department||!f.designation||!f.campus||!f.institution)
    throw new Error('All fields are required.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var h    = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();
  if (f.email) {
    var emI = h.indexOf('Email');
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][emI]||'').toLowerCase() === f.email.toLowerCase())
        throw new Error('A faculty with this email already exists.');
    }
  }
  var existing = data.slice(1).map(function(r){ return String(r[0]); });
  var id = _makeID('VMRF');
  while (existing.indexOf(id) !== -1) id = _makeID('VMRF');
  var vals = {
    'FacultyID':id,'FacultyName':f.name,'Email':f.email||'',
    'Department':f.department,'Campus':f.campus,'Institution':f.institution,
    'Designation':f.designation,'PasswordHash':'','GoogleEmail':'','Status':'Pending'
  };
  var newRow = h.map(function(col){ return vals[col]!==undefined?vals[col]:''; });
  sheet.appendRow(newRow);
  return { id:id, name:f.name };
}

// ─── ENSURE FACULTY MASTER COLUMNS ───────────────────────────────────────────
// Adds any missing columns from SCHEMA to an existing Faculty_Master sheet.
// Called before every login so stale sheets from old initializeSystem() runs
// don't cause "PasswordHash column missing" errors.
function _ensureFacultyColumns(sheet) {
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var required = SCHEMA[SH.FACULTY];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) === -1) {
      // Append missing column header at the end
      var newCol = headers.length + 1;
      sheet.getRange(1, newCol).setValue(required[i]);
      headers.push(required[i]);
    }
  }
}

// ─── DEBUG / REPAIR FACULTY ROW ──────────────────────────────────────────────
// Run this from Apps Script editor to see what's actually in the sheet.
// Returns the raw header row and all data rows so you can diagnose misalignment.
function debugFacultySheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) return { error: 'Faculty_Master sheet not found' };
  var data = sheet.getDataRange().getValues();
  var result = { headers: data[0], rows: [] };
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < data[0].length; j++) {
      row[data[0][j] || 'col_' + j] = data[i][j];
    }
    result.rows.push(row);
  }
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Rebuilds Faculty_Master with correct canonical headers.
// Detects columns by POSITION if headers are missing/wrong,
// or by NAME if headers exist. Wipes and rewrites cleanly.
// Run once from the Apps Script editor after any schema change.
function repairFacultyMaster() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) { Logger.log('Sheet not found'); return; }

  var canonical = SCHEMA[SH.FACULTY];
  var raw = sheet.getDataRange().getValues();

  Logger.log('=== BEFORE REPAIR ===');
  Logger.log('Row count: ' + raw.length);
  Logger.log('Headers: ' + JSON.stringify(raw[0]));
  for (var d = 1; d < raw.length; d++) Logger.log('Row ' + d + ': ' + JSON.stringify(raw[d]));

  var oldHeaders = raw[0].map(function(h){ return String(h).trim(); });
  var hasHeaders = oldHeaders.indexOf('FacultyID') >= 0;

  var objects = [];
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i].some(function(v){ return v !== ''; })) continue; // skip blank rows
    var obj = {};
    if (hasHeaders) {
      // Map by column name
      for (var j = 0; j < oldHeaders.length; j++) {
        if (oldHeaders[j]) obj[oldHeaders[j]] = raw[i][j];
      }
    } else {
      // Fallback: map by old 7-column positional order (pre-schema-update)
      var pos7 = ['FacultyID','FacultyName','Email','Department','Campus','Institution','Designation'];
      for (var p = 0; p < pos7.length && p < raw[i].length; p++) {
        obj[pos7[p]] = raw[i][p];
      }
    }
    if (obj['FacultyID']) objects.push(obj);
  }

  // Clear and rewrite
  sheet.clearContents();
  var newRows = [canonical];
  for (var k = 0; k < objects.length; k++) {
    var row = canonical.map(function(col){ return objects[k][col] !== undefined ? objects[k][col] : ''; });
    newRows.push(row);
  }
  sheet.getRange(1, 1, newRows.length, canonical.length).setValues(newRows);

  Logger.log('=== AFTER REPAIR ===');
  Logger.log('Faculty_Master repaired. ' + objects.length + ' rows migrated.');
  Logger.log('New headers: ' + canonical.join(', '));
  SpreadsheetApp.getUi().alert('Repair complete! ' + objects.length + ' faculty rows migrated.\nNOTE: Any rows registered before this fix had no password stored — those users must re-register.');
  return { repaired: objects.length, headers: canonical };
}


// ─── CONFIG ───────────────────────────────────────────────────────────────────
function getConfig() {
  return {
    activityTypes: ACTIVITY_TYPES,
    timeSlots:     TIME_SLOTS,
    departments:   DEPARTMENTS,
    designations:  DESIGNATIONS,
    campuses:      CAMPUSES,
    academicYears: ACADEMIC_YEARS,
    days:          DAYS,
    institutions:  INSTITUTIONS
  };
}

// ─── FACULTY REGISTER ─── see selfRegisterFaculty in AUTH section above ──




// ─── FACULTY LIST ─────────────────────────────────────────────────────────────
function getFacultyList() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY).getDataRange().getValues();
  var h = data[0];
  // Always look up columns by name — never by hardcoded index
  var idI   = h.indexOf('FacultyID');
  var nmI   = h.indexOf('FacultyName');
  var emI   = h.indexOf('Email');
  var dpI   = h.indexOf('Department');
  var cpI   = h.indexOf('Campus');
  var inI   = h.indexOf('Institution');
  var dgI   = h.indexOf('Designation');
  var stI   = h.indexOf('Status');
  var out = [];
  for(var i=1;i<data.length;i++){
    if(data[i][idI] && String(data[i][stI]||'').trim()==='Active')
      out.push({
        id:          String(data[i][idI]),
        name:        String(data[i][nmI]||''),
        email:       String(data[i][emI]||''),
        dept:        String(data[i][dpI]||''),
        campus:      String(data[i][cpI]||''),
        institution: String(data[i][inI]||''),
        designation: String(data[i][dgI]||'')
      });
  }
  return out;
}

// ─── MY SUBMISSIONS ───────────────────────────────────────────────────────────
function getMySubmissions(facultyID) {
  if(!facultyID) throw new Error('Faculty ID is required.');
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH=subD[0];
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var hodMap=_bm(hodD,hodD[0]), hoiMap=_bm(hoiD,hoiD[0]), imoMap=_bm(imoD,imoD[0]);
  var out=[];
  for(var i=1;i<subD.length;i++){
    if(String(subD[i][subH.indexOf('FacultyID')])!==String(facultyID)) continue;
    var sid=String(subD[i][subH.indexOf('SubmissionID')]);
    var hod=hodMap[sid]||{}, hoi=hoiMap[sid]||{}, imo=imoMap[sid]||{};
    out.push({
      submissionID: sid,
      semester:     String(subD[i][subH.indexOf('AcademicYearSemester')]||''),
      from:         _fmt(subD[i][subH.indexOf('ReportingFrom')]),
      to:           _fmt(subD[i][subH.indexOf('ReportingTo')]),
      submitted:    _fmtDT(subD[i][subH.indexOf('SubmittedDateTime')]),
      hodStatus:    hod['HOD_Status']||'Pending',
      hoiStatus:    hoi['HOI_Status']||'Pending',
      imoStatus:    imo['IMO_Status']||'Pending',
      hodRemark:    String(hod['HOD_Remark']||''),
      hoiRemark:    String(hoi['HOI_Remark']||''),
      imoRemark:    String(imo['IMO_Remark']||'')
    });
  }
  return out.reverse();
}

// ─── SUBMIT WEEKLY REPORT ────────────────────────────────────────────────────
function submitWeeklyReport(data) {
  if(!data.facultyID)         throw new Error('Faculty ID is required.');
  if(!data.academicYearSem)   throw new Error('Please select the Academic Year / Semester.');
  if(!data.reportingFrom)     throw new Error('Please set the Reporting From date.');
  if(!data.reportingTo)       throw new Error('Please set the Reporting To date.');
  if(data.reportingFrom>data.reportingTo) throw new Error('Reporting From cannot be after Reporting To.');
  if(!data.outcomeOfWeek)     throw new Error('Please fill in the Outcome of the Week.');
  if(!data.targetPlanNextWeek)throw new Error('Please fill in the Target Plan for Next Week.');
  if(data.declaration!=='YES')throw new Error('Declaration must be YES to submit.');

  var ss=SpreadsheetApp.getActiveSpreadsheet(), sid=_uid(), now=new Date();
  ss.getSheetByName(SH.SUBMISSION).appendRow([sid,data.facultyID,data.academicYearSem,new Date(data.reportingFrom),new Date(data.reportingTo),data.declaration,now]);

  if(data.timesheet&&data.timesheet.length){
    var tsRows=data.timesheet.map(function(e){return [sid,e.day,e.slot,e.activity,e.details||''];});
    var tsSheet=ss.getSheetByName(SH.TIMESHEET);
    var tsStart=tsSheet.getLastRow()+1;
    // Clear any stale data validation before writing
    tsSheet.getRange(tsStart,1,tsRows.length,5).clearDataValidations();
    tsSheet.getRange(tsStart,1,tsRows.length,5).setValues(tsRows);
  }
  ss.getSheetByName(SH.SELF_ASSESS).appendRow([sid,data.outcomeOfWeek,data.targetPlanNextWeek]);
  // Pre-create blank review rows so _bm lookups always find a row
  ss.getSheetByName(SH.HOD).appendRow([sid,'','','']);
  ss.getSheetByName(SH.HOI).appendRow([sid,'','','']);
  ss.getSheetByName(SH.IMO).appendRow([sid,'','','']);

  var facRow=_rowByKey(SH.FACULTY,data.facultyID,'FacultyID')||{};
  var facName=String(facRow['FacultyName']||data.facultyID);
  var facDept=String(facRow['Department']||'');
  var facPeriod=data.reportingFrom+' to '+data.reportingTo;
  // In-app notification → HOD
  _pushNotif('HOD','new_submission',
    '📋 New Submission from '+facName,
    facName+(facDept?' ('+facDept+')':'')+' submitted a weekly report for '+facPeriod+'. Awaiting your review.',
    sid, facName);
  try { _notifyHOD(ss,sid,data.facultyID); } catch(e) { Logger.log('HOD notify failed: '+e.message); }
  return { sid:sid };
}

// ─── HOD QUEUE ────────────────────────────────────────────────────────────────
function getHODQueue() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var facMap=_bmByCol(facD,'FacultyID');
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var hodH=hodD[0];
  var stI=hodH.indexOf('HOD_Status'), sbI=hodH.indexOf('SubmissionID');
  var out=[];
  for(var i=1;i<hodD.length;i++){
    var sid=String(hodD[i][sbI]||'').trim();
    var st =String(hodD[i][stI]||'').trim();
    if(!sid) continue;
    // Show in HOD queue: no decision yet (blank) or returned for revision
    if(st!==''&&st!=='Needs Revision') continue;
    var sub=subMap[sid]||{};
    var fid=String(sub['FacultyID']||'').trim();
    var fac=fid?(facMap[fid]||{}):{};
    var sa=saMap[sid]||{};
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],{hodStatus:st},null,null));
  }
  return out;
}

function submitHODReview(sid, remark, status) {
  if(!sid)    throw new Error('Submission ID missing.');
  if(!status) throw new Error('Please select a status.');
  if(status!=='Approved'&&!remark) throw new Error('Please enter your remarks.');
  _writeReview(SH.HOD,sid,remark||'',status);
  var ss2=SpreadsheetApp.getActiveSpreadsheet();
  var sub2=_rowByKey(SH.SUBMISSION,sid)||{};
  var fac2=sub2['FacultyID']?(_rowByKey(SH.FACULTY,String(sub2['FacultyID']||''),'FacultyID')||{}):{}; 
  var fn2=String(fac2['FacultyName']||sid);
  if(status==='Approved'){
    _pushNotif('HOI','hod_approved','✅ HOD Approved — Review Required',
      'Submission by '+fn2+' ('+sid+') has been approved by HOD and is awaiting your review.',sid,fn2);
    var fid2=String(sub2['FacultyID']||'').trim().toUpperCase();
    _pushNotif('FACULTY:'+fid2,'status_update','Your submission was reviewed by HOD',
      'HOD has approved your submission '+sid+'. It has been forwarded to the Head of Institution for review.',sid,fn2);
    try{_notifyHOI(ss2,sid);}catch(e){Logger.log(e.message);}
  } else {
    _pushNotif('FACULTY:'+fid2,'needs_revision','⚠️ Submission Needs Revision',
      'Your submission '+sid+' was returned by HOD'+(remark?' with the remark: '+remark:'')+'.',sid,fn2);
    try{_notifyRevision(ss2,sid,'Head of Department',remark);}catch(e){Logger.log(e.message);}
  }
  return { ok:true };
}

// ─── HOI QUEUE ────────────────────────────────────────────────────────────────
function getHOIQueue() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var facMap=_bmByCol(facD,'FacultyID');
  var hodMap=_bm(hodD,hodD[0]);
  var hoiMap=_bm(hoiD,hoiD[0]);
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var hodH=hodD[0];
  var hodStI=hodH.indexOf('HOD_Status'), hodSbI=hodH.indexOf('SubmissionID');
  var out=[];
  for(var i=1;i<hodD.length;i++){
    var sid=String(hodD[i][hodSbI]||'').trim();
    var hst=String(hodD[i][hodStI]||'').trim();
    if(!sid||hst!=='Approved') continue;
    var hoiR=hoiMap[sid]||{}, hoiSt=String(hoiR['HOI_Status']||'').trim();
    if(hoiSt!==''&&hoiSt!=='Needs Revision') continue;
    var sub=subMap[sid]||{};
    var fid=String(sub['FacultyID']||'').trim();
    var fac=fid?(facMap[fid]||{}):{};
    var sa=saMap[sid]||{}, hodR=hodMap[sid]||{};
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],
      {hodStatus:hst,hodRemark:String(hodR['HOD_Remark']||'')},
      {hoiStatus:hoiSt},null));
  }
  return out;
}

function submitHOIReview(sid, remark, status) {
  if(!sid)    throw new Error('Submission ID missing.');
  if(!status) throw new Error('Please select a status.');
  if(status!=='Approved'&&!remark) throw new Error('Please enter your remarks.');
  var hodSt=_getCell(SH.HOD,sid,'HOD_Status');
  if(hodSt!=='Approved') throw new Error('HOD must approve this submission before HOI can review it.');
  _writeReview(SH.HOI,sid,remark||'',status);
  var ss3=SpreadsheetApp.getActiveSpreadsheet();
  var sub3=_rowByKey(SH.SUBMISSION,sid)||{};
  var fac3=sub3['FacultyID']?(_rowByKey(SH.FACULTY,String(sub3['FacultyID']||''),'FacultyID')||{}):{}; 
  var fn3=String(fac3['FacultyName']||sid);
  if(status==='Approved'){
    _pushNotif('IMO','hoi_approved','✅ HOI Approved — Monitoring Required',
      'Submission by '+fn3+' ('+sid+') has been approved by HOD and HOI. Ready for your final monitoring decision.',sid,fn3);
    var fid3=String(sub3['FacultyID']||'').trim().toUpperCase();
    _pushNotif('FACULTY:'+fid3,'status_update','Your submission was reviewed by HOI',
      'HOI has approved your submission '+sid+'. It has been forwarded to the IMO for final monitoring.',sid,fn3);
    try{_notifyIMO(ss3,sid);}catch(e){Logger.log(e.message);}
  } else {
    _pushNotif('FACULTY:'+fid3,'needs_revision','⚠️ Submission Needs Revision',
      'Your submission '+sid+' was returned by HOI'+(remark?' with the remark: '+remark:'')+'.',sid,fn3);
    try{_notifyRevision(ss3,sid,'Head of Institution',remark);}catch(e){Logger.log(e.message);}
  }
  return { ok:true };
}

// ─── IMO QUEUE ────────────────────────────────────────────────────────────────
function getIMOQueue() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var facMap=_bmByCol(facD,'FacultyID');
  var hodMap=_bm(hodD,hodD[0]);
  var hoiMap=_bm(hoiD,hoiD[0]);
  var imoMap=_bm(imoD,imoD[0]);
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var hoiH=hoiD[0];
  var hoiStI=hoiH.indexOf('HOI_Status'), hoiSbI=hoiH.indexOf('SubmissionID');
  var out=[];
  for(var i=1;i<hoiD.length;i++){
    var sid=String(hoiD[i][hoiSbI]||'').trim();
    var hoiSt=String(hoiD[i][hoiStI]||'').trim();
    if(!sid||hoiSt!=='Approved') continue;
    var imoR=imoMap[sid]||{}, imoSt=String(imoR['IMO_Status']||'').trim();
    if(imoSt==='Finalised') continue;
    var sub=subMap[sid]||{};
    var fid=String(sub['FacultyID']||'').trim();
    var fac=fid?(facMap[fid]||{}):{};
    var sa=saMap[sid]||{}, hodR=hodMap[sid]||{}, hoiR2=hoiMap[sid]||{};
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],
      {hodStatus:String(hodR['HOD_Status']||''),hodRemark:String(hodR['HOD_Remark']||'')},
      {hoiStatus:hoiSt,hoiRemark:String(hoiR2['HOI_Remark']||'')},
      {imoStatus:imoSt}));
  }
  return out;
}

function submitIMOReview(sid, remark, status) {
  if(!sid)    throw new Error('Submission ID missing.');
  if(!status) throw new Error('Please select a final status.');
  var hoiSt=_getCell(SH.HOI,sid,'HOI_Status');
  if(hoiSt!=='Approved') throw new Error('HOI must approve this submission before IMO can finalise it.');
  _writeReview(SH.IMO,sid,remark||'',status);
  var sub4=_rowByKey(SH.SUBMISSION,sid)||{};
  var fac4=sub4['FacultyID']?(_rowByKey(SH.FACULTY,String(sub4['FacultyID']||''),'FacultyID')||{}):{}; 
  var fn4=String(fac4['FacultyName']||sid);
  var emoji4=status==='Finalised'?'✅':status==='Escalated'?'⚡':'ℹ️';
  var fid4=String(sub4['FacultyID']||'').trim().toUpperCase();
  _pushNotif('FACULTY:'+fid4,'imo_decision',emoji4+' IMO Decision: '+status,
    'Your submission '+sid+' has been marked "'+status+'" by the Institutional Management Office.'+(remark?' Note: '+remark:''),sid,fn4);
  try{ _notifyFinalStatus(SpreadsheetApp.getActiveSpreadsheet(),sid,status,remark); }catch(e){Logger.log(e.message);}
  return { ok:true };
}

// ─── ALL SUBMISSIONS (IMO view) ───────────────────────────────────────────────
function getAllSubmissions() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH=subD[0];
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var facMap=_bm(facD,facD[0],'FacultyID'), hodMap=_bm(hodD,hodD[0]), hoiMap=_bm(hoiD,hoiD[0]), imoMap=_bm(imoD,imoD[0]);
  var out=[];
  for(var i=1;i<subD.length;i++){
    var sid=String(subD[i][subH.indexOf('SubmissionID')]||'');
    if(!sid) continue;
    var facID=String(subD[i][subH.indexOf('FacultyID')]||'');
    var fac=facMap[facID]||{}, hod=hodMap[sid]||{}, hoi=hoiMap[sid]||{}, imo=imoMap[sid]||{};
    out.push({
      submissionID: sid,
      facultyName:  String(fac['FacultyName']||facID),
      facultyID:    facID,
      department:   String(fac['Department']||''),
      institution:  String(fac['Institution']||''),
      campus:       String(fac['Campus']||''),
      designation:  String(fac['Designation']||''),
      semester:     String(subD[i][subH.indexOf('AcademicYearSemester')]||''),
      from:         _fmt(subD[i][subH.indexOf('ReportingFrom')]),
      to:           _fmt(subD[i][subH.indexOf('ReportingTo')]),
      submitted:    _fmtDT(subD[i][subH.indexOf('SubmittedDateTime')]),
      hodStatus:    String(hod['HOD_Status']||'Pending'),
      hoiStatus:    String(hoi['HOI_Status']||'Pending'),
      imoStatus:    String(imo['IMO_Status']||'Pending')
    });
  }
  return out.reverse();
}

// ─── DASHBOARD STATS ─────────────────────────────────────────────────────────
function getDashboardStats() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH=hodD[0];
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH=hoiD[0];
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH=imoD[0];
  var pendHOD=0,pendHOI=0,pendIMO=0,finalised=0,escalated=0;
  var hodApproved=0,hodRevision=0,hodRejected=0;
  var hoiApproved=0,hoiRevision=0,hoiRejected=0;
  for(var i=1;i<hodD.length;i++){
    var hs=String(hodD[i][hodH.indexOf('HOD_Status')]||'');
    if(hs===''||hs==='Needs Revision') pendHOD++;
    if(hs==='Approved') hodApproved++;
    if(hs==='Needs Revision') hodRevision++;
    if(hs==='Rejected') hodRejected++;
  }
  for(var j=1;j<hoiD.length;j++){
    var is=String(hoiD[j][hoiH.indexOf('HOI_Status')]||'');
    if(is===''||is==='Needs Revision') pendHOI++;
    if(is==='Approved') hoiApproved++;
    if(is==='Needs Revision') hoiRevision++;
    if(is==='Rejected') hoiRejected++;
  }
  for(var k=1;k<imoD.length;k++){
    var ms=String(imoD[k][imoH.indexOf('IMO_Status')]||'');
    if(ms===''||ms==='Under Review') pendIMO++;
    if(ms==='Finalised') finalised++;
    if(ms==='Escalated') escalated++;
  }
  return {
    totalFaculty:Math.max(0,facD.length-1), totalSubmissions:Math.max(0,subD.length-1),
    pendingHOD:pendHOD, pendingHOI:pendHOI, pendingIMO:pendIMO,
    finalised:finalised, escalated:escalated,
    hodApproved:hodApproved, hodRevision:hodRevision, hodRejected:hodRejected,
    hoiApproved:hoiApproved, hoiRevision:hoiRevision, hoiRejected:hoiRejected
  };
}


// ─── IN-APP NOTIFICATIONS ─────────────────────────────────────────────────────
// getNotifications(role)              — newest notifications for a role
// markNotifRead(notifID)              — mark one as read
// markAllRead(role)                   — mark all as read for a role
// _pushNotif(role,type,title,body,sid,facultyName) — internal writer

function getNotifications(role, facultyID) {
  if (!role) throw new Error('Role required.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), h = data[0];
  // For faculty, match the specific FACULTY:ID key; for staff, match the role directly
  var matchKey = (role === 'FACULTY' && facultyID)
    ? 'FACULTY:' + String(facultyID).trim().toUpperCase()
    : role;
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[h.indexOf('ForRole')] || '') !== matchKey) continue;
    out.push({
      notifID:     String(row[h.indexOf('NotifID')]      || ''),
      type:        String(row[h.indexOf('Type')]         || ''),
      title:       String(row[h.indexOf('Title')]        || ''),
      body:        String(row[h.indexOf('Body')]         || ''),
      submissionID:String(row[h.indexOf('SubmissionID')] || ''),
      facultyName: String(row[h.indexOf('FacultyName')]  || ''),
      isRead:      String(row[h.indexOf('IsRead')]       || '') === 'YES',
      createdAt:   _fmtDT(row[h.indexOf('CreatedAt')])
    });
  }
  return out.reverse();
}

function markNotifRead(notifID) {
  if (!notifID) return { ok: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return { ok: false };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var kI = h.indexOf('NotifID'), rI = h.indexOf('IsRead');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][kI]) === String(notifID)) {
      sheet.getRange(i + 1, rI + 1).setValue('YES');
      return { ok: true };
    }
  }
  return { ok: false };
}

function markAllRead(role) {
  if (!role) return { ok: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return { ok: false };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var frI = h.indexOf('ForRole'), rI = h.indexOf('IsRead');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][frI]) === role && String(data[i][rI]) !== 'YES') {
      sheet.getRange(i + 1, rI + 1).setValue('YES');
    }
  }
  return { ok: true };
}

function _pushNotif(forRole, type, title, body, sid, facultyName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH.NOTIF);
    // Auto-create Notifications sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SH.NOTIF);
      var hdrs = SCHEMA[SH.NOTIF];
      var hr = sheet.getRange(1,1,1,hdrs.length);
      hr.setValues([hdrs]);
      hr.setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    var nid = 'N-' + new Date().getTime() + '-' + Math.random().toString(36).slice(2,5).toUpperCase();
    sheet.appendRow([nid, forRole, type, title, body, sid || '', facultyName || '', 'NO', new Date()]);
  } catch(e) { Logger.log('_pushNotif failed: ' + e.message); }
}

// ─── FRIDAY REMINDERS ────────────────────────────────────────────────────────
function sendFridayReminders() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues(),fh=facD[0];
  var eI=fh.indexOf('Email'),idI=fh.indexOf('FacultyID'),nI=fh.indexOf('FacultyName');
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(),sh=subD[0];
  var sfI=sh.indexOf('FacultyID'),sdI=sh.indexOf('ReportingFrom');
  var today=new Date(),weekAgo=new Date(today); weekAgo.setDate(today.getDate()-6);
  var done={};
  for(var i=1;i<subD.length;i++){ if(new Date(subD[i][sdI])>=weekAgo) done[String(subD[i][sfI])]=true; }
  var dateStr=Utilities.formatDate(today,Session.getScriptTimeZone(),'dd-MMM-yyyy');
  var sent=0;
  for(var j=1;j<facD.length;j++){
    var email=String(facD[j][eI]||'').trim(), facID=String(facD[j][idI]||'').trim();
    if(!email||!facID||done[facID]) continue;
    try{
      MailApp.sendEmail({to:email,subject:'VMRF: Weekly Performance Submission Due Today',
        body:'Dear '+facD[j][nI]+',\n\nYour Weekly Performance Submission for the week ending '+dateStr+' is due TODAY by 3:00 PM.\n\nPlease open the VMRF Faculty Performance System and submit your report.\n\n— Institutional Management Office, VMRF-DU'});
      sent++;
    }catch(e){Logger.log('Email failed for '+email+': '+e.message);}
  }
  return { sent:sent };
}

// ─── EMAIL HELPERS ────────────────────────────────────────────────────────────
function _notifyHOD(ss,sid,facultyID) {
  var fac=_rowByKey(SH.FACULTY,facultyID,'FacultyID'); if(!fac) return;
  var to=_prop('HOD_'+String(fac['Department']||'').replace(/[^a-zA-Z0-9]/g,'_'))||_prop('HOD_DEFAULT');
  if(!to) return;
  MailApp.sendEmail({to:to,subject:'[VMRF] New Submission Pending HOD Review — '+fac['FacultyName']+' ('+sid+')',
    body:'A new faculty weekly report is pending your review.\n\nFaculty: '+fac['FacultyName']+'\nDepartment: '+fac['Department']+'\nSubmission ID: '+sid+'\n\nPlease login to the VMRF Faculty Performance System to review.\n\n— IMO, VMRF-DU'});
}
function _notifyHOI(ss,sid) {
  var to=_prop('HOI_DEFAULT'); if(!to) return;
  MailApp.sendEmail({to:to,subject:'[VMRF] Submission Approved by HOD — HOI Review Required ('+sid+')',
    body:'Submission '+sid+' has been approved by the Head of Department and is now pending your review.\n\nPlease login to the VMRF Faculty Performance System.\n\n— IMO, VMRF-DU'});
}
function _notifyIMO(ss,sid) {
  var to=_prop('IMO_EMAIL'); if(!to) return;
  MailApp.sendEmail({to:to,subject:'[VMRF] Submission Ready for IMO Monitoring ('+sid+')',
    body:'Submission '+sid+' has been approved by both HOD and HOI. Ready for final IMO monitoring.\n\n— Automated, VMRF-DU'});
}
function _notifyRevision(ss,sid,reviewer,remark) {
  var sub=_rowByKey(SH.SUBMISSION,sid); if(!sub) return;
  var fac=_rowByKey(SH.FACULTY,String(sub['FacultyID']||''),'FacultyID'); if(!fac||!fac['Email']) return;
  MailApp.sendEmail({to:String(fac['Email']),subject:'[VMRF] Your Submission Requires Revision ('+sid+')',
    body:'Dear '+fac['FacultyName']+',\n\nYour submission '+sid+' has been returned for revision by your '+reviewer+'.\n\nRemark:\n'+remark+'\n\nPlease make the necessary changes and resubmit.\n\n— IMO, VMRF-DU'});
}
function _notifyFinalStatus(ss,sid,status,remark) {
  var sub=_rowByKey(SH.SUBMISSION,sid); if(!sub) return;
  var fac=_rowByKey(SH.FACULTY,String(sub['FacultyID']||''),'FacultyID'); if(!fac||!fac['Email']) return;
  MailApp.sendEmail({to:String(fac['Email']),subject:'[VMRF] Submission '+status+' by IMO ('+sid+')',
    body:'Dear '+fac['FacultyName']+',\n\nYour submission '+sid+' has been marked "'+status+'" by the Institutional Management Office.\n\n'+(remark?'IMO Note:\n'+remark+'\n\n':'')+' — IMO, VMRF-DU'});
}

// ─── PRIVATE UTILITIES ────────────────────────────────────────────────────────
function _uid(){
  return 'SUB-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmm')+'-'+Math.random().toString(36).slice(2,6).toUpperCase();
}
function _prop(k){ return PropertiesService.getScriptProperties().getProperty(k)||''; }
function _fmt(d){ try{ return d?Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd-MMM-yyyy'):''; }catch(e){ return ''; } }
function _fmtDT(d){ try{ return d?Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm'):''; }catch(e){ return ''; } }

// Build a lookup map from a 2D values array keyed by keyCol (default SubmissionID)
// Build a lookup map keyed by any named column — always uses header names, never positions
function _bmByCol(data, keyCol) {
  var h=data[0], kI=h.indexOf(keyCol), map={};
  if(kI<0) return map;
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'').trim();
    if(k){ var o={}; h.forEach(function(c,j){if(c)o[String(c)]=data[i][j];}); map[k]=o; }
  }
  return map;
}

function _bm(data, headers, keyCol) {
  keyCol = keyCol || 'SubmissionID';
  var kI = headers.indexOf(keyCol), map = {};
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'');
    if(k){ var o={}; headers.forEach(function(c,j){o[c]=data[i][j];}); map[k]=o; }
  }
  return map;
}

// Build a multi-row lookup map (one key → array of row objects)
function _bmMulti(data, headers) {
  var kI=headers.indexOf('SubmissionID'), map={};
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'');
    if(!k) continue;
    var o={}; headers.forEach(function(c,j){o[c]=data[i][j];});
    if(!map[k]) map[k]=[];
    map[k].push(o);
  }
  return map;
}

function _rowByKey(sheetName, keyVal, keyCol) {
  keyCol = keyCol || 'SubmissionID';
  var data=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  var h=data[0], kI=h.indexOf(keyCol);
  for(var i=1;i<data.length;i++){
    if(String(data[i][kI])===String(keyVal)){ var o={}; h.forEach(function(c,j){o[c]=data[i][j];}); return o; }
  }
  return null;
}

function _getCell(sheetName, sid, col) {
  var r=_rowByKey(sheetName,sid); return r ? String(r[col]||'') : '';
}

function _writeReview(sheetName, sid, remark, status) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data=sheet.getDataRange().getValues(), h=data[0], kI=h.indexOf('SubmissionID');
  for(var i=1;i<data.length;i++){
    if(String(data[i][kI])===String(sid)){
      sheet.getRange(i+1,2,1,3).setValues([[remark,status,new Date()]]);
      return;
    }
  }
  sheet.appendRow([sid,remark,status,new Date()]);
}

// Builds a standardized submission item object
function _buildItem(sid,sub,fac,sa,tsRows,hodInfo,hoiInfo,imoInfo) {
  return {
    submissionID: sid,
    facultyName:  String(fac['FacultyName']||sid),
    facultyID:    String(fac['FacultyID']||''),
    department:   String(fac['Department']||''),
    institution:  String(fac['Institution']||''),
    campus:       String(fac['Campus']||''),
    designation:  String(fac['Designation']||''),
    email:        String(fac['Email']||''),
    semester:     String(sub['AcademicYearSemester']||''),
    from:         _fmt(sub['ReportingFrom']),
    to:           _fmt(sub['ReportingTo']),
    submitted:    _fmtDT(sub['SubmittedDateTime']),
    outcome:      String(sa['OutcomeOfWeek']||''),
    target:       String(sa['TargetPlanNextWeek']||''),
    timesheet:    tsRows,
    hodStatus:    String((hodInfo&&hodInfo.hodStatus)||''),
    hodRemark:    String((hodInfo&&hodInfo.hodRemark)||''),
    hoiStatus:    String((hoiInfo&&hoiInfo.hoiStatus)||''),
    hoiRemark:    String((hoiInfo&&hoiInfo.hoiRemark)||''),
    imoStatus:    String((imoInfo&&imoInfo.imoStatus)||'')
  };
}
