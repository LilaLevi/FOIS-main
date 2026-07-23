function createSalarySystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const required = [
    'README',
    'Settings',
    'Driver Template',
    'Salary Database',
    'Salary Report',
    'Archive Log'
  ];

  required.forEach(name => {
    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    ss.insertSheet(name);
  });

  buildReadme_(ss.getSheetByName('README'));
  buildSettings_(ss.getSheetByName('Settings'));
  buildDriverTemplate_(ss.getSheetByName('Driver Template'));
  buildDatabase_(ss.getSheetByName('Salary Database'));
  buildReport_(ss.getSheetByName('Salary Report'));
  buildArchiveLog_(ss.getSheetByName('Archive Log'));

  ss.setActiveSheet(ss.getSheetByName('README'));
}

function buildReadme_(sh) {
  sh.getRange('A1:H1').merge().setValue('Salary System — новая архитектура');
  styleTitle_(sh.getRange('A1:H1'));
  sh.getRange('A3:B9').setValues([
    ['Лист', 'Назначение'],
    ['Settings', 'Справочник водителей, машин, ставок и типов оплаты.'],
    ['Driver Template', 'Шаблон расчёта одного водителя за одну неделю.'],
    ['Salary Database', 'Одна строка = один водитель за одну неделю.'],
    ['Salary Report', 'Автоматический отчёт по выбранной неделе.'],
    ['Archive Log', 'Контроль переноса завершённых недель в базу.'],
    ['Record Key', 'Формат: 2026-W30|709']
  ]);
  styleHeader_(sh.getRange('A3:B3'));
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 430);
}

function buildSettings_(sh) {
  sh.getRange('A1:H1').merge().setValue('Settings — Driver Directory');
  styleTitle_(sh.getRange('A1:H1'));
  sh.getRange('A3:H3').setValues([[
    'Driver ID','Driver Name','Truck','Pay Type','Rate','Dispatcher','Recruiter','Status'
  ]]);
  styleHeader_(sh.getRange('A3:H3'));
  sh.getRange('D4:D200').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['CPM','% Gross'], true).build()
  );
  sh.getRange('H4:H200').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive'], true).build()
  );
  sh.getRange('E4:E200').setNumberFormat('0.00%');
  sh.setFrozenRows(3);
}

function buildDriverTemplate_(sh) {
  sh.getRange('A1:H1').merge().setValue('Driver Payroll Calculation');
  styleTitle_(sh.getRange('A1:H1'));

  sh.getRange('A3:B12').setValues([
    ['Field','Value'],
    ['Year',2026],
    ['Week',30],
    ['Week Key',''],
    ['Driver ID',''],
    ['Driver Name',''],
    ['Truck',''],
    ['Pay Type',''],
    ['Rate',''],
    ['Record Key','']
  ]);
  styleHeader_(sh.getRange('A3:B3'));

  sh.getRange('B6').setFormula('=TEXT(B4,"0")&"-W"&TEXT(B5,"00")');
  sh.getRange('B8').setFormula('=IFERROR(XLOOKUP(B7,Settings!A:A,Settings!B:B,"Unknown"),"Unknown")');
  sh.getRange('B9').setFormula('=IFERROR(XLOOKUP(B7,Settings!A:A,Settings!C:C,""),"")');
  sh.getRange('B10').setFormula('=IFERROR(XLOOKUP(B7,Settings!A:A,Settings!D:D,""),"")');
  sh.getRange('B11').setFormula('=IFERROR(XLOOKUP(B7,Settings!A:A,Settings!E:E,""),"")');
  sh.getRange('B12').setFormula('=B6&"|"&B7');
  sh.getRange('B11').setNumberFormat('0.00%');

  sh.getRange('A15:H15').setValues([[
    'Date','Load ID','Origin','Destination','Loaded Miles','Empty Miles','Gross','Adjustment'
  ]]);
  styleHeader_(sh.getRange('A15:H15'));
  sh.getRange('A16:A60').setNumberFormat('mm/dd/yyyy');
  sh.getRange('G16:H60').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  sh.getRange('A63:B70').setValues([
    ['Summary','Amount'],
    ['Total Loaded Miles',''],
    ['Total Empty Miles',''],
    ['Total Miles',''],
    ['Gross',''],
    ['Adjusted Gross',''],
    ['Driver Pay',''],
    ['Deductions',0]
  ]);
  styleHeader_(sh.getRange('A63:B63'));
  sh.getRange('B64').setFormula('=SUM(E16:E60)');
  sh.getRange('B65').setFormula('=SUM(F16:F60)');
  sh.getRange('B66').setFormula('=B64+B65');
  sh.getRange('B67').setFormula('=SUM(G16:G60)');
  sh.getRange('B68').setFormula('=SUM(G16:G60)+SUM(H16:H60)');
  sh.getRange('B69').setFormula('=IF(B10="CPM",B66*B11,IF(B10="% Gross",B68*B11,0))');
  sh.getRange('B70').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  sh.getRange('A72:B73').setValues([
    ['Net Pay',''],
    ['Status','Open']
  ]);
  sh.getRange('B72').setFormula('=B69-B70');
  sh.getRange('B72').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  sh.getRange('B73').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Open','Ready','Archived'], true).build()
  );

  const headers = [[
    'Year','Week','Week Key','Driver ID','Driver Name','Truck','Pay Type','Rate',
    'Loaded Miles','Empty Miles','Total Miles','Gross','Adjusted Gross','Driver Pay',
    'Deductions','Net Pay','Record Key','Status'
  ]];
  sh.getRange('D63:U63').setValues(headers);
  styleHeader_(sh.getRange('D63:U63'));
  sh.getRange('D64:U64').setFormulas([[
    '=B4','=B5','=B6','=B7','=B8','=B9','=B10','=B11',
    '=B64','=B65','=B66','=B67','=B68','=B69','=B70','=B72','=B12','=B73'
  ]]);
  sh.setFrozenRows(15);
}

function buildDatabase_(sh) {
  sh.getRange('A1:R1').merge().setValue('Salary Database — one row per driver per week');
  styleTitle_(sh.getRange('A1:R1'));
  sh.getRange('A3:R3').setValues([[
    'Year','Week','Week Key','Driver ID','Driver Name','Truck','Pay Type','Rate',
    'Loaded Miles','Empty Miles','Total Miles','Gross','Adjusted Gross','Driver Pay',
    'Deductions','Net Pay','Record Key','Status'
  ]]);
  styleHeader_(sh.getRange('A3:R3'));
  sh.getRange('H4:H1000').setNumberFormat('0.00%');
  sh.getRange('L4:P1000').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  sh.setFrozenRows(3);
}

function buildReport_(sh) {
  sh.getRange('A1:I1').merge().setValue('Weekly Salary Report');
  styleTitle_(sh.getRange('A1:I1'));
  sh.getRange('A3:B5').setValues([
    ['Parameter','Value'],
    ['Week Key','2026-W30'],
    ['Rows found','']
  ]);
  styleHeader_(sh.getRange('A3:B3'));
  sh.getRange('B5').setFormula('=COUNTIF(\'Salary Database\'!C:C,B4)');
  sh.getRange('A8:I8').setValues([[
    'Driver ID','Driver Name','Truck','Total Miles','Adjusted Gross',
    'Driver Pay','Deductions','Net Pay','Status'
  ]]);
  styleHeader_(sh.getRange('A8:I8'));
  sh.getRange('A9').setFormula(
    '=IFERROR(FILTER({\'Salary Database\'!D4:F,\'Salary Database\'!K4:K,\'Salary Database\'!M4:P,\'Salary Database\'!R4:R},\'Salary Database\'!C4:C=$B$4),"No records")'
  );
  sh.setFrozenRows(8);
}

function buildArchiveLog_(sh) {
  sh.getRange('A1:F1').merge().setValue('Archive Log');
  styleTitle_(sh.getRange('A1:F1'));
  sh.getRange('A3:F3').setValues([[
    'Archived At','Week Key','Driver ID','Record Key','Archived By','Notes'
  ]]);
  styleHeader_(sh.getRange('A3:F3'));
  sh.getRange('A4:A1000').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setFrozenRows(3);
}

function archiveActiveDriverWeek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();
  if (!source.getName().startsWith('D_')) {
    throw new Error('Откройте водительский лист с названием D_<Driver ID>.');
  }

  const status = source.getRange('B73').getValue();
  if (status !== 'Ready') {
    throw new Error('Перед архивированием установите Status = Ready.');
  }

  const row = source.getRange('D64:U64').getValues()[0];
  const recordKey = row[16];
  if (!recordKey) throw new Error('Record Key пустой.');

  const db = ss.getSheetByName('Salary Database');
  const existing = db.getRange('Q4:Q').getValues().flat();
  if (existing.includes(recordKey)) {
    throw new Error('Запись с Record Key ' + recordKey + ' уже существует.');
  }

  db.appendRow(row);
  source.getRange('B73').setValue('Archived');

  const log = ss.getSheetByName('Archive Log');
  log.appendRow([
    new Date(),
    row[2],
    row[3],
    recordKey,
    Session.getActiveUser().getEmail(),
    'Archived from ' + source.getName()
  ]);
}

function createDriverSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('New Driver Sheet', 'Введите Driver ID:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const driverId = result.getResponseText().trim();
  if (!driverId) return;

  const name = 'D_' + driverId;
  if (ss.getSheetByName(name)) throw new Error('Лист ' + name + ' уже существует.');

  const template = ss.getSheetByName('Driver Template');
  const copy = template.copyTo(ss).setName(name);
  copy.getRange('B7').setValue(driverId);
  ss.setActiveSheet(copy);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Salary System')
    .addItem('Create driver sheet', 'createDriverSheet')
    .addItem('Archive active driver week', 'archiveActiveDriverWeek')
    .addToUi();
}

function styleTitle_(range) {
  range
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(15)
    .setVerticalAlignment('middle');
}

function styleHeader_(range) {
  range
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}
