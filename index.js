// consts
var ssRss = SpreadsheetApp.openById(
  '12mf_TSPbZmNwxtN0UkQQ-Tcq7-C8H63WlUTCLdIvG_s'
);
var ssStudent = SpreadsheetApp.openById(
  '1DlLsKS5E37rczxCixJ1KZcyj2DuHd3CTP9CbXaP_Td0'
);
var ssData = SpreadsheetApp.openById(
  '1Bp5-ux_dqgJKDBFaifyfB-Ux4IEcgDmKkOawMR8OnC8'
);

const INDUSTRIES = [
  'GD',
  'メーカー',
  'IT・コンサル',
  '建設',
  '小売・流通・印刷',
  '金融',
  'IT人材・教育・サービス業・エンタメ・コンサル',
];
const TIMETABLE = [
  '1コマ(9:00-10:30)',
  '2コマ(10:45-12:15)',
  '3コマ(13:05-14:35)',
  '4コマ(14:50-16:20)',
  '5コマ(16:35-18:05)',
];
const MAX_SESSION_NUM = 3;
const LABELS_TT_ROW = [['時間割'], ['セッション']];
const LABELS_INPUT_ROW_RSS = [['名前'], ['学籍番号'], ['詳しい業界']];
const LABELS_INPUT_ROW_STUDENT = [
  ['名前'],
  ['学籍番号'],
  ['相談内容'],
  ['担当RSSの詳しい業界'],
];
const CONSULTING_CONTENT = ['面接対策', '相談', 'ES添削'];
const LABEL_COLOR_BASE = '#d9ead3';
const LABEL_COLOR_SESSION = '#93c47d';
const COLOR_AVAILABLE = '#ffffff';
const COLOR_UNAVAILABLE = '#666666';
const OFFSET_ROW = 2;
const OFFSET_COL = 1;
const CHECK_RANGE = {
  from: {
    rowIdx: OFFSET_ROW + LABELS_TT_ROW.length + 1,
    colIdx: OFFSET_COL + 2,
  },
};
const DATA_COLMUNS = [
  'id',
  'student_id',
  'student_name',
  'rss_id',
  'rss_name',
  'date',
  'period',
  'session_num',
  'created_at',
  'updated_at',
];

const CONSULTING_CONTENT_RULE = SpreadsheetApp.newDataValidation()
  .requireValueInList(CONSULTING_CONTENT)
  .build();

// triggers
function setTriggers() {
  delTriggers();
  setOnEditRSS();
  setOnEditStudent();
}

function setOnEditRSS() {
  ScriptApp.newTrigger('onEditRss').forSpreadsheet(ssRss).onEdit().create();
}

function setOnEditStudent() {
  ScriptApp.newTrigger('onEditStudent')
    .forSpreadsheet(ssStudent)
    .onEdit()
    .create();
}

function delTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

function onEditRss(e) {
  // validate check range
  const range = e.range;
  const rowIdx = range.getRow();
  const colIdx = range.getColumn();
  if (rowIdx < CHECK_RANGE.from.rowIdx || colIdx < CHECK_RANGE.from.colIdx)
    return;

  const sheetRss = e.source.getActiveSheet();
  const sheetStudent = ssStudent.getSheetByName(sheetRss.getSheetName());
  const sheetData = ssData.getSheetByName(sheetRss.getSheetName());

  const sectionNum = calcSectionNum(rowIdx, 'rss');
  const questionNum = calcQuestionNum(rowIdx, 'rss');

  const rowIdxStartFromRss = calcSectionRowStartFrom(sectionNum, 'rss');
  const rowIdxStartFromStudent = calcSectionRowStartFrom(sectionNum, 'student');
  const rowIdxLastData = sheetData.getLastRow();

  const date = sheetRss
    .getRange(rowIdxStartFromRss, 1, LABELS_INPUT_ROW_RSS.length, 1)
    .getValue(); // suppose to be Date object
  const dataId = calcDateId(date, colIdx);

  // find dataId and row idx of data
  const dataIdx = getDataIdx(sheetData, dataId);
  const existDataId = dataIdx !== -1;
  const rowIdxData = existDataId ? dataIdx + 2 : rowIdxLastData + 1;

  // create id if necessry
  if (!existDataId) {
    sheetData.getRange(rowIdxData, 1).setValue(dataId); // id
    sheetData.getRange(rowIdxData, 9).setValue(new Date()); // created_at
  }

  // set data
  let colIdxData = null;
  if (questionNum === 1) colIdxData = 5;
  if (questionNum === 2) colIdxData = 4;
  if (colIdxData) {
    sheetData.getRange(rowIdxData, colIdxData).setValue(e.value);
    sheetData.getRange(rowIdxData, 10).setValue(new Date()); // updated_at
  }

  if (isAllValuesSet(sheetRss, rowIdxStartFromRss, colIdx)) {
    // for "担当RSSの詳しい業界"
    const outputRowIdx =
      rowIdxStartFromStudent + LABELS_INPUT_ROW_STUDENT.length - 1;
    const value = sheetRss
      .getRange(rowIdxStartFromRss + LABELS_INPUT_ROW_RSS.length - 1, colIdx)
      .getValues();
    sheetStudent.getRange(outputRowIdx, colIdx).setValues(value);
    sheetStudent
      .getRange(rowIdxStartFromStudent, colIdx, 2, 1)
      .setFontColor('black')
      .setBackground(LABEL_COLOR_BASE);
    sheetStudent
      .getRange(rowIdxStartFromStudent + 2, colIdx, 2, 1)
      .setFontColor('black')
      .setBackground(COLOR_AVAILABLE);
  } else {
    sheetStudent
      .getRange(
        rowIdxStartFromStudent,
        colIdx,
        LABELS_INPUT_ROW_STUDENT.length,
        1
      )
      .setFontColor(COLOR_UNAVAILABLE)
      .setBackground(COLOR_UNAVAILABLE);
  }
}

function onEditStudent(e) {
  const range = e.range;
  const rowIdx = range.getRow();
  const colIdx = range.getColumn();
  if (rowIdx < CHECK_RANGE.from.rowIdx || colIdx < CHECK_RANGE.from.colIdx)
    return;

  const sheetStudent = e.source.getActiveSheet();
  const sheetData = ssData.getSheetByName(sheetStudent.getSheetName());

  const sectionNum = calcSectionNum(rowIdx, 'student');
  const questionNum = calcQuestionNum(rowIdx, 'student');
  const rowIdxStartFromStudent = calcSectionRowStartFrom(sectionNum, 'student');
  const rowIdxStudentId = rowIdxStartFromStudent + 1;
  const rowIdxLastData = sheetData.getLastRow();

  const date = sheetStudent
    .getRange(rowIdxStartFromStudent, 1, LABELS_INPUT_ROW_STUDENT.length, 1)
    .getValue(); // suppose to be Date object
  const dataId = calcDateId(date, colIdx);

  // find dataId and row idx of data
  const dataIdx = getDataIdx(sheetData, dataId);
  const existDataId = dataIdx !== -1;
  const rowIdxData = existDataId ? dataIdx + 2 : rowIdxLastData + 1;

  // create id if necessry
  if (!existDataId) {
    sheetData.getRange(rowIdxData, 1).setValue(dataId); // id
    sheetData.getRange(rowIdxData, 9).setValue(new Date()); // created_at
  }

  // set data
  let colIdxData = null;
  console.log(questionNum);
  if (questionNum === 1) colIdxData = 3;
  if (questionNum === 2) colIdxData = 2;
  if (!colIdxData) return;

  const period = calcPeriod(colIdx);
  const session = calcSession(colIdx);
  sheetData.getRange(rowIdxData, colIdxData).setValue(e.value);
  sheetData.getRange(rowIdxData, 6).setValue(date);
  sheetData.getRange(rowIdxData, 7).setValue(period);
  sheetData.getRange(rowIdxData, 8).setValue(session);
  sheetData.getRange(rowIdxData, 10).setValue(new Date()); // updated_at

  // hide student id
  if (rowIdx === rowIdxStudentId)
    sheetStudent.getRange(rowIdx, colIdx).setValue('*******');
}

function isAllValuesSet(sheet, initRowIdx, initColIdx) {
  let isSet = true;
  const values = sheet
    .getRange(initRowIdx, initColIdx, LABELS_INPUT_ROW_RSS.length, 1)
    .getValues();
  values.forEach((v) => {
    if (v.length === 1 && v[0] === '') isSet = false;
  });
  return isSet;
}

function calcSectionNum(rowIdx, type) {
  // define a set of LABELS_INPUT_ROW_** as section
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return (
    Math.floor((rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) / labelsLen) +
    1
  );
}

function calcQuestionNum(rowIdx, type) {
  // define a set of LABELS_INPUT_ROW_** as section
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return ((rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) % labelsLen) + 1;
}

function calcSectionRowStartFrom(sectionNum, type) {
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  return (sectionNum - 1) * labelsLen + LABELS_TT_ROW.length + OFFSET_ROW + 1;
}

function calcDateId(date, colIdx) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  const period = calcPeriod(colIdx);
  const session = calcSession(colIdx);
  return `${year}${month}${day}_${period}_${session}`;
}

function calcPeriod(colIdx) {
  return Math.floor((colIdx - OFFSET_COL - 2) / MAX_SESSION_NUM) + 1;
}

function calcSession(colIdx) {
  return ((colIdx - OFFSET_COL - 2) % MAX_SESSION_NUM) + 1;
}

function getDataIdx(sheetData, dataId) {
  const rowIdxLastData = sheetData.getLastRow();
  if (rowIdxLastData < 2) return -1;
  const dataIds = sheetData.getRange(2, 1, rowIdxLastData - 1, 1).getValues();
  let concated = [];
  dataIds.forEach((id) => (concated = concated.concat(id)));
  return concated.indexOf(dataId);
}

// create sheets
function createSheets() {
  recreateSheets();
  const sheetsRss = ssRss.getSheets();
  sheetsRss.forEach((sheet) => setSheetCommon(sheet));
  const sheetsStudent = ssStudent.getSheets();
  sheetsStudent.forEach((sheet) => setSheetCommon(sheet));
  const sheetsData = ssData.getSheets();
  sheetsData.forEach((sheet) => setSheetData(sheet));
}

function recreateSheets() {
  const sheetNamesRss = ssRss.getSheets().map((s) => s.getSheetName());
  const sheetNamesStudent = ssStudent.getSheets().map((s) => s.getSheetName());
  const sheetNamesData = ssData.getSheets().map((s) => s.getSheetName());
  INDUSTRIES.forEach((name) => {
    if (sheetNamesRss.some((s) => s === name))
      ssRss.deleteSheet(ssRss.getSheetByName(name));
    ssRss.insertSheet(name);

    if (sheetNamesStudent.some((s) => s === name))
      ssStudent.deleteSheet(ssStudent.getSheetByName(name));
    ssStudent.insertSheet(name);

    if (sheetNamesData.some((s) => s === name)) return;
    ssData.insertSheet(name);
  });
}

const DATE_TO_ADD = {
  year: 2021,
  month: 2,
};
function addBlankData() {
  INDUSTRIES.forEach((name) => {
    const date = new Date(DATE_TO_ADD.year, DATE_TO_ADD.month - 1, 1);
    for (let i = 1; i < 8; i++) {
      const day = date.getDate();
      const dayOfTheWeek = date.getDay();
      if (dayOfTheWeek > 0 && dayOfTheWeek < 6) {
        const sheetRss = ssRss.getSheetByName(name);
        setSheetRss(sheetRss, DATE_TO_ADD.year, DATE_TO_ADD.month, day);
        const sheetStudent = ssStudent.getSheetByName(name);
        setSheetStudent(sheetStudent, DATE_TO_ADD.year, DATE_TO_ADD.month, day);
      }
      date.setDate(day + 1);
      if (date.getMonth() === DATE_TO_ADD.month) break;
    }
  });
}

function setSheetCommon(sheet) {
  const initCol = OFFSET_COL + 1;
  let lastCol = initCol + 1;
  sheet
    .getRange(OFFSET_ROW + 1, initCol, LABELS_TT_ROW.length)
    .setValues(LABELS_TT_ROW)
    .setHorizontalAlignment('center');
  TIMETABLE.forEach((t) => {
    sheet
      .getRange(OFFSET_ROW + 1, lastCol, 1, MAX_SESSION_NUM)
      .merge()
      .setValue(t)
      .setHorizontalAlignment('center');
    for (let i = 0; i < MAX_SESSION_NUM; i++) {
      sheet
        .getRange(OFFSET_ROW + 2, lastCol + i)
        .setValue(i + 1)
        .setHorizontalAlignment('center');
    }
    lastCol += MAX_SESSION_NUM;
  });
  sheet
    .getRange(
      OFFSET_ROW + LABELS_TT_ROW.length,
      initCol,
      1,
      TIMETABLE.length * MAX_SESSION_NUM + 1
    )
    .setBackground(LABEL_COLOR_SESSION);
}

function setSheetRss(sheet, year, month, day) {
  const initRow = sheet.getLastRow() + 1;
  const initCol = OFFSET_COL + 1;

  // set date
  setDate(sheet, initRow, year, month, day, 'rss');

  // set color
  sheet
    .getRange(initRow, initCol, 2, TIMETABLE.length * MAX_SESSION_NUM + 1)
    .setBackground(LABEL_COLOR_BASE);

  // set labels
  sheet
    .getRange(initRow, initCol, LABELS_INPUT_ROW_RSS.length)
    .setValues(LABELS_INPUT_ROW_RSS)
    .setHorizontalAlignment('center');
}

function setSheetStudent(sheet, year, month, day) {
  const initCol = OFFSET_COL + 1;
  const initRow = sheet.getLastRow() + 1;

  // set date
  setDate(sheet, initRow, year, month, day, 'student');

  // set color
  sheet.getRange(initRow, initCol, 2, 1).setBackground(LABEL_COLOR_BASE);
  sheet
    .getRange(
      initRow,
      initCol + 1,
      LABELS_INPUT_ROW_STUDENT.length,
      TIMETABLE.length * MAX_SESSION_NUM
    )
    .setBackground(COLOR_UNAVAILABLE);

  // set labels and default data
  sheet.setColumnWidth(initCol, 300);
  sheet
    .getRange(initRow, initCol, LABELS_INPUT_ROW_STUDENT.length)
    .setValues(LABELS_INPUT_ROW_STUDENT)
    .setHorizontalAlignment('center');
  LABELS_INPUT_ROW_STUDENT.forEach((_, idx) => {
    if (idx !== 2) return;
    const cell = sheet.getRange(
      initRow + idx,
      initCol + 1,
      1,
      TIMETABLE.length * MAX_SESSION_NUM
    );
    cell.setDataValidation(CONSULTING_CONTENT_RULE);
  });
}

function setSheetData(sheet) {
  sheet.getRange(1, 1, 1, DATA_COLMUNS.length).setValues([DATA_COLMUNS]);
}

function setDate(sheet, initRow, year, month, day, type) {
  const labelsLen =
    type === 'rss'
      ? LABELS_INPUT_ROW_RSS.length
      : LABELS_INPUT_ROW_STUDENT.length;
  const date = year + '/' + month + '/' + day;
  sheet
    .getRange(initRow, 1, labelsLen, 1)
    .merge()
    .setValue(date)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}
