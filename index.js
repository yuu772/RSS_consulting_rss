// consts
var ssRss = SpreadsheetApp.openById(
  '12mf_TSPbZmNwxtN0UkQQ-Tcq7-C8H63WlUTCLdIvG_s'
);
var ssStudent = SpreadsheetApp.openById(
  '1DlLsKS5E37rczxCixJ1KZcyj2DuHd3CTP9CbXaP_Td0'
);
// var ssData = SpreadsheetApp.openById(
//   ''
// );

const INDUSTRIES = [
  'GD',
  'メーカー',
  'IT・コンサル',
  '建設',
  '小売・流通・印刷',
  '金融',
  'I人材・教育・サービス業・エンタメT・コンサル',
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
const CHECK_RANGE = {
  from: {
    rowIdx: 5,
    colIdx: 4,
  },
};
const OFFSET_ROW = 2;
const OFFSET_COL = 2;

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
  ScriptApp.newTrigger('onEditStudent').forSpreadsheet(ssStudent).onEdit().create();
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

  // define a set of LABELS_INPUT_ROW_** as section
  const sectionNum =
    Math.floor(
      (rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) /
        LABELS_INPUT_ROW_RSS.length
    ) + 1;
  const initRowIdxRss =
    (sectionNum - 1) * LABELS_INPUT_ROW_RSS.length +
    LABELS_TT_ROW.length +
    OFFSET_ROW +
    1;
  const initRowIdxStudent =
    (sectionNum - 1) * LABELS_INPUT_ROW_STUDENT.length +
    LABELS_TT_ROW.length +
    OFFSET_ROW +
    1;
  if (isAllValuesSet(sheetRss, initRowIdxRss, colIdx)) {
    // for "担当RSSの詳しい業界"
    const outputRowIdx =
      initRowIdxStudent + LABELS_INPUT_ROW_STUDENT.length - 1;
    const value = sheetRss
      .getRange(initRowIdxRss + LABELS_INPUT_ROW_RSS.length - 1, colIdx)
      .getValues();
    sheetStudent.getRange(outputRowIdx, colIdx).setValues(value);
    sheetStudent
      .getRange(initRowIdxStudent, colIdx, 2, 1)
      .setFontColor("black")
      .setBackground(LABEL_COLOR_BASE);
    sheetStudent
      .getRange(initRowIdxStudent + 2, colIdx, 2, 1)
      .setFontColor("black")
      .setBackground(COLOR_AVAILABLE);
  } else {
    sheetStudent
      .getRange(initRowIdxStudent, colIdx, LABELS_INPUT_ROW_STUDENT.length, 1)
      .setFontColor(COLOR_UNAVAILABLE)
      .setBackground(COLOR_UNAVAILABLE);
  }
}

function onEditStudent(){
  const range = e.range;
  const rowIdx = range.getRow();
  const colIdx = range.getColumn();
  if (rowIdx < CHECK_RANGE.from.rowIdx || colIdx < CHECK_RANGE.from.colIdx)
    return;

  const sheetRss = e.source.getActiveSheet();
  const sheetStudent = ssStudent.getSheetByName(sheetRss.getSheetName());

  // define a set of LABELS_INPUT_ROW_** as section
  const sectionNum =
    Math.floor(
      (rowIdx - (LABELS_TT_ROW.length + OFFSET_ROW) - 1) /
        LABELS_INPUT_ROW_RSS.length
    ) + 1;
  const initRowIdxRss =
    (sectionNum - 1) * LABELS_INPUT_ROW_RSS.length +
    LABELS_TT_ROW.length +
    OFFSET_ROW +
    1;
  const initRowIdxStudent =
    (sectionNum - 1) * LABELS_INPUT_ROW_STUDENT.length +
    LABELS_TT_ROW.length +
    OFFSET_ROW +
    1;
  if (isAllValuesSet(sheetRss, initRowIdxRss, colIdx)) {
    // for "担当RSSの詳しい業界"
    const outputRowIdx =
      initRowIdxStudent + LABELS_INPUT_ROW_STUDENT.length - 1;
    const value = sheetRss
      .getRange(initRowIdxRss + LABELS_INPUT_ROW_RSS.length - 1, colIdx)
      .getValues();
    sheetStudent
      .getRange(initRowIdxStudent, colIdx, 2, 1)
      .setFontColor("black")
      .setBackground(LABEL_COLOR_BASE);
    sheetStudent
      .getRange(initRowIdxStudent + 2, colIdx, 2, 1)
      .setFontColor("black")
      .setBackground(COLOR_AVAILABLE);
    sheetStudent.getRange(outputRowIdx, colIdx).setValues(value);
  } else {
    sheetStudent
      .getRange(initRowIdxStudent, colIdx, LABELS_INPUT_ROW_STUDENT.length, 1)
      .setFontColor(COLOR_UNAVAILABLE)
      .setBackground(COLOR_UNAVAILABLE);
  } 
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


// create sheets
function createSheets() {
  recreateSheets();
  const sheetsRss = ssRss.getSheets();
  sheetsRss.forEach((sheet) => setSheetRss(sheet));
  const sheetsStudent = ssStudent.getSheets();
  sheetsStudent.forEach((sheet) => setSheetStudent(sheet));
}

function recreateSheets() {
  const sheetNamesRss = ssRss.getSheets().map((s) => s.getSheetName());
  const sheetNamesStudent = ssStudent.getSheets().map((s) => s.getSheetName());
  INDUSTRIES.forEach((name) => {
    //delete sheets
    if (sheetNamesRss.some((s) => s === name))
      ssRss.deleteSheet(ssRss.getSheetByName(name));
    if (sheetNamesStudent.some((s) => s === name))
      ssStudent.deleteSheet(ssStudent.getSheetByName(name));
    //recreate sheets
    ssRss.insertSheet(name);
    ssStudent.insertSheet(name);
  });
}

function setSheetCommon(sheet, initColumn, lastColumn) {
  TIMETABLE.forEach((t) => {
    sheet
      .getRange(OFFSET_ROW + 1, lastColumn, 1, MAX_SESSION_NUM)
      .merge()
      .setValue(t)
      .setHorizontalAlignment('center');
    for (let i = 0; i < MAX_SESSION_NUM; i++) {
      sheet
        .getRange(OFFSET_ROW + 2, lastColumn + i)
        .setValue(i + 1)
        .setHorizontalAlignment('center');
    }
    lastColumn += MAX_SESSION_NUM;
  });
  sheet
    .getRange(
      OFFSET_ROW + LABELS_TT_ROW.length,
      initColumn,
      1,
      lastColumn - MAX_SESSION_NUM
    )
    .setBackground(LABEL_COLOR_SESSION);
  sheet
    .getRange(
      OFFSET_ROW + LABELS_TT_ROW.length + 1,
      initColumn,
      2,
      lastColumn - MAX_SESSION_NUM
    )
    .setBackground(LABEL_COLOR_BASE);
  return lastColumn;
}

function setSheetRss(sheet) {
  const initColumn = OFFSET_COL + 1;
  let lastRow = OFFSET_ROW + 1;
  let lastColumn = initColumn + 1;

  // set labels virtically
  const ttLabelRowLen = LABELS_TT_ROW.length;
  sheet
    .getRange(lastRow, initColumn, ttLabelRowLen)
    .setValues(LABELS_TT_ROW)
    .setHorizontalAlignment('center');
  lastRow += ttLabelRowLen;

  const inputLabelRowLen = LABELS_INPUT_ROW_RSS.length;
  sheet
    .getRange(lastRow, initColumn, inputLabelRowLen)
    .setValues(LABELS_INPUT_ROW_RSS)
    .setHorizontalAlignment('center');
  lastRow += inputLabelRowLen;

  setSheetCommon(sheet, initColumn, lastColumn);
}

function setSheetStudent(sheet) {
  const initColumn = OFFSET_COL + 1;
  let lastRow = OFFSET_ROW + 1;
  let lastColumn = initColumn + 1;

  // set labels virtically
  sheet.setColumnWidth(initColumn, 300);

  const ttLabelRowLen = LABELS_TT_ROW.length;
  sheet
    .getRange(lastRow, initColumn, ttLabelRowLen)
    .setValues(LABELS_TT_ROW)
    .setHorizontalAlignment('center');
  lastRow += ttLabelRowLen;

  const inputLabelRowLen = LABELS_INPUT_ROW_STUDENT.length;
  sheet
    .getRange(lastRow, initColumn, inputLabelRowLen)
    .setValues(LABELS_INPUT_ROW_STUDENT)
    .setHorizontalAlignment('center');
  lastRow += inputLabelRowLen;

  lastColumn = setSheetCommon(sheet, initColumn, lastColumn);
  // LABELS_INPUT_ROW_STUDENT.forEach((_, idx) => {
  //   const cell = sheet.getRange(
  //     OFFSET_ROW + LABELS_TT_ROW.length + idx + 1,
  //     initColumn + 1,
  //     1,
  //     lastColumn - MAX_SESSION_NUM
  //   );
  //   cell.setBackground(COLOR_UNAVAILABLE);
  //   if (idx === 2) cell.setDataValidation(CONSULTING_CONTENT_RULE);
  // });
}
