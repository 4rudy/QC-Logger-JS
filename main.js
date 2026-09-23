let fullSheet = SpreadsheetApp.getActiveSpreadsheet()
let curSheet = fullSheet.getActiveSheet();
let allUnits = []

const COLS = {
  date: { letter: "A", num: 1 },
  pid: { letter: "B", num: 2 },
  iRun: { letter: "C", num: 3 },
  failPoint: { letter: "D", num: 4 },
  failType: { letter: "E", num: 5 },
  fRun: { letter: "H", num: 8 },
  status: { letter: "I", num: 9 },
  updated: { letter: "J", num: 10 },
  qcMode: { letter: "L", num: 12 },
}

function onOpen() {
  sheetUI.createMenu("Logger")
    // .addItem("SCAN PID", "openGUI")
    // .addSeparator()
    .addItem("QC MODE", "qcLogMode")
    .addItem("PREP MODE", "prepLogMode")
    .addItem("FIND ALL DUPES", "findAllDupes")
    .addItem("CLEAR ALL NOTES", "clearAllNotes")
    .addToUi();

  changeLogMode("log")
}

function onEdit(e) {
  if (!e || !e.range) return;

  let curCell = e.range
  let curColumn = curCell.getColumn()
  curSheet = curCell.getSheet() || fullSheet.getActiveSheet();
  allUnits = getDataArray()

  if (isNull(e.value)) {
    return deleteUnit(curCell, e.oldValue);
  }

  switch (curColumn) {
    case COLS.pid.num:
      return initializeUnit(curCell);
    case COLS.iRun.num:
      return failUnit(curCell)
    case COLS.status.num:
      return lastUpdated(curCell)
    case COLS.failPoint.num:
      return showFailures(curCell)
    case COLS.fRun.num:
      return shelveUnit(curCell)
  }
}
