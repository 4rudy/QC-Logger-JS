let fullSheet = SpreadsheetApp.getActiveSpreadsheet()
let curSheet = fullSheet.getActiveSheet();

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
    .addToUi();

  changeLogMode("log")
}

function onEdit(e) {
  if (!e || !e.range) return;

  let curCell = e.range
  curSheet = curCell.getSheet() || fullSheet.getActiveSheet();

  if (isNull(e.value)) {
    deleteUnit(curCell, e.oldValue);
  } else if (curCell.getColumn() == COLS.pid.num) {
    initializeUnit(curCell);
  } else if (curCell.getColumn() == COLS.iRun.num) {
    failUnit(curCell)
  } else if (curCell.getColumn() == COLS.status.num) {
    lastUpdated(curCell)
  } else if (curCell.getColumn() == COLS.failPoint.num) {
    showFailures(curCell)
  } else if (curCell.getColumn() == COLS.fRun.num) {
    shelveUnit(curCell)
  }
}
