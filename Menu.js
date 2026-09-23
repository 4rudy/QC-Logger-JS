const sheetUI = SpreadsheetApp.getUi()
const scProperties = PropertiesService.getScriptProperties()
let curLogMode = scProperties.getProperty("logMode")

const qcModes = {
  log: { code: "qcLogMode", name: "QC MODE", tColor: "#F2D0A7", bColor: "#7A6048" },
  prep: { code: "prepLogMode", name: "PREP MODE", tColor: "#D5E5D0", bColor: "#58705A" }
}

function openGUI() {
  let widget = HtmlService.createHtmlOutput("<h1>Sidebar</h1>")
  sheetUI.showSidebar(widget);
}

function notify(msg) {
  SpreadsheetApp.getActive().toast(msg)
}

function changeLogMode(newMode) {
  curSheet = fullSheet.getActiveSheet();
  scProperties.setProperty("logMode", qcModes[newMode].code)

  let qcModeCell = getCell(`${COLS.qcMode.letter}1`)
  qcModeCell.setValue(qcModes[newMode].name)
  qcModeCell.setFontColor(qcModes[newMode].tColor)
  qcModeCell.setBackground(qcModes[newMode].bColor)
}

function qcLogMode() {
  changeLogMode("log")
}

function prepLogMode() {
  changeLogMode("prep")
}

function findAllDupes() {
  allUnits = getDataArray();

  allUnits.forEach(unit => {
    let dupedCells = isUnitDuped(unit.pid)
    if (dupedCells) { dupeUnit(dupedCells) }
  })
}

function clearAllNotes() {
  allUnits = getDataArray();
  allUnits.forEach(unit => {
    let cell = getCell(`${COLS.pid.letter}${unit.row}`)
    cell.clearNote()
  })
}

function allCaps() {
  allUnits = getDataArray();
  allUnits.forEach(unit => {
    let cell = getCell(`${COLS.pid.letter}${unit.row}`)
    cell.setValue(String(unit.pid).toUpperCase())
  })
}
