const STATUS = {
  cGrade: "C GRADE", pending: "PENDING", surplus: "SURPLUS", shelf: "SHELVED", ready: "READY TO LIST",
}

function initializeUnit(curCell) {
  if (curCell.getColumn() !== COLS.pid.num) return;

  let curRow = curCell.getRow()
  let dateCell = getCell(`${COLS.date.letter}${curRow}`)
  let iRunCell = getCell(`${COLS.iRun.letter}${curRow}`)
  let fRunCell = getCell(`${COLS.fRun.letter}${curRow}`)
  let statusCell = getCell(`${COLS.status.letter}${curRow}`)
  let statVal = ((curLogMode == qcModes.prep.code) && STATUS.ready) || STATUS.shelf

  dateCell.setValue(new Date())
  iRunCell.setValue("PASS")
  fRunCell.setValue("PASS")
  timeUpdate(curRow)
  statusCell.setValue(statVal)

  let dupedCells = isUnitDuped(curCell.getValue())
  if (dupedCells) { dupeUnit(curCell, dupedCells) }
}

function deleteUnit(curCell, oldVal) {
  if (curCell.getColumn() !== COLS.pid.num) return;

  let curRow = curCell.getRow()
  let dateCell = getCell(`${COLS.date.letter}${curRow}`)
  let dateVal = dateCell.getValue()

  if (isNull(dateVal) && isNull(oldVal)) return;

  if (oldVal) {
    let dupedCells = isUnitDuped(oldVal)
    if (dupedCells) { deleteCellNotes(dupedCells) }
  }

  curSheet.deleteRow(curRow)
}

function deleteCellNotes(cells) {
  cells.forEach(cell => cell.clearNote())
}

function dupeUnit(curCell, dupedCells) {
  dupedCells.push(curCell)

  let entries = "ENTRIES: \n"

  dupedCells.forEach(cell => {
    let cellRow = cell.getRow()
    let statusVal = getCell(`${COLS.status.letter}${cellRow}`).getValue()
    let dateVal = getCell(`${COLS.date.letter}${cellRow}`).getValue()
    let date = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM/dd/yy")

    let txt = `• [${date}] [R:${cellRow}] → ${statusVal} \n`
    entries += txt
  })

  dupedCells.forEach(cell => {
    cell.setNote(entries)
  })
}

function failUnit(curCell) {
  if (curCell.getColumn() !== COLS.iRun.num) return;

  let curRow = curCell.getRow()
  let pidCell = getCell(`${COLS.pid.letter}${curRow}`)
  let pidVal = pidCell.getValue()

  if (isNull(pidVal)) return;

  let iRunCell = getCell(`${COLS.iRun.letter}${curRow}`)
  let iRunVal = iRunCell.getValue()

  if (iRunVal !== "FAIL") return;

  let fRunCell = getCell(`${COLS.fRun.letter}${curRow}`)
  let statusCell = getCell(`${COLS.status.letter}${curRow}`)

  fRunCell.setValue("FAIL")
  statusCell.setValue(STATUS.pending)
  timeUpdate(curRow)
}

function shelveUnit(curCell) {
  if (curCell.getColumn() !== COLS.fRun.num) return;

  let curRow = curCell.getRow()
  let fRunCell = getCell(`${COLS.fRun.letter}${curRow}`)
  let fRunVal = fRunCell.getValue()

  if ((isNull(fRunVal)) || (fRunVal !== "PASS")) return;

  let statusCell = getCell(`${COLS.status.letter}${curRow}`)
  statusCell.setValue(STATUS.shelf)
  timeUpdate(curRow)
}

function lastUpdated(curCell) {
  if (curCell.getColumn() !== COLS.status.num) return;

  timeUpdate(curCell.getRow())
}

function showFailures(curCell) {
  let failureSheet = fullSheet.getSheetByName("Failures")
  let headers = failureSheet.getRange("D2:R2")
  let failPoint = curCell.getValue()

  let failPointCell = findMatch(headers, failPoint, false)
  if (!failPointCell) return;

  let startRow = failPointCell.getRow() + 1
  let startCol = failPointCell.getColumn()
  let numRows = failureSheet.getLastRow() - startRow + 1
  let numCols = 1
  let failArray = failureSheet.getRange(startRow, startCol, numRows, numCols).getValues()

  let failTypes = failArray.map(failType => {
    return failType[0]
  }).filter(failType => {
    return String(failType).trim() !== ""
  })

  let validRule = SpreadsheetApp.newDataValidation().requireValueInList(failTypes).build()
  curCell.offset(0, 1).setDataValidation(validRule)
}
