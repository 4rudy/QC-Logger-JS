function getCell(cell) {
  return qcSheet.getRange(cell) //getRange is the actual cell object
}

function isDupedUnit(unitPID) {
  let lastRow = qcSheet.getLastRow() - 1
  let unitPIDS = getCell(`B2:B${lastRow}`)
  let dupedCells = findMatch(unitPIDS, unitPID, true)

  return (dupedCells.length>=1 && dupedCells) || false;
}

function findMatch(searchRange, searchValue, findAll) {
  let matches = searchRange.createTextFinder(searchValue)
    .matchEntireCell(true)
    .matchCase(false)

  return (findAll && matches.findAll()) || matches.findNext()
}

function isNull(val) {
  return (!val || String(val).trim() === "")
}

function timeUpdate(curRow) {
  let updatedCell = getCell(`${COLS.updated.letter}${curRow}`)
  updatedCell.setValue(new Date())
}
