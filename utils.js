function getCell(cell) {
  return curSheet.getRange(cell) //getRange is the actual cell object
}

function isUnitDuped(unitPID) {
  let dupedCells = allUnits.filter((unit) => (unit.pid == unitPID))
  return (dupedCells.length > 1 && dupedCells) || false;
}

function findMatch(searchRange, searchValue) {
  let matches = searchRange.createTextFinder(searchValue)
    .matchEntireCell(true)
    .matchCase(false)

  return matches.findNext()
}

function isNull(val) {
  return (!val || String(val).trim() === "")
}

function timeUpdate(curRow) {
  let updatedCell = getCell(`${COLS.updated.letter}${curRow}`)
  updatedCell.setValue(new Date())
}
