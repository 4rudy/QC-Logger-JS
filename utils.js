function getCell(cell){
  return qcSheet.getRange(cell) //getRange is the actual cell object
}

function dupedUnit(unitPID) {
  let lastRow = qcSheet.getLastRow()-1

  let unitPIDS = qcSheet.getRange(`B2:B${lastRow}`)
  let dupedCell = unitPIDS.createTextFinder(unitPID)
                              .matchEntireCell(true)
                              .matchCase(false)
                              .findNext()

  return dupedCell != null;
}

function isNull(val){
  return (!val || String(val).trim() === "")
}

function timeUpdate(curRow){
    let updatedCell = getCell(`${COLS.updated.letter}${curRow}`)
    updatedCell.setValue(new Date())
}