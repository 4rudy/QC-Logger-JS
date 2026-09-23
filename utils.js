function getDataArray() {
  let data = curSheet.getDataRange().getValues()
  data.shift()

  data.forEach((unit, idx) => {
    if (String(unit[0]).trim() !== "") {
      let row = idx + 2

      allUnits.push({
        row: row,
        date: unit[0],
        pid: unit[1],
        iRun: unit[2],
        failPoint: unit[3],
        failType: unit[4],
        otherFails: unit[5],
        notes: unit[6],
        fRun: unit[7],
        status: unit[8],
        updated: unit[9],
      })
    }
  })

  return allUnits
}

function addUnit(curRow, pid, statVal) {
  allUnits.push({
    row: curRow,
    date: new Date(),
    pid: pid,
    iRun: "PASS",
    failPoint: "",
    failType: "",
    otherFails: "",
    notes: "",
    fRun: "PASS",
    status: statVal,
    updated: new Date(),
  })
}

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
