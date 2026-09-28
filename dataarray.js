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

function findUnit(unitRow) {
    return allUnits.find(unit => unit.row == unitRow)
}
