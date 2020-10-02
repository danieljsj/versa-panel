/* shim for gsheets. todo = add eslint exception */ if ('undefined' === typeof SpreadsheetApp) { var SpreadsheetApp = {}} //prettier-ignore
const WIDTHS_ROW = 9
const WIDTHS_COL = WIDTHS_ROW
const LOG_ROW = 7
const LOG_COL = LOG_ROW
const A1_pixelsPerInch = 'H6'

// eslint-disable-next-line
function refreshLayoutStripWidths() {
  const layoutSheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter((sheet) => sheet.getName().match(/^LAYOUT:/))

  layoutSheets.forEach((sheet) => {
    const {
      getA1Val,
      getRowColVal,
      logToCol,
      logToRow,
    } = new SheetHelpers(sheet)

    const pixelsPerInch = getA1Val(A1_pixelsPerInch)

    const numFrozenCols = sheet.getFrozenColumns()
    const colCount = sheet.getMaxColumns()

    /// COLUMNS:
    for (
      let col = numFrozenCols + 1;
      col <= colCount;
      col++
    ) {
      let msg = `${now()}: `
      const inches = getRowColVal(WIDTHS_ROW, col)
      if (inches) {
        const pixels = inches * pixelsPerInch
        sheet.setColumnWidth(col, pixels)
        msg += `inches: ${inches}. `
      } else {
        msg += 'no width. '
      }
      logToCol(col, msg)
    }

    /// ROWS:
    const numFrozenRows = sheet.getFrozenRows()
    const rowCount = sheet.getMaxRows()
    for (
      let row = numFrozenRows + 1;
      row <= rowCount;
      row++
    ) {
      let msg = `${now()}: `
      const inches = getRowColVal(WIDTHS_COL, row)
      if (inches) {
        const pixels = inches * pixelsPerInch
        sheet.setRowHeight(row, pixels)
        msg += `inches: ${inches}. `
      } else {
        msg += 'no width. '
      }
      logToRow(row, msg)
    }
  })
}

// eslint-disable-next-line
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const menuItem =
    'Refresh Layout Display to Match Specified Widths'
  const entries = [
    {
      name: menuItem,
      functionName: 'refreshLayoutStripWidths',
    },
  ]
  ss.addMenu('REFRESH!', entries)
}

function SheetHelpers(sheet) {
  return {
    getA1Val: (a1) => sheet.getRange(a1).getValue(),
    getRowColVal: (y, x) => sheet.getRange(y, x).getValue(),
    logToCol: (col, msg) => {
      sheet.getRange(LOG_ROW, col).setValue(msg)
    },
    logToRow: (row, msg) => {
      sheet.getRange(row, LOG_COL).setValue(msg)
    },
  }
}

function now() {
  return new Date().toTimeString()
}
