/* shim for gsheets. todo = add eslint exception */ if ('undefined' === typeof SpreadsheetApp) { var SpreadsheetApp = {}} //prettier-ignore
const WIDTHS_ROW = 9
const WIDTHS_COL = WIDTHS_ROW
const LOG_ROW = 7
const LOG_COL = LOG_ROW
const A1_pixelsPerInch = 'H6'
const DEFAULT_CELL_INCHES = 3.5

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
      let msg = ''
      let inches = getRowColVal(WIDTHS_ROW, col)
      /// SAME ///
      if (inches) {
        msg += `inches: ${inches}. `
        msg += `pixels: ${inches * pixelsPerInch}. `
      } else {
        msg += 'no width. '
        inches = DEFAULT_CELL_INCHES
      }
      /// END SAME ///
      const pixels = inches * pixelsPerInch
      sheet.setColumnWidth(col, pixels)
      logToCol(col, msg + ` (${now()})`)
    }

    /// ROWS:
    const numFrozenRows = sheet.getFrozenRows()
    const rowCount = sheet.getMaxRows()
    for (
      let row = numFrozenRows + 1;
      row <= rowCount;
      row++
    ) {
      let msg = ''
      let inches = getRowColVal(row, WIDTHS_COL)
      /// SAME ///
      if (inches) {
        msg += `inches: ${inches}. `
        msg += `pixels: ${inches * pixelsPerInch}. `
      } else {
        msg += 'no width. '
        inches = DEFAULT_CELL_INCHES
      }
      /// END SAME ///
      const pixels = inches * pixelsPerInch
      sheet.setRowHeight(row, pixels)
      logToRow(row, msg + ` (${now()})`)
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
  return `${new Date().getHours()}:${new Date().getMinutes()}:${new Date().getSeconds()}`
}
