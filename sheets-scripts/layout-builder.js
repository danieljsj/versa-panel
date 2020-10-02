if ('undefined' === typeof SpreadsheetApp) { var SpreadsheetApp = {}} //prettier-ignore

const WIDTHS_ROW_NUM = 9
const WIDTHS_COL_NUM = 9
const A1_pixelsPerInch = 'H6'

// eslint-disable-next-line
function refreshLayoutStripWidths() {
  const layoutSheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter((sheet) => sheet.getName().match(/^LAYOUT:/))

  layoutSheets.forEach((sheet) => {
    const { getA1Val, getRowColVal } = new SheetHelpers(
      sheet
    )
    const pixelsPerInch = getA1Val(A1_pixelsPerInch)

    const numFrozenCols = sheet.getFrozenColumns()
    const colCount = sheet.getMaxColumns()

    /// COLUMNS:
    for (
      let colNum = numFrozenCols + 1;
      colNum <= colCount;
      colNum++
    ) {
      const inches = getRowColVal(WIDTHS_ROW_NUM, colNum)
      if (inches) {
        const pixels = inches * pixelsPerInch
        sheet.setColumnWidth(colNum, pixels)
      }
    }

    /// ROWS:
    const numFrozenRows = sheet.getFrozenRows()
    const rowCount = sheet.getMaxRows()
    for (
      let rowNum = numFrozenRows + 1;
      rowNum <= rowCount;
      rowNum++
    ) {
      const inches = getRowColVal(WIDTHS_COL_NUM, rowNum)
      if (inches) {
        const pixels = inches * pixelsPerInch
        sheet.setRowHeight(rowNum, pixels)
      }
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
  }
}
