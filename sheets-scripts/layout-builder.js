if ('undefined' === typeof SpreadsheetApp) { var SpreadsheetApp = {}} //prettier-ignore

const WIDTHS_ROW_NUM = 9
const WIDTHS_COL_NUM = 9
const SCALE_A1 = [9, 1]

// eslint-disable-next-line
function refreshLayoutStripWidths() {
  const layoutSheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter((sheet) => sheet.getName().match(/^LAYOUT:/))

  layoutSheets.forEach((sheet) => {
    const { getA1Val, getRowColVal } = new SheetHelpers(
      sheet
    )
    const scale = getA1Val(...SCALE_A1)

    const colCount = sheet.getMaxColumns()
    for (let colNum = 1; colNum <= colCount; colNum++) {
      const colInches = getRowColVal(WIDTHS_ROW_NUM, colNum)
      const colPixels = colInches * scale
      sheet.setColumnWidth(colNum, colPixels)
    }
    const rowCount = sheet.getMaxRows()
    for (let rowNum = 1; rowNum <= rowCount; rowNum++) {
      const rowInches = getRowColVal(WIDTHS_COL_NUM, rowNum)
      const rowPixels = rowInches * scale
      sheet.setColumnWidth(rowNum, rowPixels)
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
  ss.addMenu('Refresh Layout Widths', entries)
}

function SheetHelpers(sheet) {
  return {
    getA1Val: (a1) => sheet.getRange(a1).getValue(),
    getRowColVal: (y, x) => sheet.getRange(y, x).getValue(),
  }
}
