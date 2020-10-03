/* shim for gsheets. todo = add eslint exception */ if ('undefined' === typeof SpreadsheetApp) { var SpreadsheetApp = {}} //prettier-ignore
const WIDTHS_ROW = 9
const WIDTHS_COL = WIDTHS_ROW
const LOG_ROW = 7
const LOG_COL = LOG_ROW
const A1_zoom = 'J3'
const PIXELS_PER_INCH_PER_ZOOM = 10
const DEFAULT_CELL_INCHES = 7
const LAYER_SHEET_NAME_MATCH_STR = 'LAYER_' // 'LAYER_' also hardcoded below
const MIN_WIDTH = 21

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

  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu('Custom Menu')
    .addItem('Show alert', 'showAlert')
    .addToUi()
}

// eslint-disable-next-line
function refreshLayoutStripWidths() {
  // // just going to do one for now.
  // const layoutSheets = SpreadsheetApp.getActiveSpreadsheet()
  //   .getSheets()
  //   .filter((sheet) => sheet.getName().match(/LAYOUT:/))

  // layoutSheets.forEach()

  const sheet = SpreadsheetApp.getActiveSheet()

  if (sheet.getName().match(LAYER_SHEET_NAME_MATCH_STR)) {
    refreshSheetLayoutStripWidths(sheet)
  } else {
    // eslint-disable-next-line
    aSheetNameWhoseTextContainsLAYER_() // originally: `Refresh canceled; Design tabs must contain "${LAYER_SHEET_NAME_MATCH_STR}" in their title.`
  }
}

function refreshSheetLayoutStripWidths(sheet) {
  const {
    getA1Val,
    getRowColVal,
    logToCol,
    logToRow,
  } = new SheetHelpers(sheet)
  const zoom = getA1Val(A1_zoom)
  const pixelsPerInch = zoom * PIXELS_PER_INCH_PER_ZOOM

  const rowCount = sheet.getMaxRows()
  const colCount = sheet.getMaxColumns()
  const numFrozenRows = sheet.getFrozenRows()
  const numFrozenCols = sheet.getFrozenColumns()

  /// C O L U M N S :
  for (
    let col = numFrozenCols + 1;
    col <= colCount;
    col++
  ) {
    let inches = getRowColVal(WIDTHS_ROW, col)
    /// SAME!!!!!!!!!!!!!! ///
    let msg = `${now()}: `
    if (inches) {
      msg += `inches: ${inches}. `
    } else {
      inches = DEFAULT_CELL_INCHES
    }
    const pixels = Math.max(
      inches * pixelsPerInch,
      MIN_WIDTH
    )
    msg += `pixels: ${pixels}. `
    /// END SAME!!!!!!!!!!! ///
    sheet.setColumnWidth(col, pixels)
    logToCol(col, msg)
  }

  /// R O W S :
  for (
    let row = numFrozenRows + 1;
    row <= rowCount;
    row++
  ) {
    let inches = getRowColVal(row, WIDTHS_COL)
    /// SAME!!!!!!!!!!!!!! ///
    let msg = `${now()}: `
    if (inches) {
      msg += `inches: ${inches}. `
    } else {
      inches = DEFAULT_CELL_INCHES
    }
    const pixels = Math.max(
      inches * pixelsPerInch,
      MIN_WIDTH
    )
    msg += `pixels: ${pixels}. `
    /// END SAME!!!!!!!!!!! ///

    sheet.setRowHeight(row, pixels)
    logToRow(row, msg)
  }
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

// NOT WORKING?????
// eslint-disable-next-line
function showAlert() {
  var ui = SpreadsheetApp.getUi() // Same variations.

  var result = ui.alert(
    'Please confirm',
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  )

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.')
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.')
  }
}

/*** return sheet names for active document (DID THIS MESSAGE SHOW UP???? ARE YOU READING THIS???)
 * @customfunction
 */
function sheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (sheet) {
      return sheet.getName()
    })
}
/*** return LAYER sheet names for active document (DID THIS MESSAGE SHOW UP???? ARE YOU READING THIS???)
 * @customfunction
 */
//eslint-disable-next-line
function layerSheetNames() {
  return sheetNames.filter((sheetName) =>
    sheetName.match(LAYER_SHEET_NAME_MATCH_STR)
  )
}

// THIS WASNT WORKING... BUT IT MIGHT BE USEFUL LATER
// sheet
//   .getRange({
//     row: numFrozenRows + 1,
//     column: col,
//     numRows: rowCount - numFrozenRows,
//     numColumns: 1,
//   })
//   .setTextRotation(-90)
