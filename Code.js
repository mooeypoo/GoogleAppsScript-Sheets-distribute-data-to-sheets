/**
 * This is a script that takes an expected format of a live data sheet from Asana
 * and distributes it to individual sheets based on search terms.
 * 
 * The search terms are defined in the SEARCHTERMS sheet, and the headers for the
 * individual sheets are defined in the HEADERS sheet.
 * 
 * @autor Moriel Schottlender
 * @version 1.0
 * @see https://github.com/mooeypoo/GoogleAppsScript-Sheets-distribute-data-to-sheets
 */
// Defaults
const DEFINITION = {
  headers: [],
  terms: {
    // 'search term / short name' -> 'sheet name'
    'WE1': 'WE1',
    'WE2': 'WE2',
    // 'WE3': 'WE3',
    // 'WE4': 'WE4',
    // 'WE5': 'WE5',
    // 'WE6': 'WE6'
  },
  Cols: [
    '', // <-- nothing so the count in 'indexOf' can start from 1
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
    'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U' ,'V', 'W','X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE'
  ]
}
const activeSpreadsheet = SpreadsheetApp.getActive();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('✨ Distribute Live Data ✨')
      .addItem('✈️ Distribute Live Data', 'copySourceToSheets')
      .addSeparator()
      .addItem('🗑️ Delete all unprotected sheets', 'deleteUnprotectedSheets')
      .addToUi();
}

/**
 * Deletes all unprotected sheets.
 * 
 * @returns Boolean Whether the operation should proceed (true) or was stopped (false)
 */
function deleteUnprotectedSheets() {
  // Verify action!
  if (!checkConfirmation(`This action will delete all generated (non-locked) sheets and regenerate them. Any current data in these sheets will be forever lost.
  
  ** This action CANNOT be undone. **
  Please make sure you have a backup of this document if you want to retain this information.
  
  Are you sure you want to proceed?`)) {
    Logger.log('User did not confirm. Stopping operation.')
    // Break operation
    return false
  }

  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  toast('Deleting unprotected sheets...')
  for (let s=0; s < sheets.length; s++) {
    const protection = sheets[s].getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    if (
      protection?.canEdit() ||
      sheets[s].getName() === 'Overview' ||
      sheets[s].getName() === 'Asana portfolio data'
    ) {
      continue;
    }
    console.log(`Deleting sheet: "${sheets[s].getName()}"`)
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheets[s])
  }

  return true
}

/**
 * Run the operation to copy the data from source sheet
 * into the individual sheets.
 * 
 * This is the entrypoint method that includes all the
 * sub operations and is called from the menu.
 */
function copySourceToSheets() {
  Logger.log('Start operation')

  const resetAction = deleteUnprotectedSheets()
  if (!resetAction) {
    toast('You have not confirmed. Operation stopped without action.')
    // User did not confirm. Break.
    return
  }

  toast('Starting distribution: analyzing live data.')
  const termMap = getTermDefinitions()
  const valueRowMap = processLiveDataToMap(termMap)
  const rowMapkeys = Object.keys(valueRowMap).sort((a, b) => a > b)

  // Map the rows in the live data sheet into the search term results
  // that they fit into, so we can then create sheets with those
  // results pasted into the individual sheets.
  toast('Distributing data to sheets.')
  rowMapkeys.forEach(key => {
    // Get the target sheet
    const targetSheet = getTargetSheet(key, termMap)
    const targetData = valueRowMap[key]

    // Check the range for the data
    let cellRange = targetSheet.getRange(
      2, // Row (row 2 to account for headers)
      1, // Col
      targetData.length, // Number of rows
      targetData[0].length // Number of cols
    );
    Logger.log('Inserting data to sheet', key)

    // Style
    cellRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    cellRange.setVerticalAlignment('top')

    // Insert data
    cellRange.setValues(targetData)

    // SORT the rows by "Posted on" (col U = ) in descending order

    // First, set the formatting on the "U" column to date so that the
    // sorting operation understands what it's doing.
    targetSheet.getRange('U:U').setNumberFormat("MM/DD/YYYY")
    // Apply sort on the entire range based on the "U" column
    cellRange.sort({
      column: DEFINITION.Cols.indexOf('U'),
      ascending: false
    })
    console.log('Sorted data based on Col U(' + DEFINITION.Cols.indexOf('U') + ')')

    Logger.log('Formatting columns for sheet', key)
    // Insert checkbox to column Y
    cellRange = targetSheet.getRange(
      2, // Skip for headers,
      25, // Col 25 = Y (count starts from 1)
      targetData.length, // Num rows
      1 // Num cols
    )
    cellRange.insertCheckboxes()

    // Format content for Status update -- FOR EACH LINE
    replaceAndFormatCell(
      targetSheet,
      19, // Col 19 = S (count starts from 1)
      24, // Col 24 = X (count starts from 1)
      targetData.length // Number of rows to repeat
    )
    // Format content for Description -- FOR EACH LINE
    replaceAndFormatCell(
      targetSheet,
      5, // Col 5 = E (count starts from 1)
      5, // Col 5 = E (count starts from 1)
      targetData.length // Number of rows to repeat
    )

    // Hide columns
    targetSheet.hideColumns(10, 8) // J - Q
    targetSheet.hideColumns(19, 1) // S

    // Resize cols to fit data
    targetSheet.setColumnWidth(
      1, // Column 1 - Name
      300 // pixels
    )
    targetSheet.setColumnWidth(
      5, // Column 5 - Description
      550 // pixels
    )
    targetSheet.setColumnWidth(
      19, // Column 19 - LATEST STATUS UPDATE
      550 // pixels
    )
    targetSheet.setColumnWidth(
      24, // Column 24 - 
      550 // pixels
    )

    // Freeze headers
    targetSheet.setFrozenRows(1) 
  })

  toast('Operation complete.')
}


/**
 * Process the live data sheet and map values into the search terms
 * that they fit into.
 */
function processLiveDataToMap(termMap) {
  const sourceSheet = activeSpreadsheet.getSheetByName('Live source data');
  const allSearchTerms = Object.keys(termMap)

  Logger.log('Analyzing source data.')
  const valuesRowMap = {}
  const values = sourceSheet.getDataRange().getValues().filter(row => !!row[0])

  Logger.log('Searching for term matches.')
  // Go over the source data, map line by line to the key it belongs to
  // W = item 25 in the array, whcih is values[r][24]
  for (let r = 0; r < values.length; r++) {
      if (values[r][22]) {
        for (let t = 0; t < allSearchTerms.length; t++) {
          const term = allSearchTerms[t]
          // console.log('check match (term / values[r][22])', term, values[r][22])
          try {
            if (
              values[r][22]
                .toString()
                .toLowerCase()
                .includes(term.toLowerCase())
            ) {
              console.log('matching', term, values[r][22])
              // Found. Insert into value map
              if (!valuesRowMap[term]) {
                valuesRowMap[term] = [] // instantiate if needed
              }
              valuesRowMap[term].push(values[r])
            }
          } catch(err) {
            SpreadsheetApp.getActive().toast('Error, could not read column ' + r + ', ' + c)
          }
        }
    }
  }

  return valuesRowMap
}

/**
 * This takes text from one cell, formats it to replace spaces that Asana produces
 * with linebreaks for better readability. This repeats for all available rows.
 * The cell could also be the same cell, replacing its own data.
 * 
 * Remember: Cell numbers start from 1.
 * 
 * @param Sheet targetSheet The Sheet object to perform this replacement on
 * @param number fromCellNum The column count for the needed cell to copy info from
 * @param number fromCellNum The column count for the needed cell to copy info to
 * @param number rowNumber The total number of rows that this should repeat on
 */
function replaceAndFormatCell(targetSheet, fromCellNum, toCellNum, rowNumber) {
  // Format content for Status update -- FOR EACH LINE
  for (let i=1; i<= rowNumber; i++) {
    const pasteToCell = targetSheet.getRange(
      i + 1, // Skip for headers,
      toCellNum, // Col number (count starts from 1)
      1, // Num rows
      1 // Num cols
    )
    const readFromCell = targetSheet.getRange(
      i + 1, // Skip for headers,
      fromCellNum, // Col number (count starts from 1)
      1, // Num rows
      1 // Num cols
    )

    try {
      const value = readFromCell.getValues()[0][0]
      pasteToCell.setValue(value.replaceAll('    ', '\n'))
    } catch(err) {
      console.log(' Could not format information from Cell ' + fromCellNum + ' to cell ' + toCellNum + '.', err)
    }
  }
}

/**
 * Fetch the term definition that maps search terms to sheet names.
 * Fall back on base definition.
 */
function getTermDefinitions() {
  Logger.log('getTermDefinitions')
  // Get from sheet
  let searchTermMap = DEFINITION.terms // Fall back on definition
  console.log('searchTermMap from DEFINITION', searchTermMap)
  try {
    const searchTermsSheet = activeSpreadsheet.getSheetByName('SEARCHTERMS')
    if (searchTermsSheet) {
      const searchTerms = searchTermsSheet.getDataRange().getValues()
      // Create the map
      // Remove first row
      searchTerms.shift()
      console.log('Search terms from spreadsheet (without header row)', searchTerms)
      console.log('Mapping search terms values into searchtermMap')
      searchTermMap = {} // reset
      for (let rowIndex = 0; rowIndex < searchTerms.length; rowIndex++) {
        console.log('for loop', rowIndex,searchTerms[rowIndex][0],searchTerms[rowIndex][1] )
        searchTermMap[searchTerms[rowIndex][0]] = searchTerms[rowIndex][1]
      }
    }
  } catch (err) {
    console.log('Error reading search terms from SEARCHTERMS sheet.', err)
  }
  console.log('Search terms map ready', searchTermMap)
  return searchTermMap
}

/**
 * Show a confirmation prompt to the user and check
 * if the answer is true (confirmed) or false (not confirmed)
 * 
 * @param String text The text to appear in the confirmation prompt
 */
function checkConfirmation(text) {
  const ui = SpreadsheetApp.getUi()
  const result = ui.alert(
      text,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  return result === ui.Button.YES
}


/**
 * Get the target sheet for data.
 * Return the existing sheet matching the given key or create a sheet
 * if one does not already exist.
 * If creating a new sheet, insert the proper headers.
 * 
 * @param string key A key for the sheet; should be in the definition
 */
function getTargetSheet(key, termMap) {
  const targetSheetName = termMap[key]
  let targetSheet = activeSpreadsheet.getSheetByName(targetSheetName)
  if (!targetSheet) {
    // Target sheet doesn't exist; create it
    targetSheet = activeSpreadsheet.insertSheet();
    targetSheet.setName(targetSheetName);

    // Get headers for new sheet
    let headers = DEFINITION.headers // Fall back on definition
    Logger.log('headers from DEFINITION', headers)
    try {
      const headersSheet = activeSpreadsheet.getSheetByName('HEADERS')
      if (headersSheet) {
        headers = headersSheet.getDataRange().getValues()
        // Remove first cell
        headers[0].shift()
        Logger.log('headers from spreadsheet', headers)
      }
    } catch (err) {
      Logger.log('Error reading headers from HEADERS sheet.', err)
    }

    // Insert headers
    try {
      const cellRange = targetSheet.getRange(1,1,headers.length, headers[0].length);
      // Style headers
      cellRange.setFontWeight('bold')
      cellRange.setBorder(true, true, true, true, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
      cellRange.setValues(headers)
    } catch(err) {
      toast(`Could not insert headers to new sheet: "${targetSheetName}"`, 'warning')
    }
  }

  return targetSheet
}

function toast(msg, type) {
  let icon
  switch (type) {
    case 'warning':
      icon = '⚠️'
      break;
    case 'error':
      icon = '❗'
      break;
    default:
      icon = 'ℹ️' // info (default)
      break;
  }
  SpreadsheetApp.getActive().toast(icon + ' ' + msg)
}
