function rosterSort() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var rosterSheet = ss.getSheetByName("Roster")
  var keySheet = ss.getSheetByName("Sorting Code")
  var sortColumn = keySheet.getRange('E4:E28')

  // Variables
  var rosterSortColumn = 26
  var rosterFirstRow = 9
  var rosterLastRow = 33
  var rosterFirstColumn = 4
  var numberOfColumns = 23 // Hidden column
  var regSlots = 25

  // Sort
  sort(rosterSheet, sortColumn, rosterSortColumn, rosterFirstRow, rosterLastRow, regSlots, rosterFirstColumn, numberOfColumns)

  // TODO: Call other functions
  subCompanyUpdate()
}

function subCompanySort() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var subCompanySheet = ss.getSheetByName("Sub Companies")
  var keySheet = ss.getSheetByName("Sorting Code")
  var sortCom1 = keySheet.getRange('F4:F8')
  var sortCom2 = keySheet.getRange('F11:F15')
  var sortCom3 = keySheet.getRange('F18:F22')


  // Variables
  var subComSortColumn = 14
  var subComFirstColumn = 4
  var numberOfColumns = 11
  var numberOfSlots = 5
  var com1FirstRow = 10
  var com1LastRow = 14
  var com2FirstRow = 27
  var com2LastRow = 31
  var com3FirstRow = 44
  var com3LastRow = 48


  // Sort
  sort(subCompanySheet, sortCom1, subComSortColumn, com1FirstRow, com1LastRow, numberOfSlots, subComFirstColumn, numberOfColumns)
  sort(subCompanySheet, sortCom2, subComSortColumn, com2FirstRow, com2LastRow, numberOfSlots, subComFirstColumn, numberOfColumns)
  sort(subCompanySheet, sortCom3, subComSortColumn, com3FirstRow, com3LastRow, numberOfSlots, subComFirstColumn, numberOfColumns)
}

function subCompanyUpdate() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var subCompanySheet = ss.getSheetByName("Sub Companies")
  var sortingSheet = ss.getSheetByName("Sorting Code")
  var com1Names = subCompanySheet.getRange('D9:D14')
  var com2Names = subCompanySheet.getRange('D26:D31')
  var com3Names = subCompanySheet.getRange('D43:D48')
  var arcNames = subCompanySheet.getRange('D60:D61')

  // Variables
  var com1Removed = sortingSheet.getRange('L5:L10')
  var com2Removed = sortingSheet.getRange('L12:L17')
  var com3Removed = sortingSheet.getRange('L19:L24')
  var arcRemoved = sortingSheet.getRange('L26:L27')
  var comSlots = 6
  var arcSlots = 2
  
  // Do removals
  removeCompany(com1Names, com1Removed, comSlots)
  removeCompany(com2Names, com2Removed, comSlots)
  removeCompany(com3Names, com3Removed, comSlots)
  removeCompany(arcNames, arcRemoved, arcSlots)

  subCompanySort()

  // Add variables
  var toAdd = sortingSheet.getRange('K5:K29')
  var companies = sortingSheet.getRange('B36:B39')
  var names = sortingSheet.getRange('J5:J29')
  var slots = 25
  
  // Do adds
  addCompany(names, toAdd, companies, slots, com1Names, com2Names, com3Names, arcNames, comSlots, arcSlots)

  // Sort
  subCompanySort()
}

function benchmarkingSort() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var benchmarkingSheet = ss.getSheetByName("Benchmarking Review")
  var sortingSheet = ss.getSheetByName("Sorting Code")
  var sortColumn = sortingSheet.getRange('G4:G29')

  // Variables
  var benchSortColumn = 29
  var benchFirstRow = 9
  var benchLastRow = 33
  var benchFirstColumn = 4
  var numberOfColumns = 26
  var regSlots = 25

  // Sort
  sort(benchmarkingSheet, sortColumn, benchSortColumn, benchFirstRow, benchLastRow, regSlots, benchFirstColumn, numberOfColumns)


}

function benchmarkingUpdate() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var benchmarkingSheet = ss.getSheetByName("Benchmarking Review")
  var sortingSheet = ss.getSheetByName("Sorting Code")

  // Variables
  var toRemove = sortingSheet.getRange('M5:M30')
  var slots = 25
  var benchFirstRow = 9
  var benchFirstColumn = 4
  var numberOfColumns = 26

  // Removals
  for (var i = 0; i < slots; i++) {
    if (toRemove.getCell(i + 1, 1).getValue() === false) {
      var clear = benchmarkingSheet.getRange(benchFirstRow + i, benchFirstColumn, 1, numberOfColumns)
      clear.clearContent()
      clear.clearNote()
    }
  }

  // sort
  benchmarkingSort()
}

function taskSort() {
  // Set up sheets to be used
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var taskSheet = ss.getSheetByName("Task Checklist")
  var sortSheet = ss.getSheetByName("Sorting Code")
  var sortSpec = sortSheet.getRange('H4:H13')
  var sortSerg = sortSheet.getRange('H15:H26')
  var sortWO = sortSheet.getRange('I4:I10')


  // Variables
  var taskSortColumn = 23
  var taskFirstColumn = 3
  var numberOfColumns = 21
  var specSlots = 10 // Stupid merge cell so +1

  var sergSlots = 12
  var sergFirstRow = 23
  var sergLastRow = 34
  var woSlots = 7
  var woFirstRow = 40
  var woLastRow = 46

  // Make a new table to sort for spec
  var specTable = taskSheet.getRange('C8:V17')
  var newTableFirstC = 4
  var newTableLastC = 23
  var newTableFirstR = 36
  var newTableLastR = 45
  var specSortColumn = 24

  // Sort specialist in temp
  specTable.copyValuesToRange(sortSheet, newTableFirstC, newTableLastC, newTableFirstR, newTableLastR)

  // Sort the rest
  sort(sortSheet, sortSpec, specSortColumn, newTableFirstR, newTableLastR, specSlots, newTableFirstC, numberOfColumns)
  sort(taskSheet, sortSerg, taskSortColumn, sergFirstRow, sergLastRow, sergSlots, taskFirstColumn, numberOfColumns)
  sort(taskSheet, sortWO, taskSortColumn, woFirstRow, woLastRow, woSlots, taskFirstColumn, numberOfColumns)

  // Transfer temp spec table to actual (Two parts since merge)
  var sortedTempSpecTable1 = sortSheet.getRange('D36:D41')
  var sortedTempSpecTable1p2 = sortSheet.getRange('H36:W41')
  var sortedTempSpecTable2 = sortSheet.getRange('D42:D44')
  var sortedTempSpecTable2p2 = sortSheet.getRange('H42:W44')

  // (Some variables got from above)
  var specFirstRow = 8
  var specLastRow = 17
  var taskLastColumn = 22
  var mergedRow = 13
  var rowAfterMerge = 15
  var columnAfterCode = 7 // No before since same as name cell

  // Also have to avoid copying the middle cells as code won't transfer
  sortedTempSpecTable1.copyValuesToRange(taskSheet, taskFirstColumn, taskFirstColumn, specFirstRow, mergedRow)
  sortedTempSpecTable1p2.copyValuesToRange(taskSheet, columnAfterCode, taskLastColumn, specFirstRow, mergedRow)
  sortedTempSpecTable2.copyValuesToRange(taskSheet, taskFirstColumn, taskFirstColumn, rowAfterMerge, specLastRow)
  sortedTempSpecTable2p2.copyValuesToRange(taskSheet, columnAfterCode, taskLastColumn, rowAfterMerge, specLastRow)
}

function sort(sheet, sortColumn, newSortColumn, firstRow, lastRow, numberOfRows, firstColumn, numberOfColumns) {
  sortColumn.copyValuesToRange(sheet, newSortColumn, newSortColumn, firstRow, lastRow)
  var range = sheet.getRange(firstRow, firstColumn, numberOfRows, numberOfColumns)
  range.sort({column: newSortColumn, ascending: false})

  var clear = sheet.getRange(firstRow, newSortColumn, numberOfRows, 1)
  clear.clearContent()
}

function removeCompany(comNames, comRemoved, slots) {
  for (var i = 1; i < slots + 1; i++) {
    if (comRemoved.getCell(i, 1).getValue() === false) {
      comNames.getCell(i, 1).clearContent()
    }
  }
}

function addCompany(names, addList, companies, slots, com1Names, com2Names, com3Names, arcNames, comSlots, arcSlots) {
  var com1Count = comSlots
  var com2Count = comSlots
  var com3Count = comSlots
  var arcCount = arcSlots
  for (var i = 1; i < slots + 1; i++) {
    if (!addList.getCell(i, 1).isBlank()) {
      var wordArray = addList.getCell(i, 1).getValue().split(" ")
      var company = wordArray[0]
      var name = names.getCell(i, 1).getValue()
      switch (company) {
        case companies.getCell(1, 1).getValue():
          com1Names.getCell(com1Count, 1).setValue(name)
          com1Count--
          break;
        case companies.getCell(2, 1).getValue():
          com2Names.getCell(com2Count, 1).setValue(name)
          com2Count--
          break;
        case companies.getCell(3, 1).getValue():
          com3Names.getCell(com3Count, 1).setValue(name)
          com3Count--
          break;
        case companies.getCell(4, 1).getValue():
          arcNames.getCell(arcCount, 1).setValue(name)
          arcCount--
          break;
      }
    }
  }
}



