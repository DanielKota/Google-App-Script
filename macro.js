/** @OnlyCurrentDoc */

function script(){
  
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    //var ui = SpreadsheetApp.getUi();
    var allRecords = spreadsheet.getSheets()[1]
    var lastNames = allRecords.getRange("A2:A192322")
    var startingRow = 2
    var endingRow = 420

    var paidRecords = spreadsheet.getSheets()[2]

    for (i = startingRow; i < (endingRow+1); i++) {
        var lastNameString = paidRecords.getRange(i, 1).getDisplayValue()
        var firstNameString = paidRecords.getRange(i, 2).getDisplayValue()

        var textFinder = lastNames.createTextFinder(lastNameString)
        var searchResult = textFinder.matchEntireCell(true).findAll()
        var numOfLastNameMatches = searchResult.length
        var ui = SpreadsheetApp.getUi();

        if (numOfLastNameMatches > 0) {
            var firstRowLastNameMatch = searchResult[0].getRow()
            var firstNames = allRecords.getRange("B" + firstRowLastNameMatch + ":B" + (firstRowLastNameMatch + numOfLastNameMatches))
            textFinder = firstNames.createTextFinder(firstNameString)
            searchResult = textFinder.matchEntireCell(true).findAll()
            var numOfFirstNameMatches = searchResult.length
            if (numOfFirstNameMatches > 0) {
                var firstRowFirstNameMatch = searchResult[0].getRow()

                //var resultsSummary =    lastNameString + ", " + firstNameString + "\n" +
                //                        "Row number of the first last name match: " + firstRowLastNameMatch + "\n" + 
                //                        "Number of last name matches: " + numOfLastNameMatches + "\n" +
                //                        "Row number of the first first name match: " + firstRowFirstNameMatch + "\n" +
                //                        "Number of first name matches: " + numOfFirstNameMatches

                if (numOfFirstNameMatches == 1) {
                    var cell = paidRecords.getRange(i, 29)
                    cell.setValue("x")
                    allRecords.getRange("A" + firstRowFirstNameMatch + ":O" + firstRowFirstNameMatch).copyTo(paidRecords.getRange("L" + i + ":Z" + i))
                    // ui.alert(   resultsSummary + "\n" + 
                    //             "Perfect Match")
                }
                else if (numOfFirstNameMatches > 1) {
                    var cell = paidRecords.getRange(i, 28)
                    cell.setValue("x")
                    // ui.alert(   resultsSummary + "\n" + 
                    //             numOfFirstNameMatches + " Duplicate Matches")
                }
            }
            else {
                var cell = paidRecords.getRange(i, 30)
                cell.setValue("x")
                // ui.alert(   lastNameString + ", " + firstNameString + "\n" +
                //             "There were last name matches, but no first name matches")
            }
        }
        else {
            var cell = paidRecords.getRange(i, 30)
            cell.setValue("x")
            // ui.alert(   lastNameString + ", " + firstNameString + "\n" +
            //             "No last name Matches")
        }
    }
}