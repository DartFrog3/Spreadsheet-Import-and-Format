function onInstall(event) {
    onOpen(event);
}

function onOpen(event) {
    main();
}

function main() {
    // Setup
    var targetsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    //Variable Setup
    var nums = [];
    var list = ["PlaceHolder", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    var Acolumn = "";
    var count = 7;
    
    //UI Input
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Datasheet URL', 'Please input the URL of the Datasheet', ui.ButtonSet.YES_NO);
    
    // Process the user's response
    if (response.getSelectedButton() === ui.Button.YES) {
        var datasheet_link = response.getResponseText();
    } else if (response.getSelectedButton() === ui.button.NO) {
        Logger.log('The user clicked the close button in the dialog\'s title bar.');
    } else {
        Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
    
    //Datasheet Setup
    var datasheet = SpreadsheetApp.openByUrl(datasheet_link);
    var sheet = datasheet.getActiveSheet();
    
    // Sizing
    var LastRow = datasheet.getLastRow();
    var LastColumn = datasheet.getLastColumn();
    for (var i = 1; i <= 26; i++) {
        list.push("A" + list[i]);
    }
    
    //Back-Up Sizing
    if (LastColumn > 52) {
        for (i = 1; i <= 26; i++) {
            list.push("B" + list[i]);
        }
    }
    
    // Create Range
    for (x = 1; x <= 52; x++) {
        nums.push(x);
    }
    
    //A1-ify
    for each (number in nums) {
        if (number === LastColumn + 1) {
            Acolumn = list[number];
            LastColumn = list[number - 1];
        }
    }
    
    //A1 Format Size Variable
    var size = "A1:" + LastColumn + LastRow;
    
    //Recording && Inputing Data
    var test = targetsheet.getRange("A1").getValue();
    var data = sheet.getRange(size).getValues();
    targetsheet.getRange(size).setValues(data);
    
    //Adding Padding
    targetsheet.insertRowsBefore(1, 4);
    targetsheet.insertColumnBefore(1);
    
    //Adjust Size Variable
    var LastRow = targetsheet.getLastRow();
    var LastColumn = targetsheet.getLastColumn();
    //A1-ify
    for each (number in nums) {
        if (number === LastColumn) {
            Acolumn = list[number];
            var Bcolumn = list[number + 1];
        }
    }
    var Asize = "A1:" + Acolumn + LastRow;
    
    //Remove Gridlines
    targetsheet.setHiddenGridlines(true);
    
    //Decide Type of Sheet
    var end = false;
    var q = 1;
    while (end === false) {
        if (sheet.getRange(list[q] + 1).getValue() === "Client" || sheet.getRange(list[q] + 1).getValue() === "Project") {
            var PRsheet = true;
            end = true;
        } else {
            q++;
        }
        
        if (q > 52 && PRsheet !== true) {
            var TMsheet = true;
            end = true;
        }
    }
    
    if (PRsheet === true) {
        
        //Formatting
        targetsheet.insertRowBefore(6);
        targetsheet.insertRowBefore(32);
        targetsheet.getRange("E7:" + Acolumn + (LastRow + 2)).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DASHED);
        
        //More Formatting
        targetsheet.getRange("E7:" + Acolumn + LastRow).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DASHED);
        targetsheet.getRange("B2").setValue("Project Resourcing").setFontColor("#0b5394").setFontWeight("bold");
        targetsheet.getRange("C32:" + Acolumn + 32).clearFormat();
        targetsheet.getRange(Asize).setFontFamily("Roboto");
        targetsheet.getRange("B2").setFontSize(30);
        targetsheet.getRange("B5:" + Acolumn + 5).setFontSize(9).setFontColor("white").setBackground("#0b5394").setFontWeight("bold");
        targetsheet.getRange("B6").setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Signed Projects");
        targetsheet.getRange("B3:" + Acolumn + 3).setBorder(true, null, null, null, null, null, "#0b5394", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B6:" + Acolumn + 6).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B32").setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Opportunities");
        targetsheet.getRange("B32:" + Acolumn + 32).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        //Conditional Formatting
        targetsheet.setColumnWidth(2, 290);
        targetsheet.setColumnWidth(3, 260);
        targetsheet.setColumnWidth(4, 160);
        targetsheet.setColumnWidths(5, LastColumn - 5, 75);
        targetsheet.setRowHeight(6, 66);
        targetsheet.setRowHeight(32, 66);
    } else if (TMsheet === true) {
        
        //Formatting
        var v = 16;
        targetsheet.insertRowBefore(5);
        targetsheet.insertRowBefore(7);
        targetsheet.insertRowBefore(11);
        targetsheet.insertRowBefore(v);
        targetsheet.insertRowBefore(32);
        targetsheet.getRange("D7:" + Acolumn + (LastRow + 5)).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DASHED).setBackground("#fce8b2");
        
        //More Formatting
        targetsheet.getRange("D7:" + Acolumn + 7).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DASHED).setBackground("#fce8b2");
        targetsheet.getRange("B2").setValue("TEAM MEMBER RESOURCING").setFontColor("#0b5394").setFontWeight("bold");
        targetsheet.getRange(Asize).setFontFamily("Roboto");
        targetsheet.getRange("B2").setFontSize(30);
        targetsheet.getRange("D7:D" + LastRow).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.DOTTED);
        targetsheet.getRange("B6:" + Acolumn + 6).setFontSize(9).setFontColor("white").setBackground("#0b5394").setFontWeight("bold");
        targetsheet.getRange("C7:" + Acolumn + 7).clearFormat();
        targetsheet.getRange("C32:" + Acolumn + 32).clearFormat();
        targetsheet.getRange("B7").setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Employees: 100% Billable");
        targetsheet.getRange("C11:" + Acolumn + 11).clearFormat();
        targetsheet.getRange("B11").setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Employees: 50% Billable");
        targetsheet.getRange("C" + v + ":" + Acolumn + v).clearFormat();
        targetsheet.getRange("B" + v).setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Contractors");
        targetsheet.getRange("B3:" + Acolumn + 3).setBorder(true, null, null, null, null, null, "#0b5394", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B7:" + Acolumn + 7).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B11:" + Acolumn + 11).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B" + v + ":" + Acolumn + v).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("B32").setFontSize(16).setFontWeight("bold").setVerticalAlignment("bottom").setValue("Placeholders");
        targetsheet.getRange("B32:" + Acolumn + 32).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        targetsheet.getRange("C4").setValue("Under-Utilized").setFontColor("#f09300");
        targetsheet.getRange("E4").setValue("On-Target (90 - 110%)").setFontColor("#0b8043");
        targetsheet.getRange("G4").setValue("Over-Utilized").setFontColor("#c53929");
        
        //Conditional Formatting
        targetsheet.setColumnWidth(2, 140);
        targetsheet.setColumnWidth(3, 160);
        targetsheet.setColumnWidths(4, LastColumn - 3, 75);
        targetsheet.setRowHeight(7, 66);
        targetsheet.setRowHeight(32, 66);
        targetsheet.setRowHeight(11, 66);
        targetsheet.setRowHeight(v, 66);
        targetsheet.setColumnWidth(1, 33);
        targetsheet.setRowHeight(1, 33);
        
        while( count < LastRow + 6) {
            if (targetsheet.getRange("B" + count).getFontWeight() === "bold") {
                count += 2;
            } else {
                targetsheet.getRange("B" + count + ":C" + count).setBackground("#d9d9d9");
                count += 2;
            }
        }
        
        //In-Sheet Conditional Formatting
        var x = 8;
        var z = 4;
        while (z <= 52) {
            if (targetsheet.getRange(list[z] + x).getValue() >= 36 && targetsheet.getRange(list[z] + x).getValue() <= 44) {
                targetsheet.getRange(list[z] + x).setBackground("#b7e1cd");
                x++;
            } else if (targetsheet.getRange(list[z] + x).getValue() > 44) {
                targetsheet.getRange(list[z] + x).setBackground("#f4c7c3");
                x++;
            } else {
                x++;
            }
            
            if (x > LastRow + 6) {
                x = 8;
                z++;
            }
        }
    }
}

