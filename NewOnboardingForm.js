
//Creates a New Onboarding Sheet for A new individual

function FullNewTeamate() {

// Prompt for New Teamates Name - https://stackoverflow.com/questions/16464342/user-input-in-google-spreadsheet-script/52314815#52314815
  var NewTeamMateName = Browser.inputBox("What is the new Teamate's name?");
  
//Check to limit name to 40 caracters to ensure proper layout on check sheet and to forgo any sheet naming errors later as Google Sheets Only Allows 40 Character String 
  NewTeamMateName = NewTeamMateName.substring(0, 40);
  
//All instructions to affect Googlesheet have to be preceded by 'SpreadsheetApp.getActive()' so this is a way to decrease added typing.
  var spreadsheet = SpreadsheetApp.getActive();
  
//Unhides the Onboarding sheet so that the remainder of the function can execute
  spreadsheet.getSheetByName('Onboarding').showSheet().activate();
  
//Points to the sheet - mimicing the user clicking on the sheet to make it the active sheet & then Duplicating it.
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Onboarding'), true);
  spreadsheet.duplicateActiveSheet();

//Set's the sheet name equal to the New Teamate's name captured in the variable above via the input box. Once that's done we active the sheet as our active sheet.
  spreadsheet.getActiveSheet().setName(NewTeamMateName);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(NewTeamMateName), true);
  
//With the right sheet avtive we save the user name from the variable to the correct place in the active sheet.
  spreadsheet.getRange('G4').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText(NewTeamMateName)
  .setTextStyle(0, 4, SpreadsheetApp.newTextStyle()
  .setFontSize(24)
  .build())
  .build());
  spreadsheet.getRange('G3').activate();
  
// Moves Sheet to a position right after the   Dashboard so that the workbook is ordered by the most recent new starts to those that have started way back when
  spreadsheet.moveActiveSheet(4);
  
//  Refocusses on the Onboarding template & Hides it - taking you back to the newly created sheet.
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Onboarding'), true);
  spreadsheet.getActiveSheet().hideSheet();
};


