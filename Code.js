function myFunction() {
  const templateDocFileLevel1 = DriveApp.getFileById("XXXX");
  const docFolderLevel1 = DriveApp.getFolderById("XXXX");
  const level1Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Level 1");

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Level 1");
  const data = currentSheet.getRange(2,1, 4, 12).getValues();
  // Logger.log(data[0][3]);
  // Set vars
  const lesson_number = "1";
  const lesson_title = data[0][3];
  const introduction_en = data[0][4];
  const introduction_ja = data[0][5];
  const conversation1_en = data[0][6];
  const conversation1_ja = data[0][7];
  const conversation2_en = data[1][6];
  const conversation2_ja = data[1][7];
  const conversation3_en = data[2][6];
  const conversation3_ja = data[2][7];
  const conversation4_en = data[3][6];
  const conversation4_ja = data[3][7];
  const vocab1_en = data[0][8];
  const vocab1_ja = data[0][9];
  const vocab2_en = data[1][8];
  const vocab2_ja = data[1][9];
  const vocab3_en = data[2][8];
  const vocab3_ja = data[2][9];
  const extra1_en = data[0][10];
  const extra1_ja = data[0][11];
  const extra2_en = data[1][10];
  const extra2_ja = data[1][11];
  const extra3_en = data[2][10];
  const extra3_ja = data[2][11];

  // const rtRuns = conversation4_en.getRuns();
  // let str = "";
  // rtRuns.forEach(run => {
  //   str += run.getText();
  //   if(run.getTextStyle().isItalic()) {
  //     Logger.log(run.getStartIndex());
  //     Logger.log(run.getEndIndex());
  //   }
  // });
  // Logger.log(str);

  // make new temporary Doc
  const newFile = templateDocFileLevel1.makeCopy("Test - Level 1 - Lesson 1", docFolderLevel1);
  const newFileDoc = DocumentApp.openById(newFile.getId());
  const body = newFileDoc.getBody();
  body.replaceText("{{lesson_number}}", lesson_number);
  body.replaceText("{{lesson_title}}", lesson_title);
  body.replaceText("{{introduction_en}}", introduction_en);
  body.replaceText("{{introduction_ja}}", introduction_ja);
  body.replaceText("{{conversation1_en}}", conversation1_en);
  body.replaceText("{{conversation1_ja}}", conversation1_ja);
  body.replaceText("{{conversation2_en}}", conversation2_en);
  body.replaceText("{{conversation2_ja}}", conversation2_ja);
  body.replaceText("{{conversation3_en}}", conversation3_en);
  body.replaceText("{{conversation3_ja}}", conversation3_ja);
  body.replaceText("{{conversation4_en}}", conversation4_en);
  body.replaceText("{{conversation4_ja}}", conversation4_ja);
  body.replaceText("{{vocab1_en}}", vocab1_en);
  body.replaceText("{{vocab1_ja}}", vocab1_ja);
  body.replaceText("{{vocab2_en}}", vocab2_en);
  body.replaceText("{{vocab2_ja}}", vocab2_ja);
  body.replaceText("{{vocab3_en}}", vocab3_en);
  body.replaceText("{{vocab3_ja}}", vocab3_ja);
  body.replaceText("{{extra1_en}}", extra1_en);
  body.replaceText("{{extra1_ja}}", extra1_ja);
  body.replaceText("{{extra2_en}}", extra2_en);
  body.replaceText("{{extra2_ja}}", extra2_ja);
  body.replaceText("{{extra3_en}}", extra3_en);
  body.replaceText("{{extra3_ja}}", extra3_ja);
  newFileDoc.saveAndClose();
  
}
