function myFunction() {
  const templateDocFileLevel1 = DriveApp.getFileById("XXXX");
  const docFolderLevel1 = DriveApp.getFolderById("XXXX");
  const level1Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Level 1");

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Level 1");

  const levels = [1];
  const level = 1;
  const units = [1, 2];

  const data = currentSheet.getRange(2, 1, 24, 12).getValues();
  // Loop over all units
  units.forEach(unit => {
    const unit_number = Number(unit);
    const lessons = (unit_number === 1) ? [1, 2] : [1, 2, 3, 4];
    // Loop over all lessons (only 2 in Unit 1)
    lessons.forEach(lesson => {
      // Set vars
      const level_number = Number(level);
      const lesson_number = Number(lesson);
      const starting_row = (unit_number === 1) ?
                            0 + ((lesson_number - 1) * 4) :
                            8 + ((unit_number - 2) * 16) + ((lesson_number - 1) * 4);
      const starting_col = 3;
      const lesson_title = data[starting_row][starting_col];
      const introduction_en = data[starting_row][starting_col + 1];
      const introduction_ja = data[starting_row][starting_col + 2];
      const conversation1_en = data[starting_row][starting_col + 3];
      const conversation1_ja = data[starting_row][starting_col + 4];
      const conversation2_en = data[starting_row + 1][starting_col + 3];
      const conversation2_ja = data[starting_row + 1][starting_col + 4];
      const conversation3_en = data[starting_row + 2][starting_col + 3];
      const conversation3_ja = data[starting_row + 2][starting_col + 4];
      const conversation4_en = data[starting_row + 3][starting_col + 3];
      const conversation4_ja = data[starting_row + 3][starting_col + 4];
      const vocab1_en = data[starting_row][starting_col + 5];
      const vocab1_ja = data[starting_row][starting_col + 6];
      const vocab2_en = data[starting_row + 1][starting_col + 5];
      const vocab2_ja = data[starting_row + 1][starting_col + 6];
      const vocab3_en = data[starting_row + 2][starting_col + 5];
      const vocab3_ja = data[starting_row + 2][starting_col + 6];
      const extra1_en = data[starting_row][starting_col + 7];
      const extra1_ja = data[starting_row][starting_col + 8];
      const extra2_en = data[starting_row + 1][starting_col + 7];
      const extra2_ja = data[starting_row + 1][starting_col + 8];
      const extra3_en = data[starting_row + 2][starting_col + 7];
      const extra3_ja = data[starting_row + 2][starting_col + 8];

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
      const newFile = templateDocFileLevel1.makeCopy(`Test - NELC${level_number}U${unit_number}L${lesson_number}`, docFolderLevel1);
      const newFileDoc = DocumentApp.openById(newFile.getId());
      const body = newFileDoc.getBody();
      body.replaceText("{{level_number}}", level_number);
      body.replaceText("{{unit_number}}", unit_number);
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
    })
  });
}
