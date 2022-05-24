// Main function
function myFunction() {
  const levels = [1, 2];
  const units = [1, 2];

  levels.forEach((level) => {
    // Tabs in the Sheet are named by level
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      `Level ${level}`
    );
    // The data in the sheet begins on the second row and first column
    // The number of rows and columns should be increased when more lessons are added
    const data = currentSheet.getRange(2, 1, 24, 12).getValues();
    // Loop over all units
    units.forEach((unit) => {
      const unit_number = Number(unit);
      // There are only 2 lessons in Unit 1, but 4 in all other units
      const lessons = unit_number === 1 ? [1, 2] : [1, 2, 3, 4];
      // Loop over all lessons
      lessons.forEach((lesson) => {
        const newFilename = `Test - NELC${level}U${unit}L${lesson}`;
        const docID = createLessonDoc(level, unit, lesson, data, newFilename);
        createPDF(level, docID);
      });
    });
  });
}

// Create the Google Docs file for one lesson
function createLessonDoc(level, unit, lesson, data, newFilename) {
  // Set template and folder vars depending on level
  // Level 1 template: XXXX
  // Level 1 Doc folder: XXXX
  // Level 1 PDF folder: XXXX
  // Level 2 template: XXXX
  // Level 2 Doc folder: XXXX
  let templateDocFile, docFolder;
  if (level === 1) {
    templateDocFile = DriveApp.getFileById(
      "XXXX"
    );
    docFolder = DriveApp.getFolderById("XXXX");
  } else if (level === 2) {
    templateDocFile = DriveApp.getFileById(
      "XXXX"
    );
    docFolder = DriveApp.getFolderById("XXXX");
  } else {
    Logger.log(`Level ${level} is not valid.`);
  }
  const level_number = Number(level);
  const unit_number = Number(unit);
  const lesson_number = Number(lesson);
  // The starting row to use from the GS data depends on the unit and lesson numbers,
  //  and is complicated by the fact that Unit 1 only has 2 lessons.
  // So if we're doing Unit 1, we only worry about the lesson number. But if it's
  //  another unit, we need to add the rows for Unit 1 and then factor in the unit.
  const starting_row =
    unit_number === 1
      ? 0 + (lesson_number - 1) * 4
      : 8 + (unit_number - 2) * 16 + (lesson_number - 1) * 4;
  // The starting column is always the third one
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

  // Check if a file already exists with this name, and move to trash if it does
  const oldFiles = docFolder.getFilesByName(newFilename);
  while (oldFiles.hasNext()) {
    oldFiles.next().setTrashed(true);
  }

  const newFile = templateDocFile.makeCopy(newFilename, docFolder);
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
  return newFile.getId();
}

// Create the PDF file for one lesson
function createPDF(level, docFileID) {
  // Set folder var depending on level
  // Level 1 PDF folder: XXXX
  // Level 2 PDF folder: XXXX
  let pdfFolder;
  if (level === 1) {
    pdfFolder = DriveApp.getFolderById("XXXX");
  } else if (level === 2) {
    pdfFolder = DriveApp.getFolderById("XXXX");
  } else {
    Logger.log(`Level ${level} is not valid.`);
  }

  const doc = DriveApp.getFileById(docFileID);

  // first, we need to make a blob which will contain the data from the Doc as PDF
  docBlob = doc.getAs("application/pdf");
  pdfFileName = doc.getName() + ".pdf";
  // Check if a file already exists with this name, and move to trash if it does
  const oldFiles = pdfFolder.getFilesByName(pdfFileName);
  while (oldFiles.hasNext()) {
    oldFiles.next().setTrashed(true);
  }
  docBlob.setName(pdfFileName);
  const file = pdfFolder.createFile(docBlob);
  return file.getId;
}
