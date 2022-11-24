//Create menu in Google Sheets UI to run functions
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Automation");
  menu.addItem("Create Docs and PDFs", "main");
  menu.addToUi();
}

// Main function
function main() {
  const levels = [1, 2];
  const units = [1, 2, 3, 4, 5, 6];

  levels.forEach((level) => {
    // Tabs in the Sheet are named by level
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      `Level ${level}`
    );
    // The data in the sheet begins on the second row and first column
    // The number of rows and columns should be increased when more lessons are added
    const data = currentSheet.getRange(2, 1, 88, 13).getValues();
    // Loop over all units
    units.forEach((unit) => {
      const unit_number = Number(unit);
      // There are only 2 lessons in Unit 1, but 4 in all other units
      const lessons = unit_number === 1 ? [1, 2] : [1, 2, 3, 4];
      // Loop over all lessons
      lessons.forEach((lesson) => {
        const newFilename = `NELC${level}U${unit}L${lesson}`;
        const docID = createLessonDoc(level, unit, lesson, data, newFilename);
        createPDF(level, docID);
      });
    });
  });
}

// Create the Google Docs file for one lesson
function createLessonDoc(level, unit, lesson, data, newFilename) {
  // Set template and folder vars depending on level
  const scriptProperties = PropertiesService.getScriptProperties();
  let trashFolder = DriveApp.getFolderById(
    scriptProperties.getProperty("TRASH_FOLDER")
  );
  let templateDocFile, docFolder;
  if (level === 1) {
    templateDocFile = DriveApp.getFileById(
      scriptProperties.getProperty("TEMPLATE_DOC_L1")
    );
    docFolder = DriveApp.getFolderById(
      scriptProperties.getProperty("DOC_FOLDER_L1")
    );
  } else if (level === 2) {
    templateDocFile = DriveApp.getFileById(
      scriptProperties.getProperty("TEMPLATE_DOC_L2")
    );
    docFolder = DriveApp.getFolderById(
      scriptProperties.getProperty("DOC_FOLDER_L2")
    );
  } else {
    Logger.log(`Level ${level} is not valid.`);
  }
  // The starting row to use from the GS data depends on the unit and lesson numbers,
  //  and is complicated by the fact that Unit 1 only has 2 lessons.
  // So if we're doing Unit 1, we only worry about the lesson number. But if it's
  //  another unit, we need to add the rows for Unit 1 and then factor in the unit.
  const starting_row =
    unit === 1 ? 0 + (lesson - 1) * 4 : 8 + (unit - 2) * 16 + (lesson - 1) * 4;
  // The starting column is always the third one
  const starting_col = 3;
  const content = {
    level_number: Number(level),
    unit_number: Number(unit),
    lesson_number: Number(lesson),
    lesson_date: data[starting_row][starting_col],
    lesson_title: data[starting_row][starting_col + 1],
    introduction_en: data[starting_row][starting_col + 2],
    introduction_ja: data[starting_row][starting_col + 3],
    conversation1_en: data[starting_row][starting_col + 4],
    conversation1_ja: data[starting_row][starting_col + 5],
    conversation2_en: data[starting_row + 1][starting_col + 4],
    conversation2_ja: data[starting_row + 1][starting_col + 5],
    conversation3_en: data[starting_row + 2][starting_col + 4],
    conversation3_ja: data[starting_row + 2][starting_col + 5],
    conversation4_en: data[starting_row + 3][starting_col + 4],
    conversation4_ja: data[starting_row + 3][starting_col + 5],
    vocab1_en: data[starting_row][starting_col + 6],
    vocab1_ja: data[starting_row][starting_col + 7],
    vocab2_en: data[starting_row + 1][starting_col + 6],
    vocab2_ja: data[starting_row + 1][starting_col + 7],
    vocab3_en: data[starting_row + 2][starting_col + 6],
    vocab3_ja: data[starting_row + 2][starting_col + 7],
    extra1_en: data[starting_row][starting_col + 8],
    extra1_ja: data[starting_row][starting_col + 9],
    extra2_en: data[starting_row + 1][starting_col + 8],
    extra2_ja: data[starting_row + 1][starting_col + 9],
    extra3_en: data[starting_row + 2][starting_col + 8],
    extra3_ja: data[starting_row + 2][starting_col + 9],
  };

  // Check if a file already exists with this name, and move to trash if it does
  const oldFiles = docFolder.getFilesByName(newFilename);
  while (oldFiles.hasNext()) {
    // Setting as trashed is only possible for the file's owner, so other editors
    //  won't be able to run the script. Instead send to a shared folder for trash.
    // oldFiles.next().setTrashed(true);
    oldFiles.next().moveTo(trashFolder);
  }

  const newFile = templateDocFile.makeCopy(newFilename, docFolder);
  const newFileDoc = DocumentApp.openById(newFile.getId());
  const header = newFileDoc.getHeader();
  header.replaceText("{{lesson_date}}", content.lesson_date);
  const body = newFileDoc.getBody();
  body.replaceText("{{level_number}}", content.level_number);
  body.replaceText("{{unit_number}}", content.unit_number);
  body.replaceText("{{lesson_number}}", content.lesson_number);
  body.replaceText("{{lesson_title}}", content.lesson_title);
  body.replaceText("{{introduction_en}}", content.introduction_en);
  body.replaceText("{{introduction_ja}}", content.introduction_ja);
  body.replaceText("{{conversation1_en}}", content.conversation1_en);
  body.replaceText("{{conversation1_ja}}", content.conversation1_ja);
  body.replaceText("{{conversation2_en}}", content.conversation2_en);
  body.replaceText("{{conversation2_ja}}", content.conversation2_ja);
  body.replaceText("{{conversation3_en}}", content.conversation3_en);
  body.replaceText("{{conversation3_ja}}", content.conversation3_ja);
  body.replaceText("{{conversation4_en}}", content.conversation4_en);
  body.replaceText("{{conversation4_ja}}", content.conversation4_ja);
  body.replaceText("{{vocab1_en}}", content.vocab1_en);
  body.replaceText("{{vocab1_ja}}", content.vocab1_ja);
  body.replaceText("{{vocab2_en}}", content.vocab2_en);
  body.replaceText("{{vocab2_ja}}", content.vocab2_ja);
  body.replaceText("{{vocab3_en}}", content.vocab3_en);
  body.replaceText("{{vocab3_ja}}", content.vocab3_ja);
  body.replaceText("{{extra1_en}}", content.extra1_en);
  body.replaceText("{{extra1_ja}}", content.extra1_ja);
  body.replaceText("{{extra2_en}}", content.extra2_en);
  body.replaceText("{{extra2_ja}}", content.extra2_ja);
  body.replaceText("{{extra3_en}}", content.extra3_en);
  body.replaceText("{{extra3_ja}}", content.extra3_ja);
  newFileDoc.saveAndClose();
  return newFile.getId();
}

// Create the PDF file for one lesson
function createPDF(level, docFileID) {
  // Set folder var depending on level
  const scriptProperties = PropertiesService.getScriptProperties();
  let trashFolder = DriveApp.getFolderById(
    scriptProperties.getProperty("TRASH_FOLDER")
  );
  let pdfFolder;
  if (level === 1) {
    pdfFolder = DriveApp.getFolderById(
      scriptProperties.getProperty("PDF_FOLDER_L1")
    );
  } else if (level === 2) {
    pdfFolder = DriveApp.getFolderById(
      scriptProperties.getProperty("PDF_FOLDER_L2")
    );
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
    // Setting as trashed is only possible for the file's owner, so other editors
    //  won't be able to run the script. Instead send to a shared folder for trash.
    // oldFiles.next().setTrashed(true);
    oldFiles.next().moveTo(trashFolder);
  }
  docBlob.setName(pdfFileName);
  const file = pdfFolder.createFile(docBlob);
  return file.getId;
}
