function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('✍️ Writer Tools')
    .addItem('📝 Import Word Counts', 'importWordCountsForAllWriters')
    .addToUi();
}


function importWordCountsForAllWriters() {
  const config = {
    sheetName: "google_sheet_name",
    startCell: "C2", // cell address to start filling the word count from
    writers: [
      { name: "writer_1", docId: "doc_ID_1" },
      { name: "writer_2", docId: "doc_ID_2" }
      //add more writers if you want
    ]
  };

  function parseDateToMMDDYYYY(text) {
    try {
      const cleaned = text
        .toLowerCase()
        .replace(/\b(\d{1,2})(st|nd|rd|th)\b/g, '$1')
        .replace(/\s+/g, ' ')
        .trim();

      const dmyMatch = cleaned.match(/^(\d{1,2}) (\w{3,}) (\d{2,4})$/);
      if (dmyMatch) {
        let [_, day, monthStr, year] = dmyMatch;
        const monthIndex = new Date(`${monthStr} 1, 2000`).getMonth();
        if (!isNaN(monthIndex)) {
          if (year.length === 2) year = '20' + year;
          const dateObj = new Date(year, monthIndex, day);
          if (!isNaN(dateObj)) return formatDate(dateObj);
        }
      }

      const mdyMatch = cleaned.match(/^(\w{3,}) (\d{1,2}),?\s*(\d{2,4})$/);
      if (mdyMatch) {
        let [_, monthStr, day, year] = mdyMatch;
        const monthIndex = new Date(`${monthStr} 1, 2000`).getMonth();
        if (!isNaN(monthIndex)) {
          if (year.length === 2) year = '20' + year;
          const dateObj = new Date(year, monthIndex, day);
          if (!isNaN(dateObj)) return formatDate(dateObj);
        }
      }

      const slashMatch = cleaned.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (slashMatch) {
        let [_, dd, mm, yy] = slashMatch;
        if (yy.length === 2) yy = '20' + yy;
        const dateObj = new Date(`${yy}-${mm}-${dd}`);
        if (!isNaN(dateObj)) return formatDate(dateObj);
      }

      const parsed = new Date(cleaned);
      if (!isNaN(parsed.getTime())) return formatDate(parsed);
    } catch (e) {
      Logger.log("Date parsing error: " + e.message);
    }
    return '';
  }

  function countWords(text) {
    const words = text.match(/[\p{L}\p{N}]+(?:[-'][\p{L}\p{N}]+)*/gu);
    return words ? words.length : 0;
  }

  function tryParseDate(value) {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
      return value;
    }
    try {
      const parsed = new Date(value);
      return isNaN(parsed) ? null : parsed;
    } catch {
      return null;
    }
  }

  function formatDate(date) {
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const yyyy = date.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
  const startRange = sheet.getRange(config.startCell);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();

  const names = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, 1).getValues().flat();
  const datesRow = sheet.getRange(1, startCol, 1, sheet.getLastColumn() - startCol + 1).getValues()[0];
  const sheetDates = datesRow.map(d => {
    const parsed = tryParseDate(d);
    return parsed ? formatDate(parsed) : null;
  });

  const updates = [];

  config.writers.forEach(writer => {
    const writerRowIndex = names.findIndex(name => name.toLowerCase().trim() === writer.name.toLowerCase().trim());
    if (writerRowIndex === -1) {
      Logger.log(`❌ Writer "${writer.name}" not found in sheet.`);
      return;
    }

    const rowNumber = startRow + writerRowIndex;
    const rowValues = sheet.getRange(rowNumber, startCol, 1, sheetDates.length).getValues()[0];

    const firstEmptyColIndex = rowValues.findIndex(val => val === "" || val === null);
    if (firstEmptyColIndex === -1) {
      Logger.log(`✅ ${writer.name} is already fully filled. Skipping.`);
      return;
    }

    const startDate = sheetDates[firstEmptyColIndex];
    if (!startDate) {
      Logger.log(`⚠️ Invalid sheet date for ${writer.name}. Skipping.`);
      return;
    }

    Logger.log(`Opening document: ${writer.docId}`);

    const doc = DocumentApp.openById(writer.docId);
    const paragraphs = doc.getBody().getParagraphs();

    const dateWordMap = {};
    let currentDate = '';
    let currentWordCount = 0;

    for (let p of paragraphs) {
      const text = p.getText().trim();
      if (!text) continue;

      if (p.getHeading() === DocumentApp.ParagraphHeading.TITLE) {
        if (currentDate && currentWordCount > 0 && currentDate >= startDate) {
          if (!dateWordMap[currentDate]) dateWordMap[currentDate] = 0;
          dateWordMap[currentDate] += currentWordCount;
        }

        currentDate = parseDateToMMDDYYYY(text);
        currentWordCount = countWords(text);
      } else {
        currentWordCount += countWords(text);
      }
    }

    if (currentDate && currentWordCount > 0 && currentDate >= startDate) {
      if (!dateWordMap[currentDate]) dateWordMap[currentDate] = 0;
      dateWordMap[currentDate] += currentWordCount;
    }

    for (const [date, count] of Object.entries(dateWordMap)) {
      const colIndex = sheetDates.findIndex(d => d === date);
      if (colIndex !== -1 && !rowValues[colIndex]) {
        updates.push({
          row: rowNumber,
          col: startCol + colIndex,
          value: count
        });
      }
    }

    Logger.log(`✅ Processed ${writer.name}`);
  });

  updates.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });

  SpreadsheetApp.getUi().alert(`✅ Word counts imported for ${config.writers.length} writers.`);
}