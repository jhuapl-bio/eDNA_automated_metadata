// ##############################################################################################
// # Copyright 2026 The Johns Hopkins University Applied Physics Laboratory LLC
// # All rights reserved.
// # Permission is hereby granted, free of charge, to any person obtaining a copy of this
// # software and associated documentation files (the "Software"), to deal in the Software
// # without restriction, including without limitation the rights to use, copy, modify,
// # merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
// # permit persons to whom the Software is furnished to do so.
// #
// # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// # INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// # PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// # LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// # TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
// # OR OTHER DEALINGS IN THE SOFTWARE.
// #

const CONFIG = {
  SHEET_NAME: 'samplesheet_v1',
  HEADER_ROW: 4,
  LINK_COLUMN_LETTER: 'B',

  // Paste your deployed Apps Script Web App URL here. (See Google Apps Script instructions for how to generate)
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycby1OIj46v4scWWeWbg8ThATgIMjngda2gY9FACy-tfiU9kloKlCwWFRop8nLFJ2mLBqmQ/exec',

  // Google Drive folder containing all BTL files.
  CTD_FOLDER_NAME: 'CTD',
  ROV_FOLDER_NAME: 'ROV',
  FLOWTHROUGH_FOLDER_NAME: 'Flowthrough',

  // DYMO label variables
  LABEL_TEXT_TOP_COLUMN: 'sample_ID',
  LABEL_TEXT_BOTTOM_COLUMN: 'project_ID',
  LABEL_TEXT_EXTRA_2_COLUMN: 'dive_or_cast_number',
  LABEL_TEXT_EXTRA_3_COLUMN: 'asset',
  LABEL_TEXT_EXTRA_4_COLUMN: 'collection_method',
  LABEL_TEXT_EXTRA_5_COLUMN: 'bottle_number',

  // Map data
  LATITUDE_HEADER: 'decimalLatitude',
  LONGITUDE_HEADER: 'decimalLongitude',
  LABEL_HEADER: 'sample_ID',

  // CTD sheet headers
  CTD_ASSET_HEADER: 'asset',
  CTD_DIVE_OR_CAST_HEADER: 'dive_or_cast_number',
  CTD_BOTTLE_NUMBER_HEADER: 'bottle_number',
  CTD_SALINITY_HEADER: 'salinity',
  CTD_DEPTH_HEADER: 'PrDM',
  CTD_ZTIME_HEADER: 'z-time',
  CTD_TEMP_HEADER: 'temp',
  CTD_O2_HEADER: 'diss_oxygen',
  CTD_O2_UNIT_HEADER: 'diss_oxygen_unit',
  CTD_CHLOR_HEADER: 'chlorophyll',
  CTD_VERBATIM_LATITUDE_HEADER: 'verbatimLatitude',
  CTD_VERBATIM_LONGITUDE_HEADER: 'verbatimLongitude',

  FLOWTHROUGH_ASSET_HEADER: 'asset',
  FLOWTHROUGH_EVENT_DATE_HEADER: 'eventDate (YYYY-MM-DD)'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Import Data')
    .addItem('Import CTD Data from BTL', 'importCTDDataFromBTL')
    .addItem('Import ROV Data from CSV/TXT', 'importROVDataFromROV')
    .addItem('Import Flowthrough Data from TXTs', 'importFlowthroughDataFromFiles')
    .addToUi();

  ui.createMenu('DYMO Labels')
    .addItem('Generate Label PDF from Selected Rows', 'showPdfSidebar')
    .addToUi();

  ui.createMenu('Sample Map')
    .addItem('Show highlighted samples in sidebar', 'showMapSidebar')
    .addItem('Open highlighted samples in new tab', 'openMapInNewTab')
    .addToUi();
}

function showPdfSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('PdfSidebar')
    .setTitle('Generate DYMO PDF');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHeaders_(sheet) {
  return sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());
}

function normalizeHeader_(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

function normalizeBTLHeader_(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function parseNmeaDdm_(text) {
  const cleaned = String(text || '').trim();
  if (!cleaned) return '';

  const match = cleaned.match(/^([+-]?\d{1,3})\s+(\d{1,2}(?:\.\d+)?)\s*([NSEW])?$/i);
  if (!match) return '';

  const degrees = Math.abs(Number(match[1]));
  const minutes = Number(match[2]);
  const hemi = String(match[3] || '').toUpperCase();

  let decimal = degrees + (minutes / 60);

  if (hemi === 'S' || hemi === 'W' || Number(match[1]) < 0) {
    decimal = -decimal;
  }

  return decimal;
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const normalized = normalizeHeader_(h);
    if (normalized) map[normalized] = i + 1;
  });
  return map;
}

function getColumnIndexOrThrow_(headerMap, headerName) {
  const col = headerMap[normalizeHeader_(headerName)];
  if (!col) {
    throw new Error(`Missing required column in sheet: ${headerName}`);
  }
  return col;
}

function getSelectedLabelRows() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error(`Select rows on the "${CONFIG.SHEET_NAME}" sheet first.`);
  }

  const range = sheet.getActiveRange();
  if (!range) throw new Error('No rows selected.');

  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const lastCol = sheet.getLastColumn();

  if (startRow <= CONFIG.HEADER_ROW) {
    throw new Error(`Select data rows below header row ${CONFIG.HEADER_ROW}.`);
  }

  const headers = getHeaders_(sheet);
  const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

  return values.map((row, i) => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx];
    });

    const actualRow = startRow + i;
    const samplePageUrl = `${CONFIG.WEB_APP_URL}?row=${actualRow}`;

    return {
      barcode_value: samplePageUrl,
      text_top: String(obj[CONFIG.LABEL_TEXT_TOP_COLUMN] || ''),
      text_bottom: String(obj[CONFIG.LABEL_TEXT_BOTTOM_COLUMN] || ''),
      text_extra_2:
        'Asset/dive: ' +
        String(obj[CONFIG.LABEL_TEXT_EXTRA_3_COLUMN] || ' ') +
        String(obj[CONFIG.LABEL_TEXT_EXTRA_2_COLUMN] || ''),
      text_extra_3:
        'Collection: ' +
        String(obj[CONFIG.LABEL_TEXT_EXTRA_4_COLUMN] || '') +
        String(obj[CONFIG.LABEL_TEXT_EXTRA_5_COLUMN] || '')
    };
  });
}

// UPDATE FOR MAP

function showMapSidebar() {
  const samples = getSelectedSamplesForMap();

  if (!samples.length) {
    throw new Error('No valid samples found in the selected rows.');
  }

  const template = HtmlService.createTemplateFromFile('MapSidebar');
  template.samplesJson = JSON.stringify(samples);

  const html = template.evaluate().setTitle('Sample Map');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSelectedSamplesForMap() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error(`Select rows on the "${CONFIG.SHEET_NAME}" sheet first.`);
  }

  const range = sheet.getActiveRange();
  if (!range) {
    throw new Error('Select one or more rows first.');
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  if (startRow <= CONFIG.HEADER_ROW) {
    throw new Error(`Select data rows below header row ${CONFIG.HEADER_ROW}.`);
  }

  const headers = getHeaders_(sheet);
  const headerIndex = {};

  headers.forEach((h, i) => {
    headerIndex[h] = i;
  });

  const latIdx = headerIndex[CONFIG.LATITUDE_HEADER];
  const lngIdx = headerIndex[CONFIG.LONGITUDE_HEADER];
  const labelIdx = headerIndex[CONFIG.LABEL_HEADER];

  if (latIdx === undefined || lngIdx === undefined) {
    throw new Error('Latitude/longitude headers not found.');
  }

  const values = sheet
    .getRange(startRow, 1, numRows, sheet.getLastColumn())
    .getValues();

  return values
    .map((row, i) => ({
      row: startRow + i,
      label: row[labelIdx] || `Row ${startRow + i}`,
      lat: Number(row[latIdx]),
      lng: Number(row[lngIdx])
    }))
    .filter(s => Number.isFinite(s.lat) && Number.isFinite(s.lng));
}

function openMapInNewTab() {
  const samples = getSelectedSamplesForMap();

  if (!samples.length) {
    throw new Error('No valid samples found in the selected rows.');
  }

  const encoded = encodeURIComponent(JSON.stringify(samples));

  const url = `${CONFIG.WEB_APP_URL}?mapData=${encoded}`;

  const html = HtmlService
    .createHtmlOutput(`
      <script>
        window.open('${url}', '_blank');
        google.script.host.close();
      </script>
    `);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Map');
}

function renderMapPage_(mapDataEncoded) {
  const html = HtmlService.createTemplateFromFile('FullPageMap');

  html.samplesJson = decodeURIComponent(mapDataEncoded);

  return html
    .evaluate()
    .setTitle('Sample Map');
}

// END OF UPDATE FOR MAP

// ==============================
// CTD / BTL IMPORT
// ==============================

function importCTDDataFromBTL() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert(`Select rows on the "${CONFIG.SHEET_NAME}" sheet first.`);
    return;
  }

  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Please highlight one or more rows first.');
    return;
  }

  const headers = getHeaders_(sheet);
  const headerMap = buildHeaderMap_(headers);

  const requiredColumns = [
    'asset',
    'dive_or_cast_number',
    'bottle_number',
    'salinity',
    'PrDM',
    'z-time',
    'temp',
    'diss_oxygen',
    'diss_oxygen_unit',
    'chlorophyll',
    'verbatimLatitude',
    'verbatimLongitude',
    'decimalLatitude',
    'decimalLongitude'
  ];

  for (const colName of requiredColumns) {
    const normalized = normalizeHeader_(colName);
    if (!(normalized in headerMap)) {
      ui.alert(`Missing required column in sheet: ${colName}`);
      return;
    }
  }

  // Resolve columns once so getRange(row, null) cannot happen
  const verbLatCol = getColumnIndexOrThrow_(headerMap, 'verbatimLatitude');
  const verbLonCol = getColumnIndexOrThrow_(headerMap, 'verbatimLongitude');
  const decLatCol = getColumnIndexOrThrow_(headerMap, 'decimalLatitude');
  const decLonCol = getColumnIndexOrThrow_(headerMap, 'decimalLongitude');

  const selectedRanges = rangeList.getRanges();
  const rowsToProcess = new Set();

  for (const range of selectedRanges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();

    for (let r = 0; r < numRows; r++) {
      const row = startRow + r;
      if (row > CONFIG.HEADER_ROW) {
        rowsToProcess.add(row);
      }
    }
  }

  if (rowsToProcess.size === 0) {
    ui.alert(`Select data rows below header row ${CONFIG.HEADER_ROW}.`);
    return;
  }

  const btlCache = {};
  const rowList = Array.from(rowsToProcess).sort((a, b) => a - b);

  for (const row of rowList) {
    const asset = String(
      sheet.getRange(
        row,
        headerMap[normalizeHeader_(CONFIG.CTD_ASSET_HEADER)]
      ).getValue()
    ).trim().toUpperCase();

    if (asset !== 'CTD') {
      ui.alert('asset for one or more samples does not equal CTD. Please check highlighted rows and try again.');
      return;
    }

    const diveOrCastValue = String(
      sheet.getRange(
        row,
        headerMap[normalizeHeader_(CONFIG.CTD_DIVE_OR_CAST_HEADER)]
      ).getDisplayValue()
    ).trim();

    if (!diveOrCastValue) {
      ui.alert(`Row ${row} is missing a dive_or_cast_number value.`);
      return;
    }

    const bottleCellValue = sheet
      .getRange(
        row,
        headerMap[normalizeHeader_(CONFIG.CTD_BOTTLE_NUMBER_HEADER)]
      )
      .getDisplayValue();

    const bottleNumbers = extractBottleNumbers_(bottleCellValue);

    if (!bottleNumbers.length) {
      continue;
    }

    const lowestBottle = Math.min(...bottleNumbers);

    if (!btlCache[diveOrCastValue]) {
      const btlFile = findBtlFileForDiveOrCast_(diveOrCastValue);
      btlCache[diveOrCastValue] = parseBTLFileFromFile_(btlFile);
    }

    const btlPackage = btlCache[diveOrCastValue];
    const btlData = btlPackage.records;
    const match = btlData[lowestBottle];

    if (!match) {
      Logger.log(`Bottle ${lowestBottle} not found in BTL file for row ${row} (dive/cast ${diveOrCastValue}).`);
      continue;
    }

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_SALINITY_HEADER)]
    ).setValue(match.sal00);

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_DEPTH_HEADER)]
    ).setValue(match.prdm);

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_ZTIME_HEADER)]
    ).setValue(match.zTime);

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_TEMP_HEADER)]
    ).setValue(match.temp);

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_O2_HEADER)]
    ).setValue(match.dissOxygen);

    if (Number.isFinite(match.dissOxygen)) {

      sheet.getRange(
        row,
        headerMap[normalizeHeader_(CONFIG.CTD_O2_UNIT_HEADER)]
      ).setValue('mmol/kg');
    }

    sheet.getRange(
      row,
      headerMap[normalizeHeader_(CONFIG.CTD_CHLOR_HEADER)]
    ).setValue(match.chlorophyll);

    sheet.getRange(row, verbLatCol).setValue(btlPackage.metadata.verbatimLatitude);
    sheet.getRange(row, verbLonCol).setValue(btlPackage.metadata.verbatimLongitude);
    sheet.getRange(row, decLatCol).setValue(btlPackage.metadata.decimalLatitude);
    sheet.getRange(row, decLonCol).setValue(btlPackage.metadata.decimalLongitude);
  }

  ui.alert('CTD data import complete.');
}

function importROVDataFromROV() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert(`Select rows on the "${CONFIG.SHEET_NAME}" sheet first.`);
    return;
  }

  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Please highlight one or more rows first.');
    return;
  }

  const headers = getHeaders_(sheet);
  const headerMap = buildHeaderMap_(headers);

  const requiredColumns = [
    'asset',
    'collection_method',
    'dive_or_cast_number',
    'bottle_number',
    'z-time',
    'temp',
    'chlorophyll',
    'diss_oxygen',
    'diss_oxygen_unit',
    'minimumDepthInMeters',
    'maximumDepthInMeters',
    'salinity'
  ];

  for (const colName of requiredColumns) {
    const normalized = normalizeHeader_(colName);
    if (!(normalized in headerMap)) {
      ui.alert(`Missing required column in sheet: ${colName}`);
      return;
    }
  }

  const assetCol = getColumnIndexOrThrow_(headerMap, 'asset');
  const collectionCol = getColumnIndexOrThrow_(headerMap, 'collection_method');
  const diveCol = getColumnIndexOrThrow_(headerMap, 'dive_or_cast_number');
  const bottleCol = getColumnIndexOrThrow_(headerMap, 'bottle_number');
  const zTimeCol = getColumnIndexOrThrow_(headerMap, 'z-time');
  const tempCol = getColumnIndexOrThrow_(headerMap, 'temp');
  const chlorophyllCol = getColumnIndexOrThrow_(headerMap, 'chlorophyll');
  const o2Col = getColumnIndexOrThrow_(headerMap, 'diss_oxygen');
  const o2UnitCol = getColumnIndexOrThrow_(headerMap, 'diss_oxygen_unit');
  const minDepthCol = getColumnIndexOrThrow_(headerMap, 'minimumDepthInMeters');
  const maxDepthCol = getColumnIndexOrThrow_(headerMap, 'maximumDepthInMeters');
  const salinityCol = getColumnIndexOrThrow_(headerMap, 'salinity');

  const decLatCol = getColumnIndexOrThrow_(headerMap, 'decimalLatitude');
  const decLonCol = getColumnIndexOrThrow_(headerMap, 'decimalLongitude');

  const selectedRanges = rangeList.getRanges();
  const rowsToProcess = new Set();

  for (const range of selectedRanges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();

    for (let r = 0; r < numRows; r++) {
      const row = startRow + r;
      if (row > CONFIG.HEADER_ROW) {
        rowsToProcess.add(row);
      }
    }
  }

  if (rowsToProcess.size === 0) {
    ui.alert(`Select data rows below header row ${CONFIG.HEADER_ROW}.`);
    return;
  }

  const rovCache = {};
  const rowList = Array.from(rowsToProcess).sort((a, b) => a - b);

  for (const row of rowList) {
    const asset = String(sheet.getRange(row, assetCol).getValue()).trim().toUpperCase();
    const collectionMethod = String(sheet.getRange(row, collectionCol).getValue()).trim().toUpperCase();

    if (asset !== 'ROV' || collectionMethod !== 'NISKIN') {
      ui.alert('asset for one or more samples does not equal ROV or collection_method does not equal Niskin. Please check highlighted rows and try again.');
      return;
    }

    const diveOrCastValue = String(sheet.getRange(row, diveCol).getDisplayValue()).trim();
    if (!diveOrCastValue) {
      ui.alert(`Row ${row} is missing a dive_or_cast_number value.`);
      return;
    }

    const bottleCellValue = sheet.getRange(row, bottleCol).getDisplayValue();
    const bottleNumbers = extractBottleNumbers_(bottleCellValue);
    if (!bottleNumbers.length) {
      continue;
    }

    const lowestBottle = Math.min(...bottleNumbers);

    if (!rovCache[diveOrCastValue]) {
      const csvFile = findRovFileForDiveOrCast_(diveOrCastValue, 'csv');
      const txtFile = findRovFileForDiveOrCast_(diveOrCastValue, 'txt');

      rovCache[diveOrCastValue] = {
        csvRecords: parseRovCsvFile_(csvFile),
        txtRows: parseRovTxtFile_(txtFile)
      };

      Logger.log(
        `ROV cache built for ${diveOrCastValue}: CSV records=${Object.keys(rovCache[diveOrCastValue].csvRecords).length}, TXT rows=${rovCache[diveOrCastValue].txtRows.length}`
      );
    }

    const rovPackage = rovCache[diveOrCastValue];
    const csvRecord = rovPackage.csvRecords[lowestBottle];

    if (!csvRecord) {
      Logger.log(`No CSV record found for bottle ${lowestBottle} in dive/cast ${diveOrCastValue}.`);
      continue;
    }

    const zTime = String(csvRecord.ts || '').trim();
    if (!zTime) {
      Logger.log(`Blank ts for bottle ${lowestBottle} in ${diveOrCastValue}.`);
      continue;
    }

    const txtMatch = findClosestRovTemp_(rovPackage.txtRows, zTime);
    const latitude = csvRecord.latitude;
    const longitude = csvRecord.longitude;

    Logger.log(
      `ROV row ${row}: bottle=${lowestBottle}, zTime=${zTime}, temp=${txtMatch ? txtMatch.temp : 'NaN'}`
    );

    sheet.getRange(row, zTimeCol).setValue(zTime);

    if (Number.isFinite(latitude)) {
      sheet.getRange(row, decLatCol).setValue(latitude);
    }

    if (Number.isFinite(longitude)) {
      sheet.getRange(row, decLonCol).setValue(longitude);
    }

    if (txtMatch) {
      if (Number.isFinite(txtMatch.temp)) {
        sheet.getRange(row, tempCol).setValue(txtMatch.temp);
      }

      if (Number.isFinite(txtMatch.chlorophyll)) {
        sheet.getRange(row, chlorophyllCol).setValue(txtMatch.chlorophyll);
      }

      if (Number.isFinite(txtMatch.dissOxygen)) {
        sheet.getRange(row, o2Col).setValue(txtMatch.dissOxygen);
        sheet.getRange(row, o2UnitCol).setValue('mmol/kg');
      }

      if (Number.isFinite(txtMatch.depth)) {
        sheet.getRange(row, minDepthCol).setValue(txtMatch.depth);
        sheet.getRange(row, maxDepthCol).setValue(txtMatch.depth);
      }

      if (Number.isFinite(txtMatch.salinity)) {
        sheet.getRange(row, salinityCol).setValue(txtMatch.salinity);
      }
    } else {
      Logger.log(`No valid temperature match for ${expectedSampleId} in ${diveOrCastValue}.`);
    }
  }

  ui.alert('ROV data import complete.');
}

function findBtlFileForDiveOrCast_(diveOrCastValue) {
  const suffix = String(diveOrCastValue).trim();
  if (!suffix) {
    throw new Error('Missing dive_or_cast_number value.');
  }

  const folderIter = DriveApp.getFoldersByName(CONFIG.CTD_FOLDER_NAME);
  const matches = [];

  while (folderIter.hasNext()) {
    const folder = folderIter.next();
    findBtlFilesInFolderTree_(folder, suffix, matches);
  }

  if (matches.length === 0) {
    throw new Error(
      `No .btl file found ending with "${suffix}.btl" inside Drive folder "${CONFIG.CTD_FOLDER_NAME}".`
    );
  }

  if (matches.length > 1) {
    Logger.log(
      `Multiple BTL matches for suffix "${suffix}": ${matches.map(f => f.getName()).join(', ')}. Using the first match.`
    );
  }

  return matches[0];
}

function findBtlFilesInFolderTree_(folder, suffix, matches) {
  const files = folder.getFiles();
  const suffixLower = `${suffix}.btl`.toLowerCase();

  while (files.hasNext()) {
    const file = files.next();
    const nameLower = String(file.getName()).toLowerCase();

    if (nameLower.endsWith(suffixLower)) {
      matches.push(file);
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    findBtlFilesInFolderTree_(subfolders.next(), suffix, matches);
  }
}

function parseBTLFileFromFile_(file) {
  const content = file.getBlob().getDataAsString();
  const lines = content.split(/\r?\n/);

  const latLine = lines.find(line => /^\*\s*NMEA Latitude\s*=/.test(String(line || '')));
  const lonLine = lines.find(line => /^\*\s*NMEA Longitude\s*=/.test(String(line || '')));

  const metadata = {
    verbatimLatitude: latLine ? String(latLine).split('=')[1].trim() : '',
    verbatimLongitude: lonLine ? String(lonLine).split('=')[1].trim() : ''
  };

  metadata.decimalLatitude = parseNmeaDdm_(metadata.verbatimLatitude);
  metadata.decimalLongitude = parseNmeaDdm_(metadata.verbatimLongitude);

  let headerIndex = -1;
  let headers = [];

  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || '').trim();
    if (!line) continue;

    if (line.includes('Bottle') && line.includes('Sal00')) {
      headerIndex = i;
      headers = line.trim().split(/\s+/);  // headers on 1+ spaces
      break;
    }
  }

  if (headerIndex === -1) {
    throw new Error('Could not find the BTL header row containing Bottle and Sal00.');
  }

  const bottleIdx = headers.findIndex(h => normalizeBTLHeader_(h) === 'bottle');
  const dateIdx = headers.findIndex(h => normalizeBTLHeader_(h) === 'date');
  const sal00Idx = headers.findIndex(h => normalizeBTLHeader_(h).startsWith('sal00'));
  const prdmIdx = headers.findIndex(h => normalizeBTLHeader_(h) === 'prdm');
  const tempIdx = headers.findIndex(h => normalizeBTLHeader_(h) === 't090c');
  const oxyIdx = headers.findIndex(h => {
    const normalized = normalizeBTLHeader_(h);
    return normalized.includes('sbox0') || normalized.includes('sbeox0mm');
  });
  const chlIdx = headers.findIndex(h => normalizeBTLHeader_(h).includes('fleco'));

  if (bottleIdx === -1) throw new Error('Could not find Bottle column in BTL file.');
  if (dateIdx === -1) throw new Error('Could not find Date column in BTL file.');
  if (sal00Idx === -1) throw new Error('Could not find first Sal00 column in BTL file.');
  if (prdmIdx === -1) throw new Error('Could not find PrDM column in BTL file.');
  if (tempIdx === -1) throw new Error('Could not find T090C column in BTL file.');
  if (oxyIdx === -1) throw new Error('Could not find Sbox0Mm/Kg column in BTL file.');
  if (chlIdx === -1) throw new Error('Could not find FlECO-AFL column in BTL file.');

  const data = {};

  for (let i = headerIndex + 1; i < lines.length; i++) {
    const line = String(lines[i] || '').trim();
    if (!line) continue;

    // Bottle average rows start with the bottle number.
    if (!/^\d+\s/.test(line)) continue;

    const parts = line.split(/\s{2,}/); // data rows on 2+ spaces

    if (
      parts.length <= Math.max(
        bottleIdx,
        dateIdx,
        sal00Idx,
        prdmIdx,
        tempIdx,
        oxyIdx,
        chlIdx
      )
    ) continue;

    const bottle = Number(parts[bottleIdx]);
    if (!Number.isFinite(bottle)) continue;

    // Time is on the next line after the average row.
    const nextLine = String(lines[i + 1] || '').trim();
    let timeText = '';

    if (nextLine) {
      const nextParts = nextLine.split(/\s{2,}/);
      const timeMatch = String(nextParts[0] || '').trim().match(/^(\d{1,2}:\d{2}:\d{2})$/);
      if (timeMatch) {
        timeText = timeMatch[1];
        i++; // skip the following time/sdev line
      }
    }

    data[bottle] = {
      sal00: Number(parts[sal00Idx]),
      prdm: Number(parts[prdmIdx]),
      zTime: buildZuluTimestamp_(parts[dateIdx], timeText),
      temp: Number(parts[tempIdx]),
      dissOxygen: Number(parts[oxyIdx]),
      chlorophyll: Number(
        String(parts[chlIdx]).replace(/\s*\(avg\)\s*$/i, '')
      )
    };
  }

  return {
    records: data,
    metadata: metadata
  };
}

function buildZuluTimestamp_(dateText, timeText) {

  const dateMatch =
    String(dateText || '')
      .trim()
      .match(/^([A-Za-z]{3,9})\s+(\d{1,2})\s+(\d{4})$/);

  const timeMatch =
    String(timeText || '')
      .trim()
      .match(/^(\d{1,2}):(\d{2}):(\d{2})$/);

  if (!dateMatch || !timeMatch) {
    return '';
  }

  const monthMap = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  const monthName = dateMatch[1].toLowerCase();
  const month = monthMap[monthName];

  const day = Number(dateMatch[2]);
  const year = Number(dateMatch[3]);

  const hour = Number(timeMatch[1]);
  const minute = Number(timeMatch[2]);
  const second = Number(timeMatch[3]);

  if (
    month === undefined ||
    !Number.isFinite(day) ||
    !Number.isFinite(year) ||
    !Number.isFinite(hour) ||
    !Number.isFinite(minute) ||
    !Number.isFinite(second)
  ) {
    return '';
  }

  return Utilities.formatDate(
    new Date(Date.UTC(year, month, day, hour, minute, second)),
    'UTC',
    "yyyy-MM-dd'T'HH:mm:ss'Z'"
  );
}

function extractBottleNumbers_(value) {
  const matches = String(value || '').match(/\d+/g);
  if (!matches) return [];
  return matches.map(Number).filter(n => Number.isFinite(n));
}

function findRovFileForDiveOrCast_(diveOrCastValue, type) {
  const folderIter = DriveApp.getFoldersByName(CONFIG.ROV_FOLDER_NAME);

  if (!folderIter.hasNext()) {
    throw new Error(`Could not find folder named "${CONFIG.ROV_FOLDER_NAME}".`);
  }

  const folder = folderIter.next();
  const matches = [];
  const diveLower = String(diveOrCastValue).trim().toLowerCase();

  findRovFilesInFolderTree_(folder, diveLower, type, matches);

  if (matches.length === 0) {
    throw new Error(`No ${type.toUpperCase()} file found for dive/cast "${diveOrCastValue}" in the ROV folder.`);
  }

  if (matches.length > 1) {
    Logger.log(`Multiple ${type.toUpperCase()} matches for dive/cast "${diveOrCastValue}": ${matches.map(f => f.getName()).join(', ')}. Using the first match.`);
  }

  return matches[0];
}

function findRovFilesInFolderTree_(folder, diveLower, type, matches) {
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const nameLower = String(file.getName()).toLowerCase();

    if (type === 'csv') {
      if (
        nameLower.endsWith('_sealogexport.csv') &&
        nameLower.includes(`_${diveLower}_`)
      ) {
        matches.push(file);
      }
    } else if (type === 'txt') {
      if (nameLower.endsWith('_data.txt') && nameLower.includes(`_${diveLower}_`)) {
        matches.push(file);
      }
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    findRovFilesInFolderTree_(subfolders.next(), diveLower, type, matches);
  }
}

function parseRovCsvFile_(file) {
  const content = file.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(content);

  if (!rows || rows.length < 2) {
    throw new Error(`CSV file "${file.getName()}" does not contain usable rows.`);
  }

  const headers = rows[0];
  const headerMap = buildLooseHeaderMap_(headers);

  const sampleIdx = headerMap[normalizeLooseHeader_('event_option.sample_id')];
  const tsIdx = headerMap[normalizeLooseHeader_('ts')];
  const latIdx = headerMap[normalizeLooseHeader_('vehicleRealtimeNavData.latitude_value')];
  const lonIdx = headerMap[normalizeLooseHeader_('vehicleRealtimeNavData.longitude_value')];

  if (sampleIdx === undefined) {
    throw new Error(`Could not find event_option.sample_id column in ${file.getName()}.`);
  }
  if (tsIdx === undefined) {
    throw new Error(`Could not find ts column in ${file.getName()}.`);
  }
  if (latIdx === undefined) {
    throw new Error(`Could not find vehicleRealtimeNavData.latitude_value column in ${file.getName()}.`);
  }
  if (lonIdx === undefined) {
    throw new Error(`Could not find vehicleRealtimeNavData.longitude_value column in ${file.getName()}.`);
  }

  const records = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const sampleId = String(row[sampleIdx] || '').trim();
    if (!sampleId) continue;

    const bottleNumbers = extractBottleNumbers_(sampleId);
    if (!bottleNumbers.length) continue;

    const bottle = Math.min(...bottleNumbers);

    records[bottle] = {
      ts: String(row[tsIdx] || '').trim(),
      latitude: parseFirstNumber_(row[latIdx]),
      longitude: parseFirstNumber_(row[lonIdx])
    };
  }

  return records;
}

function parseRovTxtFile_(file) {
  const content = file.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(content);

  if (!rows || rows.length < 2) {
    throw new Error(`TXT file "${file.getName()}" does not contain usable rows.`);
  }

  const headerIndex = rows.findIndex(row => {
    const c0 = String(row[0] || '').trim().toLowerCase();
    return c0 === 'time';
  });

  if (headerIndex === -1) {
    throw new Error(`Could not find Time header row in ${file.getName()}.`);
  }

  const parsed = [];

  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 17) continue;

    const timeText = String(row[0] || '').trim();
    const timeMs = parseUtcTimestampToMs_(timeText);

    const tempText = String(row[2] || '').trim();
    const chlorophyllText = String(row[4] || '').trim();
    const oxygenText = String(row[7] || '').trim();
    const depthText = String(row[11] || '').trim();
    const salinityText = String(row[12] || '').trim();

    const temp = parseFirstNumber_(tempText);
    const chlorophyll = parseFirstNumber_(chlorophyllText);
    const dissOxygen = parseFirstNumber_(oxygenText);
    const depth = parseFirstNumber_(depthText);
    const salinity = parseFirstNumber_(salinityText);

    if (Number.isFinite(timeMs)) {
      parsed.push({
        timeMs,
        temp,
        chlorophyll,
        dissOxygen,
        depth,
        salinity
      });
    }
  }

  return parsed;
}

function findClosestRovTemp_(txtRows, zTime) {
  const targetMs = parseUtcTimestampToMs_(zTime);
  if (!Number.isFinite(targetMs)) return null;

  let bestRow = null;
  let bestDiff = Infinity;

  for (const row of txtRows) {
    const diff = Math.abs(targetMs - row.timeMs);

    if (diff < bestDiff) {
      bestDiff = diff;
      bestRow = row;
    }
  }

  return bestRow;
}

function parseUtcTimestampToMs_(value) {
  const text = String(value || '').trim();
  if (!text) return NaN;

  // Accept either "YYYY-MM-DD HH:MM:SS.SSS" or ISO strings.
  const normalized = text.includes('T')
    ? text
    : text.replace(' ', 'T') + 'Z';

  const d = new Date(normalized);
  if (isNaN(d.getTime())) return NaN;

  return d.getTime();
}

function parseFirstNumber_(value) {
  const match = String(value || '').match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : NaN;
}

function parseAssignedNumber_(value) {
  const match = String(value || '').match(/=\s*(-?\d+(?:\.\d+)?)/);
  return match ? Number(match[1]) : NaN;
}

function zTimeToUtcSeconds_(zTime) {
  const d = (zTime instanceof Date) ? zTime : new Date(zTime);
  if (isNaN(d.getTime())) return NaN;

  return (
    d.getUTCHours() * 3600 +
    d.getUTCMinutes() * 60 +
    d.getUTCSeconds() +
    d.getUTCMilliseconds() / 1000
  );
}

function zTimeToUtcHmsMs_(zTime) {

  const d =
    (zTime instanceof Date)
      ? zTime
      : new Date(zTime);

  if (isNaN(d.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    d,
    'UTC',
    'HH:mm:ss.SSS'
  );
}

function parseHmsMsToSeconds_(timeText) {

  const match =
    String(timeText || '')
      .trim()
      .match(
        /^(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?$/
      );

  if (!match) {
    return NaN;
  }

  const h = Number(match[1]);
  const m = Number(match[2]);
  const s = Number(match[3]);

  const ms =
    match[4]
      ? Number(match[4].padEnd(3, '0'))
      : 0;

  return (
    (h * 3600) +
    (m * 60) +
    s +
    (ms / 1000)
  );
}

function circularSecondsDiff_(a, b) {

  const raw = Math.abs(a - b);

  return Math.min(
    raw,
    86400 - raw
  );
}

function normalizeLooseHeader_(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function buildLooseHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const normalized = normalizeLooseHeader_(h);
    if (normalized) map[normalized] = i;
  });
  return map;
}

// ==============================
// FLOWTHROUGH IMPORT
// ==============================

function importFlowthroughDataFromFiles() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert(`Select rows on the "${CONFIG.SHEET_NAME}" sheet first.`);
    return;
  }

  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Please highlight one or more rows first.');
    return;
  }

  const headers = getHeaders_(sheet);
  const headerMap = buildHeaderMap_(headers);

  const requiredColumns = [
    CONFIG.FLOWTHROUGH_ASSET_HEADER,
    CONFIG.FLOWTHROUGH_EVENT_DATE_HEADER,
    'z-time',
    'chlorophyll',
    'temp',
    'salinity'
  ];

  for (const colName of requiredColumns) {
    const normalized = normalizeHeader_(colName);
    if (!(normalized in headerMap)) {
      ui.alert(`Missing required column in sheet: ${colName}`);
      return;
    }
  }

  const assetCol = getColumnIndexOrThrow_(headerMap, CONFIG.FLOWTHROUGH_ASSET_HEADER);
  const eventDateCol = getColumnIndexOrThrow_(headerMap, CONFIG.FLOWTHROUGH_EVENT_DATE_HEADER);
  const zTimeCol = getColumnIndexOrThrow_(headerMap, 'z-time');
  const chlorophyllCol = getColumnIndexOrThrow_(headerMap, 'chlorophyll');
  const tempCol = getColumnIndexOrThrow_(headerMap, 'temp');
  const salinityCol = getColumnIndexOrThrow_(headerMap, 'salinity');

  const selectedRanges = rangeList.getRanges();
  const rowsToProcess = new Set();

  for (const range of selectedRanges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();

    for (let r = 0; r < numRows; r++) {
      const row = startRow + r;
      if (row > CONFIG.HEADER_ROW) {
        rowsToProcess.add(row);
      }
    }
  }

  if (rowsToProcess.size === 0) {
    ui.alert(`Select data rows below header row ${CONFIG.HEADER_ROW}.`);
    return;
  }

  const flowthroughCache = {};
  const rowList = Array.from(rowsToProcess).sort((a, b) => a - b);

  for (const row of rowList) {
    const asset = String(sheet.getRange(row, assetCol).getValue()).trim().toUpperCase();
    if (asset !== 'FLOWTHROUGH') {
      ui.alert('asset for one or more samples does not equal Flowthrough. Please check highlighted rows and try again.');
      return;
    }

    const eventDate = String(sheet.getRange(row, eventDateCol).getDisplayValue()).trim();
    if (!eventDate) {
      ui.alert(`Row ${row} is missing an eventDate value.`);
      return;
    }

    const zTime = String(sheet.getRange(row, zTimeCol).getDisplayValue()).trim();
    if (!zTime) {
      ui.alert(`Row ${row} is missing a z-time value.`);
      return;
    }

    if (!flowthroughCache[eventDate]) {
      const fluorometerFile = findFlowthroughFileForDate_(eventDate, 'fluorometer');
      const tsgFile = findFlowthroughFileForDate_(eventDate, 'tsg');

      flowthroughCache[eventDate] = {
        fluorometerRows: parseFlowthroughFluorometerTxtFile_(fluorometerFile),
        tsgRows: parseFlowthroughTsgTxtFile_(tsgFile)
      };

      Logger.log(
        `Flowthrough cache built for ${eventDate}: Fluorometer rows=${flowthroughCache[eventDate].fluorometerRows.length}, TSG rows=${flowthroughCache[eventDate].tsgRows.length}`
      );
    }

    const flowthroughPackage = flowthroughCache[eventDate];

    const closestFluorometerRow = findClosestRecordByTime_(flowthroughPackage.fluorometerRows, zTime);
    if (closestFluorometerRow && Number.isFinite(closestFluorometerRow.chlorophyll)) {
      sheet.getRange(row, chlorophyllCol).setValue(closestFluorometerRow.chlorophyll);
    }

    const closestTsgRow = findClosestRecordByTime_(flowthroughPackage.tsgRows, zTime);
    if (closestTsgRow) {
      if (Number.isFinite(closestTsgRow.temp)) {
        sheet.getRange(row, tempCol).setValue(closestTsgRow.temp);
      }

      if (Number.isFinite(closestTsgRow.salinity)) {
        sheet.getRange(row, salinityCol).setValue(closestTsgRow.salinity);
      }
    }
  }

  ui.alert('Flowthrough data import complete.');
}

function findFlowthroughFileForDate_(eventDateValue, type) {
  const folderIter = DriveApp.getFoldersByName(CONFIG.FLOWTHROUGH_FOLDER_NAME);

  if (!folderIter.hasNext()) {
    throw new Error(`Could not find folder named "${CONFIG.FLOWTHROUGH_FOLDER_NAME}".`);
  }

const folder = folderIter.next();
  const matches = [];
  const dateKey = String(eventDateValue || '').trim().toLowerCase();

  if (!dateKey) {
    throw new Error('Missing eventDate value.');
  }

  findFlowthroughFilesInFolderTree_(folder, dateKey, type, matches);

  if (matches.length === 0) {
    throw new Error(
      `No ${type.toUpperCase()} file found for event date "${eventDateValue}" in the Flowthrough folder.`
    );
  }

  if (matches.length > 1) {
    Logger.log(
      `Multiple ${type.toUpperCase()} matches for date "${eventDateValue}": ${matches.map(f => f.getName()).join(', ')}. Using the first match.`
    );
  }

  return matches[0];
}

function findFlowthroughFilesInFolderTree_(folder, dateKey, type, matches) {
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const nameLower = String(file.getName()).toLowerCase();

    if (type === 'fluorometer') {
      if (nameLower.endsWith(`_fluorometer_wetlab-${dateKey}.txt`)) {
        matches.push(file);
      }
    } else if (type === 'tsg') {
      if (nameLower.endsWith(`_tsg_sbe45_bow-${dateKey}.txt`)) {
        matches.push(file);
      }
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    findFlowthroughFilesInFolderTree_(subfolders.next(), dateKey, type, matches);
  }
}

function parseFlowthroughFluorometerTxtFile_(file) {
  const content = file.getBlob().getDataAsString();
  const lines = content.split(/\r?\n/);
  const parsed = [];

  for (const line of lines) {
    const text = String(line || '').trim();
    if (!text || !/^\d{4}-\d{2}-\d{2}T/.test(text)) continue;

    const commaIdx = text.indexOf(',');
    if (commaIdx === -1) continue;

    const timeText = text.slice(0, commaIdx).trim();
    const rest = text.slice(commaIdx + 1).trim();
    const cols = rest.split('\t');

    if (cols.length < 4) continue;

    const timeMs = parseUtcTimestampToMs_(timeText);
    const chlorophyll = parseFirstNumber_(cols[3]); // 4th tab-delimited column

    if (Number.isFinite(timeMs)) {
      parsed.push({
        timeMs,
        chlorophyll
      });
    }
  }

  return parsed;
}

function parseFlowthroughTsgTxtFile_(file) {
  const content = file.getBlob().getDataAsString();
  const lines = content.split(/\r?\n/);
  const parsed = [];

  for (const line of lines) {
    const text = String(line || '').trim();
    if (!text || !/^\d{4}-\d{2}-\d{2}T/.test(text)) continue;

    const commaIdx = text.indexOf(',');
    if (commaIdx === -1) continue;

    const timeText = text.slice(0, commaIdx).trim();
    const rest = text.slice(commaIdx + 1).trim();

    const cols = rest.split(/,\s*/);

    const t1Field = cols.find(c => /^\s*t1\s*=/.test(c));
    const sField = cols.find(c => /^\s*s\s*=/.test(c));

    const timeMs = parseUtcTimestampToMs_(timeText);

    if (Number.isFinite(timeMs)) {
      parsed.push({
        timeMs,
        temp: parseAssignedNumber_(t1Field),
        salinity: parseAssignedNumber_(sField)
      });
    }
  }

  return parsed;
}

function findClosestRecordByTime_(records, zTime) {
  const targetMs = parseUtcTimestampToMs_(zTime);
  if (!Number.isFinite(targetMs)) return null;

  let bestRow = null;
  let bestDiff = Infinity;

  for (const row of records) {
    const diff = Math.abs(targetMs - row.timeMs);
    if (diff < bestDiff) {
      bestDiff = diff;
      bestRow = row;
    }
  }

  return bestRow;
}

// ==============================
// WEB APP
// ==============================

function doGet(e) {
  if (e.parameter.mapData) {
    return renderMapPage_(e.parameter.mapData);
  }

  const row = parseInt(e.parameter.row, 10);

  if (!row) {
    return HtmlService.createHtmlOutput('Missing row parameter.');
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    return HtmlService.createHtmlOutput(`Missing sheet "${CONFIG.SHEET_NAME}".`);
  }

  const headers = getHeaders_(sheet);
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const data = {};
  headers.forEach((h, i) => {
    data[h] = values[i];
  });

  const sheetUrl =
    `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${sheet.getSheetId()}&range=A${row}`;

  const title =
    data[CONFIG.LABEL_TEXT_TOP_COLUMN] ||
    data.sample_id ||
    data.sampleID ||
    `Sample row ${row}`;

  let fieldsHtml = '';
  headers.forEach(header => {
    const value = data[header];
    if (value !== '' && value !== null && value !== undefined) {
      fieldsHtml += `
        <div class="field">
          <div class="key">${escapeHtml_(header)}</div>
          <div class="value">${escapeHtml_(value)}</div>
        </div>
      `;
    }
  });

  return HtmlService.createHtmlOutput(`
    <!doctype html>
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 16px;
            background: #f4f6f8;
            color: #111;
          }

          .card {
            max-width: 640px;
            margin: 0 auto;
            background: white;
            border-radius: 14px;
            padding: 18px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.12);
          }

          h1 {
            font-size: 22px;
            margin: 0 0 6px;
          }

          .subtitle {
            font-size: 13px;
            color: #666;
            margin-bottom: 16px;
          }

          .field {
            border-top: 1px solid #eee;
            padding: 10px 0;
          }

          .key {
            font-size: 12px;
            color: #666;
            font-weight: bold;
          }

          .value {
            font-size: 16px;
            margin-top: 3px;
            overflow-wrap: anywhere;
          }

          .button {
            display: inline-block;
            margin-top: 18px;
            padding: 11px 14px;
            background: #0b57d0;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            font-weight: bold;
          }
        </style>
      </head>

      <body>
        <div class="card">
          <h1>${escapeHtml_(title)}</h1>
          <div class="subtitle">Row ${row} · ${escapeHtml_(CONFIG.SHEET_NAME)}</div>

          ${fieldsHtml}

          <a class="button" href="${sheetUrl}" target="_blank">
            Open in Google Sheet
          </a>
        </div>
      </body>
    </html>
  `);
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
