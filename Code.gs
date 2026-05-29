function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
var SHEET_ID_PROP = 'SHEET_ID';
// Default Spreadsheet ID (from user) — can be overridden with setSheetId(id)
var DEFAULT_SHEET_ID = '14NDgvy9Q97C-o0_4qyBqpFPC9zBEBY-EtlcysEYZIhQ';

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP) || DEFAULT_SHEET_ID;
  if (id && id !== 'YOUR_SHEET_ID') {
    return SpreadsheetApp.openById(id);
  }
  // Try active spreadsheet (useful when script is bound)
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  throw new Error('Spreadsheet ID not configured. Call setSheetId(id) or set property "' + SHEET_ID_PROP + '".');
}

function setSheetId(id) {
  if (!id || typeof id !== 'string') throw new Error('Invalid sheet id');
  PropertiesService.getScriptProperties().setProperty(SHEET_ID_PROP, id);
  return { status: 'OK', message: 'SHEET_ID saved' };
}

function _normalizeCustomers(payloadCustomers) {
  if (!payloadCustomers) return [];
  if (payloadCustomers.length === 0) return [];
  // Already array of arrays
  if (Array.isArray(payloadCustomers[0])) return payloadCustomers;
  // Objects -> array of arrays
  return payloadCustomers.map(function(c) {
    return [c.name || '', c.mobile || '', c.email || '', c.shop || '', c.address || ''];
  });
}

function _normalizeJobs(payloadJobs) {
  if (!payloadJobs) return [];
  if (payloadJobs.length === 0) return [];
  if (Array.isArray(payloadJobs[0])) return payloadJobs;
  return payloadJobs.map(function(j) {
    return [
      j.id || '',
      j.relatedJobId || '',
      j.date || '',
      j.cname || '',
      j.description || '',
      j.size || '',
      (typeof j.price !== 'undefined' ? j.price : ''),
      j.status || '',
      j.payDate || '',
      j.paymentHistory ? 'TRUE' : 'FALSE'
    ];
  });
}

function getDataFromSheet() {
  var ss = getSpreadsheet();
  var customersSheet = getOrCreateSheet(ss, 'Customers');
  var jobsSheet = getOrCreateSheet(ss, 'Jobs');

  var customers = [];
  var jobs = [];

  var custData = customersSheet.getDataRange().getValues();
  if (custData.length > 1) {
    for (var i = 1; i < custData.length; i++) {
      var row = custData[i];
      if (row[0] === '' && row[1] === '' && row[2] === '' && row[3] === '' && row[4] === '') continue;
      customers.push({
        name: row[0] || '',
        mobile: row[1] || '',
        email: row[2] || '',
        shop: row[3] || '',
        address: row[4] || ''
      });
    }
  }

  var jobsData = jobsSheet.getDataRange().getValues();
  if (jobsData.length > 1) {
    for (var j = 1; j < jobsData.length; j++) {
      var row = jobsData[j];
      if (row[0] === '' && row[3] === '' && row[4] === '' && row[6] === '' && row[7] === '') continue;
      jobs.push({
        id: row[0] || '',
        relatedJobId: row[1] || null,
        date: row[2] || '',
        cname: row[3] || '',
        description: row[4] || '',
        size: row[5] || '',
        price: parseFloat(row[6]) || 0,
        status: row[7] || '',
        payDate: row[8] || null,
        paymentHistory: String(row[9]).toLowerCase() === 'true'
      });
    }
  }

  return {
    customers: customers,
    jobs: jobs
  };
}

function saveDataToSheet(payload) {
  var ss = getSpreadsheet();
  var customersSheet = getOrCreateSheet(ss, 'Customers');
  var jobsSheet = getOrCreateSheet(ss, 'Jobs');

  // Normalize incoming payloads: support both arrays-of-arrays (external sync) and arrays-of-objects (google.script.run)
  var custRows = _normalizeCustomers(payload.customers);
  var jobRows = _normalizeJobs(payload.jobs);

  if (custRows && custRows.length >= 0) {
    customersSheet.clearContents();
    customersSheet.getRange(1, 1, 1, 5).setValues([['Name', 'Mobile', 'Email', 'Shop', 'Address']]);
    if (custRows.length > 0) {
      customersSheet.getRange(2, 1, custRows.length, 5).setValues(custRows);
    }
    customersSheet.setFrozenRows(1);
  }

  if (jobRows && jobRows.length >= 0) {
    jobsSheet.clearContents();
    jobsSheet.getRange(1, 1, 1, 10).setValues([['ID', 'RelatedJobID', 'Date', 'Customer', 'Description', 'Size', 'Price', 'Status', 'PayDate', 'PaymentHistory']]);
    if (jobRows.length > 0) {
      jobsSheet.getRange(2, 1, jobRows.length, 10).setValues(jobRows);
    }
    jobsSheet.setFrozenRows(1);
  }

  return { status: 'OK' };
}

function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents || '{}');
      } catch (err) {
        // If not JSON, ignore and return error
        return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: 'Invalid JSON payload' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    var result = saveDataToSheet(payload);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
