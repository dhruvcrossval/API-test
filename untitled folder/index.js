const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

async function listVal(auth) {
  const sheets = google.sheets({version: 'v4', auth});
  
  // Get sheet metadata
  // test sheet 1 : 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
  // test sheet 2 : 1ECWv2anoZB3ksOt6HFdaY7YnHewxS2FQ-K7YU3ycgEE
  const sheetMetadata = await sheets.spreadsheets.get({
    spreadsheetId: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
  });

  const sheet = sheetMetadata.data.sheets[0];
  const sheetName = sheet.properties.title;
  const rowCount = sheet.properties.gridProperties.rowCount;
  const columnCount = sheet.properties.gridProperties.columnCount;

  // Construct the dynamic range
  const endColumn = String.fromCharCode(64 + columnCount); // Converts number to letter (e.g., 1 -> A, 2 -> B)
  const range = `${sheetName}!A1:${endColumn}${rowCount}`;

  // Fetch the data using the dynamic range
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
    range: range,
  });

  const rows = res.data.values;
  if (!rows || rows.length === 0) {
    console.log('No data found.');
    return;
  }

  console.log('Row Num, Column Values:');
  rows.forEach((row, rowIndex) => {
    const RowNum = rowIndex + 1; 
    const colValues = row.map((value, colIndex) => value !== undefined ? value : `undefined_col_${colIndex + 1}`);
    console.log(`Row ${RowNum}: ${colValues.join(', ')}`);
  });
}

authorize().then(listVal).catch(console.error);
