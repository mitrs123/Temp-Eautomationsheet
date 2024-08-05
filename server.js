const { google } = require('googleapis');
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const dotenv = require('dotenv');


// Load environment variables from .env file
dotenv.config();

const app = express();
const port = 3000;

// Google Sheets API credentials loaded from .env
const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
  private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
}

const project_id = process.env.PROJECT_ID;

// Google Sheets API version
const client = new google.auth.JWT(
  client_email,
  null,
  private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

const sheets = google.sheets({ version: 'v4', auth: client });

// Sheet IDs
const SPREADSHEET_ID_MASTER = '1xcetIqD5BesWLlAOtRbXOu7SuM7OP9pbmjBNVHkL5OE'; // Replace with your master sheet ID
const SPREADSHEET_ID_SALES1 = '1d3_LxjSA-1Esl0rKtZu7WbY18Y3SvJDn0R5sRX-Q9H0'; // Replace with your sales1 sheet ID
const SPREADSHEET_ID_SALES2 = '1tvtVQ20gSN3vJPJ6KWfs-B7QX39MZfF5GZZAuFhk6Zs'; // Replace with your sales2 sheet ID
const SPREADSHEET_ID_SALES3 = '1S0Obv2Juk7F3D8L1b-J4aaC1VWgfacm_SgcRXpBqvtM'; // Replace with your sales3 sheet ID

// Sales persons mapping
const SALES_PERSONS = {
  'User1': { name: 'Sales Person 1', sheetId: SPREADSHEET_ID_SALES1 },
  'User2': { name: 'Sales Person 2', sheetId: SPREADSHEET_ID_SALES2 },
  'User3': { name: 'Sales Person 3', sheetId: SPREADSHEET_ID_SALES3 }
};

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Function to append data to a sheet
const appendDataToSheet = async (sheetId, formData, includeSalespersonName) => {
  const newRow = [
    formData.leadId,
    Date(formData.date),
    formData.projectType,
    formData.leadOrigin,
    formData.clientName,
    +formData.expectedProjectCapacity,
    formData.contactPersonName,
    formData.designation,
    +formData.contactNumber,
    +formData.contactNumber2,
    formData.area,
    formData.city,
    formData.remarks
  ];

  if (includeSalespersonName) {
    newRow.push(formData.salespersonName); // Add salesperson name if needed
  }

  // Append to the specified sheet
  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: 'Sheet1!A:N', // Adjust the range if needed
    valueInputOption: 'RAW',
    resource: {
      values: [newRow]
    }
  });
};

// Handle form submission
app.post('/form-data', async (req, res) => {
  console.log("Endpoint called");
  const formData = req.body;

  try {
    // Add salesperson name to formData for processing
    const salesperson = SALES_PERSONS[formData.salesperson];
    if (salesperson) {
      formData.salespersonName = salesperson.name;
    } else {
      formData.salespersonName = 'Unknown';
    }

    // Append data to the master sheet
    await appendDataToSheet(SPREADSHEET_ID_MASTER, formData, true);

    // Append data to the sales person's sheet without salesperson name
    if (salesperson) {
      await appendDataToSheet(salesperson.sheetId, formData, false);
    }

    res.status(200).send('Form data received and added to master and sales sheets');
  } 
  catch (error) {
    console.error('Error processing form data:', error);
    res.status(500).send('Internal Server Error');
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Express server listening at http://localhost:${port}`);
});



// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3000;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const project_id = process.env.PROJECT_ID;

// // Google Sheets API version
// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1x9qGU9-ZpMsbYV3v-PcKznqK9CU51XzuKvveg3l_1NE'; // Replace with your master sheet ID

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Function to check for duplicate entries by comparing the entire row in the master sheet
// const checkForDuplicateEntry = async (sheetId, formData) => {
//   const response = await sheets.spreadsheets.values.get({
//     spreadsheetId: sheetId,
//     range: 'Sheet1!A:N' // Assuming the relevant data is in columns A to N
//   });

//   const rows = response.data.values || [];
  
//   return rows.some(row => 
//     row[0] === formData.timestamp &&
//     row[1] === formData.emailAddress &&
//     row[2] === formData.uniqueID &&
//     row[3] === formData.leadOrigin &&
//     row[4] === formData.clientName &&
//     row[5] === formData.expectedProjectCapacity.toString() &&
//     row[6] === formData.projectType &&
//     row[7] === formData.contactPersonName &&
//     row[8] === formData.designation &&
//     row[9] === formData.contactNumber.toString() &&
//     row[10] === formData.contactNumber2.toString() &&
//     row[11] === formData.area &&
//     row[12] === formData.city &&
//     row[13] === formData.remarks
//   );
// };

// // Function to append data to the master sheet
// const appendDataToMasterSheet = async (formData) => {
//   const newRow = [
//     formData.timestamp,
//     formData.emailAddress,
//     formData.uniqueID,
//     formData.leadOrigin,
//     formData.clientName,
//     formData.expectedProjectCapacity,
//     formData.projectType,
//     formData.contactPersonName,
//     formData.designation,
//     formData.contactNumber,
//     formData.contactNumber2,
//     formData.area,
//     formData.city,
//     formData.remarks
//   ];

//   // Append to master sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: SPREADSHEET_ID_MASTER,
//     range: 'Sheet1!A:N',
//     valueInputOption: 'RAW',
//     resource: {
//       values: [newRow]
//     }
//   });
// };

// // Handle form submission
// app.post('/form-data', async (req, res) => {
//   console.log("Endpoint called");
//   const formData = req.body;

//   try {
//     // Check for duplicate entry in master sheet
//     const isDuplicate = await checkForDuplicateEntry(SPREADSHEET_ID_MASTER, formData);

//     if (!isDuplicate) {
//       // Append data to the master sheet
//       await appendDataToMasterSheet(formData);
//       res.status(200).send('Form data received and added to master sheet');
//     } else {
//       res.status(409).send('Duplicate entry found, data not added');
//     }
//   } 
//   catch (error) {
//     console.error('Error processing form data:', error);
//     res.status(500).send('Internal Server Error');
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Express server listening at http://localhost:${port}`);
// });





// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3000;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const project_id = process.env.PROJECT_ID;

// // Google Sheets API version
// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1ZN4BOvovDxFyYz-AHVFnZB-cYrvL0s48E_oLijQojTs'; // Replace with your master sheet ID

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Function to check for duplicate entries by comparing the entire row in the master sheet
// const checkForDuplicateEntry = async (sheetId, formData) => {
//   const response = await sheets.spreadsheets.values.get({
//     spreadsheetId: sheetId,
//     range: 'Sheet1!A:N' // Assuming the relevant data is in columns A to N
//   });

//   const rows = response.data.values || [];
  
//   return rows.some(row => 
//     row[0] === formData.timestamp &&
//     row[1] === formData.emailAddress &&
//     row[2] === formData.uniqueID &&
//     row[3] === formData.leadOrigin &&
//     row[4] === formData.clientName &&
//     row[5] === formData.expectedProjectCapacity.toString() &&
//     row[6] === formData.projectType &&
//     row[7] === formData.contactPersonName &&
//     row[8] === formData.designation &&
//     row[9] === formData.contactNumber.toString() &&
//     row[10] === formData.contactNumber2.toString() &&
//     row[11] === formData.area &&
//     row[12] === formData.city &&
//     row[13] === formData.remarks
//   );
// };

// // Function to append data to the master sheet
// const appendDataToMasterSheet = async (formData) => {
//   const newRow = [
//     formData.timestamp,
//     formData.emailAddress,
//     formData.uniqueID,
//     formData.leadOrigin,
//     formData.clientName,
//     formData.expectedProjectCapacity,
//     formData.projectType,
//     formData.contactPersonName,
//     formData.designation,
//     formData.contactNumber,
//     formData.contactNumber2,
//     formData.area,
//     formData.city,
//     formData.remarks
//   ];

//   // Append to master sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: SPREADSHEET_ID_MASTER,
//     range: 'Sheet1!A:N',
//     valueInputOption: 'RAW',
//     resource: {
//       values: [newRow]
//     }
//   });
// };

// // Handle form submission
// app.post('/form-data', async (req, res) => {
//   console.log("Endpoint called");
//   const formData = req.body;

//   try {
//     // Check for duplicate entry in master sheet
//     const isDuplicate = await checkForDuplicateEntry(SPREADSHEET_ID_MASTER, formData);

//     if (!isDuplicate) {
//       // Append data to the master sheet
//       await appendDataToMasterSheet(formData);
//       res.status(200).send('Form data received and added to master sheet');
//     } else {
//       res.status(409).send('Duplicate entry found, data not added');
//     }
//   } catch (error) {
//     console.error('Error processing form data:', error);
//     res.status(500).send('Internal Server Error');
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Express server listening at http://localhost:${port}`);
// });

// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3000;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const project_id = process.env.PROJECT_ID;

// // Google Sheets API version
// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1ZN4BOvovDxFyYz-AHVFnZB-cYrvL0s48E_oLijQojTs'; // Replace with your master sheet ID

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Function to check for duplicate entries by timestamp in the master sheet
// const checkForDuplicateTimestamp = async (sheetId, timestamp) => {
//   const response = await sheets.spreadsheets.values.get({
//     spreadsheetId: sheetId,
//     range: 'Sheet1!A:A' // Assuming the timestamp is in column A
//   });

//   const rows = response.data.values;
//   return rows.some(row => row[0] === timestamp);
// };

// // Function to append data to the master sheet
// const appendDataToMasterSheet = async (formData) => {
//   const newRow = [
//     formData.timestamp,
//     formData.emailAddress,
//     formData.uniqueID,
//     formData.leadOrigin,
//     formData.clientName,
//     +formData.expectedProjectCapacity,
//     formData.projectType,
//     formData.contactPersonName,
//     formData.designation,
//     +formData.contactNumber,
//     +formData.contactNumber2,
//     formData.area,
//     formData.city,
//     formData.remarks
//   ];

//   // Append to master sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: SPREADSHEET_ID_MASTER,
//     range: 'Sheet1!A:N',
//     valueInputOption: 'RAW',
//     resource: {
//       values: [newRow]
//     }
//   });
// };

// // Handle form submission
// app.post('/form-data', async (req, res) => {
//   console.log("Endpoint called");
//   const formData = req.body;

//   try {
//     // Check for duplicate timestamp in master sheet
//     const isDuplicate = await checkForDuplicateTimestamp(SPREADSHEET_ID_MASTER, formData.timestamp);

//     if (!isDuplicate) {
//       // Append data to the master sheet
//       await appendDataToMasterSheet(formData);
//       res.status(200).send('Form data received and added to master sheet');
//     } else {
//       res.status(409).send('Duplicate timestamp found, data not added');
//     }
//   } catch (error) {
//     console.error('Error processing form data:', error);
//     res.status(500).send('Internal Server Error');
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Express server listening at http://localhost:${port}`);
// });


// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3000;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const project_id = process.env.PROJECT_ID;

// // Google Sheets API version
// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1ZN4BOvovDxFyYz-AHVFnZB-cYrvL0s48E_oLijQojTs'; // Replace with your master sheet ID
// const SPREADSHEET_ID_SALES1 = '159yCKypb4xdM0_ger3N6ls8lsT92Yob_0SIpCsnjBzQ'; // Replace with your sales1 sheet ID
// const SPREADSHEET_ID_SALES2 = '1UFJ76nUv922KOiBxu7K2Tl8nTmNZ2fAWnFP05fxeOCI'; // Replace with your sales2 sheet ID
// const SPREADSHEET_ID_SALES3 = '1KqJvOPZQsVEb26frQ32-3iyx-xW1fMgj0wNvCQH6tDo'; // Replace with your sales3 sheet ID

// // Sales persons mapping
// const SALES_PERSONS = {
//   'UniqueId@1': { name: 'Sales Person 1', sheetId: SPREADSHEET_ID_SALES1 },
//   'UniqueId@2': { name: 'Sales Person 2', sheetId: SPREADSHEET_ID_SALES2 },
//   'UniqueId@3': { name: 'Sales Person 3', sheetId: SPREADSHEET_ID_SALES3 }
// };

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Temporary storage array
// let tempDataArray = [];
// let tempDataTimeout;

// // Function to check for duplicate entries in a given sheet
// const checkForDuplicate = async (sheetId, range, formData) => {
//   const response = await sheets.spreadsheets.values.get({
//     spreadsheetId: sheetId,
//     range: range
//   });

//   const rows = response.data.values;
//   const lastRows = rows.slice(-3); // Get the last 3 rows

//   return lastRows.some(row => (
//     row[1] === formData.timestamp.split(' ')[0] &&
//     row[2] === formData.salesPersonName &&
//     row[3] === formData.leadOrigin &&
//     row[4] === formData.clientName &&
//     row[5] === formData.expectedProjectCapacity &&
//     row[6] === formData.projectType &&
//     row[7] === formData.contactPersonName &&
//     row[8] === formData.designation &&
//     row[9] === formData.contactNumber &&
//     row[10] === formData.contactNumber2 &&
//     row[11] === formData.area &&
//     row[12] === formData.city &&
//     row[13] === formData.remarks
//   ));
// };

// // Function to remove duplicate entries from a given sheet
// const removeDuplicates = async (sheetId, range) => {
//   const response = await sheets.spreadsheets.values.get({
//     spreadsheetId: sheetId,
//     range: range
//   });

//   const rows = response.data.values;
//   const uniqueRows = [];
//   const duplicateIndexes = [];

//   rows.forEach((row, index) => {
//     const rowString = JSON.stringify(row);
//     if (uniqueRows.includes(rowString)) {
//       duplicateIndexes.push(index + 1); // +1 because Sheets API is 1-indexed
//     } else {
//       uniqueRows.push(rowString);
//     }
//   });

//   if (duplicateIndexes.length > 0) {
//     await sheets.spreadsheets.batchUpdate({
//       spreadsheetId: sheetId,
//       resource: {
//         requests: duplicateIndexes.map(rowIndex => ({
//           deleteDimension: {
//             range: {
//               sheetId: 0,
//               dimension: 'ROWS',
//               startIndex: rowIndex - 1,
//               endIndex: rowIndex
//             }
//           }
//         }))
//       }
//     });
//   }
// };

// // Function to process temporary data
// const processTempData = async () => {
//   if (tempDataArray.length === 0) return;

//   const firstData = tempDataArray[0];
//   const isAllDataSame = tempDataArray.every(data => JSON.stringify(data) === JSON.stringify(firstData));

//   try {
//     if (isAllDataSame) {
//       // Append only one object to both sheets
//       await appendDataToSheets(firstData);
//     } else {
//       // Append each object separately
//       for (const data of tempDataArray) {
//         await appendDataToSheets(data);
//       }
//     }
//   } catch (error) {
//     console.error('Error processing temp data:', error);
//   } finally {
//     // Clear the temp array
//     tempDataArray = [];
//   }
// };

// // Function to append data to sheets
// const appendDataToSheets = async (formData) => {
//   const salesPersonInfo = SALES_PERSONS[formData.uniqueID];
//   formData.salesPersonName = salesPersonInfo.name;

//   const masterSheetResponse = await sheets.spreadsheets.values.get({
//     spreadsheetId: SPREADSHEET_ID_MASTER,
//     range: 'Sheet1!A:N'
//   });

//   const rows = masterSheetResponse.data.values;
//   const newSequenceNumber = rows.length > 1 ? parseInt(rows[rows.length - 1][0]) + 1 : 1;
//   // const date = formData.timestamp.split(' ')[0];

//     const dateObj = new Date(formData.timestamp);


//   // Extract date components
// const day = dateObj.getDate().toString().padStart(2, '0');
// const month = (dateObj.getMonth() + 1).toString().padStart(2, '0'); // Month is zero-based
// const year = dateObj.getFullYear().toString();

// // Format date in dd-mm-yyyy format
// const formattedDate = `${day}-${month}-${year}`;

//   const newRow = [
//     newSequenceNumber,
//     formattedDate,
//     formData.salesPersonName,
//     formData.leadOrigin,
//     formData.clientName,
//     +formData.expectedProjectCapacity,
//     formData.projectType,
//     formData.contactPersonName,
//     formData.designation,
//     +formData.contactNumber,
//     +formData.contactNumber2,
//     formData.area,
//     formData.city,
//     formData.remarks
//   ];
//   console.log(formData)

//   // Append to master sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: SPREADSHEET_ID_MASTER,
//     range: 'Sheet1!A:N',
//     valueInputOption: 'RAW',
//     resource: {
//       values: [newRow]
//     }
//   });

//   // Append to specific sales person's sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: salesPersonInfo.sheetId,
//     range: 'Sheet1!A:N', // Ensure the range includes all columns including Sales Person Name
//     valueInputOption: 'RAW',
//     resource: {
//       values: [newRow]
//     }
//   });

//   // Check for and remove any duplicates after insertion
//   // await removeDuplicates(SPREADSHEET_ID_MASTER, 'Sheet1!A:N');
//   // await removeDuplicates(salesPersonInfo.sheetId, 'Sheet1!A:N');
// };

// // Handle form submission
// app.post('/form-data', (req, res) => {
//   console.log("Endpoint called");
//   const formData = req.body;
//   tempDataArray.push(formData);

//   // If a timeout is already set, clear it
//   if (tempDataTimeout) {
//     clearTimeout(tempDataTimeout);
//   }

//   // Set a new timeout to process the data after 4 seconds
//   tempDataTimeout = setTimeout(processTempData, 4000);
//   res.status(200).send('Form data received and temporarily stored');
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Express server listening at http://localhost:${port}`);
// });

