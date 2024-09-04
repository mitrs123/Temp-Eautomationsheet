{
// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3333;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1OaKEgNWWUEi1LLHyTVCoJMMeTQ4TE7NuB2Zwy4lFTzo'; // Replace with your master sheet ID
// const SPREADSHEET_ID_SALES1 = '1_eU7YevVyWs6OlGem4js_qL7KKYXDNXlloQVVIhyApc'; // Replace with your sales1 sheet ID
// const SPREADSHEET_ID_SALES2 = '1GwKY8MY8aKEudRG6v-hKIUrjZd0WrV5KGAKpOBcvYCA'; // Replace with your sales2 sheet ID
// const SPREADSHEET_ID_SALES3 = '10yCa--HOn4mBBsQXhEsj5FGeddEGvv776ZcRG29A014'; // Replace with your sales3 sheet ID
// const SPREADSHEET_ID_SALES4 = '1nwJ-Uo7RVcXUtXuSE-PkG25kpc_GF1aA82nWN0WuwJI'; // Replace with your sales4 sheet ID
// const SPREADSHEET_ID_SALES5 = '1gB0l50xioy_-5Q7qIeZ-ZAGgHJycuh6FCLVhl6jOvcs'; // Replace with your sales5 sheet ID
// const SPREADSHEET_ID_SALES6 = '1pv_WOHLnrcXQ5VaeCr8f51Vf48UvueOOU7FOw-AJGFo'; // Replace with your sales6 sheet ID
// const SPREADSHEET_ID_SALES7 = '14j4_EKrY2NXOxnAXpu6MPUdDOjioxR4t_y_95x1hgLs'; // Replace with your sales7 sheet ID
// const SPREADSHEET_ID_SALES8 = '18b6e92gXN5w9vFiREWgOVerb7RlhR9V861NpXWWWJ0I'; // Replace with your sales7 sheet ID
// const SPREADSHEET_ID_SALES9 = '1Ql1jOzipJQHb5A_loA9AM1pvL7SidCnCV5IXvbs3Sqo'; // Replace with your sales7 sheet ID

// // Sales persons mapping
// const SALES_PERSONS = {
//   'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES1 },
//   'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES2 },
//   'hemant@enersol.co.in': { name: 'Hemant Trivedi', sheetId: SPREADSHEET_ID_SALES3 },
//   'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES4 },
//   'subhakanta.sahoo@enersol.co.in': { name: 'Shubhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES5 },
//   'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES6 },
//   'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES7 },
//   'User1': { name: 'Test 1', sheetId: SPREADSHEET_ID_SALES8 },
//   'User2': { name: 'Test 2', sheetId: SPREADSHEET_ID_SALES9 },
// };

// // Headers for Lead and ESL Sheets in Master Sheet
// const LEAD_HEADERS_MASTER = [
//   "Lead ID", "Created Date", "Sales Person Name", "Project Type", "Lead Origin", "Client Name", 
//   "Expected Tentative Capacity", "Contact Person 1", "Designation 1", "Contact Number 1", 
//   "Contact Person 2", "Designation 2", "Contact Number 2", "Area", "Pincode", 
//   "City", "Co-Ordinates", "Remarks"
// ];

// const LEAD_HEADERS_SALESPERSON = [
//   "Lead ID", "Created Date", "Project Type", "Lead Origin", "Client Name", 
//   "Expected Tentative Capacity", "Contact Person 1", "Designation 1", "Contact Number 1", 
//   "Contact Person 2", "Designation 2", "Contact Number 2", "Area", "Pincode", 
//   "City", "Co-Ordinates", "Remarks"
// ];

// const ESL_HEADERS_MASTER = [
//   "Created Date", "Lead ID", "ESL Number", "Sales Person Name", "Project Type", "Date", 
//   "Client Name", "Final Capacity", "Consumer Number", "Application Number", 
//   "Address", "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
// ];

// const ESL_HEADERS_SALESPERSON = [
//   "Created Date", "Lead ID", "ESL Number", "Project Type", "Date", 
//   "Client Name", "Final Capacity", "Consumer Number", "Application Number (Blank)", 
//   "Address", "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
// ];

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Utility to format date to DD/MM/YYYY
// const formatDate = (isoDate) => {
//   const date = new Date(isoDate);
//   return `${('0' + date.getDate()).slice(-2)}/${('0' + (date.getMonth() + 1)).slice(-2)}/${date.getFullYear()}`;
// };

// // Function to ensure the sheet exists and has headers
// const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
//   try {
//     // Check if the sheet exists
//     const sheetResponse = await sheets.spreadsheets.get({
//       spreadsheetId
//     });

//     let sheetExists = false;
//     sheetResponse.data.sheets.forEach((sheet) => {
//       if (sheet.properties.title === sheetName) {
//         sheetExists = true;
//       }
//     });

//     // Create the sheet if it doesn't exist
//     if (!sheetExists) {
//       await sheets.spreadsheets.batchUpdate({
//         spreadsheetId,
//         resource: {
//           requests: [
//             {
//               addSheet: {
//                 properties: {
//                   title: sheetName
//                 }
//               }
//             }
//           ]
//         }
//       });
//     }

//     // Ensure headers are present
//     const range = `'${sheetName}'!A1:${String.fromCharCode(65 + headers.length - 1)}1`;
//     const headerResponse = await sheets.spreadsheets.values.get({
//       spreadsheetId,
//       range
//     });

//     if (!headerResponse.data.values || headerResponse.data.values.length === 0) {
//       await sheets.spreadsheets.values.update({
//         spreadsheetId,
//         range: `'${sheetName}'!A1`,
//         valueInputOption: 'RAW',
//         resource: { values: [headers] }
//       });
//     }
//   } catch (error) {
//     console.error('Error ensuring sheet and headers:', error);
//   }
// };

// // Function to append data to a sheet
// const appendDataToSheet = async (spreadsheetId, sheetName, data, headers) => {
//   const values = headers.map(header => {
//     switch (header) {
//       case "Lead ID":
//         return data.leadId || '';
//       case "Created Date":
//         return data.date ? formatDate(data.date) : '';
//       case "Sales Person Name":
//         return data.salespersonName || '';
//       case "Project Type":
//         return data.projectType || '';
//       case "Lead Origin":
//         return data.leadOrigin || '';
//       case "Client Name":
//         return data.clientName || '';
//       case "Expected Tentative Capacity":
//         return +data.expectedProjectCapacity || '';
//       case "Contact Person 1":
//         return data.contactPersonName1 || '';
//       case "Designation 1":
//         return data.designation1 || '';
//       case "Contact Number 1":
//         return +data.contactNumber1 || '';
//       case "Contact Person 2":
//         return data.contactPersonName2 || '';
//       case "Designation 2":
//         return data.designation2 || '';
//       case "Contact Number 2":
//         return +data.contactNumber2 || '';
//       case "Area":
//         return data.area || '';
//       case "Pincode":
//         return +data.pincode || '';
//       case "City":
//         return data.city || '';
//       case "Co-Ordinates":
//         return data.coordinates || '';
//       case "Remarks":
//         return data.remarks || '';
//       default:
//         return ''; // For any unexpected headers
//     }
//   });

//   await sheets.spreadsheets.values.append({
//     spreadsheetId,
//     range: `${sheetName}!A:${String.fromCharCode(65 + headers.length - 1)}`,
//     valueInputOption: 'RAW',
//     resource: { values: [values] }
//   });
// };

// // Route to handle data submission
// app.post('/submitLead', async (req, res) => {
//   const { email, leadData } = req.body;

//   const salesperson = SALES_PERSONS[email];
//   if (!salesperson) {
//     return res.status(400).json({ error: 'Invalid salesperson email' });
//   }

//   try {
//     // Ensure the Lead sheet exists and has the proper headers
//     await ensureSheetAndHeaders(SPREADSHEET_ID_MASTER, 'Lead', LEAD_HEADERS_MASTER);
//     await ensureSheetAndHeaders(salesperson.sheetId, 'Lead', LEAD_HEADERS_SALESPERSON);

//     // Append the lead data to both the master sheet and the salesperson's individual sheet
//     await appendDataToSheet(SPREADSHEET_ID_MASTER, 'Lead', leadData, LEAD_HEADERS_MASTER);
//     await appendDataToSheet(salesperson.sheetId, 'Lead', leadData, LEAD_HEADERS_SALESPERSON);

//     res.status(200).json({ message: 'Lead data submitted successfully' });
//   } catch (error) {
//     console.error('Error submitting lead data:', error);
//     res.status(500).json({ error: 'An error occurred while submitting lead data' });
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Server running on port ${port}`);
// });
}

const { google } = require('googleapis');
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const dotenv = require('dotenv');

// Load environment variables from .env file
dotenv.config();

const app = express();
const port = 3333;

// Google Sheets API credentials loaded from .env
const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
  private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
}

const client = new google.auth.JWT(
  client_email,
  null,
  private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

const sheets = google.sheets({ version: 'v4', auth: client });

// Sheet IDs
const SPREADSHEET_ID_MASTER = '1OaKEgNWWUEi1LLHyTVCoJMMeTQ4TE7NuB2Zwy4lFTzo'; // Replace with your master sheet ID
const SPREADSHEET_ID_SALES1 = '1_eU7YevVyWs6OlGem4js_qL7KKYXDNXlloQVVIhyApc'; // Replace with your sales1 sheet ID
const SPREADSHEET_ID_SALES2 = '1GwKY8MY8aKEudRG6v-hKIUrjZd0WrV5KGAKpOBcvYCA'; // Replace with your sales2 sheet ID
const SPREADSHEET_ID_SALES3 = '10yCa--HOn4mBBsQXhEsj5FGeddEGvv776ZcRG29A014'; // Replace with your sales3 sheet ID
const SPREADSHEET_ID_SALES4 = '1nwJ-Uo7RVcXUtXuSE-PkG25kpc_GF1aA82nWN0WuwJI'; // Replace with your sales4 sheet ID
const SPREADSHEET_ID_SALES5 = '1gB0l50xioy_-5Q7qIeZ-ZAGgHJycuh6FCLVhl6jOvcs'; // Replace with your sales5 sheet ID
const SPREADSHEET_ID_SALES6 = '1pv_WOHLnrcXQ5VaeCr8f51Vf48UvueOOU7FOw-AJGFo'; // Replace with your sales6 sheet ID
const SPREADSHEET_ID_SALES7 = '14j4_EKrY2NXOxnAXpu6MPUdDOjioxR4t_y_95x1hgLs'; // Replace with your sales7 sheet ID
const SPREADSHEET_ID_SALES8 = '18b6e92gXN5w9vFiREWgOVerb7RlhR9V861NpXWWWJ0I'; // Replace with your sales7 sheet ID
const SPREADSHEET_ID_SALES9 = '1Ql1jOzipJQHb5A_loA9AM1pvL7SidCnCV5IXvbs3Sqo'; // Replace with your sales7 sheet ID

// Sales persons mapping
const SALES_PERSONS = {
  'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES1 },
  'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES2 },
  'hemant@enersol.co.in': { name: 'Hemant Trivedi', sheetId: SPREADSHEET_ID_SALES3 },
  'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES4 },
  'subhakanta.sahoo@enersol.co.in': { name: 'Shubhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES5 },
  'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES6 },
  'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES7 },
  'User1': { name: 'Test 1', sheetId: SPREADSHEET_ID_SALES8 },
  'User2': { name: 'Test 2', sheetId: SPREADSHEET_ID_SALES9 },
};

// Headers for Lead and ESL Sheets in Master Sheet
const LEAD_HEADERS_MASTER = [
  "Lead ID", "Created Date", "Sales Person Name", "Project Type", "Lead Origin", "Client Name", 
  "Expected Tentative Capacity", "Contact Person 1", "Designation 1", "Contact Number 1", 
  "Contact Person 2", "Designation 2", "Contact Number 2", "Area", "Pincode", 
  "City", "Co-Ordinates", "Remarks"
];

const LEAD_HEADERS_SALESPERSON = [
  "Lead ID", "Created Date", "Project Type", "Lead Origin", "Client Name", 
  "Expected Tentative Capacity", "Contact Person 1", "Designation 1", "Contact Number 1", 
  "Contact Person 2", "Designation 2", "Contact Number 2", "Area", "Pincode", 
  "City", "Co-Ordinates", "Remarks"
];

const ESL_HEADERS_MASTER = [
  "Created Date", "Lead ID", "ESL Number", "Sales Person Name", "Project Type", "Date", 
  "Client Name", "Final Capacity", "Consumer Number", "Application Number", 
  "Address", "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
];

const ESL_HEADERS_SALESPERSON = [
  "Created Date", "Lead ID", "ESL Number", "Project Type", "Date", 
  "Client Name", "Final Capacity", "Consumer Number", "Application Number (Blank)", 
  "Address", "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
];

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Utility to format date to DD/MM/YYYY
const formatDate = (isoDate) => {
  const date = new Date(isoDate);
  return `${('0' + date.getDate()).slice(-2)}/${('0' + (date.getMonth() + 1)).slice(-2)}/${date.getFullYear()}`;
};

// Function to ensure the sheet exists and has headers
const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
  try {
    // Check if the sheet exists
    const sheetResponse = await sheets.spreadsheets.get({
      spreadsheetId
    });

    let sheetExists = false;
    sheetResponse.data.sheets.forEach((sheet) => {
      if (sheet.properties.title === sheetName) {
        sheetExists = true;
      }
    });
 
    // Create the sheet if it doesn't exist
    if (!sheetExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetName
                }
              }
            }
          ]
        }
      });
    }

    // Ensure headers are present
    const range = `'${sheetName}'!A1:${String.fromCharCode(65 + headers.length - 1)}1`;
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

    if (!headerResponse.data.values || headerResponse.data.values.length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${sheetName}'!A1`,
        valueInputOption: 'RAW',
        resource: { values: [headers] }
      });
    }
  } catch (error) {
    console.error('Error ensuring sheet and headers:', error);
  }
};

// Function to append data to a sheet
const appendDataToSheet = async (spreadsheetId, sheetName, data, headers) => {
  const values = headers.map(header => {
    switch (header) {
      case "Lead ID":
        return data.leadId || '';
      case "Created Date":
        return data.date ? formatDate(data.date) : '';
      case "Sales Person Name":
        return data.salespersonName || '';
      case "Project Type":
        return data.projectType || '';
      case "Lead Origin":
        return data.leadOrigin || '';
      case "Client Name":
        return data.clientName || '';
      case "Expected Tentative Capacity":
        return +data.expectedProjectCapacity || '';
      case "Contact Person 1":
        return data.contactPersonName1 || '';
      case "Designation 1":
        return data.designation1 || '';
      case "Contact Number 1":
        return +data.contactNumber1 || '';
      case "Contact Person 2":
        return data.contactPersonName2 || '';
      case "Designation 2":
        return data.designation2 || '';
      case "Contact Number 2":
        return +data.contactNumber2 || '';
      case "Area":
        return data.area || '';
      case "Pincode":
        return +data.pincode || '';
      case "City":
        return data.city || '';
      case "Co-Ordinates":
        return data.coordinates || '';
      case "Remarks":
        return data.remarks || '';
      case "ESL Number":
        return data.eslNumber || '';
      case "Final Capacity":
        return +data.finalProjectCapacity || '';
      case "Consumer Number":
        return +data.consumerNumber || '';
      case "Address":
        return `${data.area || ''}, ${data.city || ''}, ${data.pincode || ''}`;
      case "Discom":
        return data.discom || '';
      case "Exe Time Contact Person":
        return data.exeTimeContactPerson || '';
      case "Exe Time Contact Number":
        return +data.exeTimeContactNumber || '';
      case "Date":
      case "Application Number (Blank)":
        return ''; // For blank columns
      default:
        return ''; // Fallback
    }
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${sheetName}'!A1`,
    valueInputOption: 'RAW',
    resource: { values: [values] }
  });
};

// Handle form submission
app.post('/form-data', async (req, res) => {
  console.log("Endpoint called");
  const formData = req.body;
  console.log(formData);

  try {
    const salesperson = SALES_PERSONS[formData.salesperson];
    formData.salespersonName = salesperson ? salesperson.name : 'Unknown';

    if (formData.eslNumber) {
      // ESL data
      await ensureSheetAndHeaders(SPREADSHEET_ID_MASTER, "ESL Sheet", ESL_HEADERS_MASTER);
      await appendDataToSheet(SPREADSHEET_ID_MASTER, "ESL Sheet", formData, ESL_HEADERS_MASTER);
      if (salesperson) {
        await ensureSheetAndHeaders(salesperson.sheetId, "ESL Sheet", ESL_HEADERS_SALESPERSON);
        await appendDataToSheet(salesperson.sheetId, "ESL Sheet", formData, ESL_HEADERS_SALESPERSON);
      }
    } else if (formData.leadId) {
      // Lead data
      await ensureSheetAndHeaders(SPREADSHEET_ID_MASTER, "Lead Sheet", LEAD_HEADERS_MASTER);
      await appendDataToSheet(SPREADSHEET_ID_MASTER, "Lead Sheet", formData, LEAD_HEADERS_MASTER);
      if (salesperson) {
        await ensureSheetAndHeaders(salesperson.sheetId, "Lead Sheet", LEAD_HEADERS_SALESPERSON);
        await appendDataToSheet(salesperson.sheetId, "Lead Sheet", formData, LEAD_HEADERS_SALESPERSON);
      }
    } else {
      throw new Error("Form data must contain either leadId or eslNumber.");
    }

    res.status(200).send('Form data received and added to the relevant sheets');
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
// const port = 3333;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1OaKEgNWWUEi1LLHyTVCoJMMeTQ4TE7NuB2Zwy4lFTzo'; // Replace with your master sheet ID
// const SPREADSHEET_ID_SALES1 = '1_eU7YevVyWs6OlGem4js_qL7KKYXDNXlloQVVIhyApc'; // Replace with your sales1 sheet ID
// const SPREADSHEET_ID_SALES2 = '1GwKY8MY8aKEudRG6v-hKIUrjZd0WrV5KGAKpOBcvYCA'; // Replace with your sales2 sheet ID
// const SPREADSHEET_ID_SALES3 = '10yCa--HOn4mBBsQXhEsj5FGeddEGvv776ZcRG29A014'; // Replace with your sales3 sheet ID
// const SPREADSHEET_ID_SALES4 = '1nwJ-Uo7RVcXUtXuSE-PkG25kpc_GF1aA82nWN0WuwJI'; // Replace with your sales4 sheet ID
// const SPREADSHEET_ID_SALES5 = '1gB0l50xioy_-5Q7qIeZ-ZAGgHJycuh6FCLVhl6jOvcs'; // Replace with your sales5 sheet ID
// const SPREADSHEET_ID_SALES6 = '1pv_WOHLnrcXQ5VaeCr8f51Vf48UvueOOU7FOw-AJGFo'; // Replace with your sales6 sheet ID
// const SPREADSHEET_ID_SALES7 = '14j4_EKrY2NXOxnAXpu6MPUdDOjioxR4t_y_95x1hgLs'; // Replace with your sales7 sheet ID

// // Sales persons mapping
// const SALES_PERSONS = {
//   'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES1 },
//   'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES2 },
//   'hemant@enersol.co.in': { name: 'Hemant Trivedi', sheetId: SPREADSHEET_ID_SALES3 },
//   'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES4 },
//   'subhakanta.sahoo@enersol.co.in': { name: 'Shubhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES5 },
//   'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES6 },
//   'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES7 },
// };

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Utility to format date to DD/MM/YYYY
// const formatDate = (isoDate) => {
//   const date = new Date(isoDate);
//   return `${('0' + date.getDate()).slice(-2)}/${('0' + (date.getMonth() + 1)).slice(-2)}/${date.getFullYear()}`;
// };

// // Function to append data to a sheet
// const appendDataToSheet = async (spreadsheetId, sheetName, data, isLead) => {
//   const requestBody = {
//     values: [
//       isLead ? [
//         data.leadId || '',
//         data.date ? formatDate(data.date) : '',
//         data.salespersonName || '',
//         data.projectType || '',
//         data.leadOrigin || '',
//         data.clientName || '',
//         data.expectedTentativeCapacity || '',
//         data.contactPersonName1 || '',
//         data.designation1 || '',
//         data.contactNumber1 || '',
//         data.contactPersonName2 || '',
//         data.designation2 || '',
//         data.contactNumber2 || '',
//         data.area || '',
//         data.pincode || '',
//         data.city || '',
//         data.coordinates || '',
//         data.remarks || ''
//       ] : [
//         data.date ? formatDate(data.date) : '',
//         data.leadId || '',
//         data.eslNumber || '',
//         data.salespersonName || '',
//         data.projectType || '',
//         '', // Blank Date field for ESL
//         data.clientName || '',
//         data.finalProjectCapacity || '',
//         data.consumerNumber || '',
//         '', // Blank Application Number for ESL
//         `${data.area || ''}, ${data.city || ''}, ${data.pincode || ''}`,
//         data.coordinates || '',
//         data.discom || '',
//         data.exeTimeContactPerson || '',
//         data.exeTimeContactNumber || ''
//       ]
//     ]
//   };

//   await sheets.spreadsheets.values.append({
//     spreadsheetId,
//     range: isLead ? `${sheetName}!A1:R1` : `${sheetName}!A1:O1`, // Adjust range based on the number of columns
//     valueInputOption: 'RAW',
//     resource: requestBody
//   });
// };

// // Handle form submission
// app.post('/form-data', async (req, res) => {
//   console.log("Endpoint called");
//   const formData = req.body;
//   console.log(formData);

//   try {
//     const salesperson = SALES_PERSONS[formData.salesperson];
//     formData.salespersonName = salesperson ? salesperson.name : 'Unknown';

//     if (formData.leadId) {
//       // Lead data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "Lead_Sheet", formData, true);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "Lead_Sheet", formData, true);
//       }
//     } else if (formData.eslNumber) {
//       // ESL data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "ESL_Sheet", formData, false);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "ESL_Sheet", formData, false);
//       }
//     } else {
//       throw new Error("Form data must contain either leadId or eslNumber.");
//     }

//     res.status(200).send('Form data received and added to the relevant sheets');
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



{
// const { google } = require('googleapis');
// const express = require('express');
// const bodyParser = require('body-parser');
// const cors = require('cors');
// const dotenv = require('dotenv');

// // Load environment variables from .env file
// dotenv.config();

// const app = express();
// const port = 3333;

// // Google Sheets API credentials loaded from .env
// const client_email = process.env.SERVICE_ACCOUNT_EMAIL;
// let private_key = process.env.SERVICE_ACCOUNT_PRIVATE_KEY;
// if (private_key.startsWith('"-----BEGIN PRIVATE KEY-----')) {
//   private_key = JSON.parse(`{"key":${private_key}}`).key; // Remove escaped quotes
// }

// const client = new google.auth.JWT(
//   client_email,
//   null,
//   private_key,
//   ['https://www.googleapis.com/auth/spreadsheets']
// );

// const sheets = google.sheets({ version: 'v4', auth: client });

// // Sheet IDs
// const SPREADSHEET_ID_MASTER = '1OaKEgNWWUEi1LLHyTVCoJMMeTQ4TE7NuB2Zwy4lFTzo'; // Replace with your master sheet ID
// const SPREADSHEET_ID_SALES1 = '1_eU7YevVyWs6OlGem4js_qL7KKYXDNXlloQVVIhyApc'; // Replace with your sales1 sheet ID
// const SPREADSHEET_ID_SALES2 = '1GwKY8MY8aKEudRG6v-hKIUrjZd0WrV5KGAKpOBcvYCA'; // Replace with your sales2 sheet ID
// const SPREADSHEET_ID_SALES3 = '10yCa--HOn4mBBsQXhEsj5FGeddEGvv776ZcRG29A014'; // Replace with your sales3 sheet ID
// const SPREADSHEET_ID_SALES4 = '1nwJ-Uo7RVcXUtXuSE-PkG25kpc_GF1aA82nWN0WuwJI'; // Replace with your sales4 sheet ID
// const SPREADSHEET_ID_SALES5 = '1gB0l50xioy_-5Q7qIeZ-ZAGgHJycuh6FCLVhl6jOvcs'; // Replace with your sales5 sheet ID
// const SPREADSHEET_ID_SALES6 = '1pv_WOHLnrcXQ5VaeCr8f51Vf48UvueOOU7FOw-AJGFo'; // Replace with your sales6 sheet ID
// const SPREADSHEET_ID_SALES7 = '14j4_EKrY2NXOxnAXpu6MPUdDOjioxR4t_y_95x1hgLs'; // Replace with your sales7 sheet ID

// // Sales persons mapping
// const SALES_PERSONS = {
//   'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES1 },
//   'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES2 },
//   'hemant@enersol.co.in': { name: 'Hemant Trivedi 3', sheetId: SPREADSHEET_ID_SALES3 },
//   'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES4 },
//   'subhakanta.sahoo@enersol.co.in': { name: 'Shubhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES5 },
//   'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES6 },
//   'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES7 },
// };

// // Define headers for Lead and ESL sheets
// const LEAD_HEADERS_MASTER = [
//   "leadId", "Date", "Sales Person Name", "Project Type", "Lead Origin", "Client Name", 
//   "Expected Tentative Capacity", "Discom", "Contact Person 1", "Designation 1", 
//   "Contact Number 1", "Contact Person 2", "Designation 2", "Contact Number 2", 
//   "Area", "Pincode", "City", "Co-Ordinates", "Remarks"
// ];

// const LEAD_HEADERS_SALESPERSON = [
//   "leadId", "Date", "Project Type", "Lead Origin", "Client Name", 
//   "Expected Tentative Capacity", "Contact Person 1", "Designation 1", "Contact Number 1", 
//   "Contact Person 2", "Designation 2", "Contact Number 2", "Area", "Pincode", 
//   "City", "Co-Ordinates", "Remarks"
// ];

// const ESL_HEADERS_MASTER = [
//   "Date", "Lead Id", "ESL Id", "Sales Person Name", "Project Type", "Date (Blank)", 
//   "Customer Name", "Final Capacity", "Consumer Number", "Application Number (Blank)", 
//   "Address", "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
// ];

// const ESL_HEADERS_SALESPERSON = [
//   "Date", "Lead Id", "ESL Id", "Project Type", "Date (Blank)", "Customer Name", 
//   "Final Capacity", "Consumer Number", "Application Number (Blank)", "Address", 
//   "Co-Ordinates", "Discom", "Exe Time Contact Person", "Exe Time Contact Number"
// ];

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Utility to format date to DD/MM/YYYY
// const formatDate = (isoDate) => {
//   const date = new Date(isoDate);
//   return `${('0' + date.getDate()).slice(-2)}/${('0' + (date.getMonth() + 1)).slice(-2)}/${date.getFullYear()}`;
// };

// // Helper function to ensure sheet and headers
// const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
//   const sheetExists = await sheets.spreadsheets.get({
//     spreadsheetId,
//   }).then(response => {
//     const sheet = response.data.sheets.find(sheet => sheet.properties.title === sheetName);
//     return !!sheet;
//   });

//   if (!sheetExists) {
//     // Create the sheet
//     await sheets.spreadsheets.batchUpdate({
//       spreadsheetId,
//       resource: {
//         requests: [
//           {
//             addSheet: {
//               properties: {
//                 title: sheetName
//               }
//             }
//           }
//         ]
//       }
//     });
//   }

//   // Check if headers exist
//   const headerRow = await sheets.spreadsheets.values.get({
//     spreadsheetId,
//     range: `${sheetName}!A1:${String.fromCharCode(65 + headers.length)}1`
//   });

//   if (!headerRow.data.values || headerRow.data.values[0].length !== headers.length) {
//     // Set headers if they don't match
//     await sheets.spreadsheets.values.update({
//       spreadsheetId,
//       range: `${sheetName}!A1`,
//       valueInputOption: 'RAW',
//       resource: {
//         values: [headers]
//       }
//     });
//   }
// };

// // Function to append data to a sheet
// const appendDataToSheet = async (spreadsheetId, sheetName, formData, headers, includeSalespersonName = false, isMaster = false) => {
//   await ensureSheetAndHeaders(spreadsheetId, sheetName, headers);

//   const newRow = headers.map(header => {
//     switch (header) {
//       case "Sales Person Name":
//         return formData.salespersonName;
//       case "Date":
//         return formData.date ? formatDate(formData.date) : ''; // Format the date to DD/MM/YYYY
//       case "Date (Blank)":
//         return ''; // For ESL Sheets, this should be blank
//       case "Address":
//         return `${formData.area}, ${formData.city}, ${formData.pincode}`; // Concatenate area, city, and pincode
//       case "Application Number (Blank)":
//         return ''; // ESL Application Number is blank
//       default:
//         return formData[header.replace(/ /g, "").replace(/\(.+?\)/g, '')] || ''; // Remove spaces and parenthesis
//     }
//   });

//   await sheets.spreadsheets.values.append({
//     spreadsheetId,
//     range: `${sheetName}!A:${String.fromCharCode(65 + headers.length)}`, // Adjust range based on headers length
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
//   console.log(formData);

//   try {
//     const salesperson = SALES_PERSONS[formData.salesperson];
//     formData.salespersonName = salesperson ? salesperson.name : 'Unknown';

//     if (formData.eslNumber) {
//       // ESL data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "ESL_Sheet", formData, ESL_HEADERS_MASTER, true, true);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "ESL_Sheet", formData, ESL_HEADERS_SALESPERSON, false);
//       }
//     } else if (formData.leadId) {
//       // Lead data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "Lead_Sheet", formData, LEAD_HEADERS_MASTER, true, true);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "Lead_Sheet", formData, LEAD_HEADERS_SALESPERSON, false);
//       }
//     } else {
//       throw new Error("Form data must contain either leadId or eslNumber.");
//     }

//     res.status(200).send('Form data received and added to the relevant sheets');
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
}

{
// // Define headers for Lead and ESL sheets
// const LEAD_HEADERS = [
//   "leadId", "date", "projectType", "leadOrigin", "clientName", "expectedProjectCapacity",
//   "contactPersonName", "designation", "contactNumber", "contactNumber2", "area", "pincode",
//   "city", "remarks", "salesperson"
// ];

// const ESL_HEADERS = [
//   "eslNumber", "date", "projectType", "projectPortalType", "leadId", "leadOrigin", "name",
//   "expectedTentativeCapacity", "finalProjectCapacity", "consumerNumber", "discom", "contactPerson",
//   "designation", "contactNumber", "contactNumber2", "area", "pincode", "city", "remarksNotes",
//   "solarPanelsType", "panelMake1", "panelMake2", "inverter", "structure", "specialRemarks",
//   "pricePerKw", "gedaCharges", "meteringCharges", "paymentTerms", "specialCommercialRemarks",
//   "leadAddedBy", "addedBy"
// ];

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Helper function to ensure sheet and headers
// const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
//   const sheetExists = await sheets.spreadsheets.get({
//     spreadsheetId,
//   }).then(response => {
//     const sheet = response.data.sheets.find(sheet => sheet.properties.title === sheetName);
//     return !!sheet;
//   });

//   if (!sheetExists) {
//     // Create the sheet
//     await sheets.spreadsheets.batchUpdate({
//       spreadsheetId,
//       resource: {
//         requests: [
//           {
//             addSheet: {
//               properties: {
//                 title: sheetName
//               }
//             }
//           }
//         ]
//       }
//     });
//   }

//   // Check if headers exist
//   const headerRow = await sheets.spreadsheets.values.get({
//     spreadsheetId,
//     range: `${sheetName}!A1:${String.fromCharCode(65 + headers.length)}1`
// });

//   if (!headerRow.data.values || headerRow.data.values[0].length !== headers.length) {
//     // Set headers if they don't match
//     await sheets.spreadsheets.values.update({
//       spreadsheetId,
//       range: `${sheetName}!A1`,
//       valueInputOption: 'RAW',
//       resource: {
//         values: [headers]
//       }
//     });
//   }
// };

// // Function to append data to a sheet
// const appendDataToSheet = async (spreadsheetId, sheetName, formData, headers, includeSalespersonName = false) => {
//   await ensureSheetAndHeaders(spreadsheetId, sheetName, headers);

//   const newRow = headers.map(header => formData[header] || '');

//   if (includeSalespersonName) {
//     newRow.push(formData.salespersonName);
//   }

//   await sheets.spreadsheets.values.append({
//     spreadsheetId,
//     range: `${sheetName}!A:${String.fromCharCode(65 + headers.length)}`, // Adjust range based on headers length
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
//   console.log(formData)

//   try {
//     const salesperson = SALES_PERSONS[formData.salesperson];
//     formData.salespersonName = salesperson ? salesperson.name : 'Unknown';

//     if (formData.eslNumber) {
//       // ESL data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "ESL_Sheet", formData, ESL_HEADERS, true);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "ESL_Sheet", formData, ESL_HEADERS, false);
//       }
//     } else if (formData.leadId) {
//       // Lead data
//       await appendDataToSheet(SPREADSHEET_ID_MASTER, "Lead_Sheet", formData, LEAD_HEADERS, true);
//       if (salesperson) {
//         await appendDataToSheet(salesperson.sheetId, "Lead_Sheet", formData, LEAD_HEADERS, false);
//       }
//     } else {
//       throw new Error("Form data must contain either leadId or eslNumber.");
//     }

//     res.status(200).send('Form data received and added to the relevant sheets');
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
}


// {const { google } = require('googleapis');
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
// const SPREADSHEET_ID_MASTER = '1xcetIqD5BesWLlAOtRbXOu7SuM7OP9pbmjBNVHkL5OE'; // Replace with your master sheet ID
// const SPREADSHEET_ID_SALES1 = '1d3_LxjSA-1Esl0rKtZu7WbY18Y3SvJDn0R5sRX-Q9H0'; // Replace with your sales1 sheet ID
// const SPREADSHEET_ID_SALES2 = '1tvtVQ20gSN3vJPJ6KWfs-B7QX39MZfF5GZZAuFhk6Zs'; // Replace with your sales2 sheet ID
// const SPREADSHEET_ID_SALES3 = '1S0Obv2Juk7F3D8L1b-J4aaC1VWgfacm_SgcRXpBqvtM'; // Replace with your sales3 sheet ID

// // Sales persons mapping
// const SALES_PERSONS = {
//   'User1': { name: 'Sales Person 1', sheetId: SPREADSHEET_ID_SALES1 },
//   'User2': { name: 'Sales Person 2', sheetId: SPREADSHEET_ID_SALES2 },
//   'User3': { name: 'Sales Person 3', sheetId: SPREADSHEET_ID_SALES3 }
// };

// // Middleware
// app.use(cors());
// app.use(bodyParser.json());

// // Function to append data to a sheet
// const appendDataToSheet = async (sheetId, formData, includeSalespersonName) => {
//   const newRow = [
//     formData.leadId,
//     (date => `${String(date.getDate()).padStart(2, '0')}-${String(date.getMonth() + 1).padStart(2, '0')}-${date.getFullYear()}`)(new Date(formData.date)),
//     formData.projectType,
//     formData.leadOrigin,
//     formData.clientName,
//     +formData.expectedProjectCapacity,
//     formData.contactPersonName,
//     formData.designation,
//     +formData.contactNumber,
//     +formData.contactNumber2,
//     formData.area,
//     formData.city,
//     formData.remarks
//   ];

//   if (includeSalespersonName) {
//     newRow.push(formData.salespersonName); // Add salesperson name if needed
//   }

//   // Append to the specified sheet
//   await sheets.spreadsheets.values.append({
//     spreadsheetId: sheetId,
//     range: 'Sheet1!A:N', // Adjust the range if needed
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
//     // Add salesperson name to formData for processing
//     const salesperson = SALES_PERSONS[formData.salesperson];
//     if (salesperson) {
//       formData.salespersonName = salesperson.name;
//     } else {
//       formData.salespersonName = 'Unknown';
//     }

//     // Append data to the master sheet
//     await appendDataToSheet(SPREADSHEET_ID_MASTER, formData, true);

//     // Append data to the sales person's sheet without salesperson name
//     if (salesperson) {
//       await appendDataToSheet(salesperson.sheetId, formData, false);
//     }

//     res.status(200).send('Form data received and added to master and sales sheets');
//   } 
//   catch (error) {
//     console.error('Error processing form data:', error);
//     res.status(500).send('Internal Server Error');
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Express server listening at http://localhost:${port}`);
// });} datetime : 11-08-2024/00:04




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


// {
// const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
//   const lastColumnLetter = String.fromCharCode(64 + headers.length);
//   const range = `${sheetName}!A1:${lastColumnLetter}1`;

//   console.log('Generated Range:', range); // Debugging statement

//   const sheetExists = await sheets.spreadsheets.get({
//     spreadsheetId,
//   }).then(response => {
//     const sheet = response.data.sheets.find(sheet => sheet.properties.title === sheetName);
//     return !!sheet;
//   });

//   if (!sheetExists) {
//     await sheets.spreadsheets.batchUpdate({
//       spreadsheetId,
//       resource: {
//         requests: [
//           {
//             addSheet: {
//               properties: {
//                 title: sheetName
//               }
//             }
//           }
//         ]
//       }
//     });
//   }

//   const headerRow = await sheets.spreadsheets.values.get({
//     spreadsheetId,
//     range: range
//   });

//   if (!headerRow.data.values || headerRow.data.values[0].length !== headers.length) {
//     await sheets.spreadsheets.values.update({
//       spreadsheetId,
//       range: `${sheetName}!A1`,
//       valueInputOption: 'RAW',
//       resource: {
//         values: [headers]
//       }
//     });
//   }
// };
// }