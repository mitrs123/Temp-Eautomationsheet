
const { google } = require("googleapis");
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const dotenv = require("dotenv");

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

const client = new google.auth.JWT(client_email, null, private_key, [
  "https://www.googleapis.com/auth/spreadsheets",
]);

const sheets = google.sheets({ version: "v4", auth: client });

// Sheet IDs
// const SPREADSHEET_ID_MASTER = '18b6e92gXN5w9vFiREWgOVerb7RlhR9V861NpXWWWJ0I'; // Replace with your master sheet ID
const SPREADSHEET_ID_MASTER = "1OaKEgNWWUEi1LLHyTVCoJMMeTQ4TE7NuB2Zwy4lFTzo"; // Replace with your master sheet ID
const SPREADSHEET_ID_SALES1 = "1_eU7YevVyWs6OlGem4js_qL7KKYXDNXlloQVVIhyApc"; // Replace with your sales1 sheet ID
const SPREADSHEET_ID_SALES2 = "1GwKY8MY8aKEudRG6v-hKIUrjZd0WrV5KGAKpOBcvYCA"; // Replace with your sales2 sheet ID
const SPREADSHEET_ID_SALES3 = "10yCa--HOn4mBBsQXhEsj5FGeddEGvv776ZcRG29A014"; // Replace with your sales3 sheet ID
const SPREADSHEET_ID_SALES4 = "1nwJ-Uo7RVcXUtXuSE-PkG25kpc_GF1aA82nWN0WuwJI"; // Replace with your sales4 sheet ID
const SPREADSHEET_ID_SALES5 = "1gB0l50xioy_-5Q7qIeZ-ZAGgHJycuh6FCLVhl6jOvcs"; // Replace with your sales5 sheet ID
const SPREADSHEET_ID_SALES6 = "1pv_WOHLnrcXQ5VaeCr8f51Vf48UvueOOU7FOw-AJGFo"; // Replace with your sales6 sheet ID
const SPREADSHEET_ID_SALES7 = "14j4_EKrY2NXOxnAXpu6MPUdDOjioxR4t_y_95x1hgLs"; // Replace with your sales7 sheet ID
const SPREADSHEET_ID_SALES8 = "18b6e92gXN5w9vFiREWgOVerb7RlhR9V861NpXWWWJ0I"; // Replace with your sales7 sheet ID
const SPREADSHEET_ID_SALES9 = "1Ql1jOzipJQHb5A_loA9AM1pvL7SidCnCV5IXvbs3Sqo"; // Replace with your sales7 sheet ID

// Sales persons mapping
const SALES_PERSONS = {
  'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES1 },
  'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES2 },
  'hemant@enersol.co.in': { name: 'Hemant Trivedi', sheetId: SPREADSHEET_ID_SALES3 },
  'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES4 },
  'subhakanta.sahoo@enersol.co.in': { name: 'Subhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES5 },
  'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES6 },
  'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES7 },
  'User1': { name: 'Test 1', sheetId: SPREADSHEET_ID_SALES8 },
  'User2': { name: 'Test 2', sheetId: SPREADSHEET_ID_SALES9 },
  "kushal@enersol.co.in": {
    name: "Kushal Bhansali",
    sheetId: SPREADSHEET_ID_SALES1,
  },
  "karan@enersol.co.in": {
    name: "Karan Bhansali",
    sheetId: SPREADSHEET_ID_SALES2,
  },
  "hemant@enersol.co.in": {
    name: "Hemant Trivedi",
    sheetId: SPREADSHEET_ID_SALES3,
  },
  "jay.chauhan@enersol.co.in": {
    name: "Jay Chauhan",
    sheetId: SPREADSHEET_ID_SALES4,
  },
  "subhakanta.sahoo@enersol.co.in": {
    name: "Subhakanta Sahoo",
    sheetId: SPREADSHEET_ID_SALES5,
  },
  "akshay.panchal@enersol.co.in": {
    name: "Akshay Panchal",
    sheetId: SPREADSHEET_ID_SALES6,
  },
  "furkan.banva@enersol.co.in": {
    name: "Furkan Banva",
    sheetId: SPREADSHEET_ID_SALES7,
  },
  User1: { name: "Test 1", sheetId: SPREADSHEET_ID_SALES8 },
  User2: { name: "Test 2", sheetId: SPREADSHEET_ID_SALES9 },
};


// Sales persons mapping
// const SALES_PERSONS = {
//   'kushal@enersol.co.in': { name: 'Kushal Bhansali', sheetId: SPREADSHEET_ID_SALES9 },
//   'karan@enersol.co.in': { name: 'Karan Bhansali', sheetId: SPREADSHEET_ID_SALES9 },
//   'hemant@enersol.co.in': { name: 'Hemant Trivedi', sheetId: SPREADSHEET_ID_SALES9 },
//   'jay.chauhan@enersol.co.in': { name: 'Jay Chauhan', sheetId: SPREADSHEET_ID_SALES9 },
//   'subhakanta.sahoo@enersol.co.in': { name: 'Shubhakanta Sahoo', sheetId: SPREADSHEET_ID_SALES9 },
//   'akshay.panchal@enersol.co.in': { name: 'Akshay Panchal', sheetId: SPREADSHEET_ID_SALES9 },
//   'furkan.banva@enersol.co.in': { name: 'Furkan Banva', sheetId: SPREADSHEET_ID_SALES9 },
//   'User1': { name: 'Test 1', sheetId: SPREADSHEET_ID_SALES9 },
//   'User2': { name: 'Test 2', sheetId: SPREADSHEET_ID_SALES9 },
// };

// Headers for Lead and ESL Sheets in Master Sheet
const LEAD_HEADERS_MASTER = [
  "Lead ID",
  "Created Date",
  "Sales Person Name",
  "Project Type",
  "Lead Origin",
  "Client Name",
  "Expected Tentative Capacity",
  "Discom",
  "Contact Person 1",
  "Designation 1",
  "Contact Number 1",
  "Contact Person 2",
  "Designation 2",
  "Contact Number 2",
  "Email",
  "Area",
  "Pincode",
  "City",
  "Co-Ordinates",
  "Remarks",
];

const LEAD_HEADERS_SALESPERSON = [
  "Lead ID",
  "Created Date",
  "Project Type",
  "Lead Origin",
  "Client Name",
  "Expected Tentative Capacity",
  "Discom",
  "Contact Person 1",
  "Designation 1",
  "Contact Number 1",
  "Contact Person 2",
  "Designation 2",
  "Contact Number 2",
  "Email",
  "Area",
  "Pincode",
  "City",
  "Co-Ordinates",
  "Remarks",
];
//  ESL headers
const ESL_HEADERS_MASTER = [
  "Created Date",
  "Lead ID",
  "ESL Number",
  "Sales Person Name",
  "Project Type",
  "Date",
  "Client Name",
  "Final Capacity",
  "Consumer Number",
  "Application Number",
  "Address",
  "Co-Ordinates",
  "Discom",
  "Contact Person",
  "Contact Number",
  "Exe Time Contact Person",
  "Exe Time Contact Number",
  "Email",
];

const ESL_HEADERS_SALESPERSON = [
  "Created Date",
  "Lead ID",
  "ESL Number",
  "Project Type",
  "Client Name",
  "Final Capacity",
  "Consumer Number",
  "Address",
  "Co-Ordinates",
  "Discom",
  "Contact Person",
  "Contact Number",
  "Exe Time Contact Person",
  "Exe Time Contact Number",
  "Email",
];

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Utility to format date to DD/MM/YYYY
const formatDate = (isoDate) => {
  const date = new Date(isoDate);
  return `${("0" + date.getDate()).slice(-2)}/${(
    "0" +
    (date.getMonth() + 1)
  ).slice(-2)}/${date.getFullYear()}`;
};

// Function to ensure the sheet exists and has headers
const ensureSheetAndHeaders = async (spreadsheetId, sheetName, headers) => {
  try {
    // Check if the sheet exists
    const sheetResponse = await sheets.spreadsheets.get({
      spreadsheetId,
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
                  title: sheetName,
                },
              },
            },
          ],
        },
      });
    }

    // Ensure headers are present
    const range = `'${sheetName}'!A1:${String.fromCharCode(
      65 + headers.length - 1
    )}1`;
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    if (
      !headerResponse.data.values ||
      headerResponse.data.values.length === 0
    ) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${sheetName}'!A1`,
        valueInputOption: "RAW",
        resource: { values: [headers] },
      });
    }
  } catch (error) {
    console.error("Error ensuring sheet and headers:", error);
  }
};

// Function to append data to a sheet
const appendDataToSheet = async (spreadsheetId, sheetName, data, headers) => {
  const values = headers.map((header) => {
    switch (header) {
      case "Lead ID":
        return data.leadId || "";
      case "Created Date":
        return data.date ? formatDate(data.date) : "";
      case "Sales Person Name":
        return data.salespersonName || "";
      case "Project Type":
        return data.projectType || "";
      case "Lead Origin":
        return data.leadOrigin || "";
      case "Client Name":
        return data.clientName || "";
      case "Expected Tentative Capacity":
        return +data.expectedProjectCapacity || "";
      case "Contact Person 1":
        return data.contactPersonName1 || "";
      case "Designation 1":
        return data.designation1 || "";
      case "Contact Number 1":
        return +data.contactNumber1 || "";
      case "Contact Person 2":
        return data.contactPersonName2 || "";
      case "Designation 2":
        return data.designation2 || "";
      case "Contact Number 2":
        return +data.contactNumber2 || "";
      case "Area":
        return data.area || "";
      case "Pincode":
        return +data.pincode || "";
      case "City":
        return data.city || "";
      case "Co-Ordinates":
        return data.coordinates || "";
      case "Remarks":
        return data.remarks || "";
      case "ESL Number":
        return data.eslNumber || "";
      case "Final Capacity":
        return +data.finalProjectCapacity || "";
      case "Consumer Number":
        return +data.consumerNumber || "";
      case "Address":
        return `${data.area || ""}, ${data.city || ""}, ${data.pincode || ""}`;
      case "Discom":
        return data.discom || "";
      case "Contact Person":
        return data.contactPerson || "";
      case "Contact Number":
        return +data.contactNumber || "";
      case "Exe Time Contact Person":
        return data.exeTimeContactPerson || "";
      case "Email":
        return data.email || "";
      case "Exe Time Contact Number":
        return +data.exeTimeContactNumber || "";
      case "Date":
      case "Application Number (Blank)":
        return ""; // For blank columns
      default:
        return ""; // Fallback
    }
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${sheetName}'!A1`,
    valueInputOption: "RAW",
    resource: { values: [values] },
  });
};

// Handle form submission
app.post("/form-data", async (req, res) => {
  console.log("Endpoint called");
  const formData = req.body;
  console.log(formData);

  try {
    const salesperson = SALES_PERSONS[formData.salesperson];
    formData.salespersonName = salesperson ? salesperson.name : "Unknown";

    if (formData.eslNumber) {
      // ESL data
      await ensureSheetAndHeaders(
        SPREADSHEET_ID_MASTER,
        "ESL Sheet",
        ESL_HEADERS_MASTER
      );
      await appendDataToSheet(
        SPREADSHEET_ID_MASTER,
        "ESL Sheet",
        formData,
        ESL_HEADERS_MASTER
      );
      if (salesperson) {
        await ensureSheetAndHeaders(
          salesperson.sheetId,
          "ESL Sheet",
          ESL_HEADERS_SALESPERSON
        );
        await appendDataToSheet(
          salesperson.sheetId,
          "ESL Sheet",
          formData,
          ESL_HEADERS_SALESPERSON
        );
      }
    } else if (formData.leadId) {
      // Lead data
      await ensureSheetAndHeaders(
        SPREADSHEET_ID_MASTER,
        "Lead Sheet",
        LEAD_HEADERS_MASTER
      );
      await appendDataToSheet(
        SPREADSHEET_ID_MASTER,
        "Lead Sheet",
        formData,
        LEAD_HEADERS_MASTER
      );
      if (salesperson) {
        await ensureSheetAndHeaders(
          salesperson.sheetId,
          "Lead Sheet",
          LEAD_HEADERS_SALESPERSON
        );
        await appendDataToSheet(
          salesperson.sheetId,
          "Lead Sheet",
          formData,
          LEAD_HEADERS_SALESPERSON
        );
      }
    } else {
      throw new Error("Form data must contain either leadId or eslNumber.");
    }

    res.status(200).send("Form data received and added to the relevant sheets");
  } catch (error) {
    console.error("Error processing form data:", error);
    res.status(500).send("Internal Server Error");
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Express server listening at http://localhost:${port}`);
});
