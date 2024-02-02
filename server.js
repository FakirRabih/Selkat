const express = require('express');
const app = express();
const port = 3000;
const { Client, LocalAuth } = require('whatsapp-web.js');
const { MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});
app.use(express.static('public'));

// Path to the session data file
const SESSION_FILE_PATH = './session.json';

app.use(express.json());

// Load the session data if it exists
let sessionData;
if (fs.existsSync(SESSION_FILE_PATH)) {
  sessionData = require(SESSION_FILE_PATH);
}

const client = new Client({
  puppeteer: {
    headless: true,
  },
  authStrategy: new LocalAuth({
    clientId: "YOUR_CLIENT_ID",
  }),
});


client.on("qr", (qr) => {
  console.log("QR RECEIVED", qr);
});

client.on("ready", async () => {
  console.log("Client is ready!");
  await openWorksheet();
});


const EXCEL_FILE_PATH = 'H:\\Selka_Sink_1.xlsx';
  const SHEET_NAME = 'Recipient';
const OUTPUT_SHEET_NAME = 'Node_Sink'; // Define OUTPUT_SHEET_NAME here


// Event listeners and logic for interacting with WhatsApp
// add  “Microsoft Scripting Runtime” in VBA --> Tools  
async function openWorksheet() {
  const EXCEL_FILE_PATH = 'H:\\Selka_Sink_1.xlsx';  
  const SHEET_NAME = 'Recipient';
  const OUTPUT_SHEET_NAME = 'Node_Sink'; // Define OUTPUT_SHEET_NAME here

  let workbook, sheet;

  try {
    workbook = await XlsxPopulate.fromFileAsync(EXCEL_FILE_PATH);
    sheet = workbook.sheet(SHEET_NAME);
  } catch (error) {
    workbook = await XlsxPopulate.fromBlankAsync();
    sheet = workbook.sheet(0);
    sheet.name(SHEET_NAME);
    sheet.cell(1, 1).value('Sender Name');
    sheet.cell(1, 2).value('Time Received');
    sheet.cell(1, 3).value('Message');
  }
  console.log('Workbook:', workbook);
  console.log('Sheet:', sheet);

  return { workbook, sheet };
}

async function updateWorksheet(senderName, time, newData, workbook, sheet) {
//  const lastRow = sheet.usedRange().endCell().rowNumber();
  let lastRow;
  if (sheet.usedRange()) {
    lastRow = sheet.usedRange().endCell().rowNumber();
  } else {
  lastRow = 1; // or wherever you want to start adding data
 }

  sheet.cell(`A${lastRow + 1}`).value(senderName);
  sheet.cell(`B${lastRow + 1}`).value(time);
  sheet.cell(`C${lastRow + 1}`).value(newData);

 // sheet.cell(lastRow + 1, 1).value(senderName);
 // sheet.cell(lastRow + 1, 2).value(time);
 // sheet.cell(lastRow + 1, 3).value(newData);

  console.log('New Data:', newData);  
  console.log('Time:', time);
  console.log('New Data:', newData);

  await workbook.toFileAsync(EXCEL_FILE_PATH);
}


async function flushChanges(workbook) {
  if (workbook) {
    await workbook.toFileAsync(EXCEL_FILE_PATH);
    console.log('Changes flushed to disk');
  }
}


app.post('/start-logging', async (req, res) => {
  try {
    const groupName = 'السلكة الأسبوعية المحمدية';
 //   const groupName = 'المحبون في بلقايد 2';
    const FLUSH_THRESHOLD = 180;

    let { workbook, sheet } = await openWorksheet();
    let messageCount = 0;
    

    client.on('message', async (msg) => {
      console.log('Received message:', msg.body);
      try {
        const chat = await msg.getChat();
        if (chat.isGroup && chat.name.includes(groupName)) {
          const contact = await msg.getContact();
          const senderName = contact.id.user;
          const time = new Date().toLocaleString();
          const messageBody = msg.body;

          await updateWorksheet(senderName, time, messageBody, workbook, sheet);

          messageCount++;

          if (messageCount % FLUSH_THRESHOLD === 0) {
            await flushChanges(workbook);
          }
        }
       // res.sendStatus(200);
      } catch (error) {
        console.error('Error:', error.message);
      }
    });
    await applyFormula(workbook);
    // Flush changes to disk on program exit
    process.on('SIGINT', async () => {
      await flushChanges(workbook);
      process.exit(0);
    });

    // Start the WhatsApp client
    client.initialize();

    res.sendStatus(200);
  } catch (error) {
    console.error('Error:', error.message);
    res.sendStatus(500);
  }
});
async function applyFormula(workbook) {
  const OUTPUT_SHEET_NAME1 = 'Node_Sink1';
  const OUTPUT_SHEET_NAME2 = 'Node_Sink2';
  const sheet = workbook.sheet(SHEET_NAME);
  
  // Check if the output sheets exist, if not create them
  let outputSheet1 = workbook.sheet(OUTPUT_SHEET_NAME1);
  if (!outputSheet1) {
    outputSheet1 = workbook.addSheet(OUTPUT_SHEET_NAME1);
  }

  let outputSheet2 = workbook.sheet(OUTPUT_SHEET_NAME2);
  if (!outputSheet2) {
    outputSheet2 = workbook.addSheet(OUTPUT_SHEET_NAME2);
  }

  let rowCounter1 = 1; // Initialize rowCounter for Table1
  let rowCounter2 = 1; // Initialize rowCounter for Table2

  // Initialize all cells in the output range to an empty string
  for (let i = 1; i <= 180; i++) {
    outputSheet1.cell(`A${i}`).value('');
    outputSheet1.cell(`B${i}`).value(''); // Initialize column B as well
    outputSheet1.cell(`C${i}`).value(''); // Initialize column C for Timestamp

    outputSheet2.cell(`A${i}`).value(''); // Initialize column A for "Buy: " results
    outputSheet2.cell(`B${i}`).value(''); // Initialize column B for Timestamp
  }

  // Loop through each cell in the range for "اطلب"
  for (let i = 2; i <= 180; i++) {
    let cell = sheet.cell(`C${i}`);
    let value = cell.value();

    if (value && (value.includes("اطلب") || value.includes("أطلب"))) {
      // Get the corresponding value in column A
      let correspondingValue = sheet.cell(`A${i}`).value();

      // Get the timestamp from column2
      let timestamp = sheet.cell(`B${i}`).value();

      // Get the number of repetitions
      let repetitions;
      if (value.includes("اطلب")) {
        repetitions = parseInt(value.split("اطلب")[1]);
      } else if (value.includes("أطلب")) {
        repetitions = parseInt(value.split("أطلب")[1]);
      }     

      // Repeat the corresponding value the specified number of times  
      for (let j = 0; j < repetitions; j++) {
        // Write the value to a cell in the same row in column A of Table1
        outputSheet1.cell(`A${rowCounter1}`).value(correspondingValue);
        // Write the number of repetitions to a cell in the same row in column B of Table1
        outputSheet1.cell(`B${rowCounter1}`).value(repetitions);
        // Write the timestamp to a cell in the same row in column C of Table1
        outputSheet1.cell(`C${rowCounter1}`).value(timestamp);
        rowCounter1++;
      }
    }
  }

  // Reset rowCounter for "Buy: " condition
  rowCounter2 = 1;

  // Loop through each cell in the range for "Buy: "
  for (let i = 2; i <= 180; i++) {
    let cell = sheet.cell(`C${i}`);
    let value = cell.value();

    // Check if "Buy: " is found in the cell
    if (value && (value.includes("قراءة") ||  value.includes("تم"))) {
      // Get the corresponding value in column A
      let correspondingValue = sheet.cell(`A${i}`).value();

      // Get the timestamp from column2
      let timestamp = sheet.cell(`B${i}`).value();

      // Get the number of repetitions
      let repetitions;
      if (value.includes("قراءة")) {
        repetitions = parseInt(value.split("قراءة")[1]);
      } else if (value.includes("تم")) {
        repetitions = parseInt(value.split("تم")[1]);
      } 

      // Repeat the corresponding value the specified number of times  
      for (let j = 0; j < repetitions; j++) {
        // Write the value to a cell in the same row in column A of Table2
        outputSheet2.cell(`A${rowCounter2}`).value(correspondingValue);
        // Write the timestamp to a cell in the same row in column B of Table2
        outputSheet2.cell(`B${rowCounter2}`).value(timestamp);
        rowCounter2++;
      }
    }
  }

  await workbook.toFileAsync(EXCEL_FILE_PATH);
}



// Define the refresh interval (in milliseconds)
const REFRESH_INTERVAL = 120 * 1000; // 60 seconds

// Set up the interval
setInterval(async () => {
  let { workbook } = await openWorksheet();
  await applyFormula(workbook);
  console.log('Data refreshed');
}, REFRESH_INTERVAL);

// Start the interval
let intervalId = setInterval(async () => {
  // Your code here...
}, REFRESH_INTERVAL);

// Stop the interval
clearInterval(intervalId);

app.get('/qr-code', async (req, res) => {
  try {
    const qr = await qrcode.toDataURL(''); // Replace 'Your QR code data' with the actual data you want to encode
    res.json({ qr });
  } catch (error) {
    console.error('Error generating QR code:', error);
    res.sendStatus(500);
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
