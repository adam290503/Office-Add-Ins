/* global Word, Office, CryptoJS */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedContent);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedContent);
    document.getElementById("writeButton").addEventListener("click", writeHelloWorlds);
    document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
    document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);
  }
});

const keys = {
  dv: "dv-secure-key",
  sc: "sc-secure-key",
  official: "official-secure-key",
};

async function serializeSelection(context, selection) {
  // Load the selection text and tables
  selection.load("text");
  selection.tables.load("items");
  await context.sync();

  const content = { text: selection.text || "", tables: [] };

  const tables = selection.tables.items;
  if (!tables || tables.length === 0) {
    return JSON.stringify(content);
  }

  // Load rows for all tables
  for (const tbl of tables) {
    tbl.rows.load("items");
  }
  await context.sync();

  for (const tbl of tables) {
    const tableData = [];
    const rows = tbl.rows.items || [];
    // Load cells for each row
    for (const row of rows) {
      row.cells.load("items");
    }
    await context.sync();

    // Now load text for each cell
    for (const row of rows) {
      const cells = row.cells.items || [];
      for (const cell of cells) {
        cell.body.load("text");
      }
    }
    await context.sync();

    // Extract the text now that it's loaded
    for (const row of rows) {
      const rowCells = row.cells.items || [];
      const rowData = rowCells.map(cell => cell.body.text || "");
      tableData.push(rowData);
    }

    content.tables.push(tableData);
  }

  return JSON.stringify(content);
}

async function deserializeAndInsert(context, selection, serializedString) {
  const content = JSON.parse(serializedString);

  // Clear selection
  selection.insertText("", Word.InsertLocation.replace);

  // Insert text first
  if (content.text) {
    selection.insertText(content.text, Word.InsertLocation.start);
  }

  // Insert tables if any
  if (content.tables && content.tables.length > 0) {
    selection.insertParagraph("", Word.InsertLocation.end);
    for (const tableData of content.tables) {
      const rows = tableData.length;
      const cols = rows > 0 ? tableData[0].length : 0;
      if (rows > 0 && cols > 0) {
        selection.insertTable(rows, cols, Word.InsertLocation.end, tableData);
        selection.insertParagraph("", Word.InsertLocation.end);
      }
    }
  }

  await context.sync();
}

async function encryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();

    // Serialize the current selection into JSON
    const serialized = await serializeSelection(context, selection);

    // If nothing selected, do nothing
    if (!serialized) {
      console.error("Nothing to encrypt.");
      return;
    }

    // Encrypt
    const encrypted = CryptoJS.AES.encrypt(serialized, key).toString();

    // Replace selection with encrypted text
    selection.insertText(encrypted, Word.InsertLocation.replace);
    await context.sync();
  }).catch(err => console.error("Error during encryption:", err));
}

async function decryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    if (!selection.text) {
      console.error("Nothing to decrypt.");
      return;
    }

    const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
    const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);
    if (!decryptedContent) {
      console.error("Decryption failed. Check the key and content.");
      return;
    }

    await deserializeAndInsert(context, selection, decryptedContent);
  }).catch(err => console.error("Error during decryption:", err));
}

// Encrypt the entire document (only text as an example)
async function encryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const serialized = JSON.stringify({ text: body.text });
    const encrypted = CryptoJS.AES.encrypt(serialized, key).toString();

    body.clear();
    body.insertText(encrypted, Word.InsertLocation.start);
    await context.sync();
  }).catch(err => console.error("Error encrypting the entire document:", err));
}

// Decrypt the entire document (only text as an example)
async function decryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const decryptedBytes = CryptoJS.AES.decrypt(body.text, key);
    const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);

    if (!decryptedContent) {
      console.error("Decryption failed. Check the key and content.");
      return;
    }

    const deserialized = JSON.parse(decryptedContent);
    body.clear();
    body.insertText(deserialized.text || "", Word.InsertLocation.start);
    await context.sync();
  }).catch(err => console.error("Error decrypting the entire document:", err));
}

// Insert test content
async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello world! Hello ! Hello world!", Word.InsertLocation.end);
    await context.sync();
  }).catch(err => console.error("Error adding Hello World paragraphs:", err));
}
