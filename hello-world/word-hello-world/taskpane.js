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
  selection.load("text");
  selection.tables.load();
  await context.sync();

  const content = {
    text: selection.text || "",
    tables: []
  };

  // Load values for each table
  for (const tbl of selection.tables.items) {
    tbl.load("values");
  }

  await context.sync();

  // Extract tables
  for (const tbl of selection.tables.items) {
    content.tables.push(tbl.values);
  }

  return JSON.stringify(content);
}

async function deserializeAndInsert(context, selection, serializedString) {
  const content = JSON.parse(serializedString);

  // Clear the current selection
  selection.insertText("", Word.InsertLocation.replace);

  // Insert the text first
  if (content.text) {
    selection.insertText(content.text, Word.InsertLocation.start);
  }

  // Insert the tables if any
  if (content.tables && content.tables.length > 0) {
    // Move selection to end of inserted text
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
    body.insertParagraph("Hello world! Hello wlrd ! Hello world!", Word.InsertLocation.end);

    // Insert a sample table for testing
    const tableValues = [
      ["Name", "Age"],
      ["Alice", "30"],
      ["Bob", "25"]
    ];
    body.insertTable(tableValues.length, tableValues[0].length, Word.InsertLocation.end, tableValues);
    await context.sync();
  }).catch(err => console.error("Error adding Hello World paragraphs:", err));
}
