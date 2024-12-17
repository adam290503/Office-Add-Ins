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

/**
 * Serialize the user's current selection, including any fully selected tables.
 */
async function serializeSelection(context, selection) {
  // Load the text and table properties of the selection
  selection.load("text");
  selection.tables.load();
  await context.sync();

  const content = {
    text: selection.text || "",
    tables: []
  };

  // If there are tables in the selection, attempt to load their values
  if (selection.tables.items.length > 0) {
    for (const tbl of selection.tables.items) {
      tbl.load("values");
    }

    await context.sync();

    // Extract table data
    for (const tbl of selection.tables.items) {
      if (tbl.values) {
        content.tables.push(tbl.values);
      }
    }
  }

  return JSON.stringify(content);
}

/**
 * Deserialize a JSON string of { text, tables } and insert into the current selection.
 */
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

/**
 * Encrypt the currently highlighted content (text + tables) with the chosen key.
 */
async function encryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

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

/**
 * Decrypt the currently highlighted content with the chosen key.
 */
async function decryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

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

/**
 * Serialize the entire document content (text + tables).
 */
async function serializeEntireDocument(context) {
  const body = context.document.body;
  body.load("text");
  body.tables.load();
  await context.sync();

  const content = {
    text: body.text || "",
    tables: []
  };

  if (body.tables.items.length > 0) {
    for (const tbl of body.tables.items) {
      tbl.load("values");
    }
    await context.sync();

    for (const tbl of body.tables.items) {
      if (tbl.values) {
        content.tables.push(tbl.values);
      }
    }
  }

  return JSON.stringify(content);
}

/**
 * Deserialize entire document content and insert into document body.
 */
async function deserializeAndInsertIntoDocument(context, serializedString) {
  const content = JSON.parse(serializedString);

  const body = context.document.body;
  body.clear();

  // Insert the text
  if (content.text) {
    body.insertText(content.text, Word.InsertLocation.start);
  }

  // Insert tables
  if (content.tables && content.tables.length > 0) {
    body.insertParagraph("", Word.InsertLocation.end);
    for (const tableData of content.tables) {
      const rows = tableData.length;
      const cols = rows > 0 ? tableData[0].length : 0;
      if (rows > 0 && cols > 0) {
        body.insertTable(rows, cols, Word.InsertLocation.end, tableData);
        body.insertParagraph("", Word.InsertLocation.end);
      }
    }
  }

  await context.sync();
}

/**
 * Encrypt the entire document (text + tables).
 */
async function encryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

  await Word.run(async (context) => {
    const serialized = await serializeEntireDocument(context);
    const encrypted = CryptoJS.AES.encrypt(serialized, key).toString();

    const body = context.document.body;
    body.clear();
    body.insertText(encrypted, Word.InsertLocation.start);
    await context.sync();
  }).catch(err => console.error("Error encrypting the entire document:", err));
}

/**
 * Decrypt the entire document (text + tables).
 */
async function decryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

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

    await deserializeAndInsertIntoDocument(context, decryptedContent);
  }).catch(err => console.error("Error decrypting the entire document:", err));
}

/**
 * Insert test content: A paragraph and a sample table.
 */
async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello world! Hello worlrd ! Hello world!", Word.InsertLocation.end);

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
