/* global Word, Office, CryptoJS */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Attach button handlers
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

// Serialize the content (text + tables) from the selection
async function serializeContent(selection) {
  const serialized = {
    text: selection.text || "",
    tables: []
  };

  // Load tables on the selection
  selection.tables.load("items");
  await selection.context.sync();

  // If there are tables, proceed
  const tables = selection.tables.items;
  if (tables && tables.length > 0) {
    for (const table of tables) {
      // Load rows for each table
      table.rows.load("items");
    }
    await selection.context.sync();

    // Now we have table rows
    for (const table of tables) {
      const tableData = [];
      const rows = table.rows.items;
      if (!rows) continue;

      // Load cells for each row
      for (const row of rows) {
        row.cells.load("items");
      }
      await selection.context.sync();

      for (const row of rows) {
        const cells = row.cells.items;
        if (!cells) continue;

        // Load text for each cell body
        for (const cell of cells) {
          cell.body.load("text");
        }
        await selection.context.sync();

        // After loading text, we can now read the cell text
        const rowData = [];
        for (const cell of cells) {
          rowData.push(cell.body.text);
        }
        tableData.push(rowData);
      }

      serialized.tables.push(tableData);
    }
  }

  return JSON.stringify(serialized);
}

// Deserialize JSON content (with text and tables) and insert into the selection
async function deserializeAndInsertContent(serializedString, selection) {
  const serialized = JSON.parse(serializedString);

  // Clear the current selection before inserting
  selection.insertText("", Word.InsertLocation.replace);

  // Insert text if present
  if (serialized.text) {
    selection.insertText(serialized.text, Word.InsertLocation.start);
  }

  // Insert tables if present
  if (serialized.tables && serialized.tables.length > 0) {
    // Insert a paragraph before starting to insert tables to ensure proper insertion point
    selection.insertParagraph("", Word.InsertLocation.end);

    for (const tableData of serialized.tables) {
      const rowCount = tableData.length;
      const columnCount = rowCount > 0 ? tableData[0].length : 0;

      if (rowCount > 0 && columnCount > 0) {
        selection.insertTable(rowCount, columnCount, Word.InsertLocation.end, tableData);
        // Insert a paragraph after each table to separate it from the next
        selection.insertParagraph("", Word.InsertLocation.end);
      }
    }
  }
}

// Encrypt highlighted content
async function encryptHighlightedContent() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.tables.load("items");
      await context.sync();

      if (!selection.text && selection.tables.items.length === 0) {
        console.error("Nothing to encrypt.");
        return;
      }

      const serializedContent = await serializeContent(selection);
      const encryptedContent = CryptoJS.AES.encrypt(serializedContent, key).toString();
      selection.insertText(encryptedContent, Word.InsertLocation.replace);
    });
  } catch (error) {
    console.error("Error during encryption:", error);
  }
}

// Decrypt highlighted content
async function decryptHighlightedContent() {
  try {
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
        console.error("Decryption failed. Check the key and encrypted content.");
        return;
      }

      await deserializeAndInsertContent(decryptedContent, selection);
      await context.sync();
    });
  } catch (error) {
    console.error("Error during decryption:", error);
  }
}

// Encrypt the entire document (Note: currently only encrypts text of entire doc)
async function encryptEntireDocument() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const serializedContent = JSON.stringify({ text: body.text });
      const encryptedContent = CryptoJS.AES.encrypt(serializedContent, key).toString();

      body.clear();
      body.insertText(encryptedContent, Word.InsertLocation.start);
    });
  } catch (error) {
    console.error("Error encrypting the entire document:", error);
  }
}

// Decrypt the entire document (Note: currently only decrypts text for entire doc)
async function decryptEntireDocument() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const decryptedBytes = CryptoJS.AES.decrypt(body.text, key);
      const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);

      if (!decryptedContent) {
        console.error("Decryption failed. Check the key and encrypted content.");
        return;
      }

      const deserialized = JSON.parse(decryptedContent);
      body.clear();
      body.insertText(deserialized.text, Word.InsertLocation.start);
    });
  } catch (error) {
    console.error("Error decrypting the entire document:", error);
  }
}

// Add "Hello World" paragraphs for testing
async function writeHelloWorlds() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello world! Hello wold! Hello world!", Word.InsertLocation.end);
    });
  } catch (error) {
    console.error("Error adding Hello World paragraphs:", error);
  }
}
