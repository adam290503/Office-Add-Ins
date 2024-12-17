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

// Serialize the content
async function serializeContent(selection) {
  const serialized = {
    text: selection.text || "",
    tables: [],
  };

  if (selection.tables && selection.tables.items.length > 0) {
    selection.tables.load("items");
    await selection.context.sync();

    for (const table of selection.tables.items) {
      table.rows.load("items");
      await selection.context.sync();

      const serializedTable = [];
      for (const row of table.rows.items) {
        row.cells.load("items/body/text");
        await selection.context.sync();

        const serializedRow = row.cells.items.map((cell) => cell.body.text);
        serializedTable.push(serializedRow);
      }
      serialized.tables.push(serializedTable);
    }
  }

  return JSON.stringify(serialized);
}

// Deserialize and insert content
async function deserializeAndInsertContent(serializedString, selection) {
  const serialized = JSON.parse(serializedString);

  // Clear the current selection
  selection.insertText("", Word.InsertLocation.replace);

  // Insert text if any
  if (serialized.text) {
    selection.insertText(serialized.text, Word.InsertLocation.start);
  }

  // Reconstruct tables if any
  if (serialized.tables && serialized.tables.length > 0) {
    for (const tableData of serialized.tables) {
      const rowCount = tableData.length;
      const columnCount = rowCount > 0 ? tableData[0].length : 0;

      if (rowCount > 0 && columnCount > 0) {
        const newTable = selection.insertTable(rowCount, columnCount, Word.InsertLocation.end, tableData);
        await newTable.context.sync();
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
      selection.load("text, tables/items");
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
    });
  } catch (error) {
    console.error("Error during decryption:", error);
  }
}

// Add "Hello World" paragraphs
async function writeHelloWorlds() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello world! Hello world Hello world!", Word.InsertLocation.end);
    });
  } catch (error) {
    console.error("Error adding Hello World paragraphs:", error);
  }
}
