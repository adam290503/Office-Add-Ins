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

  const tables = selection.tables.items;
  if (tables.length > 0) {
    selection.tables.load("items");
    await selection.context.sync();

    for (const table of tables) {
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

  if (serialized.text) {
    selection.insertText(serialized.text, Word.InsertLocation.replace);
  }

  if (serialized.tables.length > 0) {
    for (const tableData of serialized.tables) {
      const newTable = selection.insertTable(
        tableData.length,
        tableData[0].length,
        Word.InsertLocation.replace,
        tableData
      );
      await newTable.context.sync();
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

// Encrypt the entire document
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

// Decrypt the entire document
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

// Add "Hello World" paragraphs
async function writeHelloWorlds() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello world! Hello world. Hello world.", Word.InsertLocation.end);
    });
  } catch (error) {
    console.error("Error adding Hello World paragraphs:", error);
  }
}
