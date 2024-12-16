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

// Encrypt highlighted content (text and tables)
async function encryptHighlightedContent() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.load("tables/items");
      await context.sync();

      // Encrypt text if present
      if (selection.text) {
        const encryptedText = CryptoJS.AES.encrypt(selection.text, key).toString();
        selection.insertText(encryptedText, Word.InsertLocation.replace);
      }

      // Encrypt each table in the selection
      const tables = selection.tables.items;
      if (tables.length > 0) {
        for (const table of tables) {
          table.load("rows/items/cells/items");
          await context.sync();

          for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
              cell.load("body/text");
              await context.sync();

              if (cell.body.text) {
                const encryptedText = CryptoJS.AES.encrypt(cell.body.text, key).toString();
                cell.body.clear();
                cell.body.insertText(encryptedText, Word.InsertLocation.start);
              }
            }
          }
        }
      }
    });
  } catch (error) {
    console.error("Error during content encryption:", error);
  }
}

// Decrypt highlighted content (text and tables)
async function decryptHighlightedContent() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.load("tables/items");
      await context.sync();

      // Decrypt text if present
      if (selection.text) {
        const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
        const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);

        if (!decryptedText) {
          console.error("Decryption failed. Check the key and encrypted text.");
          return;
        }

        selection.insertText(decryptedText, Word.InsertLocation.replace);
      }

      // Decrypt each table in the selection
      const tables = selection.tables.items;
      if (tables.length > 0) {
        for (const table of tables) {
          table.load("rows/items/cells/items");
          await context.sync();

          for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
              cell.load("body/text");
              await context.sync();

              if (cell.body.text) {
                const decryptedBytes = CryptoJS.AES.decrypt(cell.body.text, key);
                const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);

                if (!decryptedText) {
                  console.error("Decryption failed. Check the key and encrypted text.");
                  continue;
                }

                cell.body.clear();
                cell.body.insertText(decryptedText, Word.InsertLocation.start);
              }
            }
          }
        }
      }
    });
  } catch (error) {
    console.error("Error during content decryption:", error);
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

      if (!body.text) {
        return;
      }

      const encryptedText = CryptoJS.AES.encrypt(body.text, key).toString();
      body.clear();
      body.insertText(encryptedText, Word.InsertLocation.start);
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

      if (!body.text) {
        return;
      }

      const decryptedBytes = CryptoJS.AES.decrypt(body.text, key);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);

      if (!decryptedText) {
        console.error("Decryption failed. Check the key and encrypted text.");
        return;
      }

      body.clear();
      body.insertText(decryptedText, Word.InsertLocation.start);
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
