// Ensure Office.js is initialized before any API calls
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Attach button handlers
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedText);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedText);
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

// Encrypt highlighted text
async function encryptHighlightedText() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text) {
        return;
      }

      const encryptedText = CryptoJS.AES.encrypt(selection.text, key).toString();
      selection.insertText(encryptedText, Word.InsertLocation.replace);
    });
  } catch (error) {
    console.error("Error during encryption:", error);
  }
}

// Decrypt highlighted text
async function decryptHighlightedText() {
  try {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text) {
        return;
      }

      const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);

      if (!decryptedText) {
        console.error("Decryption failed. Check the key and encrypted text.");
        return;
      }

      selection.insertText(decryptedText, Word.InsertLocation.replace);
    });
  } catch (error) {
    console.error("Error during decryption:", error);
  }
}

// Encrypt the entire document based on the selected role
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

// Decrypt the entire document based on the selected role
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

// Insert "Hello World" at the end of the document
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
