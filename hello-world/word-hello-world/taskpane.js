// Ensure Office.js is initialized before any API calls
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office.js is ready!");

    // Attach button handlers
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedText);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedText);
    document.getElementById("writeButton").addEventListener("click", writeHelloWorlds);
    document.getElementById("protectButton").addEventListener("click", protectContent);

    console.log("Button event handlers attached successfully.");
  } else {
    console.error("This add-in is not running in Word.");
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
        console.warn("No text selected for encryption.");
        return;
      }

      console.log("Selected text for encryption:", selection.text);
      const encryptedText = CryptoJS.AES.encrypt(selection.text, key).toString();
      selection.insertText(encryptedText, Word.InsertLocation.replace);
      console.log("Text encrypted and replaced.");
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
        console.warn("No text selected for decryption.");
        return;
      }

      console.log("Selected text for decryption:", selection.text);
      const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);

      if (!decryptedText) {
        console.error("Decryption failed. Check the key and encrypted text.");
        return;
      }

      selection.insertText(decryptedText, Word.InsertLocation.replace);
      console.log("Text decrypted and replaced.");
    });
  } catch (error) {
    console.error("Error during decryption:", error);
  }
}

// Insert "Hello World" at the end of the document
async function writeHelloWorlds() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(
        "Hello world! Hello world. Hello world.",
        Word.InsertLocation.end
      );
      console.log("Hello World paragraphs added.");
    });
  } catch (error) {
    console.error("Error adding Hello World paragraphs:", error);
  }
}

// Protect content by replacing all text with '1'
async function protectContent() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const protectedText = body.text.replace(/./g, "1");
      body.clear();
      body.insertText(protectedText, Word.InsertLocation.start);
      console.log("Document protected by replacing content with '1'.");
    });
  } catch (error) {
    console.error("Error protecting content:", error);
  }
}
