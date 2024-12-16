Office.initialize = () => {
  console.log("Office.js initialized via fallback.");
  initializeAddIn();
};

Office.onReady(() => {
  console.log("Office.js initialized via onReady.");
  initializeAddIn();
});

function initializeAddIn() {
  if (Office.context.host === Office.HostType.Word) {
    document.getElementById("encryptButton").onclick = encryptHighlightedText;
    document.getElementById("decryptButton").onclick = decryptHighlightedText;
    document.getElementById("writeButton").onclick = writeHelloWorlds;
    document.getElementById("protectButton").onclick = protectContent;
  }
}

// Encryption keys for clearance levels
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

      const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
      selection.insertText(decryptedText, Word.InsertLocation.replace);
    });
  } catch (error) {
    console.error("Error during decryption:", error);
  }
}
