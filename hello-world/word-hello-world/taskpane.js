Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").onclick = encryptHighlightedText;
    document.getElementById("decryptButton").onclick = decryptHighlightedText;
    document.getElementById("writeButton").onclick = writeHelloWorlds;
    document.getElementById("protectButton").onclick = protectContent;
  }
});

// Clearance level keys
const keys = {
  dv: "dv-secure-key",
  sc: "sc-secure-key",
  official: "official-secure-key",
};

// Function to encrypt highlighted text
async function encryptHighlightedText() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync(); // Ensure selected text is loaded

    const encryptedText = CryptoJS.AES.encrypt(selection.text, key).toString();
    selection.insertText(encryptedText, Word.InsertLocation.replace);
  });
}

// Function to decrypt highlighted text
async function decryptHighlightedText() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync(); // Ensure selected text is loaded

    const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
    const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
    selection.insertText(decryptedText, Word.InsertLocation.replace);
  });
}

// Function to write a paragraph of "Hello world"
async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello world! Hello world. Hello world.", Word.InsertLocation.end);
  });
}

// Function to protect content by replacing it with '1's
async function protectContent() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync(); // Necessary to load the text for replacement

    const protectedText = body.text.replace(/./g, "1");
    body.clear();
    body.insertText(protectedText, Word.InsertLocation.start);
  });
}
