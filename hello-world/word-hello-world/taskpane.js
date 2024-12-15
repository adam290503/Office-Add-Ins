Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").onclick = encryptHighlightedText;
    document.getElementById("decryptButton").onclick = decryptHighlightedText;
    document.getElementById("writeButton").onclick = writeHelloWorlds;
    document.getElementById("protectButton").onclick = protectContent;
  }
});

// Simple shift cipher for encryption/decryption
function shiftCipher(text, shift, encrypt = true) {
  const factor = encrypt ? 1 : -1;
  return text
    .split("")
    .map((char) => {
      const code = char.charCodeAt(0);
      if (code >= 32 && code <= 126) {
        return String.fromCharCode(((code - 32 + factor * shift + 95) % 95) + 32);
      }
      return char;
    })
    .join("");
}

// Function to encrypt highlighted text
async function encryptHighlightedText() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const shift = clearanceLevel === "dv" ? 5 : clearanceLevel === "sc" ? 3 : 1;

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const encryptedText = shiftCipher(selection.text, shift, true);
    selection.insertText(encryptedText, Word.InsertLocation.replace);
    await context.sync();
  });
}

// Function to decrypt highlighted text
async function decryptHighlightedText() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const shift = clearanceLevel === "dv" ? 5 : clearanceLevel === "sc" ? 3 : 1;

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const decryptedText = shiftCipher(selection.text, shift, false);
    selection.insertText(decryptedText, Word.InsertLocation.replace);
    await context.sync();
  });
}

// Function to write a paragraph of "Hello world"
async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello world. Hello world. Hello world.", Word.InsertLocation.end);
    await context.sync();
  });
}

// Function to protect content by converting it to '1's
async function protectContent() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const originalText = body.text;
    const protectedText = originalText.replace(/./g, "1");
    body.clear();
    body.insertText(protectedText, Word.InsertLocation.start);
    await context.sync();
  });
}
