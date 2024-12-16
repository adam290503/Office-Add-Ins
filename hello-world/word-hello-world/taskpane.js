Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office.js is ready!");
    document.getElementById("encryptButton").onclick = () => {
      console.log("Encrypt Button Clicked");
      encryptHighlightedText();
    };
    document.getElementById("decryptButton").onclick = () => {
      console.log("Decrypt Button Clicked");
      decryptHighlightedText();
    };
    document.getElementById("writeButton").onclick = () => {
      console.log("Write Hello Worlds Button Clicked");
      writeHelloWorlds();
    };
    document.getElementById("protectButton").onclick = () => {
      console.log("Protect Content Button Clicked");
      protectContent();
    };
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
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    console.log("Selected text for encryption:", selection.text);
    const encryptedText = CryptoJS.AES.encrypt(selection.text, key).toString();
    selection.insertText(encryptedText, Word.InsertLocation.replace);
    console.log("Text encrypted and replaced.");
  });
}

// Decrypt highlighted text
async function decryptHighlightedText() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    console.log("Selected text for decryption:", selection.text);
    const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
    const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
    selection.insertText(decryptedText, Word.InsertLocation.replace);
    console.log("Text decrypted and replaced.");
  });
}

// Insert "Hello World" at the end of the document
async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(
      "Hello world! Hello world. Hello world.",
      Word.InsertLocation.end
    );
    console.log("Hello World paragraphs added.");
  });
}

// Protect content by replacing all text with '1'
async function protectContent() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const protectedText = body.text.replace(/./g, "1");
    body.clear();
    body.insertText(protectedText, Word.InsertLocation.start);
    console.log("Document protected by replacing content with '1'.");
  });
}
