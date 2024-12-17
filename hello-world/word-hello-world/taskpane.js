/* global Word, Office, CryptoJS */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedContent);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedContent);
    document.getElementById("writeButton").addEventListener("click", writeHelloWorlds);
    document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
    document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);

    // Add a handler to update OOXML on selection change
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => copyContentWithOOXML()
    );

    // Call copyContentWithOOXML initially
    copyContentWithOOXML();
  }
});

const keys = {
  dv: "dv-secure-key",
  sc: "sc-secure-key",
  official: "official-secure-key",
};

let copiedOOXML = ""; // Declare global variable for OOXML content

function copyContentWithOOXML() {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Ooxml,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        copiedOOXML = result.value;
        showNotification("Copied", "Content copied successfully with formatting.");
        console.log("OOXML TEST:", copiedOOXML); // Log the updated content
      } else {
        showNotification("Error", result.error.message);
        console.error("Error retrieving OOXML:", result.error.message);
      }
    }
  );
}

function showNotification(title, message) {
  const notification = document.getElementById("notification");
  if (notification) {
    notification.innerText = `${title}: ${message}`;
  } else {
    console.log(`${title}: ${message}`);
  }
}

async function encryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const serialized = await serializeSelection(context, selection);

    if (!serialized) {
      console.error("Nothing to encrypt.");
      return;
    }

    const encrypted = CryptoJS.AES.encrypt(serialized, key).toString();
    selection.insertText(encrypted, Word.InsertLocation.replace);
    await context.sync();
  }).catch(err => console.error("Error during encryption:", err));
}

async function decryptHighlightedContent() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

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
      console.error("Decryption failed. Check the key and content.");
      return;
    }

    await deserializeAndInsert(context, selection, decryptedContent);
  }).catch(err => console.error("Error during decryption:", err));
}

async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello world! Hello world!", Word.InsertLocation.end);

    const tableValues = [
      ["Name", "Age"],
      ["Alice", "30"],
      ["Bob", "25"]
    ];
    body.insertTable(tableValues.length, tableValues[0].length, Word.InsertLocation.end, tableValues);
    await context.sync();
  }).catch(err => console.error("Error adding Hello World paragraphs:", err));
}

async function encryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

  await Word.run(async (context) => {
    const serialized = await serializeEntireDocument(context);
    const encrypted = CryptoJS.AES.encrypt(serialized, key).toString();

    const body = context.document.body;
    body.clear();
    body.insertText(encrypted, Word.InsertLocation.start);
    await context.sync();
  }).catch(err => console.error("Error encrypting the entire document:", err));
}

async function decryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const decryptedBytes = CryptoJS.AES.decrypt(body.text, key);
    const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);

    if (!decryptedContent) {
      console.error("Decryption failed. Check the key and content.");
      return;
    }

    await deserializeAndInsertIntoDocument(context, decryptedContent);
  }).catch(err => console.error("Error decrypting the entire document:", err));
}
