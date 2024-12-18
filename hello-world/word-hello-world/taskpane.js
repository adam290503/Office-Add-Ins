/* global Word, Office, CryptoJS */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedContent);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedContent);
    document.getElementById("writeButton").addEventListener("click", writeHelloWorlds);
    document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
    document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);
    document.getElementById("printOOXMLButton").addEventListener("click", printHighlightedOOXML);
    document.getElementById("encryptOOXMLButton").addEventListener("click", encryptHighlightedOOXML);
    document.getElementById("decryptOOXMLButton").addEventListener("click", decryptHighlightedOOXML);

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => copyContentWithOOXML()
    );

    copyContentWithOOXML();
  }
});

const keys = {
  dv: "dv-secure-key",
  sc: "sc-secure-key",
  official: "official-secure-key",
};

let copiedOOXML = "";

function copyContentWithOOXML() {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Ooxml,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        copiedOOXML = result.value;
        console.log("Copied OOXML:", copiedOOXML);
      } else {
        console.error("Error retrieving OOXML:", result.error.message);
      }
    }
  );
}

async function printHighlightedOOXML() {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Ooxml,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Highlighted OOXML:", result.value);
      } else {
        console.error("Error retrieving highlighted OOXML:", result.error.message);
      }
    }
  );
}

async function encryptHighlightedOOXML() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    return;
  }

  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Ooxml,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const ooxml = result.value;
        const encrypted = CryptoJS.AES.encrypt(ooxml, key).toString();

        Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertText(encrypted, Word.InsertLocation.replace);
          await context.sync();
        }).catch(err => console.error("Error inserting encrypted OOXML:", err));
      } else {
        console.error("Error retrieving OOXML for encryption:", result.error.message);
      }
    }
  );
}

async function decryptHighlightedOOXML() {
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

    try {
      const decryptedBytes = CryptoJS.AES.decrypt(selection.text, key);
      const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

      if (!decryptedOOXML) {
        console.error("Decryption failed. Check the key and content.");
        return;
      }

      selection.insertOoxml(decryptedOOXML, Word.InsertLocation.replace);
      await context.sync();
    } catch (err) {
      console.error("Error decrypting OOXML:", err);
    }
  });
}

async function writeHelloWorlds() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello Hello world!", Word.InsertLocation.end);

    const tableValues = [
      ["Name", "Age"],
      ["Alice", "30"],
      ["Bob", "25"]
    ];
    body.insertTable(tableValues.length, tableValues[0].length, Word.InsertLocation.end, tableValues);
    await context.sync();
  }).catch(err => console.error("Error adding Hello World paragraphs:", err));
}
