/* global Word, Office, CryptoJS */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("encryptButton").addEventListener("click", encryptHighlightedContent);
    document.getElementById("decryptButton").addEventListener("click", decryptHighlightedContent);
    document.getElementById("writeButton").addEventListener("click", writeHelloWorlds);
    document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
    document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);
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
        console.log("OOXML copied:", copiedOOXML);
      } else {
        console.error("Error retrieving OOXML:", result.error.message);
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

        const hash = CryptoJS.SHA256(ooxml).toString();
        console.log("OOXML Hash:", hash);

        const encrypted = CryptoJS.AES.encrypt(ooxml, key).toString();
        
        const abc = addCustomXml(encrypted,"Key001");
        console.log(encrypted)
        console.log("ABC =",abc)
        
        Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertText("Key001", Word.InsertLocation.replace);
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

  try {
    // Wait for the encrypted data to be retrieved
    const EncryptedData = await getSpecificXmlNode("Key001");
    console.log("Encrypted Data: ", EncryptedData);
    console.log("Decryption Key: ", key);

    if (!EncryptedData) {
      console.error("Encrypted data not found for the given key.");
      return;
    }

    await Word.run(async (context) => {
      try {
        console.log("Encrypted Data: ", EncryptedData);
        console.log("Decryption Key: ", key);

        // Decrypt the data
        const decryptedBytes = CryptoJS.AES.decrypt(EncryptedData, key);
        const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

       

        if (!decryptedOOXML) {
          console.error("Decryption failed. Check the key and content.");
          return;
        }

        // Verify the integrity of the decrypted content
        const hash = CryptoJS.SHA256(decryptedOOXML).toString();
        console.log("Decrypted OOXML Hash: ", hash);

        // Insert the decrypted content back into the Word document
        const selection = context.document.getSelection();
        selection.insertOoxml(decryptedOOXML, Word.InsertLocation.replace);
        await context.sync();
      } catch (err) {
        console.error("Error decrypting OOXML:", err);
      }
    });
  } catch (error) {
    console.error("Error retrieving encrypted data:", error);
  }
}


async function serializeSelection(context, selection) {
  selection.load("text");
  selection.tables.load();
  await context.sync();

  const content = {
    text: selection.text || "",
    tables: []
  };

  if (selection.tables.items.length > 0) {
    for (const tbl of selection.tables.items) {
      tbl.load("values");
    }

    await context.sync();

    for (const tbl of selection.tables.items) {
      if (tbl.values) {
        content.tables.push(tbl.values);
      }
    }
  }

  return JSON.stringify(content);
}

async function deserializeAndInsert(context, selection, serializedString) {
  const content = JSON.parse(serializedString);

  selection.insertText("", Word.InsertLocation.replace);

  if (content.text) {
    selection.insertText(content.text, Word.InsertLocation.start);
  }

  if (content.tables && content.tables.length > 0) {
    selection.insertParagraph("", Word.InsertLocation.end);
    for (const tableData of content.tables) {
      const rows = tableData.length;
      const cols = rows > 0 ? tableData[0].length : 0;
      if (rows > 0 && cols > 0) {
        selection.insertTable(rows, cols, Word.InsertLocation.end, tableData);
        selection.insertParagraph("", Word.InsertLocation.end);
      }
    }
  }

  await context.sync();
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

     const keypair = getKeyPair()

    const decryptedBytes = CryptoJS.AES.decrypt(keypair, key);
    const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);
    if (!decryptedContent) {
      console.error("Decryption failed. Check the key and content.");
      return;
    }

    await deserializeAndInsert(context, selection, decryptedContent);
  }).catch(err => console.error("Error during decryption:", err));
}

async function serializeEntireDocument(context) {
  const body = context.document.body;
  body.load("text");
  body.tables.load();
  body.paragraphs.load("items");
  await context.sync();

  const content = {
    text: "",
    tables: []
  };

  for (const p of body.paragraphs.items) {
    p.load("parentTableOrNullObject");
  }
  await context.sync();

  const textParagraphs = [];
  for (const p of body.paragraphs.items) {
    if (p.parentTableOrNullObject.isNullObject) {
      textParagraphs.push(p.text);
    }
  }
  content.text = textParagraphs.join("\n");

  if (body.tables.items.length > 0) {
    for (const tbl of body.tables.items) {
      tbl.load("values");
    }
    await context.sync();

    for (const tbl of body.tables.items) {
      if (tbl.values) {
        content.tables.push(tbl.values);
      }
    }
  }

  return JSON.stringify(content);
}

async function deserializeAndInsertIntoDocument(context, serializedString) {
  const content = JSON.parse(serializedString);
  const body = context.document.body;
  body.clear();

  if (content.text) {
    const paragraphs = content.text.split("\n");
    for (let i = 0; i < paragraphs.length; i++) {
      body.insertParagraph(paragraphs[i], Word.InsertLocation.end);
    }
  }

  if (content.tables && content.tables.length > 0) {
    body.insertParagraph("", Word.InsertLocation.end);
    for (const tableData of content.tables) {
      const rows = tableData.length;
      const cols = rows > 0 ? tableData[0].length : 0;
      if (rows > 0 && cols > 0) {
        body.insertTable(rows, cols, Word.InsertLocation.end, tableData);
        body.insertParagraph("", Word.InsertLocation.end);
      }
    }
  }

  await context.sync();
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

    body.insertParagraph("Start Property xml add", Word.InsertLocation.end);
    addCustomXml("test","sample");
    body.insertParagraph("Fionish Property xml add", Word.InsertLocation.end);

    const EncryptedSample = await getSpecificXmlNode("sample");
    body.insertParagraph(EncryptedSample, Word.InsertLocation.end);



    await context.sync();
  }).catch(err => console.error("Error adding Hello World paragraphs:", err));

  
}






// Property bag to store the encruypted data so the document remains workable 



function setKeyPair(encrypted,FriendlyName) {
   const keyPair = { 
    "public-key": FriendlyName,
    "EncryptedBlock": encrypted
   };

  Office.context.document.properties.custom.setAsync(
    FriendlyName, 
    JSON.stringify(keyPair), 
    (result) => { 
        if (result.status=== Office.AsyncResultStatus.Succeeded) {
              console.log("Key pair stored successfully."); 
        } 
        else { console.error("Failed to store key pair:", result.error.message); 

        }
      } 
    ); 
  }
 

function getKeyPair(FriendlyName) {
Office.context.document.properties.custom.getAsync(
  FriendlyName,
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const keyPair = JSON.parse(result.value);
      console.log("Retrieved Key Pair:", keyPair);
      console.log("EncryptedBlock");
      console.log(keyPair.EncryptedBlock);
      return keyPair.EncryptedBlock;
    } else {
      console.error("Failed to retrieve key pair:", result.error.message);
    }
  }
);
}

// We will use this if the proerpty bag does not work - AG 
async function addHiddenContentControl(encrypted,FriendlyName) {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const contentControl = range.insertContentControl();
    contentControl.title = FriendlyName;
    contentControl.tag = FriendlyName;
    contentControl.insertText(
      encrypted,
      Word.InsertLocation.replace
    );
    contentControl.appearance = "none";
    contentControl.font.hidden = true;
    await context.sync();
    console.log("Hidden content control added.");
  });
}

async function getHiddenContentControlValue(FriendlyName) {
  await Word.run(async (context) => {
    // Get all content controls
    const contentControls = context.document.contentControls;
    contentControls.load("items/tag,title,text");

    await context.sync();

    // Find the content control by tag
    const hiddenControl = contentControls.items.find(
      (control) => control.tag === FriendlyName
    );

    if (hiddenControl) {
      console.log("Hidden Content Control Value:", hiddenControl.text);
    } else {
      console.log("No content control with the specified tag found.");
    }

    // Logging for debugging 

    
      const body = context.document.body;
      body.insertParagraph("Start", Word.InsertLocation.end);
  
      body.insertParagraph(hiddenControl.title, Word.InsertLocation.end);
      body.insertParagraph(hiddenControl.appearance, Word.InsertLocation.end);
      body.insertParagraph(hiddenControl.font, Word.InsertLocation.end);
      body.insertParagraph(hiddenControl.tag, Word.InsertLocation.end);
      body.insertParagraph("End", Word.InsertLocation.end);
      
      await context.sync();
    
  });
}

async function addCustomXml(encrypted, FriendlyName) {
  const xml = `
    <Metadata xmlns="http://schemas.custom.xml">
      <Node>
        <Key>${FriendlyName}</Key>
        <Value>${encrypted}</Value>
      </Node>
    </Metadata>
  `;

  return new Promise((resolve, reject) => {
    Office.context.document.customXmlParts.addAsync(xml, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`Custom XML for "${FriendlyName}" added successfully.`);
        
        // Retrieve the added XML part for verification
        result.value.getXmlAsync((xmlResult) => {
          if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Stored Custom XML:", xmlResult.value);
          } else {
            console.error("Error retrieving stored XML:", xmlResult.error.message);
          }
        });
        resolve();
      } else {
        console.error("Error adding custom XML:", result.error.message);
        reject(result.error);
      }
    });
  });
}





async function getSpecificXmlNode(FriendlyName) {
  return new Promise((resolve, reject) => {
    Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.custom.xml", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const parts = result.value;
        console.log(`Found ${parts.length} custom XML part(s).`);

        if (parts.length === 0) {
          console.log("No custom XML parts found.");
          resolve(null);
          return;

        }

        let found = false;
        parts.forEach((part) => {
          part.getXmlAsync((xmlResult) => {
            if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
              const xml = xmlResult.value;
              console.log("Retrieved XML Part:", xml);

              // Parse the XML using DOMParser
              const parser = new DOMParser();
              const xmlDoc = parser.parseFromString(xml, "application/xml");

              // Use XPath to find the specific node by key
              const xpath = `/Metadata/Node[Key='${FriendlyName}']/Value`;
              console.log("XPath Query:", xpath);

              const node = xmlDoc.evaluate(
                xpath,
                xmlDoc,
                null,
                XPathResult.FIRST_ORDERED_NODE_TYPE,
                null
              );

              if (node.singleNodeValue) {
                console.log(`Value for Key "${FriendlyName}":`, node.singleNodeValue.textContent);
                found = true;
                resolve(node.singleNodeValue.textContent);
              } else {
                console.error(`Key "${FriendlyName}" not found in XML.`);
              }
            } else {
              console.error("Error retrieving XML:", xmlResult.error.message);
            }
          });
        });

        if (!found) {
          console.log(`Key "${FriendlyName}" not found.`);
          resolve(null);
        }
      } else {
        console.error("Error retrieving custom XML parts:", result.error.message);
        reject(result.error);
      }
    });
  });
}











