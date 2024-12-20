
/* global Word, Office, CryptoJS */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
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

/**
Handle the encryption of the content
 */
async function encryptHighlightedOOXML() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        return;
    }

    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Ooxml,
        async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const ooxml = result.value;

                const hash = CryptoJS.SHA256(ooxml).toString();
                console.log("OOXML Hash:", hash);

                // Let's delete all the custom xml parts
                const deleteParts = await deleteXmlParts();

                const encrypted = CryptoJS.AES.encrypt(ooxml, key).toString();

                const abc = await addCustomXml(encrypted, "Key001");

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

/**
 * Handle the decryption of the encrypted content.
 */
async function decryptHighlightedOOXML() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        return;
    }

    try {
        // Retrieve the encrypted data using the updated getSpecificXmlPartContent function
        const EncryptedData = await getSpecificXmlPartContent("Key001");

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

async function encryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        return;
    }

    await Word.run(async (context) => {
        try {
            const body = context.document.body;
            body.load("text");
            await context.sync();

            const documentContent = body.text; // Get the entire document content
            const encryptedContent = CryptoJS.AES.encrypt(documentContent, key).toString();

            // Add the encrypted content to custom XML with the hardcoded key name "Key001"
            await addCustomXml(encryptedContent, "Key001");

            // Replace document content with the key reference
            body.clear();
            body.insertText("Key001", Word.InsertLocation.start);
            await context.sync();

            console.log("Entire document encrypted and key reference inserted.");
        } catch (error) {
            console.error("Error encrypting the entire document:", error);
        }
    });
}


async function decryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        return;
    }

    try {
        // Retrieve the encrypted content from custom XML using the hardcoded key name "Key001"
        const encryptedContent = await getSpecificXmlPartContent("Key001");

        if (!encryptedContent) {
            console.error(`Encrypted content not found for key: Key001`);
            return;
        }

        await Word.run(async (context) => {
            try {
                const decryptedBytes = CryptoJS.AES.decrypt(encryptedContent, key);
                const decryptedContent = decryptedBytes.toString(CryptoJS.enc.Utf8);

                if (!decryptedContent) {
                    console.error("Decryption failed. Check the key and content.");
                    return;
                }

                // Replace the document content with the decrypted content
                const body = context.document.body;
                body.clear();
                body.insertText(decryptedContent, Word.InsertLocation.start);
                await context.sync();

                console.log("Entire document decrypted successfully.");
            } catch (error) {
                console.error("Error decrypting the document content:", error);
            }
        });
    } catch (error) {
        console.error("Error retrieving encrypted content:", error);
    }
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
        addCustomXml("test", "sample");
        getCustomXml();
        body.insertParagraph("Fionish Property xml add", Word.InsertLocation.end);

        await context.sync();
    }).catch(err => console.error("Error adding Hello World paragraphs:", err));


}

// Property bag to store the encruypted data so the document remains workable 
function setKeyPair(encrypted, FriendlyName) {
    const keyPair = {
        "public-key": FriendlyName,
        "EncryptedBlock": encrypted
    };

    Office.context.document.properties.custom.setAsync(
        FriendlyName,
        JSON.stringify(keyPair),
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Key pair stored successfully.");
            }
            else {
                console.error("Failed to store key pair:", result.error.message);

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
async function addHiddenContentControl(encrypted, FriendlyName) {
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

/**
 * Function to add the custom Xml part to the document
 * @param {String} encryptedKeyValue encrypted content as key value.
 * @param {String} friendlyKeyName userfriendly key name.
 */
async function addCustomXml(encryptedKeyValue, friendlyKeyName) {
    const xml = `
      <Metadata xmlns="http://schemas.custom.xml">
        <Node>
          <${friendlyKeyName}>${encryptedKeyValue}</${friendlyKeyName}>
        </Node>
      </Metadata>
    `;

    return new Promise((resolve, reject) => {
        Office.context.document.customXmlParts.addAsync(xml, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log(`Custom XML added "${xml}".`);
                console.log(`Custom XML for "${friendlyKeyName}" added successfully.`);
                resolve();
            } else {
                console.error("Error adding custom XML:", result.error.message);
                reject(result.error);
            }
        });
    });
}

/**
 * Function to add the custom Xml part to the document
 * @param {String} friendlyKeyName userfriendly key name.
 */
async function getSpecificXmlPartContent(friendlyKeyName) {
    return new Promise((resolve, reject) => {
        Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.custom.xml", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const parts = result.value;

                if (parts.length === 0) {
                    console.log("No custom XML parts found.");
                    resolve(null);
                    return;
                }

                parts.forEach((part) => {
                    part.getXmlAsync((xmlResult) => {
                        if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xml = xmlResult.value;
                            console.error("XML retrieved:", xml);
                            // Parse the XML using DOMParser
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(xml, "application/xml");
                            console.error("XML DOC:", xmlDoc);

                            // Define the namespace URI
                            const namespaceURI = "http://schemas.custom.xml";

                            // Query the Key001 node using the namespace
                            const key001Node = xmlDoc.getElementsByTagNameNS(namespaceURI, friendlyKeyName)[0];

                            // Retrieve the value
                            const key001Value = key001Node ? key001Node.textContent : null;

                            console.log(`"${friendlyKeyName}" Value:`, key001Value);

                            if (key001Value) {
                                found = true;
                                resolve(key001Value);
                            }
                        } else {
                            console.error("Error retrieving XML:", xmlResult.error.message);
                            reject(xmlResult.error.message);
                        }
                    });
                });
            } else {
                console.error("Error retrieving custom XML parts:", result.error.message);
                reject(result.error);
            }
        });
    });
}

/**
 * Function to delete the custom Xml part from the document
 */
async function deleteXmlParts() {
    return new Promise((resolve, reject) => {
        Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.custom.xml", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const parts = result.value;

                if (parts.length === 0) {
                    console.log("deleteXmlParts: No custom XML parts found.");
                    resolve(null);
                    return;
                }

                let partCount = 0;
                let partsLength = parts.length;

                parts.forEach((part) => {
                    part.deleteAsync(function (eventArgs) {
                        console.log("deleteXmlParts: The XML Part has been deleted.");
                        partCount++;
                    });
                    if (partsLength == partCount) {
                        resolve('Success');
                    }
                });
            } else {
                console.error("deleteXmlParts: Error retrieving custom XML parts:", result.error.message);
                reject(result.error);
            }
        });
    });
}
