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

function getUniqueIdentifier() {
    const uniqueIdInput = document.getElementById("unique-id");
    const uniqueId = uniqueIdInput.value.trim();
    if (!uniqueId) {
        console.error("Unique Identifier is required.");
        alert("Please enter a unique identifier for encryption.");
        return null;
    }
    
    return uniqueId;
}

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
 * Handle the encryption of the content
 */
async function encryptHighlightedOOXML() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        alert("Please select a valid clearance level.");
        return;
    }

    const uniqueId = getUniqueIdentifier();
    if (!uniqueId) return;

    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Ooxml,
        async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const ooxml = result.value;

                const hash = CryptoJS.SHA256(ooxml).toString();
                console.log("OOXML Hash:", hash);

                // Delete existing custom XML part with the same uniqueId if it exists
                await deleteSpecificXmlPart(uniqueId);

                // Encrypt the OOXML using the key
                const encrypted = CryptoJS.AES.encrypt(ooxml, key).toString();

                // Add the encrypted content as a custom XML part with the uniqueId
                await addCustomXml(encrypted, uniqueId);

                // Insert the uniqueId into the document
                Word.run(async (context) => {
                    const selection = context.document.getSelection();
                    selection.insertText(uniqueId, Word.InsertLocation.replace);
                    await context.sync();
                }).catch(err => console.error("Error inserting encrypted OOXML:", err));

                console.log(`Content encrypted with key "${uniqueId}" successfully.`);
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
        alert("Please select a valid clearance level.");
        return;
    }

    const uniqueId = getUniqueIdentifier();
    if (!uniqueId) return;

    try {
        // Retrieve the encrypted data using the uniqueId
        const encryptedData = await getSpecificXmlPartContent(uniqueId);

        if (!encryptedData) {
            console.error(`Encrypted data not found for the key "${uniqueId}".`);
            alert(`No encrypted data found for the key "${uniqueId}".`);
            return;
        }

        await Word.run(async (context) => {
            try {
                console.log("Encrypted Data: ", encryptedData);
                console.log("Decryption Key: ", key);

                // Decrypt the data
                const decryptedBytes = CryptoJS.AES.decrypt(encryptedData, key);
                const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

                if (!decryptedOOXML) {
                    console.error("Decryption failed. Check the key and content.");
                    alert("Decryption failed. Please verify your key and try again.");
                    return;
                }

                // Verify the integrity of the decrypted content
                const hash = CryptoJS.SHA256(decryptedOOXML).toString();
                console.log("Decrypted OOXML Hash: ", hash);

                // Insert the decrypted content back into the Word document
                const selection = context.document.getSelection();
                selection.insertOoxml(decryptedOOXML, Word.InsertLocation.replace);
                await context.sync();

                console.log(`Content decrypted with key "${uniqueId}" successfully.`);
                alert(`Content decrypted successfully with key "${uniqueId}".`);
            } catch (err) {
                console.error("Error decrypting OOXML:", err);
                alert("An error occurred during decryption.");
            }
        });
    } catch (error) {
        console.error("Error retrieving encrypted data:", error);
        alert("Failed to retrieve encrypted data.");
    }
}

async function encryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        alert("Please select a valid clearance level.");
        return;
    }

    const uniqueId = getUniqueIdentifier();
    if (!uniqueId) return;

    Word.run(async (context) => {
        const body = context.document.body;
        const ooxml = body.getOoxml(); // Retrieve the entire document as OOXML
        await context.sync();

        try {
            const hash = CryptoJS.SHA256(ooxml.value).toString();
            console.log("OOXML Hash:", hash);

            // Delete existing custom XML part with the same uniqueId if it exists
            await deleteSpecificXmlPart(uniqueId);

            // Encrypt the OOXML using the key
            const encrypted = CryptoJS.AES.encrypt(ooxml.value, key).toString();

            // Add the encrypted content as a custom XML part with the uniqueId
            await addCustomXml(encrypted, uniqueId);

            // Clear the document and insert the uniqueId as a reference
            body.clear();
            body.insertText(uniqueId, Word.InsertLocation.start);
            await context.sync();

            console.log(`Entire document encrypted with key "${uniqueId}" successfully.`);
            alert(`Entire document encrypted successfully with key "${uniqueId}".`);
        } catch (error) {
            console.error("Error encrypting the document:", error);
            alert("An error occurred during encryption.");
        }
    }).catch((err) => {
        console.error("Error accessing the document:", err);
        alert("Failed to access the document for encryption.");
    });
}

async function decryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
        console.error("No valid key selected.");
        alert("Please select a valid clearance level.");
        return;
    }

    const uniqueId = getUniqueIdentifier();
    if (!uniqueId) return;

    try {
        // Retrieve the encrypted data from the custom XML part using the uniqueId
        const encryptedContent = await getSpecificXmlPartContent(uniqueId);

        if (!encryptedContent) {
            console.error(`Encrypted data not found for the key "${uniqueId}".`);
            alert(`No encrypted data found for the key "${uniqueId}".`);
            return;
        }

        Word.run(async (context) => {
            try {
                console.log("Encrypted Data:", encryptedContent);
                console.log("Decryption Key:", key);

                // Decrypt the encrypted OOXML
                const decryptedBytes = CryptoJS.AES.decrypt(encryptedContent, key);
                const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

                if (!decryptedOOXML) {
                    console.error("Decryption failed. Check the key and content.");
                    alert("Decryption failed. Please verify your key and try again.");
                    return;
                }

                // Verify the integrity of the decrypted content
                const hash = CryptoJS.SHA256(decryptedOOXML).toString();
                console.log("Decrypted OOXML Hash:", hash);

                // Replace the entire document with the decrypted OOXML
                const body = context.document.body;
                body.clear();
                body.insertOoxml(decryptedOOXML, Word.InsertLocation.start);
                await context.sync();

                console.log(`Entire document decrypted with key "${uniqueId}" successfully.`);
                alert(`Entire document decrypted successfully with key "${uniqueId}".`);
            } catch (error) {
                console.error("Error decrypting the document:", error);
                alert("An error occurred during decryption.");
            }
        }).catch((err) => {
            console.error("Error accessing the document:", err);
            alert("Failed to access the document for decryption.");
        });
    } catch (error) {
        console.error("Error retrieving encrypted content:", error);
        alert("Failed to retrieve encrypted content.");
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
        getSpecificXmlPartContent("sample");
        body.insertParagraph("Finish Property xml add", Word.InsertLocation.end);

        await context.sync();
    }).catch(err => console.error("Error adding Hello World paragraphs:", err));
}

/**
 * Function to add the custom Xml part to the document
 * @param {String} encryptedKeyValue - Encrypted content as key value.
 * @param {String} friendlyKeyName - User-friendly unique key name.
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
                console.log(`Custom XML added for key "${friendlyKeyName}".`);
                resolve();
            } else {
                console.error("Error adding custom XML:", result.error.message);
                reject(result.error);
            }
        });
    });
}

/**
 * Retrieves the encrypted content from a specific custom XML part identified by the unique key.
 * @param {String} friendlyKeyName - User-friendly unique key name.
 * @returns {Promise<String|null>} - The encrypted content or null if not found.
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

                let found = false;
                let encryptedValue = null;

                const checkPart = (index) => {
                    if (index >= parts.length) {
                        resolve(encryptedValue);
                        return;
                    }

                    const part = parts[index];
                    part.getXmlAsync((xmlResult) => {
                        if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
                            const xml = xmlResult.value;
                            // Parse the XML using DOMParser
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(xml, "application/xml");

                            // Define the namespace URI
                            const namespaceURI = "http://schemas.custom.xml";

                            // Query the node with the unique key
                            const keyNode = xmlDoc.getElementsByTagNameNS(namespaceURI, friendlyKeyName)[0];

                            if (keyNode) {
                                encryptedValue = keyNode.textContent;
                                found = true;
                                resolve(encryptedValue);
                            } else {
                                // Continue searching the next part
                                checkPart(index + 1);
                            }
                        } else {
                            console.error("Error retrieving XML:", xmlResult.error.message);
                            reject(xmlResult.error.message);
                        }
                    });
                };

                checkPart(0);
            } else {
                console.error("Error retrieving custom XML parts:", result.error.message);
                reject(result.error);
            }
        });
    });
}

/**
 * Deletes a specific custom XML part identified by the unique key.
 * @param {String} friendlyKeyName - User-friendly unique key name.
 * @returns {Promise<void>}
 */
async function deleteSpecificXmlPart(friendlyKeyName) {
    return new Promise((resolve, reject) => {
        Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.custom.xml", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const parts = result.value;

                if (parts.length === 0) {
                    console.log("No custom XML parts found to delete.");
                    resolve();
                    return;
                }

                let deletePromises = [];

                parts.forEach((part) => {
                    deletePromises.push(new Promise((res, rej) => {
                        part.getXmlAsync((xmlResult) => {
                            if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
                                const xml = xmlResult.value;
                                const parser = new DOMParser();
                                const xmlDoc = parser.parseFromString(xml, "application/xml");
                                const namespaceURI = "http://schemas.custom.xml";
                                const keyNode = xmlDoc.getElementsByTagNameNS(namespaceURI, friendlyKeyName)[0];

                                if (keyNode) {
                                    part.deleteAsync((deleteResult) => {
                                        if (deleteResult.status === Office.AsyncResultStatus.Succeeded) {
                                            console.log(`Deleted custom XML part for key "${friendlyKeyName}".`);
                                            res();
                                        } else {
                                            console.error(`Error deleting custom XML part for key "${friendlyKeyName}":`, deleteResult.error.message);
                                            rej(deleteResult.error);
                                        }
                                    });
                                } else {
                                    res(); // This part does not contain the key; move to next
                                }
                            } else {
                                console.error("Error retrieving XML for deletion:", xmlResult.error.message);
                                rej(xmlResult.error);
                            }
                        });
                    }));
                });

                Promise.all(deletePromises)
                    .then(() => resolve())
                    .catch(err => reject(err));
            } else {
                console.error("Error retrieving custom XML parts for deletion:", result.error.message);
                reject(result.error);
            }
        });
    });
}
