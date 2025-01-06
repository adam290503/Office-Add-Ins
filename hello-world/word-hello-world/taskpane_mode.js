
function showNotification(message, isError = false) {
    const notificationElement = document.getElementById("notification");
    if (!notificationElement) {
      console.error("Notification element not found in the DOM.");
      return;
    }
  
    // Set text
    notificationElement.textContent = message;
  
    //notifications will be differently styled for errors
    if (isError) {
      notificationElement.style.backgroundColor = "#f8d7da"; // light red
      notificationElement.style.color = "#721c24";           // dark red
      notificationElement.style.borderColor = "#f5c6cb";
    } else {
      notificationElement.style.backgroundColor = "#d4edda"; // light green
      notificationElement.style.color = "#155724";           // dark green
      notificationElement.style.borderColor = "#c3e6cb";
    }
  
    notificationElement.style.display = "block";
  }
  
  function clearNotification() {
    const notificationElement = document.getElementById("notification");
    if (notificationElement) {
      notificationElement.textContent = "";
      notificationElement.style.display = "none";
    }
  }
  

  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
      document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);
      document.getElementById("encryptOOXMLButton").addEventListener("click", encryptHighlightedOOXML);
      document.getElementById("decryptOOXMLButton").addEventListener("click", decryptHighlightedOOXML);
      document.getElementById("displayKeysButton").addEventListener("click", displayAllKeys);
      document.getElementById("deleteKeysButton").addEventListener("click", deleteKey);
      document.getElementById("unlockAllButton").addEventListener("click", decryptAllKeys);


      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => copyContentWithOOXML()
      );

      // Optional: On load, populate the keys dropdown:
      displayAllKeys();

      copyContentWithOOXML();
    }
  });

  /********************************************************
   *               Global Variables / Keys                *
   ********************************************************/
  const keys = {
    dv: "dv-secure-key",
    sc: "sc-secure-key",
    official: "official-secure-key",
  };

  let copiedOOXML = "";

  /********************************************************
   *        Helper Functions for Unique Identifiers       *
   ********************************************************/
  /**
   * Used ONLY for encryption/creation of a new key (entered in the text field).
   */
  function getUniqueIdentifierForEncryption() {
    const uniqueIdInput = document.getElementById("unique-id");
    const uniqueId = uniqueIdInput.value.trim();
    if (!uniqueId) {
      console.error("Unique Identifier is required.");
      showNotification("Please enter a unique identifier for encryption.", true);
      return null;
    }
    return uniqueId;
  }

  /**
   * Used for decryption, deletion, etc. 
   * Reads from the dropdown of existing keys.
   */
  function getUniqueIdentifierFromDropdown() {
    const dropdown = document.getElementById("keysDropdown");
    const selectedKey = dropdown.value.trim();
    if (!selectedKey) {
      console.error("No key selected from dropdown.");
      showNotification("Please select a valid key from the dropdown.", true);
      return null;
    }
    return selectedKey;
  }

  /********************************************************
   *           Copy Selected Content as OOXML             *
   ********************************************************/
  function copyContentWithOOXML() {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Ooxml,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          copiedOOXML = result.value;
        } else {
          console.error("Error retrieving OOXML:", result.error.message);
          showNotification("Error retrieving selected OOXML content.", true);
        }
      }
    );
  }

  /********************************************************
   *        Encrypt/Decrypt Selected (Highlighted)        *
   ********************************************************/
  async function encryptHighlightedOOXML() {
    // The user must select the clearance level (e.g. dv, sc, official).
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
      console.error("No valid key selected.");
      showNotification("Please select a valid clearance level.", true);
      return;
    }

    // For encryption, we read from the text field (creating a new unique ID).
    const uniqueId = getUniqueIdentifierForEncryption();
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
          }).catch((err) => console.error("Error inserting encrypted OOXML:", err));

          console.log(`Content encrypted with key "${uniqueId}" successfully.`);
          showNotification(`Content encrypted successfully with key "${uniqueId}".`);

          // Refresh the dropdown to include this new key
          await displayAllKeys();
        } else {
          console.error("Error retrieving OOXML for encryption:", result.error.message);
          showNotification("Failed to retrieve selected content for encryption.", true);
        }
      }
    );
  }

  async function decryptHighlightedOOXML() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
      console.error("No valid key selected.");
      showNotification("Please select a valid clearance level.", true);
      return;
    }

    // For decryption, read the key (uniqueId) from the dropdown.
    const uniqueId = getUniqueIdentifierFromDropdown();
    if (!uniqueId) return;

    try {
      // Retrieve the encrypted data using the uniqueId
      const encryptedData = await getSpecificXmlPartContent(uniqueId);

      if (!encryptedData) {
        console.error(`Encrypted data not found for the key "${uniqueId}".`);
        showNotification(`No encrypted data found for key "${uniqueId}".`, true);
        return;
      }

      await Word.run(async (context) => {
        try {
          console.log("Encrypted Data: ", encryptedData);

          // Decrypt the data
          const decryptedBytes = CryptoJS.AES.decrypt(encryptedData, key);
          const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

          if (!decryptedOOXML) {
            console.error("Decryption failed. Check the key and content.");
            showNotification("Decryption failed. Please verify your key and try again.", true);
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
          showNotification(`Content decrypted successfully with key "${uniqueId}".`);
        } catch (err) {
          console.error("Error decrypting OOXML:", err);
          showNotification("An error occurred during decryption.", true);
        }
      });
    } catch (error) {
      console.error("Error retrieving encrypted data:", error);
      showNotification("Failed to retrieve encrypted data.", true);
    }
  }

  /********************************************************
   *            Encrypt/Decrypt Entire Document           *
   ********************************************************/
  async function encryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
      console.error("No valid key selected.");
      showNotification("Please select a valid clearance level.", true);
      return;
    }

    // For encryption, we read from the text field (creating a new unique ID).
    const uniqueId = getUniqueIdentifierForEncryption();
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
        showNotification(`Entire document encrypted successfully with key "${uniqueId}".`);

        // Refresh keys so that the new one shows up in the dropdown
        await displayAllKeys();
      } catch (error) {
        console.error("Error encrypting the document:", error);
        showNotification("An error occurred during encryption.", true);
      }
    }).catch((err) => {
      console.error("Error accessing the document:", err);
      showNotification("Failed to access the document for encryption.", true);
    });
  }

  async function decryptEntireDocument() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];

    if (!key) {
      console.error("No valid key selected.");
      showNotification("Please select a valid clearance level.", true);
      return;
    }

    // For decryption, read from the dropdown.
    const uniqueId = getUniqueIdentifierFromDropdown();
    if (!uniqueId) return;

    try {
      // Retrieve the encrypted data from the custom XML part using the uniqueId
      const encryptedContent = await getSpecificXmlPartContent(uniqueId);

      if (!encryptedContent) {
        console.error(`Encrypted data not found for the key "${uniqueId}".`);
        showNotification(`No encrypted data found for key "${uniqueId}".`, true);
        return;
      }

      Word.run(async (context) => {
        try {
          console.log("Encrypted Data:", encryptedContent);

          // Decrypt the encrypted OOXML
          const decryptedBytes = CryptoJS.AES.decrypt(encryptedContent, key);
          const decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);

          if (!decryptedOOXML) {
            console.error("Decryption failed. Check the key and content.");
            showNotification("Decryption failed. Please verify your key and try again.", true);
            return;
          }

          const hash = CryptoJS.SHA256(decryptedOOXML).toString();
          console.log("Decrypted OOXML Hash:", hash);

          // Replace the entire document with the decrypted OOXML
          const body = context.document.body;
          body.clear();
          body.insertOoxml(decryptedOOXML, Word.InsertLocation.start);
          await context.sync();

          console.log(`Entire document decrypted with key "${uniqueId}" successfully.`);
          showNotification(`Entire document decrypted successfully with key "${uniqueId}".`);
        } catch (error) {
          console.error("Error decrypting the document:", error);
          showNotification("An error occurred during decryption.", true);
        }
      }).catch((err) => {
        console.error("Error accessing the document:", err);
        showNotification("Failed to access the document for decryption.", true);
      });
    } catch (error) {
      console.error("Error retrieving encrypted content:", error);
      showNotification("Failed to retrieve encrypted content.", true);
    }
  }

  /********************************************************
   *        Display All Keys (populate the dropdown)      *
   ********************************************************/
  async function displayAllKeys() {
    const namespace = "http://schemas.custom.xml";

    try {
      // Retrieve all custom XML parts with the specified namespace
      const customXmlParts = await getAllCustomXmlParts(namespace);

      if (customXmlParts.length === 0) {
        showNotification("No keys found in the document.");
        // Also clear the dropdown
        const dropdown = document.getElementById("keysDropdown");
        dropdown.innerHTML = '<option value="">Select a key</option>';
        return;
      }

      let allKeys = [];

      // Iterate through each custom XML part to extract keys
      for (let part of customXmlParts) {
        try {
          // Retrieve the XML content of the custom XML part
          const xml = await new Promise((resolve, reject) => {
            part.getXmlAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
              } else {
                reject(result.error.message);
              }
            });
          });

          // Parse XML and extract keys
          const keysFound = getKeysFromXml(xml, namespace);
          allKeys = allKeys.concat(keysFound);
        } catch (err) {
          console.error("Error retrieving XML from a custom XML part:", err);
        }
      }

      if (allKeys.length === 0) {
        showNotification("No keys found in the document.");
        // Clear the dropdown
        const dropdown = document.getElementById("keysDropdown");
        dropdown.innerHTML = '<option value="">Select a key</option>';
        return;
      }

      // Remove duplicate keys, if any
      const uniqueKeys = [...new Set(allKeys)];

      // Populate the dropdown
      const dropdown = document.getElementById("keysDropdown");
      dropdown.innerHTML = '<option value="">Select a key</option>'; // Reset dropdown

      uniqueKeys.forEach((key) => {
        const option = document.createElement("option");
        option.value = key;
        option.textContent = key;
        dropdown.appendChild(option);
      });

      console.log("Dropdown has been populated with keys successfully.");
      showNotification("Keys loaded successfully.");
    } catch (error) {
      console.error("Error in displayAllKeys:", error);
      showNotification("An error occurred while retrieving keys.", true);
    }
  }

  /********************************************************
   *        Adding / Retrieving / Deleting Custom XML     *
   ********************************************************/
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

          let encryptedValue = null;

          const checkPart = (index) => {
            if (index >= parts.length) {
              // If we’ve gone through all parts, resolve with whatever we found (could be null)
              resolve(encryptedValue);
              return;
            }

            const part = parts[index];
            part.getXmlAsync((xmlResult) => {
              if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
                const xml = xmlResult.value;
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(xml, "application/xml");

                // Define the namespace URI
                const namespaceURI = "http://schemas.custom.xml";
                // Query the node with the unique key
                const keyNode = xmlDoc.getElementsByTagNameNS(namespaceURI, friendlyKeyName)[0];

                if (keyNode) {
                  encryptedValue = keyNode.textContent;
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
            deletePromises.push(
              new Promise((res, rej) => {
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
                          console.error(
                            `Error deleting custom XML part for key "${friendlyKeyName}":`,
                            deleteResult.error.message
                          );
                          rej(deleteResult.error);
                        }
                      });
                    } else {
                      // Not found in this part, just resolve
                      res();
                    }
                  } else {
                    console.error("Error retrieving XML for deletion:", xmlResult.error.message);
                    rej(xmlResult.error);
                  }
                });
              })
            );
          });

          Promise.all(deletePromises)
            .then(() => resolve())
            .catch((err) => reject(err));
        } else {
          console.error("Error retrieving custom XML parts for deletion:", result.error.message);
          reject(result.error);
        }
      });
    });
  }

  function getAllCustomXmlParts(namespace) {
    return new Promise((resolve, reject) => {
      Office.context.document.customXmlParts.getByNamespaceAsync(namespace, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error.message);
        }
      });
    });
  }

  function getKeysFromXml(xml, namespace) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xml, "application/xml");
    const keys = [];

    // Select all Node elements within the specified namespace
    const nodes = xmlDoc.getElementsByTagNameNS(namespace, "Node");

    for (let node of nodes) {
      // Iterate through child elements of <Node>
      for (let child of node.children) {
        keys.push(child.tagName);
      }
    }

    return keys;
  }

  /********************************************************
   *        Delete Key (Uses the dropdown selection)      *
   ********************************************************/
  async function deleteKey() {
    // For deletion, read from the dropdown.
    const uniqueId = getUniqueIdentifierFromDropdown();
    if (!uniqueId) return;

    try {
      await deleteSpecificXmlPart(uniqueId);

      console.log(`Key "${uniqueId}" and its value have been deleted successfully.`);
      showNotification(`Key "${uniqueId}" and its associated value have been deleted successfully.`);

      // Refresh the keys dropdown so that the deleted one is removed
      await displayAllKeys();
    } catch (error) {
      console.error(`Error deleting the key "${uniqueId}":`, error);
      showNotification(`An error occurred while trying to delete the key "${uniqueId}". Please try again.`, true);
    }
  }


  async function decryptAllKeys() {
    const clearanceLevel = document.getElementById("clearance-level").value;
    const key = keys[clearanceLevel];
  
    if (!key) {
      console.error("No valid key selected.");
      showNotification("Please select a valid clearance level.", true);
      return;
    }
  
    // Get all custom XML parts in the "http://schemas.custom.xml" namespace
    let allParts;
    try {
      allParts = await getAllCustomXmlParts("http://schemas.custom.xml");
      if (!allParts || allParts.length === 0) {
        showNotification("No encrypted data found in the document.");
        return;
      }
    } catch (error) {
      console.error("Error retrieving custom XML parts:", error);
      showNotification("Failed to retrieve encrypted data.", true);
      return;
    }
  
    // We'll track how many items we successfully decrypted
    let decryptedCount = 0;
  
    // Iterate over each part to see if we can decrypt it with the selected key
    for (const part of allParts) {
      try {
        // Get the raw XML from this customXmlPart
        const xml = await new Promise((resolve, reject) => {
          part.getXmlAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject(result.error);
            }
          });
        });
  
        // Example: <Metadata xmlns="http://schemas.custom.xml">
        //             <Node>
        //               <Key001>ENCRYPTED_STRING</Key001>
        //             </Node>
        //           </Metadata>
  
        // Parse out each child (the "uniqueId") from <Node>
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xml, "application/xml");
  
        // Each <Node> can have multiple <uniqueId> children
        const nodes = xmlDoc.getElementsByTagNameNS("http://schemas.custom.xml", "Node");
  
        for (let node of nodes) {
          for (let child of node.children) {
            const friendlyKeyName = child.tagName; // e.g. "Key001"
            const encryptedData = child.textContent; // e.g. "U2FsdGVkX1+..."
  
            if (!encryptedData) {
              continue; // no content, skip
            }
  
            // Attempt decryption
            let decryptedOOXML = "";
            try {
              const decryptedBytes = CryptoJS.AES.decrypt(encryptedData, key);
              decryptedOOXML = decryptedBytes.toString(CryptoJS.enc.Utf8);
            } catch (err) {
              // This means invalid key, skip
              console.log(`Skipping key "${friendlyKeyName}" due to decryption error.`);
              continue;
            }
  
            if (!decryptedOOXML) {
              // If it decrypts to empty, skip
              console.log(`Skipping key "${friendlyKeyName}" because the content didn't decrypt properly.`);
              continue;
            }
  
            // If we get here, we have valid OOXML
            await Word.run(async (context) => {
              const body = context.document.body;
              // Replace entire doc content with this newly-decrypted OOXML
              // If your doc uses smaller placeholders, or you only want partial replacement,
              // you might need a more specialized search-and-replace approach.
              body.clear();
              body.insertOoxml(decryptedOOXML, Word.InsertLocation.start);
              await context.sync();
            });
  
            decryptedCount++;
            console.log(`Successfully decrypted key "${friendlyKeyName}".`);
          }
        }
      } catch (err) {
        console.log("Skipping a part due to an error: ", err);
      }
    }
  
    if (decryptedCount > 0) {
      showNotification(`Successfully decrypted ${decryptedCount} item(s).`);
    } else {
      showNotification("No items were decrypted. Either they are encrypted with a different key or no data found.", true);
    }
  }
  