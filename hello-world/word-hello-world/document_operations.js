// documentOperations.js

import { showNotification } from "./notificationHelpers.js";
import {
  addCustomXml,
  getSpecificXmlPartContent,
  deleteSpecificXmlPart,
  getAllCustomXmlParts,
  getKeysFromXml,
} from "./xmlHelpers.js";

/********************************************************
 *               Global Variables / Keys                *
 ********************************************************/
const keys = {
  dv: "dv-secure-key",
  sc: "sc-secure-key",
  official: "official-secure-key",
};

// For demonstration, store the copied OOXML globally
let copiedOOXML = "";

/********************************************************
 *        Helper Functions for Unique Identifiers       *
 ********************************************************/
/**
 * Used ONLY for encryption/creation of a new key (entered in the text field).
 */
export function getUniqueIdentifierForEncryption() {
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
export function getUniqueIdentifierFromDropdown() {
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
export function copyContentWithOOXML() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      copiedOOXML = result.value;
    } else {
      console.error("Error retrieving OOXML:", result.error.message);
      showNotification("Error retrieving selected OOXML content.", true);
    }
  });
}

/********************************************************
 *        Encrypt/Decrypt Selected (Highlighted)        *
 ********************************************************/
export async function encryptHighlightedOOXML() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    showNotification("Please select a valid clearance level.", true);
    return;
  }

  const uniqueId = getUniqueIdentifierForEncryption();
  if (!uniqueId) return;

  Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml, async (result) => {
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
  });
}

export async function decryptHighlightedOOXML() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    showNotification("Please select a valid clearance level.", true);
    return;
  }

  const uniqueId = getUniqueIdentifierFromDropdown();
  if (!uniqueId) return;

  try {
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
export async function encryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    showNotification("Please select a valid clearance level.", true);
    return;
  }

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

export async function decryptEntireDocument() {
  const clearanceLevel = document.getElementById("clearance-level").value;
  const key = keys[clearanceLevel];

  if (!key) {
    console.error("No valid key selected.");
    showNotification("Please select a valid clearance level.", true);
    return;
  }

  const uniqueId = getUniqueIdentifierFromDropdown();
  if (!uniqueId) return;

  try {
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
export async function displayAllKeys() {
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
 *        Delete Key (Uses the dropdown selection)      *
 ********************************************************/
export async function deleteKey() {
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
