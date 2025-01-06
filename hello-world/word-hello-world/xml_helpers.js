
import { showNotification } from "./notificationHelpers.js";

/**
 * Retrieve all custom XML parts for a given namespace.
 */
export function getAllCustomXmlParts(namespace) {
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

/**
 * Adds a custom XML part with a friendly key name and the encrypted value.
 */
export async function addCustomXml(encryptedKeyValue, friendlyKeyName) {
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
 * Retrieves the encrypted content (value) for a specific key name from custom XML parts.
 */
export async function getSpecificXmlPartContent(friendlyKeyName) {
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
            // If we've gone through all parts, resolve with whatever we found (could be null)
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

/**
 * Deletes the custom XML part(s) containing the specified key name.
 */
export async function deleteSpecificXmlPart(friendlyKeyName) {
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

/**
 * Extracts all child tag names under <Node> (in the given namespace) from the provided XML string.
 */
export function getKeysFromXml(xml, namespace) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xml, "application/xml");
  const keys = [];

  // Select all <Node> elements within the specified namespace
  const nodes = xmlDoc.getElementsByTagNameNS(namespace, "Node");

  for (let node of nodes) {
    // Iterate through child elements of <Node>
    for (let child of node.children) {
      keys.push(child.tagName);
    }
  }

  return keys;
}
