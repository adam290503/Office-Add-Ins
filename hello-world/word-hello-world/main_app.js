// app.js

import { 
    showNotification, 
    clearNotification 
  } from "./notificationHelpers.js";
  
  import {
    copyContentWithOOXML,
    encryptHighlightedOOXML,
    decryptHighlightedOOXML,
    encryptEntireDocument,
    decryptEntireDocument,
    displayAllKeys,
    deleteKey
  } from "./documentOperations.js";
  
  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      // Attach event listeners to your buttons
      document.getElementById("protectButton").addEventListener("click", encryptEntireDocument);
      document.getElementById("unprotectButton").addEventListener("click", decryptEntireDocument);
      document.getElementById("encryptOOXMLButton").addEventListener("click", encryptHighlightedOOXML);
      document.getElementById("decryptOOXMLButton").addEventListener("click", decryptHighlightedOOXML);
      document.getElementById("displayKeysButton").addEventListener("click", displayAllKeys);
      document.getElementById("deleteKeysButton").addEventListener("click", deleteKey);
  
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => copyContentWithOOXML()
      );
  
      displayAllKeys();
  
      // Copy content once on load
      copyContentWithOOXML();
    }
  });
  