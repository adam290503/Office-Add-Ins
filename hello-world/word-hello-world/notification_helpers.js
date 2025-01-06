
// Shows an in-page notification with optional error styling
export function showNotification(message, isError = false) {
    const notificationElement = document.getElementById("notification");
    if (!notificationElement) {
      console.error("Notification element not found in the DOM.");
      return;
    }
  
    // Set the text
    notificationElement.textContent = message;
  
    if (isError) {
      // Error styling
      notificationElement.style.backgroundColor = "#f8d7da"; // light red
      notificationElement.style.color = "#721c24";           // dark red
      notificationElement.style.borderColor = "#f5c6cb";
    } else {
      // Success/info styling
      notificationElement.style.backgroundColor = "#d4edda"; // light green
      notificationElement.style.color = "#155724";           // dark green
      notificationElement.style.borderColor = "#c3e6cb";
    }
  
    notificationElement.style.display = "block";
  }
  
  // Clears/hides the notification area
  export function clearNotification() {
    const notificationElement = document.getElementById("notification");
    if (notificationElement) {
      notificationElement.textContent = "";
      notificationElement.style.display = "none";
    }
  }
  