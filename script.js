    Office.onReady(function (info) {
      logMessage("Office.onReady() invoked!");
      // Ensure the DOM is ready
      document.addEventListener("DOMContentLoaded", function () {
        // Call the initialization function for your add-in
        initializeAddin();
      });
    });

    function initializeAddin() {
      logMessage("initializeAddin() invoked!");
      // Add event handlers to interact with the add-in and Outlook
      document.getElementById("btnReadMessage").addEventListener("click", readMessage);
    }

    function readMessage() {
      logMessage("readMessage() invoked!");
      // Get the current item (email) from the Office API
      Office.context.mailbox.item.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var messageBody = result.value;
          logMessage(`messageBody retrieve SUCCESS: ${result}`);

          // Do something with the message body, e.g., display it in a div
          document.getElementById("messageContent").innerText = messageBody;
        } else {
          // Handle error
          console.error("Error reading message:", result.error.message);
          logMessage(`messageBody retrieve FAILED: ${result}`);
        }
      });
    }

    function logMessage(messageText) {
      document.getElementById("consoleContent").innerText += `${Date.now().toString()}: ${messageText}<hr/>`;
    }