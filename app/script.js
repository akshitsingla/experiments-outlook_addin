    // Ensure the DOM is ready
    document.addEventListener("DOMContentLoaded", function () {
      // dev.logString("DOMContentLoaded triggered!");
      Office.onReady(function (info) {
        // dev.logString("Office.onReady() invoked!");
        
        // Check for scopes
        // dev.logString(`Object Office: ${Office}`);
        // dev.logString(`Object Office.context: ${Office.context}`);
        // dev.logString(`Object Office.context.mailbox.item: ${Object.keys(Office.context.mailbox.item)}`);
        // dev.logString(`Object Office.context.mailbox.item.itemType: ${Office.context.mailbox.item.itemType}`);
        
        initializeAddin();
      });
    });

    function initializeAddin() {
      // dev.logString("initializeAddin() invoked!");

      // Check for dev mode
      if (true) {   
        var devElements = document.querySelectorAll(".dev")
        for (devElement of devElements) {
          devElement.style.display = "block";
        }
      }
      
      document.getElementById("btnReadEmail").addEventListener("click", emailOps.readEmail);

      dev.logString("initializeAddin() invoked!");
    }

    const dev = {
      logString: function(logText) {
        document.getElementById("consoleContent").innerHTML += `${Date.now().toString()}: ${logText} <hr/>`;
      }  
    };

    const emailOps = {
      // API DOCS: https://learn.microsoft.com/en-us/javascript/api/outlook/office.item?view=outlook-js-preview
        readEmail: function() {
          // API DOCS: https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread?view=outlook-js-preview
          dev.logString("readEmail() invoked!");
        },

        composeEmail: function() {
          // API DOCS: https://learn.microsoft.com/en-us/javascript/api/outlook/office.messagecompose?view=outlook-js-preview
          dev.logString("composeEmail() invoked!");
        },

        moveEmail: function() {
          dev.logString("moveEmail() invoked!");
        },

        sendEmail: function() {
          dev.logString("sendEmail() invoked!");
        },

        setEmailFlag: function() {
          dev.logString("setEmailFlag() invoked!");
        },

        setEmailPriority: function() {
          dev.logString("setEmailPriority() invoked!");
        }
      };