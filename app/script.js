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
      // API DOCS: https://learn.microsoft.com/en-us/javascript/api/outlook/office.message?view=outlook-js-preview
        readEmail: function() {
          // API DOCS: https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread?view=outlook-js-preview
          // dev.logString("readEmail() invoked!");

          const emailContext = Office.context.mailbox.item;
          const emailContentDiv = document.getElementById("emailContent");
          
          var htmlStr = "<table>";
          htmlStr += `<tr><td><b>to</b></td><td>${emailContext.to}</td></tr>`;
          htmlStr += `<tr><td><b>cc</b></td><td>${emailContext.cc}</td></tr>`;
          htmlStr += `<tr><td><b>from</b></td><td>${emailContext.from}</td></tr>`;
          htmlStr += `<tr><td><b>sender</b></td><td>${emailContext.sender}</td></tr>`;
          htmlStr += `<tr><td><b>subject</b></td><td>${emailContext.subject} </td></tr>`;
          htmlStr += `<tr><td><b>body</b></td><td>PLACEHOLDER_BODY</td></tr>`;
          htmlStr += `<tr><td><b>categories</b></td><td>${emailContext.categories}</td></tr>`;
          htmlStr += `<tr><td><b>notificationMessages</b></td><td>${emailContext.notificationMessages}</td></tr>`;
          // htmlStr += `<tr><td><b>recurrence</b></td><td>${emailContext.recurrence}</td></tr>`;
          // htmlStr += `<tr><td><b>start</b></td><td>${emailContext.start}</td></tr>`;
          // htmlStr += `<tr><td><b>seriesId</b></td><td>${emailContext.seriesId}</td></tr>`;
          // htmlStr += `<tr><td><b>attachments</b></td><td>${emailContext.attachments}</td></tr>`;
          // htmlStr += `<tr><td><b>conversationId</b></td><td>${emailContext.conversationId}</td></tr>`;
          // htmlStr += `<tr><td><b>dateTimeCreated</b></td><td>${emailContext.dateTimeCreated}</td></tr>`;
          // htmlStr += `<tr><td><b>dateTimeModified</b></td><td>${emailContext.dateTimeModified}</td></tr>`;
          // htmlStr += `<tr><td><b>display</b></td><td>${emailContext.display}</td></tr>`;
          // htmlStr += `<tr><td><b>end</b></td><td>${emailContext.end}</td></tr>`;
          // htmlStr += `<tr><td><b>internetMessageId</b></td><td>${emailContext.internetMessageId}</td></tr>`;
          // htmlStr += `<tr><td><b>itemClass</b></td><td>${emailContext.itemClass}</td></tr>`;
          // htmlStr += `<tr><td><b>itemId</b></td><td>${emailContext.itemId}</td></tr>`;
          // htmlStr += `<tr><td><b>itemType</b></td><td>${emailContext.itemType}</td></tr>`;
          // htmlStr += `<tr><td><b>location</b></td><td>${emailContext.location}</td></tr>`;
          // htmlStr += `<tr><td><b>normalizedSubject</b></td><td>${emailContext.normalizedSubject}</td></tr>`;
          htmlStr += "</table>";

          emailContentDiv.innerHTML = htmlStr;

          Office.context.mailbox.item.body.getAsync("text", {},
            function(result) {
              // dev.logString(`body message extracted: ${JSON.stringify(result)}`);

              const emailContentDiv = document.getElementById("emailContent");
              emailContentDiv.innerHTML = emailContentDiv.innerHTML.replace("PLACEHOLDER_BODY", result.value);
            });
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
        },

        setEmailTags: function() {
          dev.logString("setEmailTags() invoked!");
        }
      };