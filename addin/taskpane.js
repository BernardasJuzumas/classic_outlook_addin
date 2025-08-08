Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office.js initialized - Email ID Viewer starting...");
        
        // Initialize immediately
        updateEmailId();
        
        // Add event handler for item changes
        if (Office.context.mailbox.addHandlerAsync) {
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function(eventArgs) {
                console.log("Item changed event triggered");
                setTimeout(updateEmailId, 100); // Small delay to ensure context is ready
            });
        }
        
        // Set up periodic refresh to catch any missed changes
        setInterval(updateEmailId, 2000); // Check every 2 seconds
        
        // Also try to get the current item after short delays
        setTimeout(updateEmailId, 500);
        setTimeout(updateEmailId, 1000);
        setTimeout(updateEmailId, 2000);
    }
});

let lastEmailId = null;

function updateEmailId() {
    const emailIdElement = document.getElementById("email-id");
    
    if (!emailIdElement) {
        console.log("Email ID element not found, retrying...");
        setTimeout(updateEmailId, 100);
        return;
    }

    try {
        const item = Office.context.mailbox.item;
        
        if (item && item.itemId) {
            // Only update if the email ID has actually changed
            if (lastEmailId !== item.itemId) {
                console.log("New email detected:", item.itemId);
                lastEmailId = item.itemId;
                emailIdElement.textContent = item.itemId;
                emailIdElement.style.color = "#333";
                // Send email ID to server for logging
                sendEmailIdToServer(item.itemId);
            }
        } else if (item && !item.itemId) {
            // Item exists but no ID yet (might be in compose mode or loading)
            if (lastEmailId !== "compose") {
                console.log("Compose mode detected");
                lastEmailId = "compose";
                emailIdElement.textContent = "Email ID not available (compose mode)";
                emailIdElement.style.color = "#666";
                sendEmailIdToServer(null, "Email ID not available - compose mode");
            }
        } else {
            // No item selected
            if (lastEmailId !== null) {
                console.log("No item selected");
                lastEmailId = null;
                emailIdElement.textContent = "Please select an email to view its ID";
                emailIdElement.style.color = "#999";
                sendEmailIdToServer(null, "No email selected");
            }
        }
    } catch (error) {
        console.error("Error getting email ID:", error);
        emailIdElement.textContent = "Error retrieving email ID";
        emailIdElement.style.color = "#d32f2f";
        sendEmailIdToServer(null, "Error: " + error.message);
    }
}

// Function to send email ID to server for console logging
function sendEmailIdToServer(emailId, status = null) {
    fetch('/log-email-id', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            emailId: emailId,
            status: status,
            timestamp: new Date().toISOString()
        })
    }).catch(error => {
        // Silently handle errors to avoid disrupting the add-in functionality
        console.error('Failed to log email ID to server:', error);
    });
}