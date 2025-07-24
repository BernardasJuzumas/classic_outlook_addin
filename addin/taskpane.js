Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        updateEmailId();
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, updateEmailId);
    }
});

function updateEmailId() {
    const item = Office.context.mailbox.item;
    const emailIdElement = document.getElementById("email-id");
    if (item && item.itemId) {
        emailIdElement.textContent = item.itemId;
        // Send email ID to server for logging
        sendEmailIdToServer(item.itemId);
    } else {
        emailIdElement.textContent = "No email selected or ID not available";
        // Send status to server for logging
        sendEmailIdToServer(null, "No email selected or ID not available");
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
            status: status
        })
    }).catch(error => {
        // Silently handle errors to avoid disrupting the add-in functionality
        console.error('Failed to log email ID to server:', error);
    });
}