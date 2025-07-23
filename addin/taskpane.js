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
    } else {
        emailIdElement.textContent = "No email selected or ID not available";
    }
}