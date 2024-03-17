// functions.js

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Assign event handler to the button click event
    document.getElementById("HighlightRepliesButton").onclick = () => {
      highlightReplies();
    };
  }
});

async function highlightReplies() {
  // Get the mailbox
  const mailbox = Office.context.mailbox;

  // Get the inbox folder
  const inbox = mailbox.folders.getFolderByName(Office.MailboxEnums.FolderName.Inbox);

  // Get all items in the inbox
  const items = await inbox.getEntities();

  // Loop through each item and highlight replies without "Reply All"
  items.forEach(async (item) => {
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Check if it's a reply
      const isReply = item.internetMessageHeaders.some(header => header.name === "In-Reply-To");

      // Check if it's a "Reply All"
      const isReplyAll = item.to && item.to.length > 1;

      // Check if the original email was not just sent to the same person
      const originalEmail = await getOriginalEmail(item);
      const originalEmailRecipients = originalEmail.to;

      const isOriginalSentToMultiple = originalEmailRecipients.length > 1;

      if (isReply && !isReplyAll && isOriginalSentToMultiple) {
        // Highlight the email (replace this with your actual implementation)
        await highlightEmail(item);
      }
    }
  });
}

async function getOriginalEmail(replyEmail) {
  // Use your logic to retrieve the original email based on headers or other information
  // In this example, we assume the original email is in the same folder
  const originalItemId = replyEmail.internetMessageHeaders.find(header => header.name === "In-Reply-To")?.value;
  const originalEmail = await Office.context.mailbox.item.getEntitiesByEntityId(originalItemId);

  return originalEmail[0];
}
async function highlightEmail(emailItem) {
  // Open the email in a pop-up window
  await Office.context.mailbox.displayMessageForm(
    {
      items: [emailItem.itemId],
    },
    {
      options: {
        allowEventPropagation: true,
      },
    }
  );
}
