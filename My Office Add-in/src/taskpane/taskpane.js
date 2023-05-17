// taskpane.js

Office.onReady(() => {
  const reportButton = document.querySelector("#report-button");
  reportButton.addEventListener("click", sendEmail);
});

async function sendEmail() {
  const item = Office.context.mailbox.item;

  let from;
  if (item && item.from && item.from.emailAddress) {
    from = item.from.emailAddress;
  } else {
    console.error("Could not access the email sender's details");
    return;
  }

  const subject = `Report: ${item.normalizedSubject}`;
  const body = `This email was reported:\n\nSubject: ${item.normalizedSubject}\n\nFrom: ${from}`;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["mishaalkhan135@gmail.com"],
    ccRecipients: [],
    subject: subject,
    body: body,
  });
}
