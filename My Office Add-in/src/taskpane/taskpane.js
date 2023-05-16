Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = sendEmail;
  }
});

async function sendEmail() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;

  // Create a new email
  var email = new Office.Mailbox.EmailAddressDetails();
  email.address = "mishaalkhan135@gmail.com";
  email.name = "Mishaal khan";

  // Create an email data object
  var emailData = new Office.CoercionData();
  emailData.coercionType = Office.CoercionType.Html;

  // Get the body of the current email
  item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.data = item.normalizedSubject + "<br/><br/>" + result.value;

      // Send the email
      Office.context.mailbox.item.displayReplyForm({
        htmlBody: emailData.data,
        ccRecipients: [email],
      });
    }
  });
}
