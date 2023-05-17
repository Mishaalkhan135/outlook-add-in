Office.onReady(() => {
  const reportButton = document.querySelector("#report-button");
  reportButton.addEventListener("click", sendEmail);

  // Call the crypto API when the add-in loads
  fetchCryptoData();
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

async function fetchCryptoData() {
  const response = await fetch("https://api.coincap.io/v2/assets");
  const data = await response.json();
  const cryptoTable = document.querySelector("#crypto-table tbody");

  data.data.forEach((crypto, index) => {
    const row = document.createElement("tr");
    const rankCell = document.createElement("td");
    const nameCell = document.createElement("td");
    const priceCell = document.createElement("td");

    rankCell.textContent = index + 1;
    nameCell.textContent = crypto.name;
    priceCell.textContent = parseFloat(crypto.priceUsd).toFixed(2);

    row.appendChild(rankCell);
    row.appendChild(nameCell);
    row.appendChild(priceCell);

    cryptoTable.appendChild(row);
  });
}
