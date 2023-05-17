Office.onReady(() => {
  const reportButton = document.querySelector("#report-button");
  reportButton.addEventListener("click", sendEmail);
  populateTable();
});

function sendEmail() {
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["mishaalkhan135@gmail.com"],
    ccRecipients: [],
    subject: "Report",
    body:
      "This email was reported:\n\nSubject: " +
      Office.context.mailbox.item.subject +
      "\n\nFrom: " +
      Office.context.mailbox.item.from.emailAddress,
  });
}

async function populateTable() {
  const response = await fetch("https://api.coincap.io/v2/assets");
  const data = await response.json();

  const tableBody = document.querySelector("#crypto-table tbody");

  for (const crypto of data.data) {
    const row = document.createElement("tr");

    const rankCell = document.createElement("td");
    rankCell.textContent = crypto.rank;
    row.appendChild(rankCell);

    const nameCell = document.createElement("td");
    nameCell.textContent = crypto.name;
    row.appendChild(nameCell);

    const priceCell = document.createElement("td");
    priceCell.textContent = Number(crypto.priceUsd).toFixed(2);
    row.appendChild(priceCell);

    // Add more cells for other fields here

    const changePercent24HrCell = document.createElement("td");
    changePercent24HrCell.textContent = Number(crypto.changePercent24Hr).toFixed(2) + "%";
    row.appendChild(changePercent24HrCell);

    const marketCapCell = document.createElement("td");
    marketCapCell.textContent = Number(crypto.marketCapUsd).toFixed(2);
    row.appendChild(marketCapCell);

    const volumeCell = document.createElement("td");
    volumeCell.textContent = Number(crypto.volumeUsd24Hr).toFixed(2);
    row.appendChild(volumeCell);

    const supplyCell = document.createElement("td");
    supplyCell.textContent = Number(crypto.supply).toFixed(2);
    row.appendChild(supplyCell);

    const maxSupplyCell = document.createElement("td");
    maxSupplyCell.textContent = crypto.maxSupply ? Number(crypto.maxSupply).toFixed(2) : "N/A";
    row.appendChild(maxSupplyCell);

    const circulatingSupplyCell = document.createElement("td");
    circulatingSupplyCell.textContent = Number(crypto.circulatingSupply).toFixed(2);
    row.appendChild(circulatingSupplyCell);

    tableBody.appendChild(row);
  }
}
