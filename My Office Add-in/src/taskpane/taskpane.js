//=============================================================================================
// function is an event handler that runs when the Office.js library is ready to interact with the document.
//We select the report button element from the HTML using its id (#report-button).
//We add a click event listener to the report button so that when it is clicked, the sendEmail function will be called.
//We call the populateTable function to fetch cryptocurrency data and populate it in a table.
//==============================================================================================
Office.onReady(() => {
  const reportButton = document.querySelector("#report-button");

  if (!reportButton.hasAttribute("listener")) {
    reportButton.addEventListener("click", sendEmail);
    reportButton.setAttribute("listener", "true");
  }

  populateTable();
});

function sendEmail() {
  const emailSubject = Office.context.mailbox.item.subject;
  const emailFrom = Office.context.mailbox.item.from.emailAddress;
  const emailBody = Office.context.mailbox.item.body.getAsync("text", function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log(result.error.message);
    } else {
      fetch("https://giqxzti3x1.execute-api.us-east-1.amazonaws.com/email", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: "mishaalkhan135@gmail.com",
          subject: emailSubject,
          from: emailFrom,
          body: result.value,
        }),
      })
        .then((response) => response.json())
        .then((data) => console.log(data))
        .catch((error) => {
          console.error("Error:", error);
        });
    }
  });
}

//==============================================================================================
//The populateTable function is responsible for fetching cryptocurrency data from the CoinCap API and populating it in a table.
//==============================================================================================
async function populateTable() {
  const response = await fetch("https://api.coincap.io/v2/assets");
  const data = await response.json();

  const tableBody = document.querySelector("#crypto-table tbody");

  for (const crypto of data.data) {
    const row = document.createElement("tr");
    //==============================================
    // Adding all cells for fields
    //==============================================
    const rankCell = document.createElement("td");
    rankCell.textContent = crypto.rank;
    row.appendChild(rankCell);

    const nameCell = document.createElement("td");
    nameCell.textContent = crypto.name;
    row.appendChild(nameCell);

    const priceCell = document.createElement("td");
    priceCell.textContent = Number(crypto.priceUsd).toFixed(2);
    row.appendChild(priceCell);

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
