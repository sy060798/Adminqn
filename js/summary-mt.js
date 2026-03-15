document.addEventListener("DOMContentLoaded", function() {

let data = {
    totalUser: 120,
    totalDeposit: 50000,
    totalProfit: 12000
}

document.getElementById("data-mt").innerHTML =
`
Total User: ${data.totalUser} <br>
Total Deposit: ${data.totalDeposit} <br>
Total Profit: ${data.totalProfit}
`

})
