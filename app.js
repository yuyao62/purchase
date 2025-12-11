// 模擬資料（之後可由 Excel 轉換 JSON 載入）
const inventoryData = [
  { code: "A001", name: "阿司匹林", qty: 50 },
  { code: "B002", name: "普拿疼", qty: 0 },
  { code: "C003", name: "胃藥", qty: 12 }
];

function searchDrug() {
  const keyword = document.getElementById("searchInput").value.trim();
  const outOfStockOnly = document.getElementById("outOfStockFilter").checked;
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  const results = inventoryData.filter(item => {
    const match = item.code.includes(keyword) || item.name.includes(keyword);
    const stockFilter = outOfStockOnly ? item.qty === 0 : true;
    return match && stockFilter;
  });

  results.forEach(item => {
    const row = `<tr>
      <td>${item.code}</td>
      <td>${item.name}</td>
      <td>${item.qty}</td>
      <td>${item.qty === 0 ? "缺貨" : "有庫存"}</td>
    </tr>`;
    tbody.innerHTML += row;
  });

  updateStats(results);
}

function resetSearch() {
  document.getElementById("searchInput").value = "";
  document.getElementById("outOfStockFilter").checked = false;
  searchDrug();
}

function updateStats(results) {
  const total = results.length;
  const outOfStock = results.filter(item => item.qty === 0).length;
  const rate = total > 0 ? Math.round((outOfStock / total) * 100) : 0;

  document.getElementById("totalCount").textContent = total;
  document.getElementById("outOfStockCount").textContent = outOfStock;
  document.getElementById("outOfStockRate").textContent = rate + "%";
}
