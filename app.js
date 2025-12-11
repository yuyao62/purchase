async function loadData() {
  const res = await fetch("data/meds.json");
  const data = await res.json();

  // 廠商前兩碼合併
  const vendors = {};
  data.forEach(item => {
    const code = item.廠商.slice(0,2);
    if (!vendors[code]) vendors[code] = [];
    vendors[code].push(item);
  });

  // 顯示廠商清單
  const vendorList = document.getElementById("vendors");
  Object.keys(vendors).forEach(code => {
    const btn = document.createElement("button");
    btn.textContent = code;
    btn.onclick = () => showVendor(code, vendors[code]);
    vendorList.appendChild(btn);
  });
}

function showVendor(code, items) {
  document.getElementById("vendor-title").textContent = `📋 廠商代碼 ${code} 的藥品清單`;
  const table = document.getElementById("result");
  table.innerHTML = "";

  items.forEach(item => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${item.藥品}</td>
      <td>${item.累計用量}</td>
      <td><input type="number" min="0" data-usage="${item.累計用量}"></td>
      <td class="status"></td>
    `;
    table.appendChild(row);
  });

  // 即時判斷
  table.querySelectorAll("input").forEach(input => {
    input.addEventListener("input", e => {
      const usage = parseFloat(e.target.dataset.usage);
      const stock = parseFloat(e.target.value || 0);
      const statusCell = e.target.parentElement.nextElementSibling;
      if (stock < usage) {
        statusCell.textContent = "需採購";
        statusCell.className = "status-need";
      } else {
        statusCell.textContent = "OK";
        statusCell.className = "status-ok";
      }
    });
  });
}

// 匯出 CSV
function downloadCSV() {
  const rows = [["藥品名稱","累計用量","庫存","狀態"]];
  document.querySelectorAll("#result tr").forEach(tr => {
    const cols = Array.from(tr.querySelectorAll("td")).map(td => td.textContent || td.querySelector("input").value);
    rows.push(cols);
  });
  const csvContent = rows.map(e => e.join(",")).join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "盤點結果.csv";
  link.click();
}

loadData();
