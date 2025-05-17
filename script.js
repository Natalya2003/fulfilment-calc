
function runCalculation() {
  const model = document.getElementById("model").value;
  const city = document.getElementById("city").value;
  const country = document.getElementById("country").value;
  const weight = parseFloat(document.getElementById("weight").value);
  const length = parseFloat(document.getElementById("length").value);
  const quantity = parseInt(document.getElementById("quantity").value);
  const orders = parseInt(document.getElementById("orders").value);
  const storageDays = parseInt(document.getElementById("storageDays").value) || 0;
  const declaredValue = parseFloat(document.getElementById("declaredValue").value) || 0;
  const express = document.getElementById("express").checked;
  const currency = "₽";

  if (!city || !weight || !quantity || !orders || (model === "FBO" && !length)) {
    alert("Пожалуйста, заполните все обязательные поля");
    return;
  }

  const rows = [];

  // Приемка
  const rateReception = weight <= 10 ? 10 : 13;
  rows.push({ name: "Приемка", quantity, rate: rateReception, total: rateReception * quantity });

  // Подготовка (FBO)
  if (model === "FBO") {
    const ratePrep = length <= 30 ? 120 : 160;
    rows.push({ name: "Подготовка товара", quantity, rate: ratePrep, total: ratePrep * quantity });
  }

  // Хранение
  if (storageDays > 0) {
    const rateStorage = weight <= 10 ? 3.9 : 5.85;
    const total = rateStorage * quantity * storageDays;
    rows.push({ name: "Хранение", quantity: quantity * storageDays, rate: rateStorage, total });
  }

  // Сборка
  const firstRate = (city === "Москва" || city === "Санкт-Петербург") ? 120 : 80;
  const assemblyRate = express
    ? 240 + (quantity - 1) * 26
    : firstRate + Math.max(0, weight - 5) * 13;
  rows.push({ name: express ? "Экспресс-сборка" : "Сборка", quantity: orders, rate: assemblyRate, total: assemblyRate * orders });

  // Объявленная стоимость
  if (declaredValue > 0) {
    const rate = declaredValue * 0.0001;
    rows.push({ name: "Сбор за объявленную стоимость", quantity, rate, total: rate * quantity });
  }

  // Доставка (FBS)
  if (model === "FBS") {
    const delivery = weight <= 5 ? 35 : 35 + (weight - 5) * 10;
    rows.push({ name: "Доставка (FBS)", quantity: 1, rate: delivery, total: delivery });
  }

  const totalSum = rows.reduce((sum, r) => sum + r.total, 0);
  let html = '<table><tr><th>Тип операции</th><th>Кол-во</th><th>Тариф</th><th>Итого</th></tr>';
  rows.forEach(r => {
    html += `<tr><td>${r.name}</td><td>${r.quantity}</td><td>${r.rate} ${currency}</td><td>${r.total.toFixed(2)} ${currency}</td></tr>`;
  });
  html += `</table><p><strong>Общая сумма: ${totalSum.toFixed(2)} ${currency}</strong></p>`;
  document.getElementById("result").innerHTML = html;
  window.lastResult = { breakdown: rows, total: totalSum };
}

function downloadExcel() {
  if (!window.lastResult) return;
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Тип операции", "Кол-во", "Тариф", "Итого"],
    ...window.lastResult.breakdown.map(row => [row.name, row.quantity, `${row.rate} ₽`, `${row.total.toFixed(2)} ₽`]),
    [],
    ["Общая сумма", "", "", `${window.lastResult.total.toFixed(2)} ₽`]
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Расчет");
  XLSX.writeFile(wb, "fulfillment.xlsx");
}
