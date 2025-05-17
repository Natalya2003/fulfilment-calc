
function runCalculation() {
  fetch("https://sheet.best/api/sheets/14c974b2-0133-4ee4-ae01-0ea33b2e8a93")
    .then(r => r.json())
    .then(data => {
      const model = document.getElementById("model").value;
      const city = document.getElementById("city").value;
      const weight = parseFloat(document.getElementById("weight").value);
      const length = parseFloat(document.getElementById("length").value);
      const quantity = parseInt(document.getElementById("quantity").value);
      const orders = parseInt(document.getElementById("orders").value);
      const storageDays = parseInt(document.getElementById("storageDays").value);
      const declaredValue = parseFloat(document.getElementById("declaredValue").value);
      const express = document.getElementById("express").checked;

      if (!weight || !quantity || !orders) {
        alert("Заполните обязательные поля.");
        return;
      }

      const rows = [];
      const getTariff = (type, field, compare) =>
        data.find(r => r["Тип операции"] === type && parseFloat(r[field]) >= compare);

      const reception = getTariff("Приемка", "До кг", weight);
      if (reception) {
        const rate = parseFloat(reception["Рубль (Россия)"]);
        rows.push({ name: "Приемка", quantity, rate, total: rate * quantity });
      }

      if (model === "FBO") {
        const prep = getTariff("FBO подготовка", "До см", length);
        if (prep) {
          const rate = parseFloat(prep["Рубль (Россия)"]);
          rows.push({ name: "Подготовка", quantity, rate, total: rate * quantity });
        }
      }

      if (storageDays > 0) {
        const store = getTariff("Хранение", "До кг", weight);
        if (store) {
          const rate = parseFloat(store["Рубль (Россия)"]);
          rows.push({ name: "Хранение", quantity: quantity * storageDays, rate, total: rate * quantity * storageDays });
        }
      }

      const first = (city === "Москва" || city === "Санкт-Петербург") ? 120 : 80;
      const assembly = express ? 240 + (quantity - 1) * 26 : first + Math.max(0, weight - 5) * 13;
      rows.push({ name: express ? "Экспресс-сборка" : "Сборка", quantity: orders, rate: assembly, total: assembly * orders });

      if (declaredValue > 0) {
        const rate = declaredValue * 0.0001;
        rows.push({ name: "Сбор за объявленную стоимость", quantity, rate, total: rate * quantity });
      }

      if (model === "FBS") {
        const delivery = weight <= 5 ? 35 : 35 + (weight - 5) * 10;
        rows.push({ name: "Доставка (FBS)", quantity: 1, rate: delivery, total: delivery });
      }

      const total = rows.reduce((s, r) => s + r.total, 0);
      let html = "<table><tr><th>Операция</th><th>Кол-во</th><th>Тариф</th><th>Итого</th></tr>";
      rows.forEach(r => {
        html += `<tr><td>${r.name}</td><td>${r.quantity}</td><td>${r.rate} ₽</td><td>${r.total.toFixed(2)} ₽</td></tr>`;
      });
      html += `</table><p><strong>Общая сумма: ${total.toFixed(2)} ₽</strong></p>`;
      document.getElementById("result").innerHTML = html;
      window.lastResult = { breakdown: rows, total };
    });
}

function downloadExcel() {
  if (!window.lastResult) return;
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Тип операции", "Кол-во", "Тариф", "Итого"],
    ...window.lastResult.breakdown.map(r => [r.name, r.quantity, `${r.rate} ₽`, `${r.total.toFixed(2)} ₽`]),
    [],
    ["Общая сумма", "", "", `${window.lastResult.total.toFixed(2)} ₽`]
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Расчет");
  XLSX.writeFile(wb, "fulfillment.xlsx");
}
