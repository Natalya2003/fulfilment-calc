
function getCurrency(country) {
  const map = {
    'Россия': '₽', 'Казахстан': '₸', 'Беларусь': 'Br', 'Китай': '元',
    'Азербайджан': '₼', 'Турция': '₺', 'США': '$', 'ОАЭ': 'AED',
    'Армения': '֏', 'Испания': '€', 'Кыргызстан': 'сом'
  };
  return map[country] || '₽';
}
function getColumn(country) {
  const map = {
    'Россия': 'Рубль (Россия)', 'Казахстан': 'Тенге (Казахстан)', 'Беларусь': 'Белорусский рубль (Беларусь)',
    'Китай': 'Юань  元 (Китай)', 'Азербайджан': 'Манат azn (Азербайджан)', 'Турция': 'Лира  ₺ (Турция)',
    'США': 'Доллар, $ (США)', 'ОАЭ': 'AED (ОАЭ)', 'Армения': 'AMD(Армения)', 'Испания': '€',
    'Кыргызстан': 'Кыргызский сом'
  };
  return map[country] || 'Рубль (Россия)';
}
function runCalculation() {
  const model = document.getElementById("model").value;
  const country = document.getElementById("country").value;
  const column = getColumn(country);
  const currency = getCurrency(country);
  const city = document.getElementById("city").value;
  const weight = parseFloat(document.getElementById("weight").value);
  const length = parseFloat(document.getElementById("length").value);
  const quantity = parseInt(document.getElementById("quantity").value);
  const orders = parseInt(document.getElementById("orders").value);
  const storageDays = parseInt(document.getElementById("storageDays").value) || 0;
  const declaredValue = parseFloat(document.getElementById("declaredValue").value) || 0;
  const express = document.getElementById("express").checked;

  fetch("https://api.sheetbest.com/sheets/10732c74-4e4f-4c68-91ca-de2aadf98e31")
    .then(res => res.json())
    .then(data => {
      const rows = [];
      const getRate = (type, field, value) =>
        data.find(r => r["Тип операции"] === type && parseFloat(r[field]) >= value);

      const rcp = getRate("Приемка", "До кг", weight);
      if (rcp) {
        const rate = parseFloat(rcp[column]);
        rows.push({ name: "Приемка", quantity, rate, total: rate * quantity });
      }

      if (model === "FBO") {
        const prep = getRate("FBO подготовка", "До см", length);
        if (prep) {
          const rate = parseFloat(prep[column]);
          rows.push({ name: "Подготовка", quantity, rate, total: rate * quantity });
        }
      }

      if (storageDays > 0) {
        const store = getRate("Хранение", "До кг", weight);
        if (store) {
          const rate = parseFloat(store[column]);
          rows.push({ name: "Хранение", quantity: quantity * storageDays, rate, total: rate * quantity * storageDays });
        }
      }

      const first = (city === "Москва" || city === "Санкт-Петербург") ? 120 : 80;
      const assembly = express
        ? 240 + (quantity - 1) * 26
        : first + Math.max(0, weight - 5) * 13;
      rows.push({ name: express ? "Экспресс-сборка" : "Сборка", quantity: orders, rate: assembly, total: assembly * orders });

      if (declaredValue > 0) {
        const rate = declaredValue * 0.0001;
        rows.push({ name: "Сбор за объявленную стоимость", quantity, rate, total: rate * quantity });
      }

      if (model === "FBS") {
        const delivery = weight <= 5 ? 35 : 35 + (weight - 5) * 10;
        rows.push({ name: "Доставка (FBS)", quantity: 1, rate: delivery, total: delivery });
      }

      const total = rows.reduce((sum, r) => sum + r.total, 0);
      let html = '<table><tr><th>Тип операции</th><th>Кол-во</th><th>Тариф</th><th>Итого</th></tr>';
      rows.forEach(r => {
        html += `<tr><td>${r.name}</td><td>${r.quantity}</td><td>${r.rate} ${currency}</td><td>${r.total.toFixed(2)} ${currency}</td></tr>`;
      });
      html += `</table><p><strong>Общая сумма: ${total.toFixed(2)} ${currency}</strong></p>`;
      document.getElementById("result").innerHTML = html;
      window.lastResult = { breakdown: rows, total };
    });
}
function downloadExcel() {
  if (!window.lastResult) return;
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Тип операции", "Кол-во", "Тариф", "Итого"],
    ...window.lastResult.breakdown.map(r => [r.name, r.quantity, r.rate, r.total.toFixed(2)]),
    [],
    ["Общая сумма", "", "", window.lastResult.total.toFixed(2)]
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Расчет");
  XLSX.writeFile(wb, "fulfillment.xlsx");
}
