
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
    'Кыргызстан': 'Кыргызский сом (Кыргызстан)'
  };
  return map[country] || 'Рубль (Россия)';
}
function parseLimit(value) {
  const match = /\d+(\.\d+)?/.exec(value);
  return match ? parseFloat(match[0]) : NaN;
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

  fetch("https://script.google.com/macros/s/AKfycbzlnU77HvUMHMW41fGuKl1-gQ3k6s_qSzDYQ_t1IlTu85GGHEtMDSP3Gwm2KX5IPMSZ/exec")
    .then(res => res.json())
    .then(data => {
      const rows = [];

      const getRate = (type, limit, value) => {
        const candidates = data.filter(r => r["Тип операции"] === type);
        for (let row of candidates) {
          const threshold = parseLimit(row["Unnamed: 1"]);
          if (!isNaN(threshold) && value <= threshold) {
            return parseFloat(row[column]);
          }
        }
        return null;
      };

      // Приемка
      const rcp = getRate("Приемка", "Unnamed: 1", weight);
      if (rcp !== null) {
        rows.push({ name: "Приемка", quantity, rate: rcp, total: rcp * quantity });
      }

      // Подготовка товара
      if (model === "FBO") {
        const prep = getRate("FBO подготовка", "Unnamed: 1", length);
        if (prep !== null) {
          rows.push({ name: "Подготовка товара", quantity, rate: prep, total: prep * quantity });
        }
      }

      // Хранение
      if (storageDays > 0) {
        const store = getRate("Хранение", "Unnamed: 1", weight);
        if (store !== null) {
          rows.push({ name: "Хранение", quantity: quantity * storageDays, rate: store, total: store * quantity * storageDays });
        }
      }

      // Сборка / Экспресс-сборка
      const first = (city === "Москва" || city === "Санкт-Петербург") ? 120 : 80;
      const assembly = express
        ? 240 + (quantity - 1) * 26
        : first + Math.max(0, weight - 5) * 13;
      rows.push({ name: express ? "Экспресс-сборка" : "Сборка", quantity: orders, rate: assembly, total: assembly * orders });

      // Объявленная стоимость
      if (declaredValue > 0) {
        const rate = declaredValue * 0.0001;
        rows.push({ name: "Сбор за объявленную стоимость", quantity, rate, total: rate * quantity });
      }

      // Доставка
      if (model === "FBS") {
        const delivery = getRate("Доставка", "Unnamed: 1", weight);
        if (delivery !== null) {
          rows.push({ name: "Доставка", quantity: 1, rate: delivery, total: delivery });
        }
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
