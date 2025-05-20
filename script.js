
const currencies = {
  'Россия': '₽',
  'Казахстан': '₸',
  'Беларусь': 'Br',
  'Китай': '¥',
  'Турция': '₺'
};

function getTariff(table, val) {
  const found = table.find(t => val <= t.max);
  return found ? found.value : null;
}

function calculateFulfillment(options) {
  const {
    country,
    city,
    model,
    quantity,
    weight,
    length,
    storageDays = 0,
    declaredValue = 0,
    express = false
  } = options;

  const tariffs = {
    reception: {
      'Россия': [
        { max: 5, value: 7 }, { max: 10, value: 10 },
        { max: 20, value: 13 }, { max: 30, value: 16 },
        { max: 50, value: 20 }, { max: 100, value: 30 }
      ],
      'Казахстан': [
        { max: 10, value: 50 }
      ],
      'Беларусь': [
        { max: 10, value: 0.4 }
      ],
      'Китай': [
        { max: 10, value: 1 }
      ],
      'Турция': [
        { max: 10, value: 3.5 }
      ]
    },
    preparation: [
      { max: 10, value: 16 }, { max: 20, value: 19 }, { max: 30, value: 24 },
      { max: 40, value: 32 }, { max: 50, value: 39 }, { max: 60, value: 56 },
      { max: 70, value: 68 }, { max: 80, value: 83 }, { max: 90, value: 98 },
      { max: 100, value: 108 }, { max: 110, value: 118 }, { max: 120, value: 128 }
    ],
    storage: [
      { max: 0.1, value: 0.1 }, { max: 0.5, value: 0.38 }, { max: 1, value: 0.6 },
      { max: 3, value: 1.3 }, { max: 5, value: 1.7 }, { max: 10, value: 3.9 },
      { max: 20, value: 5.85 }, { max: 30, value: 11 }, { max: 50, value: 23.4 },
      { max: 100, value: 39 }
    ]
  };

  const receptionRate = getTariff(tariffs.reception[country] || [], weight);
  const preparationRate = model === 'FBO' ? getTariff(tariffs.preparation, length) : 0;
  const storageRate = getTariff(tariffs.storage, weight);
  if (!receptionRate || (model === 'FBO' && !preparationRate) || !storageRate) {
    return { error: 'Один из тарифов отсутствует. Возможно, вес или длина превышают допустимые пределы.' };
  }

  const receptionTotal = receptionRate * quantity;
  const preparationTotal = preparationRate * quantity;
  const storageTotal = storageRate * quantity * storageDays;

  let assemblyRate;
  if (express) {
    assemblyRate = 240 + (quantity - 1) * 26 + Math.max(0, weight - 5) * 26;
  } else {
    const firstUnit = (city === 'Москва' || city === 'Санкт-Петербург') ? 120 : 80;
    assemblyRate = firstUnit + Math.max(0, weight - 5) * 13;
  }

  const declaredTotal = declaredValue ? declaredValue * 0.0001 * quantity : 0;

  const total = receptionTotal + preparationTotal + storageTotal + assemblyRate + declaredTotal;

  return {
    currency: currencies[country] || '₽',
    breakdown: [
      { name: 'Приемка', quantity, rate: receptionRate, total: receptionTotal },
      model === 'FBO' ? { name: 'Подготовка товара', quantity, rate: preparationRate, total: preparationTotal } : null,
      storageDays > 0 ? { name: 'Хранение', quantity: quantity * storageDays, rate: storageRate, total: storageTotal } : null,
      { name: express ? 'Экспресс-сборка заказа' : 'Сборка заказа', quantity: 1, rate: assemblyRate, total: assemblyRate },
      declaredValue > 0 ? { name: 'Сбор за объявленную стоимость', quantity, rate: declaredValue * 0.0001, total: declaredTotal } : null
    ].filter(Boolean),
    total: Math.round(total * 100) / 100
  };
}

function runCalculation() {
  const result = calculateFulfillment({
    country: document.getElementById('country').value,
    city: document.getElementById('city').value,
    model: document.getElementById('model').value,
    quantity: parseInt(document.getElementById('quantity').value),
    weight: parseFloat(document.getElementById('weight').value),
    length: parseInt(document.getElementById('length').value),
    storageDays: parseInt(document.getElementById('storageDays').value) || 0,
    declaredValue: parseFloat(document.getElementById('declaredValue').value) || 0,
    express: document.getElementById('express').checked
  });

  const resultContainer = document.getElementById('result');
  if (result.error) {
    resultContainer.innerHTML = `<p style='color:red'>${result.error}</p>`;
    return;
  }

  let html = '<table><tr><th>Тип операции</th><th>Кол-во</th><th>Тариф</th><th>Итого</th></tr>';
  result.breakdown.forEach(item => {
    html += `<tr><td>${item.name}</td><td>${item.quantity}</td><td>${item.rate} ${result.currency}</td><td>${item.total.toFixed(2)} ${result.currency}</td></tr>`;
  });
  html += `</table><p><strong>Общая сумма: ${result.total} ${result.currency}</strong></p>`;
  resultContainer.innerHTML = html;
  window.lastResult = result;
}

function downloadPDF() {
  if (!window.lastResult) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  doc.text('Расчет стоимости фулфилмента', 14, 16);
  doc.autoTable({
    startY: 20,
    head: [['Тип операции', 'Кол-во', 'Тариф', 'Итого']],
    body: window.lastResult.breakdown.map(item => [
      item.name,
      item.quantity,
      `${item.rate} ${window.lastResult.currency}`,
      `${item.total.toFixed(2)} ${window.lastResult.currency}`
    ])
  });
  doc.text(`Общая сумма: ${window.lastResult.total} ${window.lastResult.currency}`, 14, doc.lastAutoTable.finalY + 10);
  doc.save('fulfillment.pdf');
}

function downloadExcel() {
  if (!window.lastResult) return;
  const wb = XLSX.utils.book_new();
  const ws_data = [
    ['Тип операции', 'Кол-во', 'Тариф', 'Итого'],
    ...window.lastResult.breakdown.map(item => [
      item.name,
      item.quantity,
      `${item.rate} ${window.lastResult.currency}`,
      `${item.total.toFixed(2)} ${window.lastResult.currency}`
    ]),
    [],
    ['Общая сумма', '', '', `${window.lastResult.total} ${window.lastResult.currency}`]
  ];
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'Расчет');
  XLSX.writeFile(wb, 'fulfillment.xlsx');
}
