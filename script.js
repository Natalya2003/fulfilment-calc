
function runCalculation() {
  const city = document.getElementById("city").value;
  const weight = parseFloat(document.getElementById("weight").value);
  const quantity = parseInt(document.getElementById("quantity").value);
  const length = parseFloat(document.getElementById("length").value);
  const model = document.getElementById("model").value;

  document.querySelectorAll('input').forEach(el => el.style.borderColor = '');

  let error = false;
  function mark(id, msg) {
    document.getElementById(id).style.borderColor = 'red';
    alert(msg);
    error = true;
  }

  if (!city) mark("city", "Введите город");
  if (!weight || weight <= 0) mark("weight", "Введите корректный вес");
  if (!quantity || quantity <= 0) mark("quantity", "Введите количество");
  if (model === "FBO" && (!length || length <= 0)) mark("length", "Введите длину для FBO");

  if (error) return;

  document.getElementById("result").innerHTML = "<p><strong>Расчет выполнен (заглушка)</strong></p>";
}

function downloadExcel() {
  alert("Экспорт в Excel будет позже.");
}
