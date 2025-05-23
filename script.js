
function runCalculation() {
  fetch("https://script.google.com/macros/s/AKfycbzlnU77HvUMHMW41fGuKl1-gQ3k6s_qSzDYQ_t1IlTu85GGHEtMDSP3Gwm2KX5IPMSZ/exec")
    .then(res => res.json())
    .then(data => {
      const types = [...new Set(data.map(r => r["Тип операции"]))];
      console.log("🔍 Найденные типы операций:", types);
    });
}
