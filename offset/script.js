function calculate() {
  const tirazh = parseInt(document.getElementById('tirazh').value);
  const kolirnist = parseInt(document.getElementById('kolirnist').value);
  const paper = document.getElementById('paper').value;

  const paperPrices = {
    'Крейда 90': { density: 90, price: 73 },
    'Крейда 115': { density: 115, price: 73 },
    'Крейда 130': { density: 130, price: 73 },
    'Крейда 150': { density: 150, price: 73 },
    'Крейда 170': { density: 170, price: 73 },
    'Крейда 200': { density: 200, price: 73 },
    'Крейда 250': { density: 250, price: 73 },
    'Крейда 300': { density: 300, price: 73 },
    'Крейда 350': { density: 350, price: 73 },
    'Офсетка 80': { density: 80, price: 73 },
    'Офсетка 120': { density: 120, price: 73 },
    'Офсетка 160': { density: 160, price: 73 },
    'Офсетка 250': { density: 250, price: 73 },
    'Картон 220': { density: 220, price: 73 },
    'Картон 235': { density: 235, price: 73 },
    'Крафт': { density: 45, price: 45 },
    'Папір клієнта': { density: 30, price: 30 }
  };

  const paperData = paperPrices[paper];
  const paperDensity = paperData.density;
  const paperPrice = paperData.price;
  const platesPrice = 175;

  const result = Math.round(
    eval(
      `Math.ceil((${tirazh} + ${paperDensity}) * ${paperPrice} + (${platesPrice} * ${kolirnist}))`
    ) / 100
  ) * 100;

  document.getElementById('result').textContent = `Вартість: ${result} ₴`;
}

function calculateTirazh() {
  const width = parseInt(document.getElementById('width').value);
  const height = parseInt(document.getElementById('height').value);
  const requiredTirazh = parseInt(document.getElementById('requiredTirazh').value);

  const listCount = Math.max(
    Math.floor(425 / width) * Math.floor(600 / height),
    Math.floor(430 / height) * Math.floor(620 / width)
  );

  const printTirazh = Math.ceil(requiredTirazh / listCount / 10) * 10;

  document.getElementById('tirazhResult').textContent = `Тираж друку: ${printTirazh}`;
}