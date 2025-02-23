document.addEventListener('DOMContentLoaded', () => {
    fetchExcelFile();
});

let workbook;

async function fetchExcelFile() {
    try {
        const response = await fetch('/Головний прайс.xlsx');
        const data = await response.arrayBuffer();
        workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
        populateMaterials();
    } catch (error) {
        console.error('Error fetching the Excel file:', error);
    }
}

function populateMaterials() {
    const sheet = workbook.Sheets['Вартість матеріалів і послуг'];
    const materials = [];

    for (let i = 5; i <= 100; i++) {
        const material = sheet[`A${i}`]?.v;
        if (material) {
            materials.push(material);
        }
    }

    const materialSelect = document.getElementById('material');
    materials.forEach(material => {
        const option = document.createElement('option');
        option.value = material;
        option.textContent = material;
        materialSelect.appendChild(option);
    });

    // Додаємо слухачі подій для автоматичного розрахунку
    document.getElementById('width').addEventListener('input', calculate);
    document.getElementById('height').addEventListener('input', calculate);
    document.getElementById('material').addEventListener('change', calculate);
    document.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
        checkbox.addEventListener('change', calculate);
    });
}

function calculate() {
    if (!workbook) {
        alert('Помилка завантаження файлу Excel!');
        return;
    }

    const width = parseFloat(document.getElementById('width').value);
    const height = parseFloat(document.getElementById('height').value);
    const material = document.getElementById('material').value;

    if (isNaN(width) || isNaN(height) || !material) {
        return;
    }

    const materialPrice = getMaterialPrice(material);
    const materialCost = (width / 100) * (height / 100) * materialPrice;

    const cmyk = document.getElementById('cmyk').checked ? calculateCheckbox('CMYK', width, height) : 0;
    const cmyk2 = document.getElementById('cmyk2').checked ? calculateCheckbox('CMYK+CMYK', width, height) : 0;
    const cmykWhite = document.getElementById('cmykWhite').checked ? calculateCheckbox('CMYK+White', width, height) : 0;
    const mattOracal = document.getElementById('mattOracal').checked ? calculateOracal('матовий оракал', width, height) : 0;
    const colorOracal = document.getElementById('colorOracal').checked ? calculateOracal('кольоровий оракал', width, height) : 0;
    const scotch = document.getElementById('scotch').checked ? calculateScotch(width, height) : 0;
    const bend1 = document.getElementById('bend1').checked ? calculateBend(1, width, height) : 0;
    const bend2 = document.getElementById('bend2').checked ? calculateBend(2, width, height) : 0;
    const laserCut = document.getElementById('laserCut').checked ? calculateCut(15, width, height) : 0;
    const knifeCut = document.getElementById('knifeCut').checked ? calculateCut(5, width, height) : 0;

    const totalCost = materialCost + cmyk + cmyk2 + cmykWhite + mattOracal + colorOracal + scotch + bend1 + bend2 + laserCut + knifeCut;

    let additionalServices = '';
    if (cmyk > 0) additionalServices += `Друк CMYK: ${cmyk.toFixed(2)} грн<br>`;
    if (cmyk2 > 0) additionalServices += `Друк CMYK+CMYK: ${cmyk2.toFixed(2)} грн<br>`;
    if (cmykWhite > 0) additionalServices += `Друк CMYK+White: ${cmykWhite.toFixed(2)} грн<br>`;
    if (mattOracal > 0) additionalServices += `Матовий оракал: ${mattOracal.toFixed(2)} грн<br>`;
    if (colorOracal > 0) additionalServices += `Кольоровий оракал: ${colorOracal.toFixed(2)} грн<br>`;
    if (scotch > 0) additionalServices += `Скотч листовий: ${scotch.toFixed(2)} грн<br>`;
    if (bend1 > 0) additionalServices += `1 гнуття: ${bend1.toFixed(2)} грн<br>`;
    if (bend2 > 0) additionalServices += `2 гнуття: ${bend2.toFixed(2)} грн<br>`;
    if (laserCut > 0) additionalServices += `Лазерна порізка: ${laserCut.toFixed(2)} грн<br>`;
    if (knifeCut > 0) additionalServices += `порізка ножем: ${knifeCut.toFixed(2)} грн<br>`;

    document.getElementById('result').innerHTML = `
        <strong>Загальна вартість: ${totalCost.toFixed(2)} грн</strong><br><br>
        <strong>Додаткові послуги:</strong><br>
        Вартість матеріалу: ${materialCost.toFixed(2)} грн<br>
        ${additionalServices}
    `;
}

function getMaterialPrice(material) {
    const sheet = workbook.Sheets['Вартість матеріалів і послуг'];
    for (let i = 5; i <= 100; i++) {
        if (sheet[`A${i}`]?.v === material) {
            return sheet[`B${i}`]?.v || 0;
        }
    }
    return 0;
}

function calculateCheckbox(type, width, height) {
    const sheet = workbook.Sheets['УФ друк'];
    const base = (width / 100) * (height / 100);
    const isLarge = width > 40 && height > 40;

    switch (type) {
        case 'CMYK':
            return isLarge ? base * (sheet['C20']?.v || 0) : base * (sheet['C27']?.v || 0);
        case 'CMYK+CMYK':
            return isLarge ? base * (sheet['C20']?.v || 0) * 2 : base * (sheet['C27']?.v || 0) * 2;
        case 'CMYK+White':
            return isLarge ? base * (sheet['C21']?.v || 0) : base * (sheet['C26']?.v || 0);
        default:
            return 0;
    }
}

function calculateOracal(type, width, height) {
    const sheet = workbook.Sheets['Вартість матеріалів і послуг'];
    const base = (width / 100) * (height / 100);

    let price1 = 0;
    let price2 = 0;

    if (type === 'матовий оракал') {
        price1 = sheet['B6']?.v || 0;
        price2 = sheet['G42']?.v || 0;
    } else if (type === 'кольоровий оракал') {
        price1 = sheet['B7']?.v || 0;
        price2 = sheet['G42']?.v || 0;
    }

    return base * price1 + base * price2;
}

function calculateScotch(width, height) {
    const sheet = workbook.Sheets['Вартість матеріалів і послуг'];
    const base = (width / 100) * (height / 100);
    const price1 = sheet['G42']?.v || 0;
    const price2 = sheet['B13']?.v || 0;

    return base * price1 + base * price2;
}

function calculateBend(count, width, height) {
    const sheet = workbook.Sheets['Вартість матеріалів і послуг'];
    const minDimension = Math.min(width, height);
    const price = sheet['G28']?.v || 0;

    return (minDimension / 100) * price * count;
}

function calculateCut(pricePerMeter, width, height) {
    return ((width + height) * 2 / 100) * pricePerMeter;
}