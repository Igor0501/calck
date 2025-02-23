document.addEventListener('DOMContentLoaded', () => {
    fetchExcelFile();
});

let workbook;

async function fetchExcelFile() {
    try {
        const response = await fetch('/path/to/your/excel/file.xlsx');
        const data = await response.arrayBuffer();
        workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
    } catch (error) {
        console.error('Error fetching the Excel file:', error);
    }
}

function calculate() {
    if (!workbook) {
        alert('Помилка завантаження файлу Excel!');
        return;
    }

    const tirazh = parseInt(document.getElementById('tirazh').value);
    const kolirnist = parseInt(document.getElementById('kolirnist').value);
    const paper = document.getElementById('paper').value;

    if (isNaN(tirazh) || isNaN(kolirnist) || !paper) {
        alert('Будь ласка, заповніть всі поля!');
        return;
    }

    const paperPrice = getPaperPrice(paper);
    const cost = tirazh * kolirnist * paperPrice;

    document.getElementById('result').textContent = `Вартість друку: ${cost.toFixed(2)} грн`;
}

function getPaperPrice(paper) {
    const sheet = workbook.Sheets['Sheet1'];
    for (let i = 1; i <= 100; i++) {
        if (sheet[`A${i}`]?.v === paper) {
            return sheet[`B${i}`]?.v || 0;
        }
    }
    return 0;
}

function calculateTirazh() {
    const width = parseFloat(document.getElementById('width').value);
    const height = parseFloat(document.getElementById('height').value);
    const requiredTirazh = parseInt(document.getElementById('requiredTirazh').value);

    if (isNaN(width) || isNaN(height) || isNaN(requiredTirazh)) {
        alert('Будь ласка, заповніть всі поля!');
        return;
    }

    const paperArea = (width / 1000) * (height / 1000);
    const tirazh = Math.ceil(requiredTirazh / paperArea);

    document.getElementById('tirazhResult').textContent = `Необхідний тираж: ${tirazh} листів`;
}