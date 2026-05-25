const ExcelJS = require('exceljs');
const path = require('path');

async function analyzeTemplate() {
  const filePath = path.join('C:', 'Users', 'augus', 'Downloads', 'NOVOPAGLIMP.xlsx');
  const workbook = new ExcelJS.Workbook();
  
  try {
    await workbook.xlsx.readFile(filePath);
  } catch(e) {
    console.log('Erro lendo arquivo. Tentando com caminho relativo...');
    await workbook.xlsx.readFile('../../NOVOPAGLIMP.xlsx');
  }
  
  const sheet = workbook.worksheets[0];
  
  console.log('📊 MAPEAMENTO DO NOVO TEMPLATE PEG LIMP\n');
  console.log('='.repeat(80));
  
  // Analisar linhas importantes
  for (let row = 1; row <= 35; row++) {
    const r = sheet.getRow(row);
    let content = '';
    
    for (let col = 1; col <= 6; col++) {
      const cell = r.getCell(col);
      if (cell.value) {
        const val = String(cell.value).substring(0, 25);
        content += `[${String.fromCharCode(64 + col)}:${val}] `;
      }
    }
    
    if (content) {
      console.log(`L${row.toString().padStart(2)}: ${content}`);
    }
  }
  
  console.log('\n' + '='.repeat(80));
  console.log('📍 ESTRUTURA IDENTIFICADA:\n');
  
  // Verificar campos específicos
  const toCheck = [
    'B6', 'E6', 'B7', 'B8', 'B9', // Customer data + date
    'A13', 'B13', 'C13', 'D13', 'E13', // Table headers
    'A14', // First product row
    'A26', 'A27', 'A28', 'A29', // Totals
    'B31', 'E31' // Payment/Delivery
  ];
  
  for (const addr of toCheck) {
    const cell = sheet.getCell(addr);
    if (cell.value) {
      console.log(`${addr.padEnd(5)}: "${String(cell.value).substring(0, 40)}"`);
    }
  }
}

analyzeTemplate().catch(console.error);
