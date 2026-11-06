import ExcelJS from 'exceljs';
import { COMPANIES } from '../data/empresas';
import { Quote } from '../model/orcamento.model';
import { ExcelGeneratorOmaService } from './excel-generator-oma.service';

describe('ExcelGeneratorOmaService', () => {
  const oma = COMPANIES.find((company) => company.id === 'oma')!;

  function createTemplateWorkbook(): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cotação');

    worksheet.mergeCells('E3:F3');
    worksheet.mergeCells('G3:L3');
    worksheet.mergeCells('C7:F7');
    worksheet.mergeCells('C8:F8');

    for (let row = 13; row <= 26; row++) {
      worksheet.mergeCells(`B${row}:E${row}`);
    }

    worksheet.mergeCells('G28:H28');
    worksheet.mergeCells('I28:J28');
    worksheet.mergeCells('K28:L28');

    for (let row = 30; row <= 32; row++) {
      worksheet.mergeCells(`E${row}:F${row}`);
      worksheet.mergeCells(`G${row}:H${row}`);
      worksheet.mergeCells(`I${row}:J${row}`);
      worksheet.mergeCells(`K${row}:L${row}`);
    }

    worksheet.mergeCells('B34:L34');
    worksheet.mergeCells('B35:E35');
    worksheet.getCell('F28').value = 'VALOR TOTAL';
    worksheet.getCell('E30').value = 'CONDIÇÕES DE PAGAMENTO';
    worksheet.getCell('E31').value = 'FORMA DE PAGAMENTO';
    worksheet.getCell('E32').value = 'PRAZO DE ENTREGA';
    worksheet.getCell('B34').value =
      'PRODUTOS ORÇADOS PELA INTERNET SOFREM ALTERAÇÕES NO PREÇO !!!';
    worksheet.getRow(13).height = 16.5;
    worksheet.pageSetup.printArea = 'A1:P40';

    return workbook;
  }

  function createQuote(itemCount = 2): Quote {
    return {
      label: 'Orçamento A',
      company: oma,
      deliveryTime: '5 dias úteis',
      paymentMethod: 'Boleto 28 dias',
      customerData: {
        client: 'Condomínio Exemplo',
        address: 'Rua Um, 10',
        district: 'Centro',
        city: 'São Paulo',
        state: 'SP',
        postalCode: '01000-000',
        salesperson: 'Maria',
      },
      items: Array.from({ length: itemCount }, (_, index) => ({
        id: index + 1,
        product: `Produto ${index + 1}`,
        quantity: index + 2,
        unit: 'unid.',
        unitPrice: 10,
        discount: index === 0 ? 10 : 0,
      })),
      adjustment: 0,
      source: 'Manual',
      date: new Date(2026, 10, 5),
    };
  }

  it('preenche os dados do orçamento no primeiro bloco do modelo OMA', () => {
    const service = new ExcelGeneratorOmaService();
    const workbook = createTemplateWorkbook();

    service.populateWorkbook(workbook, createQuote());

    const worksheet = workbook.worksheets[0];
    expect(worksheet.getCell('G3').value).toBe('Orçamento A - 05/11/2026');
    expect(worksheet.getCell('C7').value).toBe('Condomínio Exemplo');
    expect(worksheet.getCell('C8').value).toBe('Rua Um, 10 - Centro - São Paulo - SP - 01000-000');
    expect(worksheet.getCell('H7').value).toBe('Maria');
    expect(worksheet.getCell('G9').value).toBe('OMA');
    expect(worksheet.getCell('B13').value).toBe('Produto 1');
    expect(worksheet.getCell('F13').value).toBe(2);
    expect(worksheet.getCell('G13').value).toBe(10);
    expect(worksheet.getCell('H13').value).toEqual({
      formula: 'G13*F13*(1-10/100)',
      result: 18,
    });
    expect(worksheet.getCell('G28').value).toEqual({
      formula: 'SUM(H13:H27)',
      result: 48,
    });
    expect(worksheet.getCell('G31').value).toBe('Boleto 28 dias');
    expect(worksheet.getCell('G32').value).toBe('5 dias úteis');
    expect(worksheet.getCell('I28').value).toBeNull();
    expect(worksheet.getColumn('G').hidden).toBe(false);
    expect(worksheet.getColumn('H').hidden).toBe(false);
    expect(worksheet.getColumn('I').hidden).toBe(true);
    expect(worksheet.getColumn('L').hidden).toBe(true);
    expect(worksheet.pageSetup.printArea).toBe('A1:H40');
    expect(worksheet.model.merges).toContain('G3:H3');
    expect(worksheet.model.merges).not.toContain('G3:L3');
    expect(worksheet.model.merges).toContain('B34:H34');
    expect(worksheet.model.merges).not.toContain('I9:J9');
  });

  it('expande a tabela e move totais e condições sem truncar itens', () => {
    const service = new ExcelGeneratorOmaService();
    const workbook = createTemplateWorkbook();

    service.populateWorkbook(workbook, createQuote(17));

    const worksheet = workbook.worksheets[0];
    expect(worksheet.getCell('B29').value).toBe('Produto 17');
    expect(worksheet.getCell('G30').value).toEqual({
      formula: 'SUM(H13:H29)',
      result: 1698,
    });
    expect(worksheet.getCell('G33').value).toBe('Boleto 28 dias');
    expect(worksheet.getCell('G34').value).toBe('5 dias úteis');
    expect(worksheet.pageSetup.printArea).toBe('A1:H42');
    expect(worksheet.model.merges).toContain('B29:E29');
    expect(worksheet.model.merges).toContain('G30:H30');
  });
});
