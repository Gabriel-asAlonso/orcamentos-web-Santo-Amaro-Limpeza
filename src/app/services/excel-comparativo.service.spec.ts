import ExcelJS from 'exceljs';
import { ExcelComparativoService, type ComparativeQuote } from './excel-comparativo.service';

describe('ExcelComparativoService', () => {
  function createTemplateWorkbook(): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Comparativo');
    const border: Partial<ExcelJS.Borders> = {
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
    };

    worksheet.views = [{ showGridLines: false }];
    worksheet.pageSetup = {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
    };

    worksheet.getColumn(1).width = 13.69;
    worksheet.getColumn(2).width = 67.38;
    worksheet.getColumn(3).width = 14.75;

    for (let column = 4; column <= 8; column++) {
      worksheet.getColumn(column).width = 18.38;
      worksheet.getCell(4, column).value = `Modelo ${column}`;
      worksheet.getCell(5, column).value = 'V. Total';
    }

    for (let row = 1; row <= 39; row++) {
      for (let column = 1; column <= 8; column++) {
        worksheet.getCell(row, column).border = border;
      }
    }

    worksheet.getCell('A1').value = 'Planilha - Mapa Comparativo';
    worksheet.getCell('B3').value = 'Cliente';
    worksheet.getCell('D3').value = 'Produtos de Limpeza';
    worksheet.getCell('C5').value = 'Quantidade';
    worksheet.getCell('A39').value = 'Fim';

    worksheet.mergeCells('A1:H2');
    worksheet.mergeCells('A4:A5');
    worksheet.mergeCells('B4:B5');
    worksheet.mergeCells('A29:C29');
    worksheet.mergeCells('A30:C30');
    worksheet.mergeCells('A31:C31');
    worksheet.mergeCells('A32:C32');
    worksheet.mergeCells('A34:H34');
    worksheet.mergeCells('A38:B39');
    worksheet.mergeCells('C38:H39');

    return workbook;
  }

  function createQuotes(count: number): ComparativeQuote[] {
    return Array.from({ length: count }, (_, index) => {
      const supplierNumber = index + 1;
      const firstUnitPrice = 10 + supplierNumber;
      const secondUnitPrice = 20 + supplierNumber;

      return {
        label: `Orçamento ${supplierNumber}`,
        company: {
          id: `supplier-${supplierNumber}`,
          name: `Fornecedor ${supplierNumber}`,
        },
        customerData: { client: 'Condomínio Exemplo' },
        paymentTerm: `Pagamento ${supplierNumber}`,
        deliveryTerm: `${supplierNumber} dias`,
        items: [
          {
            id: '1',
            product: 'Produto 1',
            quantity: 2,
            unitPrice: firstUnitPrice,
            total: firstUnitPrice * 2,
          },
          {
            id: '2',
            product: 'Produto 2',
            quantity: 3,
            unitPrice: secondUnitPrice,
            total: secondUnitPrice * 3,
          },
        ],
      };
    });
  }

  it.each([1, 2, 3])(
    'mantém somente %i seção(ões) visível(is) de valor unitário',
    (supplierCount) => {
      const service = new ExcelComparativoService();
      const workbook = createTemplateWorkbook();

      service.populateWorkbook(workbook, createQuotes(supplierCount));

      const worksheet = workbook.worksheets[0];
      const lastSupplierColumn = 3 + supplierCount;

      expect(workbook.worksheets).toHaveLength(1);
      expect(worksheet.getCell('D4').value).toBe('Fornecedor 1');
      expect(worksheet.getCell('D5').value).toBe('V. Unitário');
      expect(worksheet.getCell('D6').value).toBe(11);
      expect(worksheet.getCell('D9').value).toBe(85);

      for (let column = 4; column <= lastSupplierColumn; column++) {
        expect(worksheet.getColumn(column).hidden).toBe(false);
      }

      for (let column = lastSupplierColumn + 1; column <= 8; column++) {
        expect(worksheet.getColumn(column).hidden).toBe(true);
      }

      expect(worksheet.pageSetup.printArea).toBe(
        `A1:${worksheet.getColumn(lastSupplierColumn).letter}19`,
      );
    },
  );

  it('pagina fornecedores adicionais sem limitar a quantidade total', () => {
    const service = new ExcelComparativoService();
    const workbook = createTemplateWorkbook();

    service.populateWorkbook(workbook, createQuotes(7));

    const firstPage = workbook.worksheets[0];
    const secondPage = workbook.worksheets[1];

    expect(workbook.worksheets).toHaveLength(2);
    expect(firstPage.name).toBe('Comparativo 1');
    expect(secondPage.name).toBe('Comparativo 2');
    expect(firstPage.getCell('H4').value).toBe('Fornecedor 5');
    expect(secondPage.getCell('D4').value).toBe('Fornecedor 6');
    expect(secondPage.getCell('E4').value).toBe('Fornecedor 7');
    expect(secondPage.getCell('D5').value).toBe('V. Unitário');
    expect(secondPage.getCell('D6').value).toBe(16);
    expect(secondPage.getColumn('D').hidden).toBe(false);
    expect(secondPage.getColumn('E').hidden).toBe(false);
    expect(secondPage.getColumn('F').hidden).toBe(true);
    expect(firstPage.pageSetup.printArea).toBe('A1:H19');
    expect(secondPage.pageSetup.printArea).toBe('A1:E19');
  });

  it('inclui itens existentes apenas em orçamentos posteriores', () => {
    const service = new ExcelComparativoService();
    const workbook = createTemplateWorkbook();
    const quotes = createQuotes(2);

    quotes[1].items.push({
      id: '3',
      product: 'Produto exclusivo',
      quantity: 1,
      unitPrice: 99,
      total: 99,
    });

    service.populateWorkbook(workbook, quotes);

    const worksheet = workbook.worksheets[0];

    expect(worksheet.getCell('B8').value).toBe('Produto exclusivo');
    expect(worksheet.getCell('D8').value).toBe('');
    expect(worksheet.getCell('E8').value).toBe(99);
    expect(worksheet.getCell('D10').value).toBe(85);
    expect(worksheet.getCell('E10').value).toBe(189);
  });
});
