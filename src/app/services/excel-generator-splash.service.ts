import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorSplashService {
  private templatePath = 'assets/templates/cotacao-template-Splash.xlsx';

  async generateQuoteExcel(quote: Quote) {
    const workbook = new ExcelJS.Workbook();

    const templateResponse = await fetch(this.templatePath);

    if (!templateResponse.ok) {
      throw new Error(`Modelo de Excel não encontrado: ${this.templatePath}`);
    }

    const templateBuffer = await templateResponse.arrayBuffer();
    await workbook.xlsx.load(templateBuffer);

    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('Aba da planilha não encontrada');
    }

    this.prepareWorksheet(worksheet);
    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const companyAddress = quote.company.address || '';
    const customer = quote.customerData;

    // Logo position: A4:B5 (as specified)
    // Company name: A2:F3 (as specified)
    this.safeMergeCells(worksheet, 'A2:F3');
    worksheet.getCell('A2').value = quote.company.name;
    worksheet.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A2').font = { bold: true, size: 11 };

    // Aplicar dados do cliente
    this.fillCustomerData(worksheet, customer, companyDocument, companyAddress, quote);
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private prepareWorksheet(worksheet: ExcelJS.Worksheet) {
    // Limpar agrupamento e imagens
    this.clearWorksheetRowGrouping(worksheet);
    this.clearWorksheetImages(worksheet);

    // Ajustar largura das colunas
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 14;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 12;
    worksheet.getColumn(5).width = 12;
    worksheet.getColumn(6).width = 12;
    worksheet.getColumn(7).width = 14;
    worksheet.getColumn(8).width = 10;
    worksheet.getColumn(9).width = 10;
    worksheet.getColumn(10).width = 12;
    worksheet.getColumn(11).width = 12;
    worksheet.getColumn(12).width = 14;
  }

  private fillCustomerData(
    worksheet: ExcelJS.Worksheet,
    customer: any,
    companyDocument: string,
    companyAddress: string,
    quote: Quote
  ) {
    // Os campos devem ser preenchidos conforme nomeados no modelo
    // Ajuste as células conforme necessário

    if (customer) {
      worksheet.getCell('A7').value = 'Cliente:';
      worksheet.getCell('B7').value = customer.client ?? '';

      worksheet.getCell('A8').value = 'Fantasia:';
      worksheet.getCell('B8').value = customer.fantasyName ?? customer.client ?? '';

      worksheet.getCell('A9').value = 'Endereço:';
      worksheet.getCell('B9').value = customer.address ?? '';

      worksheet.getCell('A10').value = 'Bairro:';
      worksheet.getCell('B10').value = customer.district ?? '';

      worksheet.getCell('A11').value = 'Cidade:';
      worksheet.getCell('B11').value = customer.city ?? '';

      worksheet.getCell('E11').value = 'UF:';
      worksheet.getCell('F11').value = customer.state ?? '';

      worksheet.getCell('A12').value = 'CEP:';
      worksheet.getCell('B12').value = customer.postalCode ?? '';

      worksheet.getCell('A13').value = 'CNPJ/CPF:';
      worksheet.getCell('B13').value = customer.document ?? '';

      worksheet.getCell('E13').value = 'Insc. Est.:';
      worksheet.getCell('F13').value = customer.stateRegistration ?? '';

      worksheet.getCell('A14').value = 'Fone:';
      worksheet.getCell('B14').value = customer.phone ?? '';

      worksheet.getCell('E14').value = 'WhatsApp:';
      worksheet.getCell('F14').value = customer.whatsapp ?? '';

      worksheet.getCell('A15').value = 'Contato:';
      worksheet.getCell('B15').value = customer.contact ?? '';

      worksheet.getCell('A16').value = 'Email:';
      worksheet.getCell('B16').value = customer.email ?? '';

      worksheet.getCell('A17').value = 'Região:';
      worksheet.getCell('B17').value = customer.region ?? '';

      worksheet.getCell('E17').value = 'Tabela Preço:';
      worksheet.getCell('F17').value = customer.priceTable ?? '';

      worksheet.getCell('A18').value = 'Vendedor:';
      worksheet.getCell('B18').value = customer.salesperson ?? '';

      worksheet.getCell('E18').value = 'Ven. Externo:';
      worksheet.getCell('F18').value = customer.externalSalesperson ?? '';
    }

    // Data do orçamento
    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    const dateCell = worksheet.getCell('G2');
    dateCell.value = `Data: ${quoteDate.toLocaleDateString('pt-BR')}`;
    dateCell.alignment = { horizontal: 'right', vertical: 'middle' };
    dateCell.font = { bold: true, size: 10 };

    // Prazo de entrega
    const deliveryCell = worksheet.getCell('G4');
    deliveryCell.value = `Prazo: ${quote.deliveryTime}`;
    deliveryCell.alignment = { horizontal: 'right', vertical: 'middle' };
    deliveryCell.font = { bold: false, size: 9 };

    // Forma de pagamento
    const paymentCell = worksheet.getCell('G5');
    paymentCell.value = `Pagamento: ${quote.paymentMethod}`;
    paymentCell.alignment = { horizontal: 'right', vertical: 'middle' };
    paymentCell.font = { bold: false, size: 9 };
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 20;
    const quantidadeProdutos = items.length;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalComDesconto = subtotal - subtotal * (discount / 100);
      return sum + totalComDesconto;
    }, 0);

    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = 20;

      // Preencher células
      linha.getCell(1).value = item.code ?? index + 1;
      linha.getCell(2).value = item.reference ?? '';
      linha.getCell(3).value = item.product;
      linha.getCell(4).value = item.quantity;
      linha.getCell(5).value = item.unit;
      linha.getCell(6).value = item.unitPrice;
      linha.getCell(7).value = discount / 100;
      linha.getCell(8).value = totalProduto;

      // Formatação
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(3).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
      linha.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };

      linha.getCell(1).font = { size: 9 };
      linha.getCell(2).font = { size: 9 };
      linha.getCell(3).font = { size: 9 };
      linha.getCell(4).font = { size: 9 };
      linha.getCell(5).font = { size: 9 };
      linha.getCell(6).font = { size: 9 };
      linha.getCell(7).font = { size: 9 };
      linha.getCell(8).font = { size: 9 };

      // Formatação numérica
      linha.getCell(6).numFmt = 'R$ #,##0.00';
      linha.getCell(7).numFmt = '0.00%';
      linha.getCell(8).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Linha de total
    const linhaTotal = linhaInicial + quantidadeProdutos;
    const totalRow = worksheet.getRow(linhaTotal);
    totalRow.height = 24;

    totalRow.getCell(1).value = 'TOTAL:';
    totalRow.getCell(1).font = { bold: true, size: 11 };
    totalRow.getCell(1).alignment = { horizontal: 'right', vertical: 'middle' };

    totalRow.getCell(8).value = total;
    totalRow.getCell(8).font = { bold: true, size: 11 };
    totalRow.getCell(8).numFmt = 'R$ #,##0.00';
    totalRow.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };

    totalRow.commit();
  }

  private clearWorksheetImages(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).drawing = null;
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;
  }

  private async insertCompanyLogo(
    workbook: ExcelJS.Workbook,
    worksheet: ExcelJS.Worksheet,
    logoPath: string
  ) {
    try {
      const logoResponse = await fetch(logoPath);
      if (!logoResponse.ok) throw new Error('Logo not found');

      const logoBuffer = await logoResponse.arrayBuffer();
      const logoImage = workbook.addImage({
        buffer: logoBuffer,
        extension: 'png'
      });

      // Logo posicionada em A4:B5
      worksheet.addImage(logoImage, 'A4:B5');
    } catch {
      // Logo não encontrada, continuar sem ela
    }
  }

  private safeMergeCells(worksheet: ExcelJS.Worksheet, range: string) {
    try {
      worksheet.mergeCells(range);
    } catch {
      // Célula já foi mesclada ou range inválido
    }
  }
}
