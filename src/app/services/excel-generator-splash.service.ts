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

    this.clearWorksheetRowGrouping(worksheet);
    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // Data (canto superior direito)
    worksheet.getCell('K2').value = `Data: ${quoteDate.toLocaleDateString('pt-BR')}`;
    worksheet.getCell('K2').font = { bold: true, size: 10 };
    worksheet.getCell('K2').alignment = { horizontal: 'right', vertical: 'middle' };

    // Prazo de entrega (row 4, entre colunas C-D)
    worksheet.getCell('C4').value = 'Prazo:';
    worksheet.getCell('D4').value = `ATÉ ${quote.deliveryTime.toUpperCase()}`;

    // Forma de pagamento (row 5, entre colunas C-D)
    worksheet.getCell('C5').value = 'Pagamento:';
    worksheet.getCell('D5').value = quote.paymentMethod.toUpperCase();

    // ===== DADOS DO CLIENTE (Row 7+) =====
    if (customer) {
      // Row 7: Cliente
      worksheet.getCell('A7').value = 'Cliente:';
      worksheet.getCell('B7').value = customer.client ?? '';
      worksheet.getCell('G7').value = 'Fantasia:';

      // Row 8: Fantasia
      worksheet.getCell('A8').value = 'Fantasia:';
      worksheet.getCell('B8').value = customer.fantasyName ?? customer.client ?? '';
      worksheet.getCell('G8').value = 'Bairro:';
      worksheet.getCell('I8').value = customer.district ?? '';

      // Row 9: Endereço
      worksheet.getCell('A9').value = 'Endereço:';
      worksheet.getCell('B9').value = customer.address ?? '';
      worksheet.getCell('G9').value = 'UF / CEP:';
      worksheet.getCell('I9').value = `${customer.state ?? ''} / ${customer.postalCode ?? ''}`;

      // Row 10: Bairro (segunda menção)
      worksheet.getCell('A10').value = 'Bairro:';
      worksheet.getCell('B10').value = customer.district ?? '';
      worksheet.getCell('G10').value = 'Insc. Est.:';
      worksheet.getCell('I10').value = customer.stateRegistration ?? '';

      // Row 11: Cidade
      worksheet.getCell('A11').value = 'Cidade:';
      worksheet.getCell('B11').value = customer.city ?? '';
      worksheet.getCell('G11').value = 'Tel e fone / WhatsApp:';
      worksheet.getCell('I11').value = customer.phone ?? '';

      // Row 12: CEP
      worksheet.getCell('A12').value = 'CEP:';
      worksheet.getCell('B12').value = customer.postalCode ?? '';
      worksheet.getCell('G12').value = 'Região:';
      worksheet.getCell('I12').value = customer.region ?? '';

      // Row 13: CNPJ/CPF
      worksheet.getCell('A13').value = 'CNPJ/CPF:';
      worksheet.getCell('B13').value = customer.document ?? '';
      worksheet.getCell('G13').value = 'Insc. Est.:';
      worksheet.getCell('I13').value = customer.stateRegistration ?? '';

      // Row 14: Fone
      worksheet.getCell('A14').value = 'Fone:';
      worksheet.getCell('B14').value = customer.phone ?? '';
      worksheet.getCell('G14').value = 'WhatsApp:';
      worksheet.getCell('I14').value = customer.whatsapp ?? '';

      // Row 15: CONDIÇÕES DE ENTREGA (header)
      // Esta é uma seção de header que deve estar no template

      // Row 16: Email e Forma de pagamento
      worksheet.getCell('A16').value = 'Email:';
      worksheet.getCell('B16').value = customer.email ?? '';
      worksheet.getCell('G16').value = 'Forma de pagamento:';
      worksheet.getCell('I16').value = quote.paymentMethod;

      // Row 17: Região e Tabela de preço
      worksheet.getCell('A17').value = 'Região:';
      worksheet.getCell('B17').value = customer.region ?? '';
      worksheet.getCell('G17').value = 'Tabela Preço:';
      worksheet.getCell('I17').value = customer.priceTable ?? '';

      // Row 18: Vendedor
      worksheet.getCell('A18').value = 'Vendedor:';
      worksheet.getCell('B18').value = customer.salesperson ?? '';
      worksheet.getCell('G18').value = 'Ven. Externo:';
      worksheet.getCell('I18').value = customer.externalSalesperson ?? '';
    }

    // ===== PRODUTOS (Row 20+) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 22;
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

      // Colunas conforme template: A, B, C, D, E, F, G, H
      linha.getCell(1).value = item.code ?? index + 1; // A - Código
      linha.getCell(2).value = item.reference ?? ''; // B - Referência
      linha.getCell(3).value = item.product; // C - Descrição/Produto
      linha.getCell(4).value = item.quantity; // D - Qtd
      linha.getCell(5).value = item.unit; // E - Unidade
      linha.getCell(6).value = item.unitPrice; // F - Unitário
      linha.getCell(7).value = discount / 100; // G - Desconto %
      linha.getCell(8).value = totalProduto; // H - Total

      // Formatação
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(3).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
      linha.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };

      // Fontes
      for (let col = 1; col <= 8; col++) {
        linha.getCell(col).font = { size: 9 };
      }

      // Formatação numérica
      linha.getCell(6).numFmt = 'R$ #,##0.00';
      linha.getCell(7).numFmt = '0.00%';
      linha.getCell(8).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Linha de total
    const linhaTotal = linhaInicial + quantidadeProdutos + 1;
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

      // Logo será inserida onde o template indicar
      // O template já tem placeholder para logo
    } catch {
      // Logo não encontrada, continuar sem ela
    }
  }
}
