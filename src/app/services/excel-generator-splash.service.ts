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

    // Ajustar altura das linhas para logo
    worksheet.getRow(1).height = 18;
    worksheet.getRow(2).height = 18;
    worksheet.getRow(3).height = 18;
    worksheet.getRow(4).height = 18;

    // Ajustar largura das colunas
    worksheet.getColumn('A').width = 12;
    worksheet.getColumn('B').width = 15;
    worksheet.getColumn('C').width = 15;
    worksheet.getColumn('D').width = 15;
    worksheet.getColumn('E').width = 15;
    worksheet.getColumn('F').width = 15;

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== CABEÇALHO - DADOS DA EMPRESA (C1:E4) =====
    worksheet.getCell('C1').value = quote.company.name;
    worksheet.getCell('C2').value = quote.company.address ?? '';
    worksheet.getCell('C3').value = quote.company.contact ?? '';
    worksheet.getCell('C4').value = quote.company.email ?? '';

    // ===== DATA DO ORÇAMENTO (F6) =====
    worksheet.getCell('F6').value = quoteDate.toLocaleDateString('pt-BR');

    // ===== DADOS DO CLIENTE (Rows 6-9) =====
    if (customer) {
      // Row 6: NOME - B6:E6
      worksheet.getCell('B6').value = customer.client ?? '';

      // Row 7: CNPJ - B7:E7
      worksheet.getCell('B7').value = customer.document ?? '';

      // Row 8: END. - B8:E8
      worksheet.getCell('B8').value = customer.address ?? '';

      // Row 9: A/C - B9:C9
      worksheet.getCell('B9').value = customer.contact ?? '';

      // Row 9: Fone - D9:E9
      worksheet.getCell('D9').value = customer.phone ?? '';
    }

    // ===== CONDIÇÕES COMERCIAIS (Row 11) =====
    // Forma de pagamento - C11
    worksheet.getCell('C11').value = quote.paymentMethod ?? '';

    // Prazo de entrega - F11
    worksheet.getCell('F11').value = quote.deliveryTime ?? '';

    // ===== PRODUTOS (Rows 15+) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 15;
    const linhaTotal = 56;
    const quantidadeProdutos = items.length;
    const linhasFinal = linhaTotal - 1; // Última linha de produto pode ser no máximo 55
    const quantidadeLinhasDisponivel = linhasFinal - linhaInicial + 1;

    // Ajustar quantidade de linhas
    const linhasAtualmente = quantidadeLinhasDisponivel;
    const diferenca = quantidadeProdutos - linhasAtualmente;

    if (diferenca > 0) {
      // Precisamos adicionar linhas
      worksheet.spliceRows(linhaTotal, 0, ...Array.from({ length: diferenca }, () => []));
    } else if (diferenca < 0) {
      // Precisamos remover linhas
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, -diferenca);
    }

    const linhaModelo = worksheet.getRow(linhaInicial);

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalComDesconto = subtotal - subtotal * (discount / 100);
      return sum + totalComDesconto;
    }, 0);

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = 16;

      // Copiar estilo da linha modelo
      for (let col = 1; col <= 6; col++) {
        const cell = linha.getCell(col);
        const modelCell = linhaModelo.getCell(col);
        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      // Preencher dados - novo layout
      // A: Item (número sequencial)
      linha.getCell(1).value = index + 1;
      // B: Quantidade
      linha.getCell(2).value = item.quantity;
      // C: Unidade
      linha.getCell(3).value = item.unit ?? '';
      // D: Descrição do produto
      linha.getCell(4).value = item.product ?? '';
      // E: Valor unitário
      linha.getCell(5).value = item.unitPrice;
      // F: Valor total
      linha.getCell(6).value = totalProduto;

      // Formatação
      for (let col of [1, 2, 3, 4, 5, 6]) {
        linha.getCell(col).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        linha.getCell(col).font = { size: 9, bold: false };
      }

      // Formatação numérica
      linha.getCell(2).numFmt = '0.00'; // Quantidade
      linha.getCell(5).numFmt = 'R$ #,##0.00'; // Valor unitário
      linha.getCell(6).numFmt = 'R$ #,##0.00'; // Valor total

      linha.commit();
    });

    // Preencher o valor do total na linha 56
    const totalRow = worksheet.getRow(linhaTotal);
    totalRow.getCell(6).value = total;
    totalRow.getCell(6).numFmt = 'R$ #,##0.00';
    totalRow.commit();
  }

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
    target.font = { ...source.font };
    target.alignment = { ...source.alignment };
    target.border = { ...source.border };
    target.fill = { ...source.fill };
    target.numFmt = source.numFmt;
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

      // Posicionar logo em A1:B4
      worksheet.addImage(logoImage, {
        tl: { col: 0, row: 0 } as any,
        br: { col: 2, row: 4 } as any
      });

    } catch {
      // Logo não encontrada, continuar sem ela
    }
  }
}
