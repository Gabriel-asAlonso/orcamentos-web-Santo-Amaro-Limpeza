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
    worksheet.views = [{ showGridLines: false }];

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

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

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
    const linhaTotalOriginal = 56;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos ANTES de qualquer alteração
    const estiloProduto = Array.from({ length: 6 }, (_, i) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(i + 1))
    );
    const estiloTotalValor = this.snapshotStyle(worksheet.getCell(`F${linhaTotalOriginal}`));
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Remover todos os merges da região dinâmica
    this.unmergeFrom(worksheet, linhaInicial);

    // Ajustar a quantidade de linhas
    if (linhasParaAdicionar > 0) {
      worksheet.spliceRows(
        linhaTotalOriginal,
        0,
        ...Array.from({ length: linhasParaAdicionar }, () => [])
      );
    }

    if (linhasParaRemover > 0) {
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, linhasParaRemover);
    }

    const linhaTotal = linhaInicial + quantidadeProdutos;

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      // Aplicar o estilo capturado
      for (let col = 1; col <= 6; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      // Preencher dados
      linha.getCell(1).value = index + 1; // A - Item
      linha.getCell(2).value = item.quantity; // B - Quantidade
      linha.getCell(3).value = item.unit ?? ''; // C - Unidade
      linha.getCell(4).value = item.product ?? ''; // D - Descrição
      linha.getCell(5).value = item.unitPrice; // E - Valor unitário
      linha.getCell(6).value = totalProduto; // F - Valor total

      // Formatação numérica
      linha.getCell(2).numFmt = '#,##0.##';
      linha.getCell(5).numFmt = 'R$ #,##0.00';
      linha.getCell(6).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Limpar total + rodapé
    this.clearFrom(worksheet, linhaTotal);

    // Reconstruir o total
    const totalValor = worksheet.getCell(`F${linhaTotal}`);
    totalValor.value = total;
    this.applyStyle(totalValor, estiloTotalValor);
    totalValor.numFmt = 'R$ #,##0.00';

    worksheet.getRow(linhaTotal).height = 16;
  }

  private snapshotStyle(cell: ExcelJS.Cell) {
    return {
      font: cell.font ? { ...cell.font } : undefined,
      alignment: cell.alignment ? { ...cell.alignment } : undefined,
      border: cell.border ? { ...cell.border } : undefined,
      fill: cell.fill ? { ...cell.fill } : undefined,
      numFmt: cell.numFmt
    };
  }

  private applyStyle(
    cell: ExcelJS.Cell,
    style: ReturnType<ExcelGeneratorSplashService['snapshotStyle']>
  ) {
    if (style.font) cell.font = { ...style.font };
    if (style.alignment) cell.alignment = { ...style.alignment };
    if (style.border) cell.border = { ...style.border };
    if (style.fill) cell.fill = { ...style.fill };
    if (style.numFmt) cell.numFmt = style.numFmt;
  }

  private clearFrom(worksheet: ExcelJS.Worksheet, startRow: number) {
    const ultimaLinha = Math.max(worksheet.rowCount, startRow) + 50;

    for (let rowNumber = startRow; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 6; colNumber++) {
        const cell = row.getCell(colNumber);
        cell.value = null;
        cell.style = {};
      }
    }
  }

  private unmergeFrom(worksheet: ExcelJS.Worksheet, startRow: number) {
    const merges: string[] = [...(((worksheet as any).model?.merges as string[]) ?? [])];

    for (const range of merges) {
      const match = /^[A-Z]+(\d+):[A-Z]+\d+$/.exec(range);

      if (match && Number(match[1]) >= startRow) {
        try {
          worksheet.unMergeCells(range);
        } catch {}
      }
    }
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
