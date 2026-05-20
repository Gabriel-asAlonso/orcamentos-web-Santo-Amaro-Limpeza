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

    // ===== CABEÇALHO - DADOS DA EMPRESA =====
    worksheet.getCell('C1').value = quote.company.name;
    worksheet.getCell('C2').value = quote.company.address ?? '';
    worksheet.getCell('C3').value = quote.company.contact ?? '';
    worksheet.getCell('C4').value = quote.company.email ?? '';

    // ===== DATA DO ORÇAMENTO =====
    worksheet.getCell('F6').value = quoteDate.toLocaleDateString('pt-BR');

    // ===== DADOS DO CLIENTE =====
    if (customer) {
      worksheet.getCell('B6').value = customer.client ?? '';
      worksheet.getCell('B7').value = customer.document ?? '';
      worksheet.getCell('B8').value = customer.address ?? '';
      worksheet.getCell('B9').value = customer.contact ?? '';
      worksheet.getCell('C9').value = customer.phone ?? '';
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    worksheet.getCell('C11').value = quote.paymentMethod ?? '';
    worksheet.getCell('F11').value = quote.deliveryTime ?? '';

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 20;
    const linhaTotalOriginal = this.findTotalRow(worksheet) ?? 27;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    const estiloProduto = Array.from({ length: 6 }, (_, i) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(i + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    this.unmergeFrom(worksheet, linhaInicial);

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

    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      for (let col = 1; col <= 6; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      linha.getCell(1).value = index + 1;
      linha.getCell(2).value = item.quantity;
      linha.getCell(3).value = item.unit ?? '';
      linha.getCell(4).value = item.product ?? '';
      linha.getCell(5).value = item.unitPrice;
      linha.getCell(6).value = totalProduto;

      linha.getCell(2).numFmt = '#,##0.##';
      linha.getCell(5).numFmt = 'R$ #,##0.00';
      linha.getCell(6).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    this.clearFrom(worksheet, linhaTotal);

    const totalValor = worksheet.getCell(`F${linhaTotal}`);
    totalValor.value = total;
    totalValor.numFmt = 'R$ #,##0.00';
    worksheet.getRow(linhaTotal).height = 16;
  }

  private findTotalRow(worksheet: ExcelJS.Worksheet) {
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 6; colNumber++) {
        const value = String(row.getCell(colNumber).value ?? '').trim().toLowerCase();

        if (value.includes('total geral')) {
          return rowNumber;
        }
      }
    }

    return undefined;
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
}
