import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorService {
  private templatePath = 'assets/templates/cotacao-template.xlsx';
  private fallbackLogoPath = '/logoSantoAmaroLimpeza.jpg';

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

    this.prepareWorksheetColumns(worksheet);
    this.clearWorksheetImages(worksheet);

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const companyDocument = quote.company.cnpj || quote.company.taxId || '';

    worksheet.getCell('D4').value = new Date().toLocaleDateString('pt-BR');
    worksheet.getCell('D5').value = companyDocument;

    worksheet.getCell('C7').value = quote.company.name;
    worksheet.getCell('C8').value = quote.company.name;
    worksheet.getCell('C9').value = quote.company.address;
    worksheet.getCell('C11').value = companyDocument;

    worksheet.getCell('C12').value = '';
    worksheet.getCell('N12').value = quote.paymentMethod;
    worksheet.getCell('C14').value = quote.company.identity ?? '';

    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 19;
    const quantidadeLinhasModelo = 4;
    const linhaEmBrancoOriginal = 23;
    const linhaTotalOriginal = 24;
    const quantidadeProdutos = items.length;
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    const linhaModelo = worksheet.getRow(linhaInicial);
    const linhaTotalModelo = worksheet.getRow(linhaTotalOriginal);
    const alturaProduto = 16;
    const alturaTotal = Math.max(linhaTotalModelo.height ?? 24, 26);

    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalComDesconto = subtotal - subtotal * (discount / 100);

      return sum + totalComDesconto;
    }, 0);

    if (linhasParaAdicionar > 0) {
      worksheet.spliceRows(
        linhaEmBrancoOriginal,
        0,
        ...Array.from({ length: linhasParaAdicionar }, () => [])
      );
    }

    if (linhasParaRemover > 0) {
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, linhasParaRemover);
    }

    const linhaEmBranco = linhaInicial + quantidadeProdutos;
    const linhaTotal = linhaEmBranco + 1;

    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      for (let colNumber = 1; colNumber <= 16; colNumber++) {
        const cell = linha.getCell(colNumber);
        const modelCell = linhaModelo.getCell(colNumber);

        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      this.applyProductRowLayout(worksheet, rowNumber);

      linha.getCell(1).value = item.code ?? index + 1;
      linha.getCell(2).value = item.reference ?? '';
      linha.getCell(4).value = item.product;
      linha.getCell(12).value = item.unit;
      linha.getCell(13).value = discount / 100;
      linha.getCell(14).value = item.quantity;
      linha.getCell(15).value = item.unitPrice;
      linha.getCell(16).value = totalProduto;

      linha.getCell(13).numFmt = '0.00%';
      linha.getCell(15).numFmt = 'R$ #,##0.00';
      linha.getCell(16).numFmt = 'R$ #,##0.00';

      linha.getCell(1).font = { size: 9 };
      linha.getCell(2).font = { size: 9 };
      linha.getCell(4).font = { size: 9 };
      linha.getCell(12).font = { size: 9 };
      linha.getCell(13).font = { size: 9 };
      linha.getCell(14).font = { size: 9 };
      linha.getCell(15).font = { size: 9 };
      linha.getCell(16).font = { size: 9 };

      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      linha.getCell(4).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(12).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(13).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(14).alignment = { horizontal: 'center', vertical: 'middle' };
      linha.getCell(15).alignment = { horizontal: 'right', vertical: 'middle' };
      linha.getCell(16).alignment = { horizontal: 'right', vertical: 'middle' };

      linha.commit();
    });

    const blankRow = worksheet.getRow(linhaEmBranco);

    blankRow.height = 8;

    for (let colNumber = 1; colNumber <= 16; colNumber++) {
      const cell = blankRow.getCell(colNumber);

      cell.value = null;
      cell.border = undefined as any;
      cell.font = undefined as any;
      cell.fill = undefined as any;
      cell.numFmt = '';
      cell.alignment = undefined as any;
    }

    blankRow.commit();

    const totalRow = worksheet.getRow(linhaTotal);

    totalRow.height = alturaTotal;

    for (let colNumber = 1; colNumber <= 16; colNumber++) {
      const cell = totalRow.getCell(colNumber);
      const modelCell = linhaTotalModelo.getCell(colNumber);

      cell.value = null;
      this.copyCellStyle(modelCell, cell);
      cell.border = undefined as any;
      cell.fill = undefined as any;
    }

    this.applyTotalRowLayout(worksheet, linhaTotal);

    worksheet.getCell(`A${linhaTotal}`).value = 'Total Produtos:';
    worksheet.getCell(`P${linhaTotal}`).value = total;

    worksheet.getCell(`A${linhaTotal}`).font = { bold: true, size: 11 };
    worksheet.getCell(`P${linhaTotal}`).font = { bold: true, size: 11 };
    worksheet.getCell(`P${linhaTotal}`).numFmt = 'R$ #,##0.00';

    worksheet.getCell(`A${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    worksheet.getCell(`P${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    totalRow.commit();

    this.removeRowsAfterTotal(worksheet, linhaTotal);
  }

  private prepareWorksheetColumns(worksheet: ExcelJS.Worksheet) {
    worksheet.getColumn(1).width = 7;
    worksheet.getColumn(2).width = 8;
    worksheet.getColumn(3).width = 8;

    worksheet.getColumn(4).width = 7;
    worksheet.getColumn(5).width = 7;
    worksheet.getColumn(6).width = 7;
    worksheet.getColumn(7).width = 7;
    worksheet.getColumn(8).width = 7;
    worksheet.getColumn(9).width = 7;
    worksheet.getColumn(10).width = 7;
    worksheet.getColumn(11).width = 7;

    worksheet.getColumn(12).width = 8;
    worksheet.getColumn(13).width = 7;
    worksheet.getColumn(14).width = 7;
    worksheet.getColumn(15).width = 12;
    worksheet.getColumn(16).width = 13;

    worksheet.pageSetup = {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: {
        left: 0.2,
        right: 0.2,
        top: 0.25,
        bottom: 0.25,
        header: 0.1,
        footer: 0.1
      }
    };

    worksheet.views = [
      {
        showGridLines: false
      }
    ];
  }

  private applyProductRowLayout(worksheet: ExcelJS.Worksheet, rowNumber: number) {
    this.safeMergeCells(worksheet, `B${rowNumber}:C${rowNumber}`);
    this.safeMergeCells(worksheet, `D${rowNumber}:K${rowNumber}`);

    const row = worksheet.getRow(rowNumber);

    for (let colNumber = 1; colNumber <= 16; colNumber++) {
      row.getCell(colNumber).border = undefined as any;
    }

    this.setOuterBorder(row.getCell(1), true, true, true, true);
    this.setOuterBorder(row.getCell(2), true, true, true, false);
    this.setOuterBorder(row.getCell(3), true, false, true, true);
    this.setOuterBorder(row.getCell(4), true, true, true, false);

    for (let colNumber = 5; colNumber <= 10; colNumber++) {
      this.setOuterBorder(row.getCell(colNumber), true, false, true, false);
    }

    this.setOuterBorder(row.getCell(11), true, false, true, true);
    this.setOuterBorder(row.getCell(12), true, true, true, true);
    this.setOuterBorder(row.getCell(13), true, true, true, true);
    this.setOuterBorder(row.getCell(14), true, true, true, true);
    this.setOuterBorder(row.getCell(15), true, true, true, true);
    this.setOuterBorder(row.getCell(16), true, true, true, true);
  }

  private applyTotalRowLayout(worksheet: ExcelJS.Worksheet, rowNumber: number) {
    this.safeMergeCells(worksheet, `A${rowNumber}:O${rowNumber}`);

    const row = worksheet.getRow(rowNumber);

    for (let colNumber = 1; colNumber <= 16; colNumber++) {
      const cell = row.getCell(colNumber);

      cell.border = undefined as any;
      cell.fill = undefined as any;
    }

    row.getCell(1).border = {
      top: { style: 'thin' }
    };

    row.getCell(16).border = {
      top: { style: 'thin' }
    };
  }

  private removeRowsAfterTotal(worksheet: ExcelJS.Worksheet, linhaTotal: number) {
    const startRow = linhaTotal + 1;
    const finalRow = Math.max(worksheet.rowCount + 80, startRow + 120);

    for (let rowNumber = startRow; rowNumber <= finalRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 20; colNumber++) {
        const cell = row.getCell(colNumber);

        cell.value = null;
        cell.border = undefined as any;
        cell.font = undefined as any;
        cell.fill = undefined as any;
        cell.numFmt = '';
        cell.alignment = undefined as any;
      }

      row.height = 0;
      row.hidden = true;
      row.commit();
    }
  }

  private clearWorksheetImages(worksheet: ExcelJS.Worksheet) {
    const sheet = worksheet as any;

    if (sheet._images?.length) {
      sheet._images = [];
    }

    if (sheet._media?.length) {
      sheet._media = [];
    }

    if (sheet._drawing) {
      sheet._drawing = undefined;
    }
  }

  private setOuterBorder(
    cell: ExcelJS.Cell,
    top: boolean,
    left: boolean,
    bottom: boolean,
    right: boolean
  ) {
    const border: Partial<ExcelJS.Borders> = {};

    if (top) border.top = { style: 'thin' };
    if (left) border.left = { style: 'thin' };
    if (bottom) border.bottom = { style: 'thin' };
    if (right) border.right = { style: 'thin' };

    cell.border = border;
  }

  private safeMergeCells(worksheet: ExcelJS.Worksheet, range: string) {
    try {
      worksheet.unMergeCells(range);
    } catch {}

    try {
      worksheet.mergeCells(range);
    } catch {}
  }

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
    target.style = { ...source.style };

    if (source.numFmt) target.numFmt = source.numFmt;
    if (source.alignment) target.alignment = { ...source.alignment };
    if (source.border) target.border = { ...source.border };
    if (source.font) target.font = { ...source.font };
    if (source.fill) target.fill = { ...source.fill };
  }

  private async insertCompanyLogo(
    workbook: ExcelJS.Workbook,
    worksheet: ExcelJS.Worksheet,
    logoPath: string
  ) {
    let resolvedLogoPath = logoPath;
    let logoResponse = await fetch(resolvedLogoPath);

    if (!logoResponse.ok && resolvedLogoPath !== this.fallbackLogoPath) {
      resolvedLogoPath = this.fallbackLogoPath;
      logoResponse = await fetch(resolvedLogoPath);
    }

    if (!logoResponse.ok) return;

    const logoBlob = await logoResponse.blob();
    const logoBase64 = await this.blobToBase64(logoBlob);

    const logoId = workbook.addImage({
      base64: logoBase64,
      extension: this.getImageExtension(resolvedLogoPath)
    });

worksheet.addImage(logoId, {
  tl: { col: 0.15, row: 1.2 } as any,
  br: { col: 2.15, row: 5.85 } as any
});
  }

  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise(resolve => {
      const reader = new FileReader();

      reader.onloadend = () => resolve(reader.result as string);
      reader.readAsDataURL(blob);
    });
  }

  private getImageExtension(path: string): 'png' | 'jpeg' {
    return path.toLowerCase().endsWith('.jpg') || path.toLowerCase().endsWith('.jpeg') ? 'jpeg' : 'png';
  }
}