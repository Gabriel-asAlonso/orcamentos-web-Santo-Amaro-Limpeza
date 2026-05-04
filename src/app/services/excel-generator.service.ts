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

await this.insertHeaderBorder(workbook, worksheet);
await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const customer = quote.customerData;

    worksheet.getCell('A2').value = '';
    worksheet.getCell('A3').value = quote.company.name;
    worksheet.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A3').font = { bold: true, size: 11 };
    worksheet.getCell('C4').value = '';
    worksheet.getCell('D4').value = '';
    this.safeMergeCells(worksheet, 'M4:O4');
    worksheet.getCell('M4').value = `Data: ${new Date().toLocaleDateString('pt-BR')}`;
    worksheet.getCell('M4').alignment = { horizontal: 'right', vertical: 'middle' };
    worksheet.getCell('M4').font = { bold: true, size: 10 };
    worksheet.getCell('N4').font = { bold: true, size: 10 };
    worksheet.getCell('D5').value = companyDocument;

    worksheet.getCell('C7').value = customer?.client ?? '';
    worksheet.getCell('C8').value = customer?.fantasyName ?? customer?.client ?? '';
    worksheet.getCell('C9').value = customer?.address ?? '';
    worksheet.getCell('N9').value = customer?.district ?? '';
    worksheet.getCell('C10').value = customer?.city ?? '';
    worksheet.getCell('M10').value = customer?.state ?? '';
    worksheet.getCell('O10').value = customer?.postalCode ?? '';
    worksheet.getCell('C11').value = customer?.document ?? '';
    worksheet.getCell('H11').value = customer?.stateRegistration ?? '';
    worksheet.getCell('N11').value = customer?.region ?? '';
    worksheet.getCell('C12').value = customer?.freight ?? '';
    worksheet.getCell('L12').value = 'Forma Pagto. :';
    worksheet.getCell('N12').value = quote.paymentMethod;
    worksheet.getCell('C13').value = customer?.phone ?? '';
    worksheet.getCell('H13').value = customer?.whatsapp ?? '';
    worksheet.getCell('N13').value = customer?.priceTable ?? '';
    worksheet.getCell('C14').value = customer?.contact ?? '';
    worksheet.getCell('C15').value = customer?.email ?? '';
    worksheet.getCell('C16').value = customer?.salesperson ?? '';
    worksheet.getCell('C17').value = customer?.externalSalesperson ?? '';

    this.fillProducts(worksheet, quote.items, quote.deliveryTime);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[], deliveryTime: string) {
    const linhaInicial = 19;
    const linhaTotalOriginal = this.findTotalProductsRow(worksheet) ?? 25;
    const quantidadeProdutos = items.length;
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
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

    const totalRow = worksheet.getRow(linhaTotal);

    totalRow.height = alturaTotal;
    totalRow.hidden = false;
    totalRow.outlineLevel = 0;

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

    const lastVisibleRow = deliveryTime ? linhaTotal + 1 : linhaTotal;

    if (deliveryTime) {
      this.applyDeliveryTimeRow(worksheet, lastVisibleRow, deliveryTime);
    }

    this.removeRowsAfterTotal(worksheet, lastVisibleRow);
  }

  private prepareWorksheetColumns(worksheet: ExcelJS.Worksheet) {
    worksheet.getColumn(1).width = 12;
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

    if (worksheet.rowCount >= startRow) {
      worksheet.spliceRows(startRow, worksheet.rowCount - linhaTotal);
    }
  }

  private applyDeliveryTimeRow(worksheet: ExcelJS.Worksheet, rowNumber: number, deliveryTime: string) {
    this.safeMergeCells(worksheet, `D${rowNumber}:P${rowNumber}`);

    const row = worksheet.getRow(rowNumber);
    const labelCell = row.getCell(1);
    const cell = row.getCell(4);

    row.height = 22;
    row.hidden = false;
    row.outlineLevel = 0;
    labelCell.value = 'Prazo de Entrega:';
    cell.value = deliveryTime;
    labelCell.font = { bold: true, size: 10 };
    cell.font = { bold: true, size: 10 };
    labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    labelCell.border = {
      top: { style: 'thin' }
    };
    cell.border = {
      top: { style: 'thin' }
    };

    for (let colNumber = 2; colNumber <= 16; colNumber++) {
      const rowCell = row.getCell(colNumber);

      if (colNumber === 4) continue;

      rowCell.value = null;
      rowCell.border = undefined as any;
      rowCell.fill = undefined as any;
    }

    row.commit();
  }

  private findTotalProductsRow(worksheet: ExcelJS.Worksheet) {
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 16; colNumber++) {
        const value = String(row.getCell(colNumber).value ?? '').trim().toLowerCase();

        if (value.includes('total produtos')) {
          return rowNumber;
        }
      }
    }

    return undefined;
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

  private async insertHeaderBorder(workbook: ExcelJS.Workbook, worksheet: ExcelJS.Worksheet) {
  const borderResponse = await fetch('assets/templates/header-border.png');

  if (!borderResponse.ok) return;

  const borderBlob = await borderResponse.blob();
  const borderBase64 = await this.blobToBase64(borderBlob);

  const borderId = workbook.addImage({
    base64: borderBase64,
    extension: 'png'
  });

  worksheet.addImage(borderId, {
    tl: { col: 0.05, row: 1.05 } as any,
    br: { col: 16.3, row: 6.15 } as any
  });
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
  tl: { col: 0.99, row: 2.20 } as any,
  ext: { width: 86, height: 86 }
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

private applyHeaderBorder(worksheet: ExcelJS.Worksheet) {
  const clearStartRow = 2;
  const clearEndRow = 7;
  const clearStartCol = 1;
  const clearEndCol = 20;

  const borderStartRow = 2;
  const borderEndRow = 6;
  const borderStartCol = 1;
  const borderEndCol = 16;

  for (let rowNumber = clearStartRow; rowNumber <= clearEndRow; rowNumber++) {
    const row = worksheet.getRow(rowNumber);

    for (let colNumber = clearStartCol; colNumber <= clearEndCol; colNumber++) {
      const cell = row.getCell(colNumber);

      cell.border = undefined as any;
    }
  }

  for (let colNumber = borderStartCol; colNumber <= borderEndCol; colNumber++) {
    const topCell = worksheet.getRow(borderStartRow).getCell(colNumber);
    const bottomCell = worksheet.getRow(borderEndRow).getCell(colNumber);

    topCell.border = {
      ...topCell.border,
      top: { style: 'medium' }
    };

    bottomCell.border = {
      ...bottomCell.border,
      bottom: { style: 'medium' }
    };
  }

  for (let rowNumber = borderStartRow; rowNumber <= borderEndRow; rowNumber++) {
    const leftCell = worksheet.getRow(rowNumber).getCell(borderStartCol);
    const rightCell = worksheet.getRow(rowNumber).getCell(borderEndCol);

    leftCell.border = {
      ...leftCell.border,
      left: { style: 'medium' }
    };

    rightCell.border = {
      ...rightCell.border,
      right: { style: 'medium' }
    };
  }
}
}
