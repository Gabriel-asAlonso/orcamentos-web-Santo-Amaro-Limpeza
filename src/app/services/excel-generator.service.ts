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

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    worksheet.getCell('B6').value = quote.company.name;
    worksheet.getCell('B7').value = quote.company.taxId;
    worksheet.getCell('B8').value = quote.company.address;
    worksheet.getCell('D9').value = quote.company.identity;

    worksheet.getCell('F6').value = new Date().toLocaleDateString('pt-BR');
    worksheet.getCell('B11').value = `FORMA DE PAG.: ${quote.paymentMethod}`;
    worksheet.getCell('D11').value = `PRAZO DE ENTREGA: ${quote.deliveryTime}`;

    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 15;
    const linhaTotalOriginal = 16;
    const quantidadeProdutos = items.length;
    const linhaEmBranco = linhaInicial + quantidadeProdutos;
    const linhaTotal = linhaEmBranco + 1;

    const linhaModelo = worksheet.getRow(linhaInicial);
    const linhaTotalModelo = worksheet.getRow(linhaTotalOriginal);

    const alturaProduto = linhaModelo.height ?? 18;
    const alturaTotal = 30;

    const total = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);

    if (quantidadeProdutos > 1) {
      worksheet.spliceRows(
        linhaTotalOriginal,
        0,
        ...Array.from({ length: quantidadeProdutos - 1 }, () => [])
      );
    }

    worksheet.spliceRows(linhaEmBranco, 0, []);

    items.forEach((item, index) => {
      const linha = worksheet.getRow(linhaInicial + index);

      linha.height = alturaProduto;

      for (let colNumber = 1; colNumber <= 6; colNumber++) {
        const cell = linha.getCell(colNumber);
        const modelCell = linhaModelo.getCell(colNumber);

        cell.value = null;
        this.copyCellStyle(modelCell, cell);
        this.applyTableBorder(cell);
      }

      linha.getCell(1).value = index + 1;
      linha.getCell(2).value = item.quantity;
      linha.getCell(3).value = item.unit;
      linha.getCell(4).value = item.product;
      linha.getCell(5).value = item.unitPrice;
      linha.getCell(6).value = item.quantity * item.unitPrice;

      linha.getCell(5).numFmt = 'R$ #,##0.00';
      linha.getCell(6).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    const blankRow = worksheet.getRow(linhaEmBranco);

    blankRow.height = alturaProduto;

    for (let colNumber = 1; colNumber <= 6; colNumber++) {
      const cell = blankRow.getCell(colNumber);
      const modelCell = linhaModelo.getCell(colNumber);

      cell.value = null;
      this.copyCellStyle(modelCell, cell);
      this.applyTableBorder(cell);
    }

    blankRow.commit();

    const totalRow = worksheet.getRow(linhaTotal);

    totalRow.height = alturaTotal;

    for (let colNumber = 1; colNumber <= 6; colNumber++) {
      const cell = totalRow.getCell(colNumber);
      const modelCell = linhaTotalModelo.getCell(colNumber);

      cell.value = null;
      this.copyCellStyle(modelCell, cell);
      this.applyTableBorder(cell);
    }

    worksheet.mergeCells(`A${linhaTotal}:D${linhaTotal}`);

    worksheet.getCell(`E${linhaTotal}`).value = 'Total Geral';
    worksheet.getCell(`F${linhaTotal}`).value = total;

    worksheet.getCell(`E${linhaTotal}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${linhaTotal}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${linhaTotal}`).numFmt = 'R$ #,##0.00';

    worksheet.getCell(`E${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    worksheet.getCell(`F${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    totalRow.commit();

    const linhaDepoisDoTotal = linhaTotal + 1;
    const rowDepoisDoTotal = worksheet.getRow(linhaDepoisDoTotal);

    for (let colNumber = 1; colNumber <= 6; colNumber++) {
      const cell = rowDepoisDoTotal.getCell(colNumber);

      cell.value = null;
      cell.border = undefined as any;
      cell.font = undefined as any;
      cell.fill = undefined as any;
      cell.numFmt = '';
    }

    rowDepoisDoTotal.height = 0;
    rowDepoisDoTotal.hidden = true;
    rowDepoisDoTotal.commit();
  }

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
    target.style = { ...source.style };

    if (source.numFmt) target.numFmt = source.numFmt;
    if (source.alignment) target.alignment = { ...source.alignment };
    if (source.border) target.border = { ...source.border };
    if (source.font) target.font = { ...source.font };
    if (source.fill) target.fill = { ...source.fill };
  }

  private applyTableBorder(cell: ExcelJS.Cell) {
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  }

  private async insertCompanyLogo(workbook: ExcelJS.Workbook, worksheet: ExcelJS.Worksheet, logoPath: string) {
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
      tl: { col: 0, row: 0 } as any,
      br: { col: 3, row: 4 } as any
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
