import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorService {
  private templatePath = 'assets/templates/cotacao-template.xlsx';
  private fallbackLogoPath = '/logoSantoAmaroLimpeza.png';

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
    this.clearWorksheetRowGrouping(worksheet);

    await this.insertHeaderBorder(workbook, worksheet);
    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const companyAddress = quote.company.address || '';
    const customer = quote.customerData;

    // DEBUG: imprime o customer recebido
    console.log('[Laguna/Default] Customer recebido:', customer);
    console.log('[Laguna/Default] Quote info:', {
      label: quote.label,
      company: quote.company.id,
      source: (quote as any).source,
      paymentMethod: quote.paymentMethod,
      deliveryTime: quote.deliveryTime
    });

    this.applyCustomerDataLayout(worksheet);

    worksheet.getCell('A2').value = '';
    worksheet.getCell('A3').value = quote.company.name;
    worksheet.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A3').font = { bold: true, size: 11 };
    worksheet.getCell('C4').value = '';
    worksheet.getCell('D4').value = '';
    this.safeMergeCells(worksheet, 'M4:O4');
    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    worksheet.getCell('M4').value = `Data: ${quoteDate.toLocaleDateString('pt-BR')}`;
    worksheet.getCell('M4').alignment = { horizontal: 'right', vertical: 'middle' };
    worksheet.getCell('M4').font = { bold: true, size: 10 };
    worksheet.getCell('N4').font = { bold: true, size: 10 };
    worksheet.getCell('D5').value = companyDocument;
    this.safeMergeCells(worksheet, 'D6:L6');
    worksheet.getCell('C6').value = companyAddress ? 'Endereço:' : '';
    worksheet.getCell('D6').value = companyAddress;
    worksheet.getCell('C6').alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
    worksheet.getCell('D6').alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    worksheet.getCell('C6').font = { bold: false, size: 9 };
    worksheet.getCell('D6').font = { bold: false, size: 9 };

    worksheet.getCell('C7').value = customer?.client ?? '';
    worksheet.getCell('C8').value = customer?.fantasyName ?? customer?.client ?? '';
    worksheet.getCell('C9').value = customer?.address ?? '';
    worksheet.getCell('K9').value = 'Bairro:';
    worksheet.getCell('M9').value = customer?.district ?? '';
    worksheet.getCell('C10').value = customer?.city ?? '';
    worksheet.getCell('K10').value = 'UF:';
    worksheet.getCell('M10').value = customer?.state ?? '';
    worksheet.getCell('C11').value = customer?.document ?? '';
    worksheet.getCell('H11').value = customer?.stateRegistration ?? '';
    worksheet.getCell('K11').value = 'CEP:';
    worksheet.getCell('M11').value = customer?.postalCode ?? '';
    worksheet.getCell('C12').value = customer?.freight ?? '';
    worksheet.getCell('K12').value = 'Forma Pagto. :';
    worksheet.getCell('M12').value = quote.paymentMethod;
    worksheet.getCell('C13').value = customer?.phone ?? '';
    worksheet.getCell('K13').value = 'Região:';
    worksheet.getCell('M13').value = customer?.region ?? '';
    worksheet.getCell('C14').value = customer?.contact ?? '';
    worksheet.getCell('K14').value = 'WhatsApp';
    worksheet.getCell('M14').value = customer?.whatsapp ?? '';
    worksheet.getCell('C15').value = customer?.email ?? '';
    worksheet.getCell('K15').value = 'Tabela Preço:';
    worksheet.getCell('M15').value = customer?.priceTable ?? '';
    worksheet.getCell('C16').value = customer?.salesperson ?? '';
    worksheet.getCell('C17').value = customer?.externalSalesperson ?? '';
    worksheet.getCell('C17').font = { bold: false, size: 10 };
    worksheet.getCell('C17').alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    this.applyCustomerDataFont(worksheet);
    this.applyDeliveryTimeHeader(worksheet, quote.deliveryTime);
    this.applyProductHeaderLayout(worksheet);

    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
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
      this.showRowWithoutGrouping(linha);

      for (let colNumber = 1; colNumber <= 16; colNumber++) {
        const cell = linha.getCell(colNumber);
        const modelCell = linhaModelo.getCell(colNumber);

        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      this.applyProductRowLayout(worksheet, rowNumber);

      // Formato Accounting: cifrão à esquerda, número à direita (justificado)
      const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

      linha.getCell(1).value = item.code ?? index + 1;
      // Descrição agora ocupa B:K mergido (era D:K com B:C de referência)
      // IMPORTANTE: não tocar nas células filhas (D, E, ..., K) — qualquer
      // escrita nelas apaga o conteúdo do master (B) no ExcelJS
      linha.getCell(2).value = item.product;
      // Unidade: garantir que é texto (não número) e sem quebra
      linha.getCell(12).value = String(item.unit ?? 'unid.');
      linha.getCell(13).value = discount / 100;
      linha.getCell(14).value = Math.round(item.quantity ?? 0);
      linha.getCell(15).value = item.unitPrice;
      linha.getCell(16).value = totalProduto;

      linha.getCell(12).numFmt = '@'; // texto
      linha.getCell(13).numFmt = '0.00%';
      linha.getCell(14).numFmt = '0';
      linha.getCell(15).numFmt = accountingFmt;
      linha.getCell(16).numFmt = accountingFmt;

      linha.getCell(1).font = { size: 9 };
      linha.getCell(2).font = { size: 9 };
      linha.getCell(12).font = { size: 9 };
      linha.getCell(13).font = { size: 9 };
      linha.getCell(14).font = { size: 9 };
      linha.getCell(15).font = { size: 9 };
      linha.getCell(16).font = { size: 9 };

      // Alinhamento consistente em todas as linhas (justificação uniforme)
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      // Descrição B:K mergida
      linha.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(12).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(13).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(14).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(15).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
      linha.getCell(16).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };

      linha.commit();
    });

    const totalRow = worksheet.getRow(linhaTotal);

    totalRow.height = alturaTotal;
    this.showRowWithoutGrouping(totalRow);

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
    worksheet.getCell(`P${linhaTotal}`).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    worksheet.getCell(`A${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    worksheet.getCell(`P${linhaTotal}`).alignment = {
      horizontal: 'right',
      vertical: 'middle'
    };

    totalRow.commit();

    const lastVisibleRow = linhaTotal;

    this.clearWorksheetRowGrouping(worksheet, linhaInicial, lastVisibleRow);
    this.removeRowsAfterTotal(worksheet, lastVisibleRow);
  }

  private prepareWorksheetColumns(worksheet: ExcelJS.Worksheet) {
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 8;
    worksheet.getColumn(3).width = 8;

    worksheet.getColumn(4).width = 7;
    worksheet.getColumn(5).width = 14;
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

    worksheet.properties.outlineLevelRow = 0;
  }

  private applyCustomerDataLayout(worksheet: ExcelJS.Worksheet) {
    [
      'C9:L9',
      'C9:K9',
      'L9:M9',
      'N9:P9',
      'N9:Q9',
      'C10:K10',
      'L10:M10',
      'N10:P10',
      'O10:R10',
      'I11:L11',
      'I11:K11',
      'L11:M11',
      'N11:P11',
      'L12:M12',
      'N12:P12',
      'N12:Q12',
      'E13:G13',
      'H13:J13',
      'K13:L13',
      'L13:M13',
      'N13:P13',
      'M13:Q13',
      'C14:L14',
      'C14:K14',
      'L14:M14',
      'N14:P14',
      'C15:K15',
      'L15:M15',
      'N15:P15',
      'C15:Q15'
    ].forEach(range => this.safeUnmergeCells(worksheet, range));

    this.safeMergeCells(worksheet, 'C9:J9');
    this.safeMergeCells(worksheet, 'K9:L9');
    this.safeMergeCells(worksheet, 'M9:P9');
    this.safeMergeCells(worksheet, 'C10:J10');
    this.safeMergeCells(worksheet, 'K10:L10');
    this.safeMergeCells(worksheet, 'M10:P10');
    this.safeMergeCells(worksheet, 'C11:F11');
    this.safeMergeCells(worksheet, 'G11:H11');
    this.safeMergeCells(worksheet, 'I11:J11');
    this.safeMergeCells(worksheet, 'K11:L11');
    this.safeMergeCells(worksheet, 'M11:P11');
    this.safeMergeCells(worksheet, 'K12:L12');
    this.safeMergeCells(worksheet, 'M12:P12');
    this.safeMergeCells(worksheet, 'K13:L13');
    this.safeMergeCells(worksheet, 'M13:P13');
    this.safeMergeCells(worksheet, 'C14:J14');
    this.safeMergeCells(worksheet, 'K14:L14');
    this.safeMergeCells(worksheet, 'M14:P14');
    this.safeMergeCells(worksheet, 'C15:J15');
    this.safeMergeCells(worksheet, 'K15:L15');
    this.safeMergeCells(worksheet, 'M15:P15');

    ['E13', 'H13', 'N9', 'N10', 'N11', 'N12', 'N13', 'N14', 'N15', 'O10'].forEach(address => {
      worksheet.getCell(address).value = '';
    });

    for (let rowNumber = 9; rowNumber <= 15; rowNumber++) {
      const labelCell = worksheet.getCell(`K${rowNumber}`);
      const valueCell = worksheet.getCell(`M${rowNumber}`);

      labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
      valueCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    }
  }

  private applyCustomerDataFont(worksheet: ExcelJS.Worksheet) {
    for (let rowNumber = 7; rowNumber <= 17; rowNumber++) {
      for (let colNumber = 1; colNumber <= 16; colNumber++) {
        this.setNormalFont(worksheet.getRow(rowNumber).getCell(colNumber));
      }
    }
  }

  private applyProductRowLayout(worksheet: ExcelJS.Worksheet, rowNumber: number) {
    // Coluna Referência foi removida do template — descrição agora ocupa B:K inteiro
    // (era: B:C = Referência, D:K = Descrição)
    this.safeUnmergeCells(worksheet, `B${rowNumber}:C${rowNumber}`);
    this.safeUnmergeCells(worksheet, `D${rowNumber}:K${rowNumber}`);
    this.safeMergeCells(worksheet, `B${rowNumber}:K${rowNumber}`);

    const row = worksheet.getRow(rowNumber);

    for (let colNumber = 1; colNumber <= 16; colNumber++) {
      row.getCell(colNumber).border = undefined as any;
    }

    this.setOuterBorder(row.getCell(1), true, true, true, true);
    // Descrição B:K mergida - apenas a célula B precisa de bordas top/bottom/left
    this.setOuterBorder(row.getCell(2), true, true, true, false);

    for (let colNumber = 3; colNumber <= 10; colNumber++) {
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

  private applyDeliveryTimeHeader(worksheet: ExcelJS.Worksheet, deliveryTime: string) {
    this.safeUnmergeCells(worksheet, 'L17:M17');
    this.safeUnmergeCells(worksheet, 'N17:P17');
    this.safeMergeCells(worksheet, 'K17:L17');
    this.safeMergeCells(worksheet, 'M17:P17');

    const labelCell = worksheet.getCell('K17');
    const valueCell = worksheet.getCell('M17');

    labelCell.value = 'Prazo de Entrega:';
    valueCell.value = deliveryTime;

    labelCell.font = { bold: false, size: 10 };
    valueCell.font = { bold: false, size: 10 };

    labelCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
    valueCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };

    for (let colNumber = 11; colNumber <= 16; colNumber++) {
      const cell = worksheet.getRow(17).getCell(colNumber);

       cell.border = { bottom: { style: 'thin' } };
      cell.fill = undefined as any;
    }
  }

  private applyProductHeaderLayout(worksheet: ExcelJS.Worksheet) {
    const k18 = worksheet.getCell('K18');

    k18.border = {
      ...k18.border,
      top: { style: 'thin' }
    };
  }

  private clearWorksheetRowGrouping(
    worksheet: ExcelJS.Worksheet,
    startRow = 1,
    endRow = worksheet.rowCount
  ) {
    worksheet.properties.outlineLevelRow = 0;

    for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
      this.showRowWithoutGrouping(worksheet.getRow(rowNumber));
    }
  }

  private showRowWithoutGrouping(row: ExcelJS.Row) {
    row.hidden = false;
    row.outlineLevel = 0;

    const rowModel = row.model as ExcelJS.RowModel & { collapsed?: boolean };
    rowModel.hidden = false;
    rowModel.outlineLevel = 0;
    rowModel.collapsed = false;
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

  private safeUnmergeCells(worksheet: ExcelJS.Worksheet, range: string) {
    try {
      worksheet.unMergeCells(range);
    } catch {}
  }

  private setNormalFont(cell: ExcelJS.Cell, size = 10) {
    cell.font = {
      ...cell.font,
      bold: false,
      size
    };
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

    // Usar twoCellAnchor (tl + br) em vez de oneCellAnchor (tl + ext).
    // O oneCellAnchor com 'ext' fixo em pixels é interpretado de forma
    // inconsistente entre Excel desktop e Google Sheets ao exportar PDF
    // (Excel reinterpreta width/height como pontos × DPI, distorcendo a
    // imagem). Com twoCellAnchor a imagem fica ancorada à grade e mantém
    // proporção igual em todos os apps.
    worksheet.addImage(logoId, {
      tl: { col: 0.7, row: 1.95 } as any,
      br: { col: 2.3, row: 5.45 } as any
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
