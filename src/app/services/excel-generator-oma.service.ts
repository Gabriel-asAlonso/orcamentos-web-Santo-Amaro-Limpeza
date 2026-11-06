import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root',
})
export class ExcelGeneratorOmaService {
  private readonly templatePath = 'assets/templates/cotacao-template-oma.xlsx';
  private readonly firstProductRow = 13;
  private readonly lastTemplateProductRow = 27;
  private readonly templateTotalRow = 28;
  private readonly firstHiddenSupplierColumn = 9;
  private readonly lastTemplateColumn = 16;

  async generateQuoteExcel(quote: Quote): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    const templateResponse = await fetch(this.templatePath);

    if (!templateResponse.ok) {
      throw new Error(`Modelo de Excel não encontrado: ${this.templatePath}`);
    }

    await workbook.xlsx.load(await templateResponse.arrayBuffer());
    this.populateWorkbook(workbook, quote);

    const output = await workbook.xlsx.writeBuffer();
    const filename = `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`;

    saveAs(
      new Blob([output], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      filename,
    );
  }

  populateWorkbook(workbook: ExcelJS.Workbook, quote: Quote): ExcelJS.Workbook {
    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('Aba da planilha não encontrada');
    }

    const productEndRow = this.prepareProductRows(worksheet, quote.items.length);
    const totalRow = productEndRow + 1;

    worksheet.views = [{ showGridLines: false }];
    this.fillHeader(worksheet, quote);
    this.fillSupplier(worksheet, quote);
    this.fillProducts(worksheet, quote.items, productEndRow);
    this.fillSummary(worksheet, quote, totalRow);
    this.configureSingleSupplierLayout(worksheet, totalRow);
    this.updateConditionalFormatting(worksheet, totalRow);
    this.updatePrintArea(worksheet, totalRow);

    return workbook;
  }

  private fillHeader(worksheet: ExcelJS.Worksheet, quote: Quote): void {
    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    const formattedDate = this.formatDate(quoteDate);

    worksheet.getCell('G3').value = `${quote.label} - ${formattedDate}`;
    // O modelo identifica este campo como código interno do condomínio.
    // O orçamento base não possui esse dado, portanto ele permanece vazio.
    worksheet.getCell('C6').value = null;
    worksheet.getCell('C7').value = this.cleanText(customer?.client);
    worksheet.getCell('C8').value = this.buildCustomerAddress(quote);
    worksheet.getCell('H7').value = this.cleanText(
      customer?.salesperson || customer?.externalSalesperson,
    );
  }

  private fillSupplier(worksheet: ExcelJS.Worksheet, quote: Quote): void {
    const supplierColumns = ['G', 'I', 'K'];

    supplierColumns.forEach((column) => {
      worksheet.getCell(`${column}9`).value = null;
      worksheet.getCell(`${column}10`).value = null;
      worksheet.getCell(`${column}11`).value = null;
    });

    worksheet.getCell('G9').value = this.cleanText(
      quote.company.aliases?.[0] || quote.company.name,
    );
    worksheet.getCell('G10').value = this.cleanText(quote.company.phone);
    worksheet.getCell('G11').value = this.cleanText(quote.company.contact);
  }

  private fillProducts(
    worksheet: ExcelJS.Worksheet,
    items: QuoteItem[],
    productEndRow: number,
  ): void {
    const accountingFormat = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    for (let rowNumber = this.firstProductRow; rowNumber <= productEndRow; rowNumber++) {
      this.ensureDescriptionMerge(worksheet, rowNumber);

      worksheet.getCell(`B${rowNumber}`).value = null;
      worksheet.getCell(`F${rowNumber}`).value = null;

      for (const column of ['G', 'H', 'I', 'J', 'K', 'L']) {
        worksheet.getCell(`${column}${rowNumber}`).value = null;
      }
    }

    items.forEach((item, index) => {
      const rowNumber = this.firstProductRow + index;
      const total = this.getItemTotal(item);

      worksheet.getCell(`B${rowNumber}`).value = this.cleanText(item.product);
      worksheet.getCell(`F${rowNumber}`).value = item.quantity ?? 0;
      worksheet.getCell(`F${rowNumber}`).numFmt = '0.##';
      worksheet.getCell(`G${rowNumber}`).value = item.unitPrice ?? 0;
      worksheet.getCell(`G${rowNumber}`).numFmt = accountingFormat;
      worksheet.getCell(`H${rowNumber}`).value = {
        formula: this.getItemFormula(item, rowNumber),
        result: total,
      };
      worksheet.getCell(`H${rowNumber}`).numFmt = accountingFormat;
    });
  }

  private fillSummary(worksheet: ExcelJS.Worksheet, quote: Quote, totalRow: number): void {
    const productEndRow = totalRow - 1;
    const total = quote.items.reduce((sum, item) => sum + this.getItemTotal(item), 0);
    const accountingFormat = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    worksheet.getCell(`G${totalRow}`).value = {
      formula: `SUM(H${this.firstProductRow}:H${productEndRow})`,
      result: total,
    };
    worksheet.getCell(`G${totalRow}`).numFmt = accountingFormat;
    worksheet.getCell(`I${totalRow}`).value = null;
    worksheet.getCell(`K${totalRow}`).value = null;

    const paymentConditionsRow = totalRow + 2;
    const paymentMethodRow = totalRow + 3;
    const deliveryRow = totalRow + 4;

    for (const rowNumber of [paymentConditionsRow, paymentMethodRow, deliveryRow]) {
      for (const column of ['G', 'I', 'K']) {
        worksheet.getCell(`${column}${rowNumber}`).value = null;
      }
    }

    worksheet.getCell(`G${paymentMethodRow}`).value = this.cleanText(quote.paymentMethod);
    worksheet.getCell(`G${deliveryRow}`).value = this.cleanText(quote.deliveryTime);
  }

  private prepareProductRows(worksheet: ExcelJS.Worksheet, itemCount: number): number {
    const templateCapacity = this.lastTemplateProductRow - this.firstProductRow + 1;
    const requiredRows = Math.max(itemCount, templateCapacity);
    const extraRows = requiredRows - templateCapacity;

    if (extraRows > 0) {
      this.unmergeFromRow(worksheet, this.templateTotalRow);
      worksheet.spliceRows(
        this.templateTotalRow,
        0,
        ...Array.from({ length: extraRows }, () => []),
      );

      for (let offset = 0; offset < extraRows; offset++) {
        const rowNumber = this.templateTotalRow + offset;
        const targetRow = worksheet.getRow(rowNumber);
        const sourceRow = worksheet.getRow(this.firstProductRow);

        targetRow.height = sourceRow.height ?? 16.5;

        for (let column = 2; column <= 12; column++) {
          this.copyCellStyle(sourceRow.getCell(column), targetRow.getCell(column));
        }
      }

      this.recreateSummaryMerges(worksheet, this.templateTotalRow + extraRows);
    }

    const productEndRow = this.firstProductRow + requiredRows - 1;

    for (let rowNumber = this.firstProductRow; rowNumber <= productEndRow; rowNumber++) {
      this.ensureDescriptionMerge(worksheet, rowNumber);
    }

    return productEndRow;
  }

  private recreateSummaryMerges(worksheet: ExcelJS.Worksheet, totalRow: number): void {
    for (const range of [
      `G${totalRow}:H${totalRow}`,
      `I${totalRow}:J${totalRow}`,
      `K${totalRow}:L${totalRow}`,
    ]) {
      this.safeMerge(worksheet, range);
    }

    for (let offset = 2; offset <= 4; offset++) {
      this.safeMerge(worksheet, `E${totalRow + offset}:F${totalRow + offset}`);
      this.safeMerge(worksheet, `G${totalRow + offset}:H${totalRow + offset}`);
      this.safeMerge(worksheet, `I${totalRow + offset}:J${totalRow + offset}`);
      this.safeMerge(worksheet, `K${totalRow + offset}:L${totalRow + offset}`);
    }

    this.safeMerge(worksheet, `B${totalRow + 6}:L${totalRow + 6}`);
    this.safeMerge(worksheet, `B${totalRow + 7}:E${totalRow + 7}`);
  }

  private updateConditionalFormatting(worksheet: ExcelJS.Worksheet, totalRow: number): void {
    const conditionalFormattings = (worksheet as any).conditionalFormattings ?? [];

    if (conditionalFormattings[0]) {
      conditionalFormattings[0].ref = `G${totalRow}`;
    }

    if (conditionalFormattings[1]) {
      conditionalFormattings[1].ref = `G${totalRow}:H${totalRow}`;
    }
  }

  private updatePrintArea(worksheet: ExcelJS.Worksheet, totalRow: number): void {
    const templateLastPrintRow = 40;
    const extraRows = totalRow - this.templateTotalRow;

    worksheet.pageSetup.printArea = `A1:H${templateLastPrintRow + extraRows}`;
  }

  private configureSingleSupplierLayout(worksheet: ExcelJS.Worksheet, totalRow: number): void {
    this.safeUnmerge(worksheet, 'G3:L3');
    this.safeMerge(worksheet, 'G3:H3');

    for (const rowNumber of [9, 10, 11, totalRow, totalRow + 2, totalRow + 3, totalRow + 4]) {
      this.safeUnmerge(worksheet, `I${rowNumber}:J${rowNumber}`);
      this.safeUnmerge(worksheet, `K${rowNumber}:L${rowNumber}`);
    }

    const warningRow = totalRow + 6;
    this.safeUnmerge(worksheet, `B${warningRow}:L${warningRow}`);
    this.safeMerge(worksheet, `B${warningRow}:H${warningRow}`);

    worksheet.getColumn('G').hidden = false;
    worksheet.getColumn('H').hidden = false;

    for (
      let columnIndex = this.firstHiddenSupplierColumn;
      columnIndex <= Math.max(this.lastTemplateColumn, worksheet.columnCount);
      columnIndex++
    ) {
      worksheet.getColumn(columnIndex).hidden = true;
    }
  }

  private safeUnmerge(worksheet: ExcelJS.Worksheet, range: string): void {
    if (!(worksheet.model.merges ?? []).includes(range)) return;

    try {
      worksheet.unMergeCells(range);
    } catch {}
  }

  private ensureDescriptionMerge(worksheet: ExcelJS.Worksheet, rowNumber: number): void {
    this.safeMerge(worksheet, `B${rowNumber}:E${rowNumber}`);
  }

  private unmergeFromRow(worksheet: ExcelJS.Worksheet, startRow: number): void {
    const merges = [...(worksheet.model.merges ?? [])];

    for (const range of merges) {
      const match = /^[A-Z]+(\d+):[A-Z]+\d+$/.exec(range);

      if (match && Number(match[1]) >= startRow) {
        try {
          worksheet.unMergeCells(range);
        } catch {}
      }
    }
  }

  private safeMerge(worksheet: ExcelJS.Worksheet, range: string): void {
    if ((worksheet.model.merges ?? []).includes(range)) return;

    try {
      worksheet.mergeCells(range);
    } catch {}
  }

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell): void {
    target.style = JSON.parse(JSON.stringify(source.style ?? {}));
  }

  private getItemTotal(item: QuoteItem): number {
    const subtotal = (item.quantity ?? 0) * (item.unitPrice ?? 0);
    const discount = item.discount ?? 0;

    return subtotal - subtotal * (discount / 100);
  }

  private getItemFormula(item: QuoteItem, rowNumber: number): string {
    const discount = item.discount ?? 0;

    if (discount === 0) {
      return `G${rowNumber}*F${rowNumber}`;
    }

    return `G${rowNumber}*F${rowNumber}*(1-${discount}/100)`;
  }

  private buildCustomerAddress(quote: Quote): string {
    const customer = quote.customerData;

    return [
      customer?.address,
      customer?.district,
      customer?.city,
      customer?.state,
      customer?.postalCode,
    ]
      .map((value) => this.cleanText(value))
      .filter(Boolean)
      .join(' - ');
  }

  private cleanText(value: unknown): string {
    return value == null
      ? ''
      : String(value)
          .replace(/[\r\n]+/g, ' ')
          .trim();
  }

  private formatDate(date: Date): string {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');

    return `${day}/${month}/${date.getFullYear()}`;
  }
}
