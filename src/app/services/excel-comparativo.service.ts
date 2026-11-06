import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export type ComparativeCompany = {
  id?: string;
  name?: string;
  nome?: string;
  cnpj?: string;
  taxId?: string;
};

export type ComparativeItem = {
  id?: string;
  product?: string;
  descricao?: string;
  description?: string;
  name?: string;
  quantity?: number;
  quantidade?: number;
  unitPrice?: number;
  valorUnitario?: number;
  total?: number;
  valorTotal?: number;
};

export type ComparativeQuote = {
  label?: string;
  company?: ComparativeCompany;
  customerData?: {
    client?: string;
    fantasyName?: string;
  };
  items: ComparativeItem[];
  paymentTerm?: string;
  deliveryTerm?: string;
  adjustment?: number;
};

type ComparativeLayout = {
  productEndRow: number;
  totalRow: number;
  paymentRow: number;
  deliveryRow: number;
  approvedRow: number;
  infoRow: number;
  dateRow: number;
  lastSupplierColumn: number;
  lastStyledSupplierColumn: number;
};

@Injectable({
  providedIn: 'root'
})
export class ExcelComparativoService {
  private readonly templatePath = 'assets/templates/planilha-comparativa.xlsx';
  private readonly firstProductRow = 6;
  private readonly templateProductRows = 22;
  private readonly templateTotalRow = 29;
  private readonly firstSupplierColumn = 4;
  private readonly templateSupplierColumns = 5;
  private readonly suppliersPerPrintedPage = 5;

  async gerarPlanilhaComparativa(quotes: ComparativeQuote[]): Promise<void> {
    if (!quotes || quotes.length < 2) {
      throw new Error('É necessário ter o Orçamento A e pelo menos um orçamento derivado.');
    }

    const workbook = new ExcelJS.Workbook();
    const response = await fetch(this.templatePath);

    if (!response.ok) {
      throw new Error('Não foi possível carregar o modelo da planilha comparativa.');
    }

    const buffer = await response.arrayBuffer();

    await workbook.xlsx.load(buffer);

    this.populateWorkbook(workbook, quotes);

    const outputBuffer = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([outputBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }),
      `planilha-comparativa-${this.formatarDataArquivo(new Date())}.xlsx`
    );
  }

  populateWorkbook(workbook: ExcelJS.Workbook, quotes: ComparativeQuote[]): ExcelJS.Workbook {
    if (!quotes?.length) {
      throw new Error('É necessário informar pelo menos um orçamento.');
    }

    const templateWorksheet = workbook.worksheets[0];

    if (!templateWorksheet) {
      throw new Error('A planilha modelo não possui nenhuma aba.');
    }

    const supplierGroups = this.chunkQuotes(quotes);
    const worksheets = [templateWorksheet];

    for (let index = 1; index < supplierGroups.length; index++) {
      worksheets.push(this.cloneTemplateWorksheet(workbook, templateWorksheet, `Comparativo ${index + 1}`));
    }

    if (supplierGroups.length > 1) {
      templateWorksheet.name = 'Comparativo 1';
    }

    const comparativeItems = this.collectComparativeItems(quotes);

    supplierGroups.forEach((supplierQuotes, index) => {
      const worksheet = worksheets[index];
      const layout = this.prepararLayout(worksheet, comparativeItems.length, supplierQuotes.length);

      this.preencherCabecalho(worksheet, supplierQuotes, quotes[0]);
      this.preencherProdutos(worksheet, supplierQuotes, comparativeItems, layout);
      this.preencherResumo(worksheet, supplierQuotes, layout);
      this.ajustarBordasDoTemplate(worksheet, layout);
      this.configurarImpressao(worksheet, layout.lastSupplierColumn, layout.dateRow + 1);
    });

    return workbook;
  }

  private chunkQuotes(quotes: ComparativeQuote[]): ComparativeQuote[][] {
    const groups: ComparativeQuote[][] = [];

    for (let index = 0; index < quotes.length; index += this.suppliersPerPrintedPage) {
      groups.push(quotes.slice(index, index + this.suppliersPerPrintedPage));
    }

    return groups;
  }

  private collectComparativeItems(quotes: ComparativeQuote[]): ComparativeItem[] {
    const items: ComparativeItem[] = [];
    const knownItems = new Set<string>();

    quotes.forEach(quote => {
      (quote.items ?? []).forEach((item, itemIndex) => {
        const key = this.getItemKey(item, itemIndex);

        if (knownItems.has(key)) return;

        knownItems.add(key);
        items.push(item);
      });
    });

    return items;
  }

  private cloneTemplateWorksheet(workbook: ExcelJS.Workbook, source: ExcelJS.Worksheet, name: string): ExcelJS.Worksheet {
    const target = workbook.addWorksheet(name);
    const targetId = target.id;
    const model = this.deepClone(source.model);

    model.id = targetId;
    model.name = name;
    target.model = model;

    return target;
  }

  private deepClone<T>(value: T): T {
    if (value instanceof Date) {
      return new Date(value.getTime()) as T;
    }

    if (Array.isArray(value)) {
      return value.map(entry => this.deepClone(entry)) as T;
    }

    if (value && typeof value === 'object') {
      return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, this.deepClone(entry)])) as T;
    }

    return value;
  }

  private prepararLayout(worksheet: ExcelJS.Worksheet, itemCount: number, supplierCount: number): ComparativeLayout {
    const productRows = Math.max(itemCount, 1);
    const rowOffset = productRows - this.templateProductRows;
    const lastSupplierColumn = this.firstSupplierColumn + supplierCount - 1;
    const templateLastSupplierColumn = this.firstSupplierColumn + this.templateSupplierColumns - 1;
    const lastStyledSupplierColumn = Math.max(lastSupplierColumn, templateLastSupplierColumn);

    this.removerMesclagens(worksheet);

    if (rowOffset > 0) {
      worksheet.spliceRows(this.firstProductRow + this.templateProductRows, 0, ...Array.from({ length: rowOffset }, () => []));
    }

    if (rowOffset < 0) {
      worksheet.spliceRows(this.firstProductRow + productRows, Math.abs(rowOffset));
    }

    const productEndRow = this.firstProductRow + productRows - 1;
    const totalRow = this.templateTotalRow + rowOffset;
    const paymentRow = totalRow + 1;
    const deliveryRow = totalRow + 2;
    const approvedRow = totalRow + 3;
    const infoRow = totalRow + 5;
    const dateRow = totalRow + 9;

    this.copiarEstiloLinhasProduto(worksheet, productEndRow, lastStyledSupplierColumn);
    this.copiarEstiloColunasFornecedor(worksheet, lastStyledSupplierColumn, dateRow + 1);

    const layout = {
      productEndRow,
      totalRow,
      paymentRow,
      deliveryRow,
      approvedRow,
      infoRow,
      dateRow,
      lastSupplierColumn,
      lastStyledSupplierColumn
    };

    this.recriarMesclagens(worksheet, layout);

    return layout;
  }

  private preencherCabecalho(worksheet: ExcelJS.Worksheet, quotes: ComparativeQuote[], customerQuote: ComparativeQuote): void {
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'Planilha - Mapa Comparativo';
    titleCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: false
    };
    titleCell.border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    worksheet.getCell('B3').value = this.getCustomerName(customerQuote) || 'Cliente';
    worksheet.getCell('D3').value = 'Produtos de Limpeza';

    quotes.forEach((quote, index) => {
      const columnIndex = this.firstSupplierColumn + index;
      const companyName = this.getCompanyName(quote);

      worksheet.getCell(4, columnIndex).value = companyName || quote.label || `Fornecedor ${index + 1}`;
      worksheet.getCell(5, columnIndex).value = 'V. Unitário';

      this.aplicarEstiloCabecalhoFornecedor(worksheet.getCell(4, columnIndex));
      this.aplicarEstiloCabecalhoFornecedor(worksheet.getCell(5, columnIndex));
    });

    this.limparColunasFornecedorNaoUsadas(worksheet, quotes.length, 4, 5);
  }

  private preencherProdutos(worksheet: ExcelJS.Worksheet, quotes: ComparativeQuote[], baseItems: ComparativeItem[], layout: ComparativeLayout): void {
    const productRows = Math.max(baseItems.length, 1);

    for (let itemIndex = 0; itemIndex < productRows; itemIndex++) {
      const item = baseItems[itemIndex];
      const rowNumber = this.firstProductRow + itemIndex;
      const row = worksheet.getRow(rowNumber);

      row.getCell(1).value = item ? itemIndex + 1 : '';
      row.getCell(2).value = item ? this.getItemDescription(item) : '';
      row.getCell(3).value = item ? this.getQuantity(item) : '';

      quotes.forEach((quote, quoteIndex) => {
        const columnIndex = this.firstSupplierColumn + quoteIndex;
        const quoteItem = item ? this.findQuoteItem(quote, item, itemIndex) : undefined;

        row.getCell(columnIndex).value = quoteItem ? this.getUnitPrice(quoteItem) : '';
        row.getCell(columnIndex).numFmt = 'R$ #,##0.00';
      });

      this.limparColunasFornecedorNaoUsadas(worksheet, quotes.length, rowNumber, rowNumber);
      this.aplicarEstiloLinhaProduto(row, layout.lastSupplierColumn);
    }
  }

  private preencherResumo(worksheet: ExcelJS.Worksheet, quotes: ComparativeQuote[], layout: ComparativeLayout): void {
    worksheet.getCell(layout.totalRow, 1).value = 'TOTAL:';
    worksheet.getCell(layout.paymentRow, 1).value = 'FORMA DE PAGAMENTO:';
    worksheet.getCell(layout.deliveryRow, 1).value = 'PRAZO DE ENTREGA:';
    worksheet.getCell(layout.approvedRow, 1).value = 'EMPRESA APROVADA';

    this.limparColunasFornecedorNaoUsadas(worksheet, quotes.length, layout.totalRow, layout.approvedRow);

    quotes.forEach((quote, index) => {
      const columnIndex = this.firstSupplierColumn + index;
      worksheet.getCell(layout.totalRow, columnIndex).value = quote.items.reduce((sum, item) => sum + this.getItemTotal(item), 0);
      worksheet.getCell(layout.totalRow, columnIndex).numFmt = 'R$ #,##0.00';
      worksheet.getCell(layout.paymentRow, columnIndex).value = quote.paymentTerm || '';
      worksheet.getCell(layout.deliveryRow, columnIndex).value = quote.deliveryTerm || '';
      worksheet.getCell(layout.deliveryRow, columnIndex).alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };
      worksheet.getCell(layout.approvedRow, columnIndex).value = '(   )';
    });

    worksheet.getCell(layout.infoRow, 1).value = 'Informações importantes sobre esta cotação a serem observadas:';

    worksheet.getCell(layout.dateRow, 1).value =
      'Devolver ao Deptº de Compras                                                                             Elaborado em';
    worksheet.getCell(layout.dateRow, 3).value = new Date();
    worksheet.getCell(layout.dateRow, 3).numFmt = 'dd/mm/yyyy';

    for (let rowIndex = layout.totalRow; rowIndex <= layout.dateRow + 1; rowIndex++) {
      this.aplicarEstiloLinhaResumo(worksheet.getRow(rowIndex), layout.lastSupplierColumn);
    }

    worksheet.getRow(layout.paymentRow).height = 26.25;
    worksheet.getRow(layout.deliveryRow).height = 60;
  }

  private recriarMesclagens(worksheet: ExcelJS.Worksheet, layout: ComparativeLayout): void {
    this.safeMergeCells(worksheet, 1, 1, 2, layout.lastSupplierColumn);
    this.safeMergeCells(worksheet, 4, 1, 5, 1);
    this.safeMergeCells(worksheet, 4, 2, 5, 2);
    this.safeMergeCells(worksheet, layout.totalRow, 1, layout.totalRow, 3);
    this.safeMergeCells(worksheet, layout.paymentRow, 1, layout.paymentRow, 3);
    this.safeMergeCells(worksheet, layout.deliveryRow, 1, layout.deliveryRow, 3);
    this.safeMergeCells(worksheet, layout.approvedRow, 1, layout.approvedRow, 3);
    this.safeMergeCells(worksheet, layout.infoRow, 1, layout.infoRow, layout.lastSupplierColumn);
    this.safeMergeCells(worksheet, layout.dateRow, 1, layout.dateRow + 1, 2);
    this.safeMergeCells(worksheet, layout.dateRow, 3, layout.dateRow + 1, layout.lastSupplierColumn);
  }

  private removerMesclagens(worksheet: ExcelJS.Worksheet): void {
    [...(worksheet.model.merges ?? [])].forEach(range => {
      try {
        worksheet.unMergeCells(range);
      } catch {}
    });
  }

  private copiarEstiloLinhasProduto(worksheet: ExcelJS.Worksheet, productEndRow: number, lastSupplierColumn: number): void {
    const sourceRow = worksheet.getRow(this.firstProductRow);

    for (let rowNumber = this.firstProductRow; rowNumber <= productEndRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      row.height = sourceRow.height || 18;

      for (let columnIndex = 1; columnIndex <= lastSupplierColumn; columnIndex++) {
        const sourceColumn = columnIndex <= this.firstSupplierColumn ? columnIndex : this.firstSupplierColumn;
        this.copyCellStyle(sourceRow.getCell(sourceColumn), row.getCell(columnIndex));
      }
    }
  }

  private copiarEstiloColunasFornecedor(worksheet: ExcelJS.Worksheet, lastSupplierColumn: number, lastRow: number): void {
    for (let columnIndex = this.firstSupplierColumn; columnIndex <= lastSupplierColumn; columnIndex++) {
      worksheet.getColumn(columnIndex).width = 18;
      worksheet.getColumn(columnIndex).hidden = false;

      if (columnIndex <= this.firstSupplierColumn + this.templateSupplierColumns - 1) {
        continue;
      }

      for (let rowNumber = 1; rowNumber <= lastRow; rowNumber++) {
        this.copyCellStyle(worksheet.getCell(rowNumber, this.firstSupplierColumn), worksheet.getCell(rowNumber, columnIndex));
      }
    }
  }

  private aplicarEstiloCabecalhoFornecedor(cell: ExcelJS.Cell): void {
    cell.font = {
      bold: false,
      size: 10
    };

    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };

    cell.border = this.getBorder();
  }

  private aplicarEstiloLinhaProduto(row: ExcelJS.Row, lastColumn: number): void {
    row.height = 18;

    for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
      const cell = row.getCell(columnIndex);

      cell.alignment = {
        horizontal: columnIndex === 2 ? 'left' : 'center',
        vertical: 'middle',
        wrapText: true
      };

      cell.font = {
        ...cell.font,
        bold: false,
        size: columnIndex === 2 || columnIndex === 3 ? 9 : 10
      };
    }
  }

  private aplicarEstiloLinhaResumo(row: ExcelJS.Row, lastColumn: number): void {
    for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
      const cell = row.getCell(columnIndex);

      cell.alignment = {
        horizontal: columnIndex === 1 ? 'left' : 'center',
        vertical: 'middle',
        wrapText: true
      };
    }
  }

  private ajustarBordasDoTemplate(worksheet: ExcelJS.Worksheet, layout: ComparativeLayout): void {
    const lastRow = layout.dateRow + 1;

    this.padronizarBordasExistentes(worksheet, layout.lastSupplierColumn, lastRow);
    this.aplicarContornoExterno(worksheet, layout.lastSupplierColumn, lastRow);
    this.limparBordasAbaixoDoConteudo(worksheet, layout.lastStyledSupplierColumn, lastRow);
    this.removerBordasVerticaisDaLinha2(worksheet, layout.lastSupplierColumn);
    this.adicionarBordaDireitaLinha2(worksheet, layout.lastSupplierColumn);
    this.adicionarBordaDireitaLinha18(worksheet, layout);
    this.normalizarEstiloCabecalho(worksheet, layout.lastSupplierColumn);
  }

  private padronizarBordasExistentes(worksheet: ExcelJS.Worksheet, lastColumn: number, lastRow: number): void {
    for (let rowNumber = 1; rowNumber <= lastRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
        const cell = row.getCell(columnIndex);

        if (!cell.border) continue;

        cell.border = this.normalizeBorder(cell.border);
      }
    }
  }

  private aplicarContornoExterno(worksheet: ExcelJS.Worksheet, lastColumn: number, lastRow: number): void {
    for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
      this.mergeBorder(worksheet.getCell(1, columnIndex), { top: { style: 'thin' } });
      this.mergeBorder(worksheet.getCell(lastRow, columnIndex), { bottom: { style: 'thin' } });
    }

    for (let rowNumber = 1; rowNumber <= lastRow; rowNumber++) {
      this.mergeBorder(worksheet.getCell(rowNumber, 1), { left: { style: 'thin' } });
      this.mergeBorder(worksheet.getCell(rowNumber, lastColumn), { right: { style: 'thin' } });
    }
  }

  private limparBordasAbaixoDoConteudo(worksheet: ExcelJS.Worksheet, lastColumn: number, lastContentRow: number): void {
    for (let rowNumber = lastContentRow + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
        worksheet.getCell(rowNumber, columnIndex).border = undefined as any;
      }
    }
  }

  private removerBordasVerticaisDaLinha2(worksheet: ExcelJS.Worksheet, lastColumn: number): void {
    const row2 = worksheet.getRow(2);

    for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
      const cell = row2.getCell(columnIndex);

      if (cell.border) {
        const updatedBorder: Partial<ExcelJS.Borders> = {};

        if (cell.border.top) updatedBorder.top = cell.border.top;
        if (cell.border.bottom) updatedBorder.bottom = cell.border.bottom;

        cell.border = updatedBorder;
      }
    }
  }

  private adicionarBordaDireitaLinha18(worksheet: ExcelJS.Worksheet, layout: ComparativeLayout): void {
    const approvedRowCell = worksheet.getCell(layout.approvedRow, layout.lastSupplierColumn);

    if (approvedRowCell.border) {
      approvedRowCell.border = {
        ...approvedRowCell.border,
        right: { style: 'thin' }
      };
    } else {
      approvedRowCell.border = {
        right: { style: 'thin' }
      };
    }
  }

  private normalizarEstiloCabecalho(worksheet: ExcelJS.Worksheet, lastColumn: number): void {
    for (let rowNumber = 4; rowNumber <= 5; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let columnIndex = 1; columnIndex <= lastColumn; columnIndex++) {
        const cell = row.getCell(columnIndex);

        if (cell.font) {
          cell.font = {
            ...cell.font,
            bold: false
          };
        }
      }
    }
  }

  private adicionarBordaDireitaLinha2(worksheet: ExcelJS.Worksheet, lastColumn: number): void {
    const lastCellRow2 = worksheet.getCell(2, lastColumn);

    if (lastCellRow2.border) {
      lastCellRow2.border = {
        ...lastCellRow2.border,
        right: { style: 'thin' }
      };
    } else {
      lastCellRow2.border = {
        right: { style: 'thin' }
      };
    }
  }

  private configurarImpressao(worksheet: ExcelJS.Worksheet, lastSupplierColumn: number, lastRow: number): void {
    worksheet.pageSetup = {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
      printArea: `A1:${worksheet.getColumn(lastSupplierColumn).letter}${lastRow}`
    };

    worksheet.getColumn(1).width = 13.69;
    worksheet.getColumn(2).width = 67.38;
    worksheet.getColumn(3).width = 14.75;
    worksheet.getColumn(5).width = 18;

    for (let columnIndex = this.firstSupplierColumn; columnIndex <= lastSupplierColumn; columnIndex++) {
      worksheet.getColumn(columnIndex).width = 18.38;
      worksheet.getColumn(columnIndex).hidden = false;
    }

    for (let columnIndex = lastSupplierColumn + 1; columnIndex <= Math.max(worksheet.columnCount, 26); columnIndex++) {
      worksheet.getColumn(columnIndex).hidden = true;
    }
  }

  private limparColunasFornecedorNaoUsadas(worksheet: ExcelJS.Worksheet, supplierCount: number, startRow: number, endRow: number): void {
    const firstUnusedColumn = this.firstSupplierColumn + supplierCount;
    const lastColumnToClear = this.firstSupplierColumn + this.templateSupplierColumns - 1;

    for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
      for (let columnIndex = firstUnusedColumn; columnIndex <= lastColumnToClear; columnIndex++) {
        const cell = worksheet.getCell(rowNumber, columnIndex);

        cell.value = null;
      }
    }
  }

  private getCompanyName(quote: ComparativeQuote): string {
    return quote.company?.name || quote.company?.nome || quote.company?.id || quote.label || '';
  }

  private getCustomerName(quote: ComparativeQuote): string {
    return quote.customerData?.fantasyName || quote.customerData?.client || '';
  }

  private getItemDescription(item: ComparativeItem): string {
    return item.product || item.descricao || item.description || item.name || '';
  }

  private getQuantity(item: ComparativeItem): number {
    return Number(item.quantity ?? item.quantidade ?? 0);
  }

  private getUnitPrice(item: ComparativeItem): number {
    return Number(item.unitPrice ?? item.valorUnitario ?? 0);
  }

  private getItemTotal(item: ComparativeItem): number {
    const total = item.total ?? item.valorTotal;

    if (total !== undefined && total !== null) {
      return Number(total);
    }

    return this.getQuantity(item) * this.getUnitPrice(item);
  }

  private getItemKey(item: ComparativeItem, itemIndex: number): string {
    if (item.id !== undefined && item.id !== null && item.id !== '') {
      return `id:${String(item.id)}`;
    }

    const description = this.normalizeText(this.getItemDescription(item));

    return description ? `description:${description}` : `position:${itemIndex}`;
  }

  private normalizeText(value: string): string {
    return value
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .toLowerCase();
  }

  private findQuoteItem(quote: ComparativeQuote, baseItem: ComparativeItem, itemIndex: number): ComparativeItem | undefined {
    if (!quote.items?.length) {
      return undefined;
    }

    if (baseItem.id !== undefined && baseItem.id !== null && baseItem.id !== '') {
      const itemById = quote.items.find(item => item.id === baseItem.id);

      if (itemById) {
        return itemById;
      }
    }

    const baseDescription = this.normalizeText(this.getItemDescription(baseItem));

    if (baseDescription) {
      const itemByDescription = quote.items.find(item => this.normalizeText(this.getItemDescription(item)) === baseDescription);

      if (itemByDescription) {
        return itemByDescription;
      }
    }

    return quote.items[itemIndex];
  }

  private safeMergeCells(worksheet: ExcelJS.Worksheet, startRow: number, startColumn: number, endRow: number, endColumn: number): void {
    try {
      worksheet.mergeCells(startRow, startColumn, endRow, endColumn);
    } catch {}
  }

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell): void {
    target.style = this.cloneStyle(source.style);

    if (source.numFmt) target.numFmt = source.numFmt;
    if (source.alignment) target.alignment = this.cloneStyle(source.alignment);
    if (source.border) target.border = this.cloneStyle(source.border);
    if (source.font) target.font = this.cloneStyle(source.font);
    if (source.fill) target.fill = this.cloneStyle(source.fill);
  }

  private cloneStyle<T>(value: T): T {
    return value ? JSON.parse(JSON.stringify(value)) : value;
  }

  private getBorder(): Partial<ExcelJS.Borders> {
    return {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  }

  private normalizeBorder(border: Partial<ExcelJS.Borders>): Partial<ExcelJS.Borders> {
    const normalizedBorder: Partial<ExcelJS.Borders> = {};

    if (border.top) normalizedBorder.top = { style: 'thin' };
    if (border.left) normalizedBorder.left = { style: 'thin' };
    if (border.bottom) normalizedBorder.bottom = { style: 'thin' };
    if (border.right) normalizedBorder.right = { style: 'thin' };

    return normalizedBorder;
  }

  private mergeBorder(cell: ExcelJS.Cell, border: Partial<ExcelJS.Borders>): void {
    cell.border = {
      ...(cell.border ?? {}),
      ...border
    };
  }

  private formatarDataArquivo(date: Date): string {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();

    return `${day}-${month}-${year}`;
  }
}
