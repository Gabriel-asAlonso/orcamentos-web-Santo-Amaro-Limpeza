import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  ArrowDown,
  ArrowUp,
  Building2,
  Download,
  FileSpreadsheet,
  FileText,
  LucideAngularModule,
  Plus,
  Trash2
} from 'lucide-angular';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { COMPANIES, type Company } from './data/empresas';

type QuoteItem = {
  id: number;
  product: string;
  quantity: number;
  unit: string;
  unitPrice: number;
};

type Quote = {
  label: string;
  company: Company;
  items: QuoteItem[];
  adjustment: number;
  source: string;
};

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule, FormsModule, LucideAngularModule],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App {
  ArrowDown = ArrowDown;
  ArrowUp = ArrowUp;
  Building2 = Building2;
  Download = Download;
  FileSpreadsheet = FileSpreadsheet;
  FileText = FileText;
  Plus = Plus;
  Trash2 = Trash2;

  templatePath = 'assets/templates/cotacao-template.xlsx';

  companies: Company[] = COMPANIES;

  mode: 'manual' | 'excel' = 'manual';
  companyId = this.companies[0].id;
  derivedCompanyId = this.companies[1]?.id ?? this.companies[0].id;
  adjustment = 0;
  quotes: Quote[] = [];
  excelStatus = 'Aguardando planilha com colunas: empresa, produto, quantidade, valor.';

  items: QuoteItem[] = [
    { id: 1, product: 'Copo térmico personalizado', quantity: 120, unit: 'peça', unitPrice: 10 },
    { id: 2, product: 'Kit boas-vindas premium', quantity: 40, unit: 'kit', unitPrice: 86.5 }
  ];

  get selectedCompany() {
    return this.companies.find(company => company.id === this.companyId) ?? this.companies[0];
  }

  get baseQuote() {
    return this.quotes[0];
  }

  get previewItems() {
    return this.baseQuote ? this.adjustedItems(this.baseQuote.items, this.adjustment) : [];
  }

  get nextLabel() {
    return `Orçamento ${String.fromCharCode(65 + this.quotes.length)}`;
  }

  get currentTotal() {
    return this.items.reduce((total, item) => total + item.quantity * item.unitPrice, 0);
  }

  get previewTotal() {
    return this.previewItems.reduce((total, item) => total + item.quantity * item.unitPrice, 0);
  }

  currency(value: number) {
    return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }

  adjustedItems(items: QuoteItem[], percentage: number) {
    return items.map(item => ({ ...item, unitPrice: item.unitPrice * (1 + percentage / 100) }));
  }

  updateItem(id: number, field: keyof QuoteItem, value: string) {
    this.items = this.items.map(item =>
      item.id === id
        ? {
            ...item,
            [field]: field === 'product' || field === 'unit' ? value : Math.max(0, Number(value) || 0)
          }
        : item
    );
  }

  sumItems(total: number, item: QuoteItem) {
    return total + item.quantity * item.unitPrice;
  }

  addItem() {
    this.items = [
      ...this.items,
      { id: Date.now(), product: 'Novo produto', quantity: 1, unit: 'unid.', unitPrice: 0 }
    ];
  }

  removeItem(id: number) {
    this.items = this.items.filter(item => item.id !== id);
  }

  trackItemById(index: number, item: QuoteItem) {
    return item.id;
  }

  generateBaseQuote(source: string) {
    this.quotes = [
      {
        label: 'Orçamento A',
        company: this.selectedCompany,
        items: this.items.map(item => ({ ...item })),
        adjustment: 0,
        source
      }
    ];
  }

  generateDerivedQuote() {
    if (!this.baseQuote) return;

    const company = this.companies.find(entry => entry.id === this.derivedCompanyId) ?? this.companies[0];

    this.quotes = [
      ...this.quotes,
      {
        label: this.nextLabel,
        company,
        items: this.previewItems.map(item => ({ ...item })),
        adjustment: this.adjustment,
        source: 'Derivado do Orçamento A'
      }
    ];
  }

  async handleExcel(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, string | number>>(firstSheet);

    const parsedItems = rows.map((row, index) => ({
      id: Date.now() + index,
      product: String(row['produto'] ?? row['Produto'] ?? row['product'] ?? `Produto ${index + 1}`),
      quantity: Number(row['quantidade'] ?? row['Quantidade'] ?? row['QUANT'] ?? row['Quant'] ?? row['quantity'] ?? 1),
      unit: String(row['unidade'] ?? row['Unidade'] ?? row['UNID'] ?? row['UNID.'] ?? row['unit'] ?? 'unid.'),
      unitPrice: Number(row['valor'] ?? row['Valor'] ?? row['preco'] ?? row['Preço'] ?? row['VALOR UNIT'] ?? row['Valor Unit'] ?? 0)
    }));

    const companyName = String(rows[0]?.['empresa'] ?? rows[0]?.['Empresa'] ?? '');

    const matchedCompany = this.companies.find(company =>
      company.name.toLowerCase().includes(companyName.toLowerCase())
    );

    if (matchedCompany) this.companyId = matchedCompany.id;
    if (parsedItems.length) this.items = parsedItems;

    this.excelStatus = `${parsedItems.length} item(ns) importados de ${file.name}.`;
  }

  async downloadQuote(quote: Quote) {
    try {
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
      worksheet.getCell('B11').value = '28 dias';
      worksheet.getCell('D11').value = '5 Dias Úteis';

      this.fillProducts(worksheet, quote.items);

      const arquivo = await workbook.xlsx.writeBuffer();

      saveAs(
        new Blob([arquivo]),
        `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
      );
    } catch (error) {
      console.error('Erro ao gerar Excel:', error);
    }
  }

fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
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

  for (let index = 0; index < quantidadeProdutos; index++) {
    const linha = worksheet.getRow(linhaInicial + index);

    linha.height = alturaProduto;

    for (let colNumber = 1; colNumber <= 6; colNumber++) {
      const cell = linha.getCell(colNumber);
      const modelCell = linhaModelo.getCell(colNumber);

      cell.value = null;
      this.copyCellStyle(modelCell, cell);
      this.applyTableBorder(cell);
    }

    const item = items[index];

    linha.getCell(1).value = index + 1;
    linha.getCell(2).value = item.quantity;
    linha.getCell(3).value = item.unit;
    linha.getCell(4).value = item.product;
    linha.getCell(5).value = item.unitPrice;
    linha.getCell(6).value = item.quantity * item.unitPrice;

    linha.getCell(5).numFmt = 'R$ #,##0.00';
    linha.getCell(6).numFmt = 'R$ #,##0.00';

    linha.commit();
  }

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

  worksheet.getCell(`E${linhaTotal}`).value = 'Total Geral';
  worksheet.getCell(`F${linhaTotal}`).value = total;

  worksheet.getCell(`E${linhaTotal}`).font = { bold: true, size: 14 };
  worksheet.getCell(`F${linhaTotal}`).font = { bold: true, size: 14 };

  worksheet.getCell(`E${linhaTotal}`).alignment = {
    horizontal: 'right',
    vertical: 'middle'
  };

  worksheet.getCell(`F${linhaTotal}`).alignment = {
    horizontal: 'right',
    vertical: 'middle'
  };

  worksheet.getCell(`F${linhaTotal}`).numFmt = 'R$ #,##0.00';

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
  copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
    target.style = { ...source.style };

    if (source.numFmt) {
      target.numFmt = source.numFmt;
    }

    if (source.alignment) {
      target.alignment = { ...source.alignment };
    }

    if (source.border) {
      target.border = { ...source.border };
    }

    if (source.font) {
      target.font = { ...source.font };
    }

    if (source.fill) {
      target.fill = { ...source.fill };
    }
  }

  applyTableBorder(cell: ExcelJS.Cell) {
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  }

  async insertCompanyLogo(workbook: ExcelJS.Workbook, worksheet: ExcelJS.Worksheet, logoPath: string) {
    const logoResponse = await fetch(logoPath);

    if (!logoResponse.ok) {
      throw new Error(`Logo não encontrada: ${logoPath}`);
    }

    const logoBlob = await logoResponse.blob();
    const logoBase64 = await this.blobToBase64(logoBlob);

    const logoId = workbook.addImage({
      base64: logoBase64,
      extension: this.getImageExtension(logoPath)
    });

    worksheet.addImage(logoId, {
      tl: { col: 0, row: 0 } as any,
      br: { col: 3, row: 4 } as any
    });
  }

  blobToBase64(blob: Blob): Promise<string> {
    return new Promise(resolve => {
      const reader = new FileReader();

      reader.onloadend = () => resolve(reader.result as string);
      reader.readAsDataURL(blob);
    });
  }

  getImageExtension(path: string): 'png' | 'jpeg' {
    return path.toLowerCase().endsWith('.jpg') || path.toLowerCase().endsWith('.jpeg') ? 'jpeg' : 'png';
  }
}