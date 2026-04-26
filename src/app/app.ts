import { Component, signal } from '@angular/core';
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

type Company = {
  id: string;
  name: string;
  address: string;
  identity: string;
  taxId: string;
};

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

  logoUrl = 'assets/logo-santo-amaro-limpeza.jpg';

  companies: Company[] = [
    {
      id: 'aurora',
      name: 'Aurora Eventos Corporativos',
      address: 'Av. das Nações, 420 — São Paulo/SP',
      identity: 'Verde institucional, selo executivo e rodapé fiscal',
      taxId: '12.584.220/0001-76'
    },
    {
      id: 'norte',
      name: 'Norte Food Service',
      address: 'Rua Itapuã, 88 — Curitiba/PR',
      identity: 'Marca âmbar, cabeçalho comercial e assinatura digital',
      taxId: '08.991.310/0001-44'
    },
    {
      id: 'delta',
      name: 'Delta Facilities',
      address: 'Alameda Central, 1370 — Belo Horizonte/MG',
      identity: 'Azul técnico, tabela compacta e dados operacionais',
      taxId: '31.773.005/0001-20'
    }
  ];

  mode: 'manual' | 'excel' = 'manual';
  companyId = this.companies[0].id;
  derivedCompanyId = this.companies[1].id;
  adjustment = 15;
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

  downloadQuote(quote: Quote) {
    const totalRow = 15 + quote.items.length + 1;

    const rows = [
      ['            SANTO AMARO LIMPEZA - ORÇAMENTO'],
      ['Distribuição de produtos de limpeza, descartáveis e higiene'],
      ['Contato: comercial@santoamaro.limpeza'],
      [null, null, null, 'Contato: Comercial'],
      [],
      ['NOME:', quote.company.name, null, null, 'DATA:', new Date().toLocaleDateString('pt-BR')],
      ['CNPJ:', quote.company.taxId],
      ['END.:', quote.company.address],
      ['At.:', quote.company.identity, null, 'Fone:'],
      [],
      ['FORMA DE PAG.:  28 dias', null, null, 'PRAZO DE ENTREGA:  5 Dias Úteis'],
      [],
      [],
      ['ITEM', 'QUANT', 'UNID.', ' DESCRIÇÃO DO PRODUTO', 'VALOR UNIT', 'VALOR TOTAL'],
      ...quote.items.map((item, index) => [
        index + 1,
        item.quantity || '',
        item.unit,
        item.product,
        item.unitPrice,
        item.quantity * item.unitPrice
      ]),
      [],
      ['Total Geral', null, null, null, null, quote.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0)]
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(rows);

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 9 },
      { wch: 13 },
      { wch: 56 },
      { wch: 15 },
      { wch: 18 }
    ];

    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } },
      { s: { r: 5, c: 1 }, e: { r: 5, c: 3 } },
      { s: { r: 6, c: 1 }, e: { r: 6, c: 2 } },
      { s: { r: 7, c: 1 }, e: { r: 7, c: 5 } },
      { s: { r: 8, c: 1 }, e: { r: 8, c: 2 } },
      { s: { r: 10, c: 0 }, e: { r: 10, c: 2 } },
      { s: { r: 10, c: 3 }, e: { r: 10, c: 5 } },
      { s: { r: totalRow, c: 0 }, e: { r: totalRow + 1, c: 4 } },
      { s: { r: totalRow, c: 5 }, e: { r: totalRow + 1, c: 5 } }
    ];

    for (let row = 15; row < 15 + quote.items.length; row += 1) {
      const unitCell = worksheet[`E${row}`];
      const totalCell = worksheet[`F${row}`];

      if (unitCell) unitCell.z = 'R$ #,##0.00';
      if (totalCell) totalCell.z = 'R$ #,##0.00';
    }

    const grandTotalCell = worksheet[`F${totalRow + 1}`];

    if (grandTotalCell) grandTotalCell.z = 'R$ #,##0.00';

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Plan1');
    XLSX.writeFile(workbook, `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`);
  }
}
