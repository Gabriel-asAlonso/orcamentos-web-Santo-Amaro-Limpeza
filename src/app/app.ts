import { ChangeDetectorRef, Component, inject } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { CommonModule } from '@angular/common';
import { ExcelImporterService } from './services/excel-importer.service';
import { FormsModule } from '@angular/forms';
import {
  ArrowDown,
  ArrowUp,
  Building2,
  Download,
  FileSpreadsheet,
  FileText,
  LucideAngularModule,
  Pencil,
  Plus,
  Trash2,
  UsersRound,
  X
} from 'lucide-angular';
import { Quote, QuoteItem } from './model/orcamento.model';
import { ExcelGeneratorService } from './services/excel-generator.service';
import { COMPANIES, type Company } from './data/empresas';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule, FormsModule, LucideAngularModule],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App {
  private excelGenerator = inject(ExcelGeneratorService);
  private excelImporter = inject(ExcelImporterService);
  private changeDetector = inject(ChangeDetectorRef);

  ArrowDown = ArrowDown;
  ArrowUp = ArrowUp;
  Building2 = Building2;
  Download = Download;
  FileSpreadsheet = FileSpreadsheet;
  FileText = FileText;
  Pencil = Pencil;
  Plus = Plus;
  Trash2 = Trash2;
  UsersRound = UsersRound;
  X = X;

  companies: Company[] = COMPANIES;

  mode: 'manual' | 'excel' = 'manual';
  companyId = this.companies[0].id;
  derivedCompanyId = this.companies[1]?.id ?? this.companies[0].id;
  baseAdjustment = 0;
  derivedAdjustment = 0;
  quotes: Quote[] = [];
  excelStatus = 'Aguardando planilha com colunas: empresa, produto, quantidade, valor.';
  isImportingExcel = false;
  deliveryTime = '';
  paymentMethod = '';
  editingQuoteIndex: number | null = null;
  editingQuote: Quote | null = null;

  items: QuoteItem[] = [
    {
      id: 1,
      code: '1',
      reference: '',
      product: 'Copo térmico personalizado',
      quantity: 120,
      unit: 'peça',
      unitPrice: 10,
      discount: 0
    },
    {
      id: 2,
      code: '2',
      reference: '',
      product: 'Kit boas-vindas premium',
      quantity: 40,
      unit: 'kit',
      unitPrice: 86.5,
      discount: 0
    }
  ];

  get selectedCompany() {
    return this.companies.find(company => company.id === this.companyId) ?? this.companies[0];
  }

  get baseQuote() {
    return this.quotes[0];
  }

  get previewItems() {
    return this.baseQuote ? this.adjustedItems(this.baseQuote.items, this.derivedAdjustment) : [];
  }

  get nextLabel() {
    return `Orçamento ${String.fromCharCode(65 + this.quotes.length)}`;
  }

  get currentTotal() {
    return this.adjustedItems(this.items, this.baseAdjustment).reduce(this.sumItems, 0);
  }

  get previewTotal() {
    return this.previewItems.reduce(this.sumItems, 0);
  }

  get editingTotal() {
    return this.editingQuote?.items.reduce(this.sumItems, 0) ?? 0;
  }

  currency(value: number) {
    return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }

  adjustedItems(items: QuoteItem[], percentage: number) {
    return items.map(item => ({
      ...item,
      unitPrice: Number((item.unitPrice * (1 + percentage / 100)).toFixed(2))
    }));
  }

  updateItem(id: number, field: keyof QuoteItem, value: string) {
    this.items = this.items.map(item =>
      item.id === id
        ? {
            ...item,
            [field]: this.isTextItemField(field) ? value : this.parsePositiveNumber(value)
          }
        : item
    );
  }

  sumItems(total: number, item: QuoteItem) {
    const discount = item.discount ?? 0;
    const subtotal = item.quantity * item.unitPrice;
    const totalWithDiscount = subtotal - subtotal * (discount / 100);

    return total + totalWithDiscount;
  }

  addItem() {
    this.items = [
      ...this.items,
      {
        id: Date.now(),
        code: String(this.items.length + 1),
        reference: '',
        product: '',
        quantity: 1,
        unit: 'unid.',
        unitPrice: 0,
        discount: 0
      }
    ];
  }

  removeItem(id: number) {
    this.items = this.items.filter(item => item.id !== id);
  }

  openEditQuote(index: number) {
    const quote = this.quotes[index];

    if (!quote) return;

    this.editingQuoteIndex = index;
    this.editingQuote = {
      ...quote,
      items: quote.items.map(item => ({ ...item }))
    };
  }

  closeEditQuote() {
    this.editingQuoteIndex = null;
    this.editingQuote = null;
  }

  updateEditingItem(id: number, field: keyof QuoteItem, value: string) {
    if (!this.editingQuote) return;

    this.editingQuote = {
      ...this.editingQuote,
      items: this.editingQuote.items.map(item =>
        item.id === id
          ? {
              ...item,
              [field]: this.isTextItemField(field) ? value : this.parsePositiveNumber(value)
            }
          : item
      )
    };
  }

  addEditingItem() {
    if (!this.editingQuote) return;

    this.editingQuote = {
      ...this.editingQuote,
      items: [
        ...this.editingQuote.items,
        {
          id: Date.now(),
          code: String(this.editingQuote.items.length + 1),
          reference: '',
          product: '',
          quantity: 1,
          unit: 'unid.',
          unitPrice: 0,
          discount: 0
        }
      ]
    };
  }

  removeEditingItem(id: number) {
    if (!this.editingQuote) return;

    this.editingQuote = {
      ...this.editingQuote,
      items: this.editingQuote.items.filter(item => item.id !== id)
    };
  }

  saveEditingQuote() {
    if (this.editingQuoteIndex === null || !this.editingQuote) return;

    const editedItems = this.editingQuote.items.map(item => ({ ...item }));

    if (this.editingQuoteIndex === 0) {
      this.quotes = this.quotes.map((quote, index) => ({
        ...quote,
        items: index === 0 ? editedItems : this.adjustedItems(editedItems, quote.adjustment)
      }));

      this.closeEditQuote();
      return;
    }

    this.quotes = this.quotes.map((quote, index) =>
      index === this.editingQuoteIndex
        ? {
            ...quote,
            items: editedItems
          }
        : quote
    );

    this.closeEditQuote();
  }

  trackItemById(index: number, item: QuoteItem) {
    return item.id;
  }

  generateBaseQuote(source: string) {
    this.quotes = [
      {
        label: 'Orçamento A',
        company: this.selectedCompany,
        deliveryTime: this.deliveryTime,
        paymentMethod: this.paymentMethod,
        items: this.adjustedItems(this.items, this.baseAdjustment),
        adjustment: this.baseAdjustment,
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
        deliveryTime: this.baseQuote.deliveryTime,
        paymentMethod: this.baseQuote.paymentMethod,
        items: this.previewItems.map(item => ({ ...item })),
        adjustment: this.derivedAdjustment,
        source: 'Derivado do Orçamento A'
      }
    ];
  }

  deleteQuote(index: number) {
    if (index === 0) {
      this.quotes = [];
      return;
    }

    this.quotes = this.quotes
      .filter((_, quoteIndex) => quoteIndex !== index)
      .map((quote, quoteIndex) => ({
        ...quote,
        label: `Orçamento ${String.fromCharCode(65 + quoteIndex)}`
      }));
  }

  async handleExcel(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (!file) return;

    this.isImportingExcel = true;
    this.excelStatus = `Importando ${file.name}...`;
    this.changeDetector.markForCheck();

    try {
      const imported = await this.excelImporter.importProducts(file);

      this.items = imported.items.map((item, index) => ({
        ...item,
        code: item.code ?? String(index + 1),
        reference: item.reference ?? '',
        discount: item.discount ?? 0
      }));

      this.quotes = [];

      if (imported.deliveryTime) {
        this.deliveryTime = imported.deliveryTime;
      }

      if (imported.paymentMethod) {
        this.paymentMethod = imported.paymentMethod;
      }

      this.excelStatus = `${this.items.length} produto(s) importado(s) de ${file.name}.`;
    } catch (error) {
      alert(error instanceof Error ? error.message : 'Erro ao importar planilha.');
      this.excelStatus = 'Não foi possível importar a planilha. Confira as colunas e tente novamente.';
    } finally {
      this.isImportingExcel = false;
      input.value = '';
      this.changeDetector.markForCheck();
    }
  }

  async downloadQuote(quote: Quote) {
    try {
      await this.excelGenerator.generateQuoteExcel(quote);
    } catch (error) {
      console.error('Erro ao gerar Excel:', error);
    }
  }

  private isTextItemField(field: keyof QuoteItem) {
    return field === 'product' || field === 'unit' || field === 'code' || field === 'reference';
  }

  private parsePositiveNumber(value: string) {
    const normalizedValue = value.replace(',', '.');
    return Math.max(0, Number(normalizedValue) || 0);
  }
}