import { ChangeDetectorRef, Component, inject } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { CommonModule } from '@angular/common';
import { ExcelImporterService } from '../../services/excel-importer.service';
import { FormsModule } from '@angular/forms';
import {
  ArrowDown,
  ArrowUp,
  Building2,
  Check,
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
import { CustomerData, Quote, QuoteItem } from '../../model/orcamento.model';
import { ExcelComparativoService, type ComparativeQuote } from '../../services/excel-comparativo.service';
import { ExcelGeneratorService } from '../../services/excel-generator.service';
import { COMPANIES, type Company } from '../../data/empresas';
import { mapProductName } from '../../config/product-mapping';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule, FormsModule, LucideAngularModule],
  templateUrl: './orcamentos.html',
  styleUrl: './orcamentos.scss'
})
export class orcamentos {
  private excelComparativo = inject(ExcelComparativoService);
  private excelGenerator = inject(ExcelGeneratorService);
  private excelImporter = inject(ExcelImporterService);
  private changeDetector = inject(ChangeDetectorRef);
  private readonly importedBaseCompanyId = 'Laguna';

  ArrowDown = ArrowDown;
  ArrowUp = ArrowUp;
  Building2 = Building2;
  Check = Check;
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
  isImportedBaseLocked = false;
  deliveryTime = '';
  paymentMethod = '';
  customerData: CustomerData | undefined;
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
    if (this.isImportedBaseLocked) {
      return this.companies.find(company => company.id === this.importedBaseCompanyId) ?? this.companies[0];
    }

    return this.companies.find(company => company.id === this.companyId) ?? this.companies[0];
  }

  get baseQuote() {
    return this.quotes[0];
  }

  get previewItems() {
    if (!this.baseQuote) return [];

    return this.adjustedItems(this.baseQuote.items, this.derivedAdjustment);
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
    if (this.isImportedBaseLocked) return;

    const mappedValue = field === 'product' ? mapProductName(value) : value;

    this.items = this.items.map(item =>
      item.id === id
        ? {
            ...item,
            [field]: this.isTextItemField(field) ? mappedValue : this.parsePositiveNumber(value)
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
    if (this.isImportedBaseLocked) return;

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
    if (this.isImportedBaseLocked) return;

    this.items = this.items.filter(item => item.id !== id);
  }

  setDerivedAdjustment(value: number | string) {
    this.derivedAdjustment = Number(value) || 0;
  }

  syncQuoteDeliveryTime(value: string) {
    this.deliveryTime = value;
    this.quotes = this.quotes.map(quote => ({
      ...quote,
      deliveryTime: value
    }));
  }

  syncQuotePaymentMethod(value: string) {
    this.paymentMethod = value;
    this.quotes = this.quotes.map(quote => ({
      ...quote,
      paymentMethod: value
    }));
  }

  openEditQuote(index: number) {
    if (this.isImportedBaseQuoteIndex(index)) return;

    const quote = this.quotes[index];

    if (!quote) return;

    this.editingQuoteIndex = index;
    this.editingQuote = {
      ...quote,
      items: quote.items.map(item => ({ ...item })),
      date: quote.date ? new Date(quote.date) : undefined
    };
  }

  closeEditQuote() {
    this.editingQuoteIndex = null;
    this.editingQuote = null;
  }

  updateEditingItem(id: number, field: keyof QuoteItem, value: string) {
    if (!this.editingQuote) return;

    const mappedValue = field === 'product' ? mapProductName(value) : value;

    this.editingQuote = {
      ...this.editingQuote,
      items: this.editingQuote.items.map(item =>
        item.id === id
          ? {
              ...item,
              [field]: this.isTextItemField(field) ? mappedValue : this.parsePositiveNumber(value)
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

  setEditingAdjustment(value: number | string) {
    if (!this.editingQuote) return;

    const nextAdjustment = Math.max(-99, Number(value) || 0);
    const currentAdjustment = this.editingQuote.adjustment || 0;
    const currentMultiplier = 1 + currentAdjustment / 100;
    const nextMultiplier = 1 + nextAdjustment / 100;
    const adjustmentFactor = currentMultiplier === 0 ? nextMultiplier : nextMultiplier / currentMultiplier;

    this.editingQuote = {
      ...this.editingQuote,
      adjustment: nextAdjustment,
      items: this.editingQuote.items.map(item => ({
        ...item,
        unitPrice: Number((item.unitPrice * adjustmentFactor).toFixed(2))
      }))
    };
  }

  updateEditingDate(value: string) {
    if (!this.editingQuote) return;

    const newDate = value ? new Date(value) : undefined;
    this.editingQuote = {
      ...this.editingQuote,
      date: newDate
    };
  }

  getFormattedDate(date?: Date): string {
    if (!date) return '';
    return date.toISOString().split('T')[0];
  }

  saveEditingQuote() {
    if (!this.editingQuote) return;

    const editedItems = this.editingQuote.items.map(item => ({ ...item }));
    const editedAdjustment = this.editingQuote.adjustment;
    const editedDate = this.editingQuote.date;

    if (this.editingQuoteIndex === null) return;

    if (this.editingQuoteIndex === 0) {
      this.quotes = this.quotes.map((quote, index) => ({
        ...quote,
        adjustment: index === 0 ? editedAdjustment : quote.adjustment,
        items: index === 0 ? editedItems : this.adjustedItems(editedItems, quote.adjustment),
        date: index === 0 ? editedDate : quote.date
      }));

      this.closeEditQuote();
      return;
    }

    this.quotes = this.quotes.map((quote, index) =>
      index === this.editingQuoteIndex
        ? {
            ...quote,
            adjustment: editedAdjustment,
            items: editedItems,
            date: editedDate
          }
        : quote
    );

    this.closeEditQuote();
  }

  trackItemById(index: number, item: QuoteItem) {
    return item.id;
  }

  isImportedBaseQuoteIndex(index: number) {
    return index === 0 && this.quotes[0]?.source === 'Excel';
  }

  generateBaseQuote(source: string) {
    if (this.isImportedBaseLocked) {
      this.companyId = this.importedBaseCompanyId;
    }

    this.quotes = [
      {
        label: 'Orçamento A',
        company: this.selectedCompany,
        deliveryTime: this.deliveryTime,
        paymentMethod: this.paymentMethod,
        customerData: this.customerData,
        items: this.adjustedItems(this.items, this.baseAdjustment),
        adjustment: this.baseAdjustment,
        source: this.isImportedBaseLocked ? 'Excel' : source,
        date: new Date()
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
        customerData: this.baseQuote.customerData,
        items: this.previewItems.map(item => ({ ...item })),
        adjustment: this.derivedAdjustment,
        source: 'Derivado do Orçamento A',
        date: this.baseQuote.date ? new Date(this.baseQuote.date) : undefined
      }
    ];
  }

  deleteQuote(index: number) {
    if (index === 0) {
      this.quotes = [];
      this.isImportedBaseLocked = false;
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
      this.customerData = imported.customerData;
      this.mode = 'excel';
      this.companyId = this.importedBaseCompanyId;
      this.isImportedBaseLocked = true;

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
      await this.excelGenerator.generateQuoteExcel({
        ...quote,
        deliveryTime: quote.deliveryTime || this.deliveryTime,
        paymentMethod: quote.paymentMethod || this.paymentMethod
      });
    } catch (error) {
      console.error('Erro ao gerar Excel:', error);
    }
  }

  async downloadComparativeQuotes() {
    if (this.quotes.length < 2) return;

    const comparativeQuotes = this.quotes.map<ComparativeQuote>(quote => ({
      label: quote.label,
      company: quote.company,
      customerData: quote.customerData,
      paymentTerm: quote.paymentMethod || this.paymentMethod,
      deliveryTerm: quote.deliveryTime || this.deliveryTime,
      adjustment: quote.adjustment,
      items: quote.items.map(item => ({
        id: String(item.id),
        product: item.product,
        quantity: item.quantity,
        unitPrice: item.unitPrice,
        total: this.getItemTotal(item)
      }))
    }));

    try {
      await this.excelComparativo.gerarPlanilhaComparativa(comparativeQuotes);
    } catch (error) {
      console.error('Erro ao gerar planilha comparativa:', error);
      alert(error instanceof Error ? error.message : 'Erro ao gerar planilha comparativa.');
    }
  }

  private isTextItemField(field: keyof QuoteItem) {
    return field === 'product' || field === 'unit' || field === 'code' || field === 'reference';
  }

  private parsePositiveNumber(value: string) {
    const normalizedValue = value.replace(',', '.');
    return Math.max(0, Number(normalizedValue) || 0);
  }

  private getItemTotal(item: QuoteItem) {
    return [item].reduce(this.sumItems, 0);
  }
}
