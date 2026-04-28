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
  Plus,
  Trash2,
  UsersRound
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
  Plus = Plus;
  Trash2 = Trash2;
  UsersRound = UsersRound;

  companies: Company[] = COMPANIES;

  mode: 'manual' | 'excel' = 'manual';
  companyId = this.companies[0].id;
  derivedCompanyId = this.companies[1]?.id ?? this.companies[0].id;
  adjustment = 0;
  quotes: Quote[] = [];
  excelStatus = 'Aguardando planilha com colunas: empresa, produto, quantidade, valor.';
  isImportingExcel = false;
  deliveryTime = '';
  paymentMethod = '';

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
      { id: Date.now(), product: '', quantity: 1, unit: 'unid.', unitPrice: 0 }
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
        deliveryTime: this.deliveryTime,
        paymentMethod: this.paymentMethod,
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
        deliveryTime: this.baseQuote.deliveryTime,
        paymentMethod: this.baseQuote.paymentMethod,
        items: this.previewItems.map(item => ({ ...item })),
        adjustment: this.adjustment,
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

      this.items = imported.items;
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
}
