import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { ExcelImportResult } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelImporterService {
  async importProducts(file: File): Promise<ExcelImportResult> {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json<Array<string | number>>(firstSheet, {
      header: 1,
      defval: ''
    });

    if (!rows.length) {
      throw new Error('A planilha está vazia.');
    }

    const headerIndex = rows.findIndex(row =>
      row.some(cell => this.normalize(cell).includes('descrição')) ||
      row.some(cell => this.normalize(cell).includes('descricao')) ||
      row.some(cell => this.normalize(cell) === 'produto')
    );

    if (headerIndex === -1) {
      throw new Error('A planilha precisa ter uma tabela com produto, quantidade, unidade e valor.');
    }

    const headers = rows[headerIndex].map(cell => this.normalize(cell));

    const produtoIndex = this.findColumn(headers, ['produto', 'descrição do produto', 'descricao do produto', 'descrição', 'descricao']);
    const quantidadeIndex = this.findColumn(headers, ['quantidade', 'quant', 'qtd']);
    const unidadeIndex = this.findColumn(headers, ['unidade', 'unid.', 'unid', 'unit']);
    const valorIndex = this.findColumn(headers, ['valor unit', 'valor unitário', 'valor unitario', 'valor', 'preço', 'preco']);

    if (
      produtoIndex === -1 ||
      quantidadeIndex === -1 ||
      unidadeIndex === -1 ||
      valorIndex === -1
    ) {
      throw new Error('A planilha precisa ter as colunas: produto, quantidade, unidade e valor.');
    }

    const dataRows = rows.slice(headerIndex + 1);

    const products = dataRows
      .filter(row => {
        const produto = String(row[produtoIndex] ?? '').trim();
        const primeiraCelula = this.normalize(row[0]);

        return produto &&
          !primeiraCelula.includes('total') &&
          !produto.toLowerCase().includes('total geral');
      })
      .map((row, index) => ({
        id: Date.now() + index,
        product: String(row[produtoIndex] ?? '').trim(),
        quantity: Number(row[quantidadeIndex] ?? 0),
        unit: String(row[unidadeIndex] ?? 'unid.').trim(),
        unitPrice: this.parseMoney(row[valorIndex])
      }))
      .filter(item => item.product && item.quantity > 0);

    if (!products.length) {
      throw new Error('Nenhum produto válido foi encontrado na planilha.');
    }

    return {
      items: products,
      deliveryTime: this.findFieldValue(rows, ['prazo de entrega', 'entrega']),
      paymentMethod: this.findFieldValue(rows, [
        'forma de pag',
        'forma de pagamento',
        'forma pagamento',
        'condicao de pagamento',
        'condicoes de pagamento',
        'pagamento'
      ])
    };
  }

  private normalize(value: string | number) {
    return String(value ?? '')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  private findColumn(headers: string[], names: string[]) {
    return headers.findIndex(header =>
      names.some(name => header === this.normalize(name))
    );
  }

  private findFieldValue(rows: Array<Array<string | number>>, labels: string[]) {
    const normalizedLabels = labels.map(label => this.normalize(label));

    for (const row of rows) {
      for (let index = 0; index < row.length; index++) {
        const cell = row[index];
        const normalizedCell = this.normalize(cell);
        const hasLabel = normalizedLabels.some(label => normalizedCell.includes(label));

        if (!hasLabel) continue;

        const valueInSameCell = this.valueAfterSeparator(cell);

        if (valueInSameCell) {
          return valueInSameCell;
        }

        const valueInNextCell = row
          .slice(index + 1)
          .map(value => String(value ?? '').trim())
          .find(Boolean);

        if (valueInNextCell) {
          return valueInNextCell;
        }
      }
    }

    return undefined;
  }

  private valueAfterSeparator(value: string | number) {
    const text = String(value ?? '').trim();
    const separatorIndex = text.indexOf(':');

    if (separatorIndex === -1) {
      return undefined;
    }

    return text.slice(separatorIndex + 1).trim() || undefined;
  }

  private parseMoney(value: string | number) {
    if (typeof value === 'number') {
      return value;
    }

    return Number(
      String(value)
        .replace('R$', '')
        .replace(/\./g, '')
        .replace(',', '.')
        .trim()
    ) || 0;
  }
}
