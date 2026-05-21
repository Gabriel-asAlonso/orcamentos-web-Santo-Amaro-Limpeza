import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { COMPANIES } from '../data/empresas';
import { CustomerData, ExcelImportResult } from '../model/orcamento.model';
import { mapProductName } from '../config/product-mapping';

type SheetRow = Array<string | number>;

@Injectable({
  providedIn: 'root'
})
export class ExcelImporterService {
  private readonly fieldLabelAliases = [
    'cliente',
    'nome',
    'fantasia',
    'endereco',
    'enderec',
    'end.',
    'end',
    'bairro',
    'cidade',
    'uf',
    'cep',
    'cnpj/cpf',
    'cnpj',
    'cpf',
    'inscr. est',
    'inscricao estadual',
    'regiao',
    'frete',
    'fone',
    'telefone',
    'whatsapp',
    'tabela preco',
    'tabela de preco',
    'contato',
    'at.',
    'at',
    'email',
    'e-mail',
    'vendedor',
    'ven. externo',
    'vendedor externo',
    'forma de pag',
    'forma de pagamento',
    'forma pagto',
    'forma pagto.',
    'forma pagamento',
    'condicao de pagamento',
    'condicoes de pagamento',
    'pagamento',
    'prazo de entrega',
    'entrega',
    'prazo de pagamento',
    'prazo pagamento',
    'data'
  ];

  async importProducts(file: File): Promise<ExcelImportResult> {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json<SheetRow>(firstSheet, {
      header: 1,
      defval: ''
    });

    if (!rows.length) {
      throw new Error('A planilha esta vazia.');
    }

    const headerIndex = rows.findIndex(row => this.isProductHeaderRow(row));

    if (headerIndex === -1) {
      throw new Error('A planilha precisa ter uma tabela com produto, quantidade e valor.');
    }

    const headers = rows[headerIndex].map(cell => this.normalize(cell));

    const codigoIndex = this.findColumn(headers, ['codigo', 'cod']);
    const referenciaIndex = this.findColumn(headers, ['referencia', 'ref']);
    const produtoIndex = this.findColumn(headers, ['produto', 'descricao do produto', 'descricao']);
    const quantidadeIndex = this.findColumn(headers, ['quantidade', 'quant', 'qtd', 'qtde']);
    // 'unit' removido propositalmente porque faz match com 'VALOR UNIT' (preço unitário)
    const unidadeIndex = this.findColumn(headers, ['unidade', 'unid.', 'unid']);
    const descontoIndex = this.findColumn(headers, ['desc %', 'desconto', 'desconto %']);
    const valorIndex = this.findColumn(headers, [
      'valor unit',
      'valor unitario',
      'vl unit',
      'vl.unit',
      'vl.unit.',
      'vl unitario',
      'valor',
      'preco'
    ]);
    const totalIndex = this.findColumn(headers, ['valor total', 'vl total', 'vl. total', 'total']);

    if (produtoIndex === -1 || quantidadeIndex === -1 || (valorIndex === -1 && totalIndex === -1)) {
      throw new Error('A planilha precisa ter as colunas: produto, quantidade e valor.');
    }

    const products = rows
      .slice(headerIndex + 1)
      .filter(row => {
        const produto = String(row[produtoIndex] ?? '').trim();
        const primeiraCelula = this.normalize(row[0]);
        const rowText = row.map(cell => this.normalize(cell)).join(' ');

        return Boolean(produto) &&
          !primeiraCelula.includes('total') &&
          !rowText.includes('total produtos') &&
          !produto.toLowerCase().includes('total geral');
      })
      .map((row, index) => {
        // Quantidade sempre inteira (nunca aceita valores quebrados como 22,22)
        const quantity = Math.round(this.parseNumber(row[quantidadeIndex]));
        const importedUnitPrice = valorIndex === -1 ? 0 : this.parseMoney(row[valorIndex]);
        const total = totalIndex === -1 ? 0 : this.parseMoney(row[totalIndex]);
        const unitPrice = importedUnitPrice || (quantity > 0 && total > 0 ? total / quantity : 0);

        return {
          id: Date.now() + index,
          code: codigoIndex === -1 ? String(index + 1) : String(row[codigoIndex] ?? '').trim(),
          reference: referenciaIndex === -1 ? '' : String(row[referenciaIndex] ?? '').trim(),
          product: mapProductName(String(row[produtoIndex] ?? '').trim()),
          quantity,
          unit: unidadeIndex === -1 ? 'unid.' : String(row[unidadeIndex] ?? 'unid.').trim() || 'unid.',
          unitPrice,
          discount: descontoIndex === -1 ? 0 : this.parsePercent(row[descontoIndex])
        };
      })
      .filter(item => item.product && item.quantity > 0);

    if (!products.length) {
      throw new Error('Nenhum produto valido foi encontrado na planilha.');
    }

    const headerRows = rows.slice(0, headerIndex);
    const customerStartIndex = this.findCustomerStartIndex(headerRows);
    const companyRows = customerStartIndex === -1 ? headerRows : headerRows.slice(0, customerStartIndex);
    const customerRows = customerStartIndex === -1 ? headerRows : headerRows.slice(customerStartIndex);

    return {
      items: products,
      customerData: this.extractCustomerData(customerRows),
      companyId: this.findCompanyId(companyRows),
      deliveryTime: this.findFieldValue(rows, [
        'prazo de entrega',
        'entrega',
        'prazo de pagamento',
        'prazo pagamento'
      ]),
      paymentMethod: this.findFieldValue(rows, [
        'forma de pag',
        'forma de pagamento',
        'forma pagto',
        'forma pagto.',
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

  private normalizeKey(value: string | number) {
    return this.normalize(value).replace(/[^a-z0-9]/g, '');
  }

  private isProductHeaderRow(row: SheetRow) {
    const keys = row.map(cell => this.normalizeKey(cell));
    const hasProduct = keys.some(key => ['produto', 'descricao', 'descricaodoproduto'].includes(key));
    const hasQuantity = keys.some(key => ['quantidade', 'quant', 'qtd', 'qtde'].includes(key));
    const hasValue = keys.some(key =>
      ['valorunit', 'valorunitario', 'vlunit', 'vlunitario', 'valor', 'preco', 'vltotal', 'valortotal'].includes(key)
    );

    return hasProduct && (hasQuantity || hasValue);
  }

  private findColumn(headers: string[], names: string[]) {
    const normalizedNames = names.map(name => this.normalize(name));
    const normalizedKeys = names.map(name => this.normalizeKey(name));

    return headers.findIndex(header =>
      normalizedNames.some(name => header === name || header.includes(name)) ||
      normalizedKeys.some(name => this.normalizeKey(header) === name)
    );
  }

  private findCustomerStartIndex(rows: SheetRow[]) {
    return rows.findIndex(row => row.some(cell => this.normalize(cell).includes('cliente')));
  }

  private findFieldValue(rows: SheetRow[], labels: string[]) {
    const normalizedLabels = labels.map(label => this.normalizeKey(label));

    for (const row of rows) {
      for (let index = 0; index < row.length; index++) {
        const cell = row[index];
        const normalizedCell = this.normalizeKey(cell);
        const hasLabel = normalizedLabels.some(label =>
          normalizedCell === label || normalizedCell.startsWith(label)
        );

        if (!hasLabel) continue;

        const valueInSameCell = this.valueAfterSeparator(cell);

        if (valueInSameCell) {
          return valueInSameCell;
        }

        const valueInNextCell = this.findNextValueBeforeAnotherLabel(row, index + 1);

        if (valueInNextCell) {
          return valueInNextCell;
        }
      }
    }

    return undefined;
  }

  private findNextValueBeforeAnotherLabel(row: SheetRow, startIndex: number) {
    for (let index = startIndex; index < row.length; index++) {
      const value = String(row[index] ?? '').trim();

      if (!value) continue;

      if (this.isFieldLabel(value)) {
        return undefined;
      }

      return value;
    }

    return undefined;
  }

  private isFieldLabel(value: string | number) {
    const normalizedValue = this.normalizeKey(value);

    return this.fieldLabelAliases.some(label => this.normalizeKey(label) === normalizedValue);
  }

  private valueAfterSeparator(value: string | number) {
    const text = String(value ?? '').trim();
    const separatorIndex = text.indexOf(':');

    if (separatorIndex === -1) {
      return undefined;
    }

    return text.slice(separatorIndex + 1).trim() || undefined;
  }

  private extractCustomerData(rows: SheetRow[]): CustomerData | undefined {
    const customerData: CustomerData = {
      client: this.findFieldValue(rows, ['cliente', 'nome']),
      fantasyName: this.findFieldValue(rows, ['fantasia']),
      address: this.findFieldValue(rows, ['endereco', 'enderec', 'end.', 'end']),
      district: this.findFieldValue(rows, ['bairro']),
      city: this.findFieldValue(rows, ['cidade']),
      state: this.findFieldValue(rows, ['uf']),
      postalCode: this.findFieldValue(rows, ['cep']),
      document: this.findFieldValue(rows, ['cnpj/cpf', 'cnpj', 'cpf']),
      stateRegistration: this.findFieldValue(rows, ['inscr. est', 'inscricao estadual']),
      region: this.findFieldValue(rows, ['regiao']),
      freight: this.findFieldValue(rows, ['frete']),
      phone: this.findFieldValue(rows, ['fone', 'telefone']),
      whatsapp: this.findFieldValue(rows, ['whatsapp']),
      priceTable: this.findFieldValue(rows, ['tabela preco', 'tabela de preco']),
      contact: this.findFieldValue(rows, ['contato', 'at.', 'at']),
      email: this.findFieldValue(rows, ['email', 'e-mail']),
      salesperson: this.findFieldValue(rows, ['vendedor']),
      externalSalesperson: this.findFieldValue(rows, ['ven. externo', 'vendedor externo'])
    };

    const cleanedData = Object.fromEntries(
      Object.entries(customerData).filter(([, value]) => Boolean(String(value ?? '').trim()))
    ) as CustomerData;

    return Object.keys(cleanedData).length ? cleanedData : undefined;
  }

  private findCompanyId(headerRows: SheetRow[]) {
    const headerText = headerRows.flat().map(cell => this.normalize(cell)).join(' ');

    return COMPANIES.find(company => {
      const companyName = this.normalize(company.name);
      const cnpj = this.normalize(company.cnpj ?? company.taxId ?? '');

      return (companyName && headerText.includes(companyName)) || (cnpj && headerText.includes(cnpj));
    })?.id;
  }

  private parseNumber(value: string | number) {
    if (typeof value === 'number') {
      return value;
    }

    return Number(
      String(value)
        .replace(/\./g, '')
        .replace(',', '.')
        .trim()
    ) || 0;
  }

  private parsePercent(value: string | number) {
    const parsed = this.parseMoney(value);

    return parsed > 0 && parsed <= 1 ? parsed * 100 : parsed;
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
