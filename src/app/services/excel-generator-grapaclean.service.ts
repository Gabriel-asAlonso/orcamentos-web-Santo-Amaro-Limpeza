import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorGrapaCleanService {
  private templatePath = 'assets/templates/cotacao-template-GrapaClean.xlsx';

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

    this.clearWorksheetRowGrouping(worksheet);
    worksheet.views = [{ showGridLines: false }];

    const customer = quote.customerData;

    // DEBUG: imprime os campos que o GrapaClean usa
    console.log('[GrapaClean] Customer recebido:', {
      client: customer?.client,
      document: customer?.document,
      address: customer?.address,
      contact: customer?.contact,
      email: customer?.email,
      phone: customer?.phone
    });

    // ===== CABEÇALHO DA EMPRESA =====
    // Vem direto do template (linhas 1-6). NÃO MEXER.

    // ===== DADOS DO CLIENTE =====
    // NOME (C7), CNPJ (C8:F8), END. (C9), CONTATO (C10), EMAIL (G9), TELEFONE (G10)
    if (customer) {
      if (customer.client) {
        worksheet.getCell('C7').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.document) {
        worksheet.getCell('C8').value = String(customer.document).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        worksheet.getCell('C9').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.contact) {
        worksheet.getCell('C10').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.email) {
        worksheet.getCell('G9').value = customer.email;
      }
      if (customer.phone) {
        worksheet.getCell('G10').value = customer.phone;
      }
    }

    // ===== ALTERAR APENAS Data de Emissão, Cond. Pagto e Prazo de Entrega =====
    // J2 = Data de Emissão | J3 = Cond. Pagto | J4 = Prazo de Entrega
    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('J2');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    worksheet.getCell('J3').value = quote.paymentMethod ?? '';
    worksheet.getCell('J4').value = quote.deliveryTime ?? '';

    // ===== PRODUTOS (linhas 14-32) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 15;
    // Encontra dinamicamente a linha do TOTAL no template (procura "FRETE:" ou "TOTAL:")
    // Estrutura esperada: produtos | linha vazia separadora | linha de totais (FRETE/DESCONTO/TOTAL)
    const linhaTotaisOriginal =
      this.findRowContaining(worksheet, 'frete:', linhaInicial, 200) ??
      this.findRowContaining(worksheet, 'total:', linhaInicial, 200) ??
      35;
    // Última linha de produto = linha do total - 2 (uma linha vazia separadora + linha de totais)
    const ultimaLinhaProdutoModelo = linhaTotaisOriginal - 2;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = ultimaLinhaProdutoModelo - linhaInicial + 1;
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilo da linha modelo (15) - colunas B (2) a J (10)
    const estiloProduto = Array.from({ length: 9 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(c + 2))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Ajustar quantidade de linhas (inserir/remover ANTES da linha vazia 34)
    if (linhasParaAdicionar > 0) {
      this.spliceRowsInChunks(
        worksheet,
        ultimaLinhaProdutoModelo + 1, // posição 34 (após último produto modelo)
        0,
        Array.from({ length: linhasParaAdicionar }, () => [])
      );
    }

    if (linhasParaRemover > 0) {
      this.spliceRowsInChunks(
        worksheet,
        linhaInicial + quantidadeProdutos,
        linhasParaRemover
      );
    }

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      // Aplicar estilo capturado - colunas B (2) a J (10)
      for (let col = 2; col <= 10; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 2]);
      }

      // Preencher dados
      // B = Código | C = Descrição | F = Qtde | H = Valor Unit | I = Desconto | J = Valor Total
      linha.getCell(2).value = item.code ?? index + 1;        // B - Código
      linha.getCell(3).value = item.product ?? '';            // C - Descrição
      linha.getCell(6).value = Math.round(item.quantity ?? 0); // F - Qtde (sempre inteiro)
      linha.getCell(8).value = item.unitPrice;                // H - Valor Unit
      linha.getCell(9).value = discount > 0 ? discount / 100 : null; // I - Desconto
      linha.getCell(10).value = totalProduto;                 // J - Valor Total

      // Formatação numérica
      linha.getCell(6).numFmt = '0';
      linha.getCell(8).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
      if (discount > 0) linha.getCell(9).numFmt = '0.00%';
      linha.getCell(10).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

      linha.commit();
    });

    // Atualizar o TOTAL na linha (linhaInicial + quantidadeProdutos + 1, que é a posição da linha 33 deslocada)
    const novaLinhaTotal = linhaInicial + quantidadeProdutos + 1;
    const totalCell = worksheet.getCell(`J${novaLinhaTotal}`);
    totalCell.value = total;
    totalCell.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
  }

  /**
   * spliceRows do ExcelJS falha silenciosamente quando o count é grande.
   * Workaround: dividir em chunks de 10 linhas.
   */
  private spliceRowsInChunks(
    worksheet: ExcelJS.Worksheet,
    start: number,
    count: number,
    inserts: any[] = []
  ) {
    const CHUNK_SIZE = 10;

    if (count === 0 && inserts.length > 0) {
      let remaining = inserts.length;
      let offset = 0;
      while (remaining > 0) {
        const chunk = Math.min(remaining, CHUNK_SIZE);
        worksheet.spliceRows(start + offset, 0, ...inserts.slice(offset, offset + chunk));
        remaining -= chunk;
        offset += chunk;
      }
      return;
    }

    let toRemove = count;
    while (toRemove > 0) {
      const chunk = Math.min(toRemove, CHUNK_SIZE);
      worksheet.spliceRows(start, chunk);
      toRemove -= chunk;
    }
  }

  /**
   * Encontra a primeira linha entre fromRow e toRow que contém o texto
   * marker (case-insensitive). Retorna undefined se não encontrar.
   */
  private findRowContaining(
    worksheet: ExcelJS.Worksheet,
    marker: string,
    fromRow: number,
    toRow: number
  ): number | undefined {
    const target = marker.toLowerCase();
    for (let r = fromRow; r <= toRow; r++) {
      for (let c = 1; c <= 16; c++) {
        const v = worksheet.getRow(r).getCell(c).value;
        let str = '';
        if (typeof v === 'string') str = v;
        else if (v && typeof v === 'object' && (v as any).richText) {
          str = (v as any).richText.map((x: any) => x.text).join('');
        }
        if (str.toLowerCase().includes(target)) return r;
      }
    }
    return undefined;
  }

  private snapshotStyle(cell: ExcelJS.Cell) {
    return {
      font: cell.font ? { ...cell.font } : undefined,
      alignment: cell.alignment ? { ...cell.alignment } : undefined,
      border: cell.border ? { ...cell.border } : undefined,
      fill: cell.fill ? { ...cell.fill } : undefined,
      numFmt: cell.numFmt
    };
  }

  private applyStyle(
    cell: ExcelJS.Cell,
    style: ReturnType<ExcelGeneratorGrapaCleanService['snapshotStyle']>
  ) {
    const newStyle: any = { ...cell.style };
    if (style.font) newStyle.font = JSON.parse(JSON.stringify(style.font));
    if (style.alignment) newStyle.alignment = JSON.parse(JSON.stringify(style.alignment));
    if (style.border) newStyle.border = JSON.parse(JSON.stringify(style.border));
    if (style.fill) newStyle.fill = JSON.parse(JSON.stringify(style.fill));
    if (style.numFmt) newStyle.numFmt = style.numFmt;
    cell.style = newStyle;
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;

    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.min(Math.max(actualRowCount, 40), 200);
    for (let rowNumber = 12; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      (row as any).outlineLevel = 0;
      row.hidden = false;
    }
  }
}
