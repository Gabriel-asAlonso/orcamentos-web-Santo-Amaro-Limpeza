import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorPegLimpService {
  private templatePath = 'assets/templates/cotacao-template-PegLimp.xlsx';

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

    worksheet.views = [{ showGridLines: false }];

    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    const customer = quote.customerData;

    // DEBUG: imprime os campos que o Peg Limp usa
    console.log('[Peg Limp] Novo Template - Customer recebido:', {
      client: customer?.client,
      document: customer?.document,
      address: customer?.address,
      contact: customer?.contact
    });

    // ===== DATA DO ORÇAMENTO =====
    // E6 = DATA (valor)
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('E6');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== DADOS DO CLIENTE =====
    // Novo layout: NOME (B6), CNPJ (B7), END. (B8), At./Contato (B9)
    if (customer) {
      if (customer.client) {
        worksheet.getCell('B6').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.document) {
        worksheet.getCell('B7').value = String(customer.document).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        worksheet.getCell('B8').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.contact) {
        worksheet.getCell('B9').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
    }

    // ===== CONDIÇÕES COMERCIAIS (linha 31) =====
    // B31 - FORMA DE PAG. (valor apenas, rótulo já no template em A31)
    // D31 - PRAZO DE ENTREGA (valor apenas, rótulo já no template em C31)

    // Limpar e preencher FORMA DE PAG.
    const formaPagCell = worksheet.getCell('B31');
    formaPagCell.value = quote.paymentMethod ?? '';

    // Preencher PRAZO DE ENTREGA (apenas o valor, sem concatenar)
    const prazoCell = worksheet.getCell('D31');
    prazoCell.value = quote.deliveryTime ?? '';

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 14;  // Primeira linha de produto
    const quantidadeProdutos = items.length;
    const linhasDisponíveis = 12; // Template tem 12 linhas de produto (14-25)
    const linhasParaAdicionar = Math.max(quantidadeProdutos - linhasDisponíveis, 0);

    // Capturar estilos das DUAS linhas de referência para preservar o padrão zebrado
    // Row 14 (offset par - tipicamente cinza) e Row 15 (offset ímpar - tipicamente branca)
    const estiloEven: any[] = [];
    const estiloOdd: any[] = [];
    for (let col = 1; col <= 5; col++) {
      estiloEven[col] = this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(col));
      estiloOdd[col] = this.snapshotStyle(worksheet.getRow(linhaInicial + 1).getCell(col));
    }
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 20;

    // Se houver mais de 12 produtos, inserir linhas extras ANTES do SUB-TOTAL
    // e copiar a formatação ALTERNADA (zebrada) das linhas existentes
    if (linhasParaAdicionar > 0) {
      // Inserir linhas após linha 25 (antes de onde está o SUB-TOTAL)
      worksheet.spliceRows(26, 0, ...Array(linhasParaAdicionar).fill([]));

      // Copiar formatação para as novas linhas, ALTERNANDO entre os dois estilos
      for (let i = 0; i < linhasParaAdicionar; i++) {
        const linhaAtual = 26 + i;
        const novaLinha = worksheet.getRow(linhaAtual);
        novaLinha.height = alturaProduto;

        // Determinar offset desde a primeira linha de produto para manter o padrão zebrado
        // Row 26 = offset 12 (par) → estilo igual ao da row 14
        // Row 27 = offset 13 (ímpar) → estilo igual ao da row 15
        const offset = linhaAtual - linhaInicial;
        const estilos = (offset % 2 === 0) ? estiloEven : estiloOdd;

        // Copiar estilo de cada coluna
        for (let col = 1; col <= 5; col++) {
          const cell = novaLinha.getCell(col);
          this.applyStyle(cell, estilos[col]);
        }
      }
    }

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // ===== PREENCHER PRODUTOS =====
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      // Preencher dados: A=Item, B=Qtde, C=Descrição, D=Valor Unit, E=Valor Total
      linha.getCell(1).value = index + 1;                       // A - ITEM
      linha.getCell(2).value = Math.round(item.quantity ?? 0);  // B - QTDE
      linha.getCell(3).value = item.product ?? '';              // C - DESCRIÇÃO
      linha.getCell(4).value = item.unitPrice;                  // D - VALOR UNIT
      linha.getCell(5).value = totalProduto;                    // E - VALOR TOTAL

      // Formatação numérica
      linha.getCell(2).numFmt = '0';
      // Formato Accounting: cifrão à esquerda, número à direita
      linha.getCell(4).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
      linha.getCell(5).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

      linha.commit();
    });

    // ===== TOTAIS (agora em linha dinâmica) =====
    const novaLinhaSubTotal = linhaInicial + quantidadeProdutos;

    // SUB-TOTAL
    worksheet.getCell(`E${novaLinhaSubTotal}`).value = total;
    worksheet.getCell(`E${novaLinhaSubTotal}`).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    // DESCONTO (linha seguinte)
    worksheet.getCell(`E${novaLinhaSubTotal + 1}`).value = null;

    // FRETE (linha seguinte)
    worksheet.getCell(`E${novaLinhaSubTotal + 2}`).value = null;

    // TOTAL (linha seguinte)
    worksheet.getCell(`E${novaLinhaSubTotal + 3}`).value = total;
    worksheet.getCell(`E${novaLinhaSubTotal + 3}`).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
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
    style: ReturnType<ExcelGeneratorPegLimpService['snapshotStyle']>
  ) {
    // Usar cell.style = {...} ao invés de propriedades individuais.
    // O ExcelJS otimiza estilos por ID interno; quando setamos propriedades
    // individuais (cell.fill = X), múltiplas células podem acabar compartilhando
    // o mesmo ID e o fill alternado é perdido. Setando cell.style força a
    // criação de um novo estilo único por célula.
    const newStyle: any = { ...cell.style };
    if (style.font) newStyle.font = JSON.parse(JSON.stringify(style.font));
    if (style.alignment) newStyle.alignment = JSON.parse(JSON.stringify(style.alignment));
    if (style.border) newStyle.border = JSON.parse(JSON.stringify(style.border));
    if (style.fill) newStyle.fill = JSON.parse(JSON.stringify(style.fill));
    if (style.numFmt) newStyle.numFmt = style.numFmt;
    cell.style = newStyle;
  }

  /**
   * O spliceRows do ExcelJS falha silenciosamente quando o count é grande
   * (descoberto via teste: remove apenas ~7 linhas quando se pede 47+).
   * Workaround: dividir em chunks pequenos.
   */
  private spliceRowsInChunks(
    worksheet: ExcelJS.Worksheet,
    start: number,
    count: number,
    inserts: any[] = []
  ) {
    const CHUNK_SIZE = 10;

    if (count === 0 && inserts.length > 0) {
      // Inserir linhas em chunks
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

    // Remover linhas em chunks
    let toRemove = count;
    while (toRemove > 0) {
      const chunk = Math.min(toRemove, CHUNK_SIZE);
      worksheet.spliceRows(start, chunk);
      toRemove -= chunk;
    }
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;

    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.min(Math.max(actualRowCount, 75), 200);
    for (let rowNumber = 12; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      (row as any).outlineLevel = 0;
      row.hidden = false;
    }
  }
}
