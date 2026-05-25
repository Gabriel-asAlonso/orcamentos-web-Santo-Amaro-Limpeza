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

    this.clearWorksheetRowGrouping(worksheet);
    worksheet.views = [{ showGridLines: false }];
    worksheet.getColumn('B').width = (worksheet.getColumn('B').width ?? 8.43) + 2;
    // Expandir coluna E para acomodar "PRAZO DE ENTREGA"
    worksheet.getColumn('E').width = (worksheet.getColumn('E').width ?? 12) + 3;

    const quoteDate = quote.date ? new Date(quote.date) : new Date();
    const customer = quote.customerData;

    // ===== DATA DO ORÇAMENTO =====
    // E6 = "DATA: " (rótulo) | F6 = data
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('F6');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== DADOS DO CLIENTE =====
    // Apenas: NOME (C6), CNPJ (C7), END. (C8), At./Contato (C9)
    if (customer) {
      if (customer.client) {
        worksheet.getCell('C6').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.document) {
        worksheet.getCell('C7').value = String(customer.document).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        worksheet.getCell('C8').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.contact) {
        worksheet.getCell('C9').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    // Procurar dinamicamente pelos campos no template
    const formaPagAddress = this.findCellContaining(worksheet, 'forma de pag');
    const prazoAddress = this.findCellContaining(worksheet, 'prazo');

    // FORMA DE PAG (suporta células mescladas)
    if (formaPagAddress) {
      const formaPagCell = worksheet.getCell(formaPagAddress);
      const paymentMethodText = quote.paymentMethod ?? '';

      // Limpar célula antes de atribuir novo valor
      formaPagCell.value = null;
      formaPagCell.value = `FORMA DE PAG.: ${paymentMethodText}`;
      formaPagCell.alignment = {
        wrapText: true,
        vertical: 'top',
        horizontal: 'left'
      };
      formaPagCell.font = { bold: false, size: 10 };
    }

    // PRAZO DE ENTREGA (suporta células mescladas)
    if (prazoAddress) {
      const prazoCell = worksheet.getCell(prazoAddress);
      const deliveryTimeText = quote.deliveryTime ?? '';

      // Limpar célula antes de atribuir novo valor
      prazoCell.value = null;
      prazoCell.value = `PRAZO DE ENTREGA: ${deliveryTimeText}`;
      prazoCell.alignment = {
        wrapText: true,
        vertical: 'top',
        horizontal: 'left'
      };
      prazoCell.font = { bold: false, size: 10 };
    }

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 14;
    // Encontra dinamicamente onde está o SUB-TOTAL (primeira linha da seção de totais)
    // Estrutura: produtos | SUB-TOTAL | DESCONTO | FRETE | TOTAL
    const linhaSubTotal =
      this.findRowContaining(worksheet, 'sub-total', linhaInicial, 200) ??
      this.findRowContaining(worksheet, 'subtotal', linhaInicial, 200) ??
      65;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaSubTotal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar dois estilos: linha branca (14) e linha cinza (19) para alternância zebra
    // Agora são 5 colunas: B (ITEM), C (QTDE), D (DESCRIÇÃO), E (VALOR UNIT), F (VALOR TOTAL)
    const estiloBranco = Array.from({ length: 5 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(14).getCell(c + 2))
    );
    const estiloCinza = Array.from({ length: 5 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(19).getCell(c + 2))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Ajustar quantidade de linhas (preserva total + rodapé que estão em 65-70)
    if (linhasParaAdicionar > 0) {
      this.spliceRowsInChunks(
        worksheet,
        linhaSubTotal,
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

      // Alternância zebra: linhas pares (0, 2, 4...) brancas, ímpares (1, 3, 5...) cinzas
      const estiloLinha = index % 2 === 0 ? estiloBranco : estiloCinza;

      // Aplicar estilo capturado - colunas B (2) a F (6)
      for (let col = 2; col <= 6; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloLinha[col - 2]);
      }

      // Preencher dados (B=Item, C=Qtde, D=Descrição, E=Vl Unit, F=Vl Total)
      linha.getCell(2).value = index + 1;                       // B - ITEM
      linha.getCell(3).value = Math.round(item.quantity ?? 0);  // C - QTDE (sempre inteiro)
      linha.getCell(4).value = item.product ?? '';    // D - DESCRIÇÃO
      linha.getCell(5).value = item.unitPrice;        // E - VALOR UNIT
      linha.getCell(6).value = totalProduto;          // F - VALOR TOTAL

      // Formatação numérica
      linha.getCell(3).numFmt = '0';
      // Formato Accounting: cifrão à esquerda, número à direita (justificado)
      linha.getCell(5).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
      linha.getCell(6).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

      linha.commit();
    });

    // Atualizar o SUB-TOTAL e TOTAL nas linhas após o ajuste
    const novaLinhaSubTotal = linhaInicial + quantidadeProdutos;

    // Atualizar SUB-TOTAL (coluna F agora) - Accounting format
    const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
    const subTotalCell = worksheet.getCell(`F${novaLinhaSubTotal}`);
    subTotalCell.value = total;
    subTotalCell.numFmt = accountingFmt;

    // Atualizar TOTAL (linha SUB-TOTAL + 3, coluna F)
    const totalCell = worksheet.getCell(`F${novaLinhaSubTotal + 3}`);
    totalCell.value = total;
    totalCell.numFmt = accountingFmt;
  }

  private findCellContaining(worksheet: ExcelJS.Worksheet, searchText: string): string | undefined {
    const target = searchText.toLowerCase();
    const merges = (worksheet as any).model?.merges || [];

    for (let row = 1; row <= Math.min(worksheet.rowCount, 100); row++) {
      for (let col = 1; col <= 16; col++) {
        const cell = worksheet.getCell(row, col);
        let cellText = '';

        if (typeof cell.value === 'string') {
          cellText = cell.value;
        } else if (cell.value && typeof cell.value === 'object' && (cell.value as any).richText) {
          cellText = (cell.value as any).richText.map((x: any) => x.text).join('');
        }

        if (cellText.toLowerCase().includes(target)) {
          // Se a célula está em um merged range, retornar apenas a master cell (primeira do range)
          const mergedRange = this.findMergedRange(merges, row, col);
          if (mergedRange) {
            return mergedRange.split(':')[0]; // Retorna apenas a primeira célula (master)
          }
          return cell.address;
        }
      }
    }
    return undefined;
  }

  private findMergedRange(merges: string[], row: number, col: number): string | undefined {
    const colLetter = String.fromCharCode(64 + col);
    const cellAddress = `${colLetter}${row}`;

    // Procura se essa célula está dentro de algum merged range
    for (const merge of merges) {
      if (this.cellIsInRange(cellAddress, merge)) {
        return merge;
      }
    }
    return undefined;
  }

  private cellIsInRange(cell: string, range: string): boolean {
    const [start, end] = range.split(':');
    if (!end) return cell === start;

    const cellNum = this.cellToNumber(cell);
    const startNum = this.cellToNumber(start);
    const endNum = this.cellToNumber(end);

    return cellNum >= startNum && cellNum <= endNum;
  }

  private cellToNumber(cell: string): number {
    const match = cell.match(/([A-Z]+)(\d+)/);
    if (!match) return 0;

    const [, col, row] = match;
    let colNum = 0;
    for (let i = 0; i < col.length; i++) {
      colNum = colNum * 26 + (col.charCodeAt(i) - 64);
    }

    return colNum * 100000 + parseInt(row);
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
