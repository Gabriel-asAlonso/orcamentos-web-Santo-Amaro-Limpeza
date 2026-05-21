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

    // ===== DADOS DO CABEÇALHO E CLIENTE =====
    // Vêm direto do template (linhas 1-9). Não mexer.

    // ===== CONDIÇÕES COMERCIAIS (linha 70) =====
    // B70 - rótulo + valor "FORMA DE PAG.: <valor>"
    const formaPagCell = worksheet.getCell('B70');
    const formaPagFont = formaPagCell.font || {};
    formaPagCell.value = {
      richText: [
        { text: 'FORMA DE PAG.:  ', font: { ...formaPagFont, bold: true } },
        { text: `${quote.paymentMethod ?? ''}`, font: { ...formaPagFont, bold: false } }
      ]
    };
    formaPagCell.alignment = {
      ...(formaPagCell.alignment || {}),
      wrapText: true,
      vertical: 'middle'
    };

    // F70 - rótulo + valor "PRAZO DE ENTREGA: <valor>"
    const prazoCell = worksheet.getCell('F70');
    const prazoFont = prazoCell.font || {};
    prazoCell.value = {
      richText: [
        { text: 'PRAZO DE ENTREGA:  ', font: { ...prazoFont, bold: true } },
        { text: `${quote.deliveryTime ?? ''}`, font: { ...prazoFont, bold: false } }
      ]
    };
    prazoCell.alignment = {
      ...(prazoCell.alignment || {}),
      wrapText: true,
      vertical: 'middle'
    };

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
    // Totais ficam em G65 (SUB-TOTAL), G66 (DESCONTO), G67 (FRETE), G68 (TOTAL)
    const linhaSubTotal = 65;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaSubTotal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos de TODAS as linhas modelo (14-64) preservando padrão zebra
    const estilosPorLinha: ReturnType<ExcelGeneratorPegLimpService['snapshotStyle']>[][] = [];
    for (let i = 0; i < quantidadeLinhasModelo; i++) {
      const row = worksheet.getRow(linhaInicial + i);
      const colEstilos = Array.from({ length: 6 }, (_, c) =>
        this.snapshotStyle(row.getCell(c + 2))
      );
      estilosPorLinha.push(colEstilos);
    }
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

      // Aplicar estilo da linha correspondente do template (preserva padrão zebra)
      // Se houver mais produtos que linhas modelo, usar cíclico
      const estiloIndex = index % estilosPorLinha.length;
      const estiloLinha = estilosPorLinha[estiloIndex];

      // Aplicar estilo capturado - colunas B (2) a G (7)
      for (let col = 2; col <= 7; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloLinha[col - 2]);
      }

      // Preencher dados (B=Item, C=Referência, D=Qtde, E=Descrição, F=Vl Unit, G=Vl Total)
      linha.getCell(2).value = index + 1;                    // B - ITEM
      linha.getCell(3).value = item.reference ?? item.code ?? ''; // C - REFERÊNCIA
      linha.getCell(4).value = item.quantity;                // D - QTDE
      linha.getCell(5).value = item.product ?? '';           // E - DESCRIÇÃO
      linha.getCell(6).value = item.unitPrice;               // F - VALOR UNIT
      linha.getCell(7).value = totalProduto;                 // G - VALOR TOTAL

      // Formatação numérica
      linha.getCell(4).numFmt = '0';
      linha.getCell(6).numFmt = 'R$ #,##0.00';
      linha.getCell(7).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Atualizar o SUB-TOTAL e TOTAL nas linhas após o ajuste
    // Após o splice, as linhas se deslocam: linhaSubTotal agora está em
    // (linhaInicial + quantidadeProdutos) se houve adição, ou no original se removeu
    const novaLinhaSubTotal = linhaInicial + quantidadeProdutos;

    // Atualizar SUB-TOTAL
    const subTotalCell = worksheet.getCell(`G${novaLinhaSubTotal}`);
    subTotalCell.value = total;
    subTotalCell.numFmt = 'R$ #,##0.00';

    // Atualizar TOTAL (linha SUB-TOTAL + 3)
    const totalCell = worksheet.getCell(`G${novaLinhaSubTotal + 3}`);
    totalCell.value = total;
    totalCell.numFmt = 'R$ #,##0.00';
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
    if (style.font) cell.font = { ...style.font };
    if (style.alignment) cell.alignment = { ...style.alignment };
    if (style.border) cell.border = { ...style.border };
    if (style.fill) cell.fill = { ...style.fill };
    if (style.numFmt) cell.numFmt = style.numFmt;
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
