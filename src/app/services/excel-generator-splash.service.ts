import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorSplashService {
  private templatePath = 'assets/templates/cotacao-template-Splash.xlsx';

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
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // DEBUG: imprime apenas os 3 campos que o Splash usa
    console.log('[Splash] Customer recebido:', {
      client: customer?.client,
      document: customer?.document,
      address: customer?.address
    });
    console.log('[Splash] Quote info:', {
      label: quote.label,
      company: quote.company.id,
      source: (quote as any).source
    });

    // ===== DATA DO ORÇAMENTO =====
    // F9 contém a data (E9 tem o rótulo "DATA:" do template)
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('F9');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== DADOS DO CLIENTE =====
    // Apenas sobrescreve se o customer tem valor real para o campo.
    // Se o campo está vazio, mantém o conteúdo do template (que pode
    // ter dados de exemplo ou ficar vazio se o template foi limpo).
    if (customer) {
      if (customer.client) {
        const b9 = worksheet.getCell('B9');
        // Limpa caracteres de quebra de linha (\r, \n) que podem
        // vir no nome importado e causar quebra visual
        b9.value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
        b9.alignment = {
          ...(b9.alignment || {}),
          wrapText: false,
          vertical: 'middle'
        };
      }
      if (customer.document) {
        worksheet.getCell('B10').value = String(customer.document).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        const b11 = worksheet.getCell('B11');
        b11.value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
        b11.alignment = {
          ...(b11.alignment || {}),
          wrapText: false,
          vertical: 'middle'
        };
      }
    }

    // Garantir altura adequada para as linhas do cliente
    worksheet.getRow(9).height = 18;
    worksheet.getRow(10).height = 16;
    worksheet.getRow(11).height = 16;
    worksheet.getRow(12).height = 16;

    // ===== CONDIÇÕES COMERCIAIS =====
    // A14:D14 mergido - rótulo em negrito + valor sem negrito (rich text)
    const formaPagCell = worksheet.getCell('A14');
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

    // E14:F14 mergido - rótulo em negrito + valor sem negrito (rich text)
    const prazoCell = worksheet.getCell('E14');
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

    // Permitir altura automática para o texto quebrar
    worksheet.getRow(14).height = 40;

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 17;
    // Encontra dinamicamente onde está o TOTAL no template (procura "Total Geral")
    const linhaTotalOriginal = this.findRowContaining(worksheet, 'total geral', linhaInicial, 100) ?? 57;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos da primeira linha de produto (17)
    const estiloProduto = Array.from({ length: 6 }, (_, i) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(i + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Capturar estilos do total (linha 57)
    const estiloTotalLabel = this.snapshotStyle(worksheet.getCell(`A${linhaTotalOriginal}`));
    const estiloTotalValor = this.snapshotStyle(worksheet.getCell(`F${linhaTotalOriginal}`));

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Desfazer merges da região dos produtos para baixo
    this.unmergeFrom(worksheet, linhaInicial);

    // Ajustar quantidade de linhas
    if (linhasParaAdicionar > 0) {
      worksheet.spliceRows(
        linhaTotalOriginal,
        0,
        ...Array.from({ length: linhasParaAdicionar }, () => [])
      );
    }

    if (linhasParaRemover > 0) {
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, linhasParaRemover);
    }

    const linhaTotal = linhaInicial + quantidadeProdutos;

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      // Aplicar estilo capturado
      for (let col = 1; col <= 6; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      // Preencher dados
      linha.getCell(1).value = index + 1;                       // A - Item
      linha.getCell(2).value = Math.round(item.quantity ?? 0);  // B - Quantidade (sempre inteiro)
      linha.getCell(3).value = item.unit ?? '';      // C - Unidade
      linha.getCell(4).value = item.product ?? '';   // D - Descrição
      linha.getCell(5).value = item.unitPrice;       // E - Valor unitário
      linha.getCell(6).value = totalProduto;         // F - Valor total

      // Formatação numérica
      linha.getCell(2).numFmt = '0';
      // Formato accounting: cifrão à esquerda, número à direita (justificado)
      linha.getCell(5).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
      linha.getCell(6).numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

      // Alinhamento da unidade (C) - centralizado, sem quebra de linha
      linha.getCell(3).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };

      linha.commit();
    });

    // Limpar resíduos abaixo da última linha de produto
    this.clearFrom(worksheet, linhaTotal);

    // Reconstruir o total (mergido A:E para o label, F para o valor)
    this.safeMerge(worksheet, `A${linhaTotal}:E${linhaTotal}`);
    const totalLabel = worksheet.getCell(`A${linhaTotal}`);
    totalLabel.value = 'Total Geral';
    this.applyStyle(totalLabel, estiloTotalLabel);
    totalLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    totalLabel.font = { ...(totalLabel.font || {}), bold: true };

    const totalValor = worksheet.getCell(`F${linhaTotal}`);
    totalValor.value = total;
    this.applyStyle(totalValor, estiloTotalValor);
    totalValor.numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
    totalValor.font = { ...(totalValor.font || {}), bold: true };

    // ===== Fortificar bordas EXTERNAS da tabela =====
    // O template tem bordas externas medium e internas thin.
    // Como o estilo da linha 15 tem F15 com right=thin (linha interna),
    // todas as linhas copiadas ficam com right=thin. Precisamos fortificar:
    // - Borda direita (col F) de todas as linhas de produto e do total = medium
    // - Borda inferior do total = medium
    const bordaMedium = { style: 'medium' as const, color: { argb: 'FF000000' } };

    // Borda direita em toda a coluna F dos produtos (linhas 15 até linhaTotal-1)
    for (let row = linhaInicial; row <= linhaTotal - 1; row++) {
      const cell = worksheet.getRow(row).getCell(6);
      cell.border = { ...(cell.border || {}), right: bordaMedium };
    }

    // Borda inferior medium em toda a linha do total (A-F)
    for (let col = 1; col <= 6; col++) {
      const cell = worksheet.getRow(linhaTotal).getCell(col);
      cell.border = { ...(cell.border || {}), bottom: bordaMedium };
    }

    // Borda direita medium na célula F do total
    const totalRightCell = worksheet.getRow(linhaTotal).getCell(6);
    totalRightCell.border = { ...(totalRightCell.border || {}), right: bordaMedium };

    worksheet.getRow(linhaTotal).height = 20;

    // Limpar agrupamento de todas as linhas após o splice
    this.clearWorksheetRowGrouping(worksheet);
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
    style: ReturnType<ExcelGeneratorSplashService['snapshotStyle']>
  ) {
    if (style.font) cell.font = { ...style.font };
    if (style.alignment) cell.alignment = { ...style.alignment };
    if (style.border) cell.border = { ...style.border };
    if (style.fill) cell.fill = { ...style.fill };
    if (style.numFmt) cell.numFmt = style.numFmt;
  }

  private clearFrom(worksheet: ExcelJS.Worksheet, startRow: number) {
    // Usar actualRowCount para limitar às linhas com conteúdo real,
    // evitando iterar milhares de linhas vazias do template
    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.max(actualRowCount, startRow) + 5;

    for (let rowNumber = startRow; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 6; colNumber++) {
        const cell = row.getCell(colNumber);
        cell.value = null;
        cell.style = {};
      }
    }
  }

  private unmergeFrom(worksheet: ExcelJS.Worksheet, startRow: number) {
    const merges: string[] = [...(((worksheet as any).model?.merges as string[]) ?? [])];

    for (const range of merges) {
      const match = /^[A-Z]+(\d+):[A-Z]+\d+$/.exec(range);

      if (match && Number(match[1]) >= startRow) {
        try {
          worksheet.unMergeCells(range);
        } catch {}
      }
    }
  }

  private safeMerge(worksheet: ExcelJS.Worksheet, range: string) {
    try {
      worksheet.unMergeCells(range);
    } catch {}

    try {
      worksheet.mergeCells(range);
    } catch {}
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;

    // Limpar agrupamento apenas das linhas da tabela de produtos pra baixo
    // (linhas 1-16 do cabeçalho não precisam ser tocadas, evitando
    // qualquer interferência com conteúdo da empresa em D1:F4)
    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.min(Math.max(actualRowCount, 60), 200);
    for (let rowNumber = 17; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      (row as any).outlineLevel = 0;
      row.hidden = false;
    }
  }
}
