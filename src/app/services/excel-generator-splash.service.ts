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

    // ===== DATA DO ORÇAMENTO =====
    // F6 contém a data (E6 tem o rótulo "DATA:" do template)
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('F6');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== DADOS DO CLIENTE =====
    if (customer) {
      // B6:D6 mergido - cliente
      worksheet.getCell('B6').value = customer.client ?? '';
      // B7:C7 mergido - cnpj
      worksheet.getCell('B7').value = customer.document ?? '';
      // B8:F8 mergido - endereço
      worksheet.getCell('B8').value = customer.address ?? '';
      // B9:C9 mergido - A/C (contato)
      worksheet.getCell('B9').value = customer.contact ?? '';
      // E9 - telefone (D9 tem rótulo "Fone: ")
      worksheet.getCell('E9').value = customer.phone ?? '';
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    // A11:C11 mergido - texto completo "FORMA DE PAG.: <valor>"
    worksheet.getCell('A11').value = `FORMA DE PAG.:  ${quote.paymentMethod ?? ''}`;
    // D11 - texto completo "PRAZO DE ENTREGA: <valor>"
    worksheet.getCell('D11').value = `PRAZO DE ENTREGA:  ${quote.deliveryTime ?? ''}`;

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 15;
    const linhaTotalOriginal = 55; // Total está em A55:E56 (rótulo) e F55:F56 (valor)
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos da primeira linha de produto (15)
    const estiloProduto = Array.from({ length: 6 }, (_, i) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(i + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Capturar estilos do total (linha 55)
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
      linha.getCell(1).value = index + 1;            // A - Item
      linha.getCell(2).value = item.quantity;        // B - Quantidade
      linha.getCell(3).value = item.unit ?? '';      // C - Unidade
      linha.getCell(4).value = item.product ?? '';   // D - Descrição
      linha.getCell(5).value = item.unitPrice;       // E - Valor unitário
      linha.getCell(6).value = totalProduto;         // F - Valor total

      // Formatação numérica
      linha.getCell(2).numFmt = '#,##0.##';
      linha.getCell(5).numFmt = 'R$ #,##0.00';
      linha.getCell(6).numFmt = 'R$ #,##0.00';

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
    totalValor.numFmt = 'R$ #,##0.00';
    totalValor.font = { ...(totalValor.font || {}), bold: true };

    worksheet.getRow(linhaTotal).height = 20;
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
    const ultimaLinha = Math.max(worksheet.rowCount, startRow) + 50;

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
  }
}
