import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorSantoAmaroService {
  private templatePath = 'assets/templates/cotacao-template-santo-amaro.xlsx';

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

    const customer = quote.customerData;
    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== DADOS DO CLIENTE (esquerda) =====
    // CLIENTE: A9 (título), B9:F9 (dado)
    worksheet.getCell('B9').value = customer?.client ?? '';

    // CNPJ / CPF: A10 (título), B10:F10 (dado)
    worksheet.getCell('B10').value = customer?.document ?? '';

    // ENDEREÇO: A11 (título), B11:F11 (dado)
    worksheet.getCell('B11').value = customer?.address ?? '';

    // CIDADE / UF: A12 (título), B12:C12 (dado)
    worksheet.getCell('B12').value = `${customer?.city ?? ''} / ${customer?.state ?? ''}`;

    // TELEFONE: D12 (título), E12:F12 (dado)
    worksheet.getCell('E12').value = customer?.phone ?? '';

    // E-MAIL: A13 (título), B13:C13 (dado)
    worksheet.getCell('B13').value = customer?.email ?? '';

    // CONTATO: D13 (título), E13:F13 (dado)
    worksheet.getCell('E13').value = customer?.contact ?? '';

    // ===== RESUMO DA PROPOSTA (meio) =====
    // DATA DO ORÇAMENTO: H9 (título), I9:J9 (dado)
    worksheet.getCell('I9').value = quoteDate.toLocaleDateString('pt-BR');

    // VALIDADE: H10 (título), I10:J10 (dado) - sem origem, limpar exemplo
    worksheet.getCell('I10').value = '';

    // VENDEDOR(A): H11 (título), I11:J11 (dado)
    worksheet.getCell('I11').value = customer?.salesperson ?? '';

    // TABELA: H12 (título), I12:J12 (dado)
    worksheet.getCell('I12').value = customer?.priceTable ?? '';

    // EMPRESA: H13 (título), I13:J13 (dado)
    worksheet.getCell('I13').value = quote.company.name;

    // ===== CONDIÇÕES (direita) =====
    // CONDIÇÃO: K9 (título), L9 (dado) - sem origem, limpar exemplo
    worksheet.getCell('L9').value = '';

    // PRAZO: K10 (título), L10 (dado)
    worksheet.getCell('L10').value = quote.deliveryTime ?? '';

    // FORMA PGTO.: K11 (título), L11 (dado)
    worksheet.getCell('L11').value = quote.paymentMethod;

    // FRETE: K12 (título), L12 (dado)
    worksheet.getCell('L12').value = customer?.freight ?? '';

    // CNPJ: K13 (título), L13 (dado)
    worksheet.getCell('L13').value = companyDocument;

    // ===== ITENS DO ORÇAMENTO (linha 16+) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 16;
    const linhaTotalOriginal = this.findTotalRow(worksheet) ?? 27;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos ANTES de qualquer alteração
    const estiloProduto = Array.from({ length: 12 }, (_, i) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(i + 1))
    );
    const estiloTotalLabel = this.snapshotStyle(worksheet.getCell(`H${linhaTotalOriginal}`));
    const estiloTotalValor = this.snapshotStyle(worksheet.getCell(`K${linhaTotalOriginal}`));
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 19.5;

    // Aumentar a largura da coluna do Código (códigos longos)
    worksheet.getColumn(1).width = 18;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Remover TODOS os merges da região dinâmica (produtos + total + rodapé)
    // para o splice não duplicar/quebrar os merges multi-linha do template
    this.unmergeFrom(worksheet, linhaInicial);

    // Ajustar a quantidade de linhas
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

    // O total fica sempre logo abaixo do último produto
    const linhaTotal = linhaInicial + quantidadeProdutos;

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      // Aplicar o estilo capturado da linha modelo
      for (let col = 1; col <= 12; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      // Recriar os merges da linha (Descrição C:F e Vl. Total K:L)
      this.safeMerge(worksheet, `C${rowNumber}:F${rowNumber}`);
      this.safeMerge(worksheet, `K${rowNumber}:L${rowNumber}`);

      // Preencher dados
      linha.getCell(1).value = item.code ?? index + 1; // A - Código
      linha.getCell(2).value = item.reference ?? ''; // B - Referência
      linha.getCell(3).value = item.product; // C - Descrição (C:F)
      linha.getCell(7).value = item.unit; // G - Unid.
      linha.getCell(8).value = discount / 100; // H - Desc %
      linha.getCell(9).value = item.quantity; // I - Qtde.
      linha.getCell(10).value = item.unitPrice; // J - Vl. Unitário
      linha.getCell(11).value = totalProduto; // K - Vl. Total (K:L)

      // Formatação numérica explícita (evita herdar formato errado do modelo)
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(8).numFmt = '0.00%';
      linha.getCell(9).numFmt = '#,##0.##';
      linha.getCell(10).numFmt = 'R$ #,##0.00';
      linha.getCell(11).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Limpar total + rodapé (resíduos do template / splice)
    this.clearFrom(worksheet, linhaTotal);

    // Reconstruir o total do zero (2 linhas de altura, como no template)
    this.safeMerge(worksheet, `H${linhaTotal}:J${linhaTotal + 1}`);
    this.safeMerge(worksheet, `K${linhaTotal}:L${linhaTotal + 1}`);

    const totalLabel = worksheet.getCell(`H${linhaTotal}`);
    totalLabel.value = 'TOTAL GERAL';
    this.applyStyle(totalLabel, estiloTotalLabel);

    const totalValor = worksheet.getCell(`K${linhaTotal}`);
    totalValor.value = total;
    this.applyStyle(totalValor, estiloTotalValor);
    totalValor.numFmt = 'R$ #,##0.00';

    worksheet.getRow(linhaTotal).height = 24;
    worksheet.getRow(linhaTotal + 1).height = 24;

    // Borda à direita da tabela de produtos (cabeçalho -> total),
    // na última coluna (L), preservando as bordas existentes
    const bordaDireita: Partial<ExcelJS.Borders> = {
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };

    for (let rowNumber = linhaInicial - 1; rowNumber <= linhaTotal + 1; rowNumber++) {
      const cell = worksheet.getRow(rowNumber).getCell(12);
      cell.border = { ...cell.border, ...bordaDireita };
    }
  }

  private clearFrom(worksheet: ExcelJS.Worksheet, startRow: number) {
    for (let rowNumber = startRow; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 12; colNumber++) {
        row.getCell(colNumber).value = null;
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
    style: ReturnType<ExcelGeneratorSantoAmaroService['snapshotStyle']>
  ) {
    if (style.font) cell.font = { ...style.font };
    if (style.alignment) cell.alignment = { ...style.alignment };
    if (style.border) cell.border = { ...style.border };
    if (style.fill) cell.fill = { ...style.fill };
    if (style.numFmt) cell.numFmt = style.numFmt;
  }

  private findTotalRow(worksheet: ExcelJS.Worksheet) {
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 12; colNumber++) {
        const value = String(row.getCell(colNumber).value ?? '').trim().toLowerCase();

        if (value.includes('total geral')) {
          return rowNumber;
        }
      }
    }

    return undefined;
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
