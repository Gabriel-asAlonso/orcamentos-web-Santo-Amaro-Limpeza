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
    const quantidadeProdutos = items.length;
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    const linhaModelo = worksheet.getRow(linhaInicial);

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalComDesconto = subtotal - subtotal * (discount / 100);
      return sum + totalComDesconto;
    }, 0);

    // Inserir linhas antes do total (empurra a linha TOTAL GERAL para baixo,
    // preservando seu estilo do template automaticamente)
    if (linhasParaAdicionar > 0) {
      worksheet.spliceRows(
        linhaTotalOriginal,
        0,
        ...Array.from({ length: linhasParaAdicionar }, () => [])
      );
    }

    // Remover linhas modelo excedentes (puxa a linha TOTAL GERAL para cima)
    if (linhasParaRemover > 0) {
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, linhasParaRemover);
    }

    // O total agora fica sempre logo abaixo do último produto
    const linhaTotal = linhaInicial + quantidadeProdutos;

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      // Copiar estilo da linha modelo (preserva formatação do template)
      for (let col = 1; col <= 12; col++) {
        const cell = linha.getCell(col);
        const modelCell = linhaModelo.getCell(col);
        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      // Manter os merges da linha (Descrição C:F e Vl. Total K:L)
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

      // Formatação numérica
      linha.getCell(8).numFmt = '0.00%';
      linha.getCell(10).numFmt = 'R$ #,##0.00';
      linha.getCell(11).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Preencher o valor do total na linha que se moveu junto com o template
    const totalRow = worksheet.getRow(linhaTotal);
    totalRow.getCell(11).value = total; // K - Total geral (K:L)
    totalRow.getCell(11).numFmt = 'R$ #,##0.00';
    totalRow.commit();

    // Limpar tudo abaixo do total (o splice quebra os merges do rodapé e
    // duplica títulos/textos; o mapeamento termina no Total)
    this.clearBelowTotal(worksheet, linhaTotal);
  }

  private clearBelowTotal(worksheet: ExcelJS.Worksheet, linhaTotal: number) {
    // O total ocupa 2 linhas (merge), preservar a linha imediatamente abaixo
    for (let rowNumber = linhaTotal + 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 12; colNumber++) {
        row.getCell(colNumber).value = null;
      }
    }
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

  private copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
    target.font = { ...source.font };
    target.alignment = { ...source.alignment };
    target.border = { ...source.border };
    target.fill = { ...source.fill };
    target.numFmt = source.numFmt;
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;
  }
}
