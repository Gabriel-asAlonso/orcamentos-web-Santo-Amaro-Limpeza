import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorLagunaService {
  private templatePath = 'assets/templates/cotacao-template-Laguna.xlsx';

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

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const customer = quote.customerData;
    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== LATERAL ESQUERDA =====
    // Empresa: A5:C5 (título), A6:C7 (dado)
    worksheet.getCell('A6').value = quote.company.name;

    // CNPJ da empresa: A9:C9 (título), A10:C11 (dado)
    worksheet.getCell('A10').value = companyDocument;

    // Data: A13:C13 (título), A14:C15 (dado)
    worksheet.getCell('A14').value = quoteDate.toLocaleDateString('pt-BR');

    // Vendedor: A17:C17 (título), A18:C19 (dado)
    worksheet.getCell('A18').value = customer?.salesperson ?? '';

    // ===== DADOS DO CLIENTE (D5:F9) =====
    if (customer) {
      // Cliente: D5 (título), E5:F5 (dado)
      worksheet.getCell('E5').value = customer.client ?? '';

      // Fantasia: D6 (título), E6:F6 (dado)
      worksheet.getCell('E6').value = customer.fantasyName ?? customer.client ?? '';

      // Endereço: D7 (título), E7:F7 (dado)
      worksheet.getCell('E7').value = customer.address ?? '';

      // Cidade / UF: D8 (título), E8:F8 (dado)
      worksheet.getCell('E8').value = `${customer.city ?? ''} / ${customer.state ?? ''}`;

      // CNPJ / CPF: D9 (título), E9:F9 (dado)
      worksheet.getCell('E9').value = customer.document ?? '';

      // ===== CONDIÇÕES E CONTATO (G5:K9) =====
      // Frete: G5 (título), H5:K5 (dado)
      worksheet.getCell('H5').value = customer.freight ?? '';

      // Forma pgto.: G6 (título), H6:K6 (dado)
      worksheet.getCell('H6').value = quote.paymentMethod;

      // Tabela: G7 (título), H7:K7 (dado)
      worksheet.getCell('H7').value = customer.priceTable ?? '';

      // Contato: G8 (título), H8:K8 (dado)
      worksheet.getCell('H8').value = customer.contact ?? '';

      // Email: G9 (título), H9:K9 (dado)
      worksheet.getCell('H9').value = customer.email ?? '';
    }

    // ===== ITENS DO ORÇAMENTO (linha 12+) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 12;
    const linhaTotalOriginal = this.findTotalRow(worksheet) ?? 24;
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

    // Preencher produtos (colunas D=4 a K=11)
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      // Copiar estilo da linha modelo (preserva formatação do template)
      for (let col = 4; col <= 11; col++) {
        const cell = linha.getCell(col);
        const modelCell = linhaModelo.getCell(col);
        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      // Preencher dados
      linha.getCell(4).value = item.code ?? index + 1; // D - Código
      linha.getCell(5).value = item.reference ?? ''; // E - Referência
      linha.getCell(6).value = item.product; // F - Descrição
      linha.getCell(7).value = item.unit; // G - Unid.
      linha.getCell(8).value = discount / 100; // H - Desc %
      linha.getCell(9).value = item.quantity; // I - Qtde.
      linha.getCell(10).value = item.unitPrice; // J - Vl. Unit.
      linha.getCell(11).value = totalProduto; // K - Vl. Total

      // Formatação numérica
      linha.getCell(8).numFmt = '0.00%';
      linha.getCell(10).numFmt = 'R$ #,##0.00';
      linha.getCell(11).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Preencher o valor do total na linha que se moveu junto com o template
    const totalRow = worksheet.getRow(linhaTotal);
    totalRow.getCell(11).value = total; // K - Total geral
    totalRow.getCell(11).numFmt = 'R$ #,##0.00';
    totalRow.commit();

    // Deixar Observações e Resumo da Proposta vazios (o splice de linhas
    // quebra os merges e duplica os títulos por toda a linha)
    this.clearObservationsAndSummary(worksheet, linhaTotal);

    // Recolorir a lateral abaixo do Vendedor (as linhas inseridas pelo
    // spliceRows vêm sem estilo, criando a faixa branca)
    this.fillSidebarBackground(worksheet, linhaTotal);
  }

  private clearObservationsAndSummary(worksheet: ExcelJS.Worksheet, linhaTotal: number) {
    for (let rowNumber = linhaTotal + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= 12; colNumber++) {
        const cell = row.getCell(colNumber);
        const text = String(cell.value ?? '').trim().toLowerCase();

        if (text.includes('observ') || text.includes('resumo')) {
          cell.value = null;
        }
      }
    }
  }

  private fillSidebarBackground(worksheet: ExcelJS.Worksheet, linhaTotal: number) {
    const sidebarColor = 'FF4F5FB8';
    const inicio = 20; // logo abaixo do bloco do Vendedor (A18:C19)

    for (let rowNumber = inicio; rowNumber <= linhaTotal; rowNumber++) {
      for (let colNumber = 1; colNumber <= 3; colNumber++) {
        worksheet.getRow(rowNumber).getCell(colNumber).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: sidebarColor }
        };
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

  private async insertCompanyLogo(
    workbook: ExcelJS.Workbook,
    worksheet: ExcelJS.Worksheet,
    logoPath: string
  ) {
    try {
      const logoResponse = await fetch(logoPath);
      if (!logoResponse.ok) throw new Error('Logo not found');

      const logoBuffer = await logoResponse.arrayBuffer();
      const logoImage = workbook.addImage({
        buffer: logoBuffer,
        extension: 'png'
      });

      // Logo em A2:C4 - abaixo do nome Laguna
      worksheet.addImage(logoImage, 'A2:C4');
    } catch {
      // Logo não encontrada, continuar sem ela
    }
  }
}
