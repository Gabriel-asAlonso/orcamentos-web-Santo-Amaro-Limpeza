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

    // Ajustar altura das linhas para logo maior
    worksheet.getRow(4).height = 60;
    worksheet.getRow(5).height = 60;

    // Ajustar largura das colunas para logo e dados
    worksheet.getColumn('A').width = 20;
    worksheet.getColumn('B').width = 15;

    await this.insertCompanyLogo(workbook, worksheet, quote.company.logoPath);

    const customer = quote.customerData;
    const companyDocument = quote.company.cnpj || quote.company.taxId || '';
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== CABEÇALHO (Rows 3-4) =====
    // Data: G3 (título), H3 (dado)
    worksheet.getCell('H3').value = quoteDate.toLocaleDateString('pt-BR');
    worksheet.getCell('H3').font = { size: 9 };

    // CNPJ da empresa: G4 (título), H4 (dado)
    worksheet.getCell('H4').value = companyDocument;
    worksheet.getCell('H4').font = { size: 9 };

    // ===== DADOS DO CLIENTE LADO ESQUERDO (Rows 7-12) =====
    if (customer) {
      // Row 7: Cliente - A7 (título), B7 (dado)
      worksheet.getCell('B7').value = customer.client ?? '';

      // Row 8: Endereço - A8 (título), B8 (dado)
      worksheet.getCell('B8').value = customer.address ?? '';

      // Row 9: Cidade - A9 (título), B9 (dado)
      worksheet.getCell('B9').value = customer.city ?? '';

      // Row 10: CNPJ/CPF - A10 (título), B10 (dado)
      worksheet.getCell('B10').value = customer.document ?? '';

      // Row 11: Contato - A11 (título), B11 (dado)
      worksheet.getCell('B11').value = customer.contact ?? '';

      // Row 12: Email - A12 (título), B12 (dado)
      worksheet.getCell('B12').value = customer.email ?? '';

      // ===== DADOS DO CLIENTE LADO DIREITO (Rows 7-12) =====
      // Row 7: Fantasia - G7 (título), I7 (dado)
      worksheet.getCell('I7').value = customer.fantasyName ?? customer.client ?? '';

      // Row 8: Bairro - G8 (título), I8 (dado)
      worksheet.getCell('I8').value = customer.district ?? '';

      // Row 9: UF / CEP - G9 (título), I9 (dado)
      worksheet.getCell('I9').value = `${customer.state ?? ''} / ${customer.postalCode ?? ''}`;

      // Row 10: Inscr. Est. - G10 (título), I10 (dado)
      worksheet.getCell('I10').value = customer.stateRegistration ?? '';

      // Row 11: Telefone / WhatsApp - G11 (título), I11 (dado)
      worksheet.getCell('I11').value = customer.phone ?? '';

      // Row 12: Região - G12 (título), I12 (dado)
      worksheet.getCell('I12').value = customer.region ?? '';

      // ===== INFORMAÇÕES ADICIONAIS (Rows 16-18) =====
      // Row 16: Frete - A16 (título), B16 (dado)
      worksheet.getCell('B16').value = customer.freight ?? '';

      // Row 16: Forma de pagamento - E16 (título), G16 (dado)
      worksheet.getCell('G16').value = quote.paymentMethod;

      // Row 17: Tabela de preço - A17 (título), B17 (dado)
      worksheet.getCell('B17').value = customer.priceTable ?? '';

      // Row 17: Vendedor - E17 (título), G17 (dado)
      worksheet.getCell('G17').value = customer.salesperson ?? '';

      // Row 18: Vendedor externo - A18 (título), B18 (dado)
      worksheet.getCell('B18').value = customer.externalSalesperson ?? '';

      // Row 18: Observação - E18 (título), G18 (dado)
      // worksheet.getCell('G18').value = quote.observation ?? ''; // Se tiver campo de observação
    }

    // ===== PRODUTOS (Rows 22+) =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 22;
    const quantidadeProdutos = items.length;
    const linhasModelo = 1; // Uma linha modelo (row 22)

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalComDesconto = subtotal - subtotal * (discount / 100);
      return sum + totalComDesconto;
    }, 0);

    // Calcular linhas a adicionar/remover
    const linhasParaAdicionar = Math.max(quantidadeProdutos - linhasModelo, 0);
    const linhasParaRemover = Math.max(linhasModelo - quantidadeProdutos, 0);
    const linhaModelo = worksheet.getRow(linhaInicial);

    // Adicionar linhas se necessário
    if (linhasParaAdicionar > 0) {
      worksheet.spliceRows(linhaInicial + linhasModelo, 0, ...Array.from({ length: linhasParaAdicionar }, () => []));
    }

    // Remover linhas se necessário
    if (linhasParaRemover > 0) {
      worksheet.spliceRows(linhaInicial + quantidadeProdutos, linhasParaRemover);
    }

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = 16;

      // Copiar estilo da linha modelo
      for (let col = 1; col <= 12; col++) {
        const cell = linha.getCell(col);
        const modelCell = linhaModelo.getCell(col);
        cell.value = null;
        this.copyCellStyle(modelCell, cell);
      }

      // Preencher dados
      linha.getCell(1).value = item.code ?? index + 1; // A - Código
      linha.getCell(2).value = item.reference ?? ''; // B - Referência
      linha.getCell(3).value = item.product; // C - Descrição
      linha.getCell(7).value = item.unit; // G - Unid.
      linha.getCell(8).value = discount / 100; // H - Desc %
      linha.getCell(9).value = item.quantity; // I - Qtde.
      linha.getCell(10).value = item.unitPrice; // J - Vl. Unit.
      linha.getCell(12).value = totalProduto; // L - Vl. Total

      // Formatação - tudo alinhado à esquerda e sem negrito
      for (let col of [1, 2, 3, 7, 8, 9, 10, 12]) {
        linha.getCell(col).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        linha.getCell(col).font = { size: 9, bold: false };
      }

      // Formatação numérica
      linha.getCell(8).numFmt = '0.00%';
      linha.getCell(10).numFmt = 'R$ #,##0.00';
      linha.getCell(12).numFmt = 'R$ #,##0.00';

      linha.commit();
    });

    // Limpar linhas após os produtos para evitar duplicação
    const lastProductRow = linhaInicial + quantidadeProdutos;

    // Remover conteúdo das linhas vazias após os produtos até linha 28 (antes do total)
    for (let row = lastProductRow + 1; row < 29; row++) {
      const emptyRow = worksheet.getRow(row);
      for (let col = 1; col <= 12; col++) {
        emptyRow.getCell(col).value = null;
      }
    }

    // Colocar total apenas em L29 (célula fixa)
    worksheet.getCell('L29').value = total;
    worksheet.getCell('L29').font = { size: 9, bold: false };
    worksheet.getCell('L29').numFmt = 'R$ #,##0.00';
    worksheet.getCell('L29').alignment = { horizontal: 'left', vertical: 'middle' };
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

      // Logo em A4:B5 - maior e centralizada
      worksheet.addImage(logoImage, 'A4:B5');
    } catch {
      // Logo não encontrada, continuar sem ela
    }
  }
}
