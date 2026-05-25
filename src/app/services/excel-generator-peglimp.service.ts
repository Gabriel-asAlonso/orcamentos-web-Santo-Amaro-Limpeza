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
    // A31 - FORMA DE PAG.
    // E31 - PRAZO DE ENTREGA (já tem o rótulo no template)

    // Atualizar apenas os valores, preservando rótulos do template
    if (quote.paymentMethod) {
      worksheet.getCell('B31').value = quote.paymentMethod;
    }

    // E31 já contém o rótulo "PRAZO DE ENTREGA:", então vamos adicionar o valor após
    if (quote.deliveryTime) {
      const prazoCell = worksheet.getCell('E31');
      const currentText = String(prazoCell.value || '');
      if (!currentText.includes(quote.deliveryTime)) {
        prazoCell.value = `${currentText} ${quote.deliveryTime}`;
      }
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
    const linhaInicial = 14;  // Primeira linha de produto
    const linhaFinal = 25;    // Última linha de produto (12 produtos)
    const linhaSubTotal = 26; // SUB-TOTAL está em linha 26

    const quantidadeProdutos = Math.max(Math.min(items.length, 12), 1); // Máximo 12 produtos

    // Capturar estilo da primeira linha de produto
    const estiloProduto = Array.from({ length: 5 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(c + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 20;

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Preencher produtos (linhas 14-25)
    items.slice(0, 12).forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

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

    // Limpar linhas não utilizadas de produtos (se menos de 12 produtos)
    for (let row = linhaInicial + quantidadeProdutos; row <= linhaFinal; row++) {
      const linha = worksheet.getRow(row);
      for (let col = 1; col <= 5; col++) {
        linha.getCell(col).value = null;
      }
    }

    // ===== TOTAIS (linhas 26-29) =====
    // Linha 26: SUB-TOTAL
    worksheet.getCell('E26').value = total;
    worksheet.getCell('E26').numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    // Linha 27: DESCONTO (deixar vazio ou com valor se houver desconto global)
    // worksheet.getCell('E27') já tem o rótulo

    // Linha 28: FRETE (deixar vazio ou com valor)
    // worksheet.getCell('E28') já tem o rótulo

    // Linha 29: TOTAL
    worksheet.getCell('E29').value = total;
    worksheet.getCell('E29').numFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
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
