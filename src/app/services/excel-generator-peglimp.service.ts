import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorPegLimpService {
  private templatePath = 'assets/templates/cotacao-template-PegLimp.xlsx';
  // Primeira linha da tabela de produtos (logo abaixo do header em linha 13).
  // A quantidade de linhas de produto, posição do SUB-TOTAL e da linha de
  // FORMA DE PAG. / PRAZO DE ENTREGA são detectadas dinamicamente pelo
  // conteúdo do template, para se adaptar ao número real de produtos.
  private readonly LINHA_INICIAL_PRODUTOS = 14;

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
    //
    // IMPORTANTE: o template traz dados de exemplo preenchidos (CONDOMINIO
    // COTES DU RHONE, 05.127.515/0001-70, AV VEREADOR ABEL FERREIRA, ROGERIO).
    // Limpar TODAS as células antes para não deixar dados do template "vazarem"
    // quando algum campo do customer vier vazio/undefined.
    const camposCliente = ['B6', 'B7', 'B8', 'B9'];
    camposCliente.forEach(ref => {
      worksheet.getCell(ref).value = null;
    });

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

    // ===== AJUSTE DE BORDA FALTANTE =====
    // A célula A12 está sem a borda esquerda no template; aplica medium para
    // ficar consistente com a espessura das bordas vizinhas
    const a12 = worksheet.getCell('A12');
    a12.border = {
      ...(a12.border ?? {}),
      left: { style: 'medium', color: { argb: 'FF000000' } }
    };

    // ===== PRODUTOS, TOTAIS E CONDIÇÕES COMERCIAIS =====
    // fillProducts redimensiona a tabela conforme o número de produtos e
    // grava SUB-TOTAL, TOTAL e FORMA DE PAG. / PRAZO DE ENTREGA nas linhas
    // corretas, localizadas dinamicamente pelo conteúdo do template.
    this.fillProducts(
      worksheet,
      quote.items,
      quote.paymentMethod ?? '',
      quote.deliveryTime ?? ''
    );

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${quote.company.name.replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(
    worksheet: ExcelJS.Worksheet,
    items: QuoteItem[],
    paymentMethod: string,
    deliveryTime: string
  ) {
    const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';
    const linhaInicial = this.LINHA_INICIAL_PRODUTOS; // 14
    const N = items.length;

    // ===== LOCALIZAÇÃO DINÂMICA DOS MARCADORES =====
    // Buscamos no template a linha real do SUB-TOTAL e a linha real de
    // FORMA DE PAG. / PRAZO DE ENTREGA. Isso permite que o template seja
    // editado livremente (mais ou menos linhas de produto) sem que a
    // posição desses campos fique fora de lugar.
    const subTotalRow = this.findRowContaining(worksheet, 'SUB-TOTAL', linhaInicial, 300);
    const formaPagRow = this.findRowContaining(worksheet, 'FORMA DE PAG', linhaInicial, 300);

    if (!subTotalRow) {
      throw new Error('Template inválido: linha do SUB-TOTAL não encontrada');
    }
    if (!formaPagRow) {
      throw new Error('Template inválido: linha de FORMA DE PAG. não encontrada');
    }
    if (formaPagRow <= subTotalRow) {
      throw new Error('Template inválido: FORMA DE PAG. deve estar após o SUB-TOTAL');
    }

    // Capturar estilos das duas primeiras linhas de produto (padrão zebrado)
    // ANTES de qualquer splice, para reaplicar em linhas novas.
    const estiloEven: any[] = [];
    const estiloOdd: any[] = [];
    for (let col = 1; col <= 5; col++) {
      estiloEven[col] = this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(col));
      estiloOdd[col] = this.snapshotStyle(worksheet.getRow(linhaInicial + 1).getCell(col));
    }
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 20;

    // ===== REDIMENSIONAR ÁREA DE PRODUTOS PARA CABER EXATAMENTE N ITENS =====
    const linhasDisponiveis = subTotalRow - linhaInicial;
    let deltaRows = 0;

    if (N > linhasDisponiveis) {
      // Inserir linhas extras imediatamente antes do SUB-TOTAL
      const toInsert = N - linhasDisponiveis;
      this.spliceRowsInChunks(worksheet, subTotalRow, 0, Array(toInsert).fill([]));

      for (let i = 0; i < toInsert; i++) {
        const linhaAtual = subTotalRow + i;
        const novaLinha = worksheet.getRow(linhaAtual);
        novaLinha.height = alturaProduto;

        const offset = linhaAtual - linhaInicial;
        const estilos = (offset % 2 === 0) ? estiloEven : estiloOdd;
        for (let col = 1; col <= 5; col++) {
          this.applyStyle(novaLinha.getCell(col), estilos[col]);
        }
      }
      deltaRows = toInsert;
    } else if (N < linhasDisponiveis) {
      // Remover o excesso de linhas a partir da posição (linhaInicial + N)
      const toRemove = linhasDisponiveis - N;
      this.spliceRowsInChunks(worksheet, linhaInicial + N, toRemove);
      deltaRows = -toRemove;
    }

    // Após o splice, SUB-TOTAL e FORMA DE PAG. estão em novas posições
    const novaLinhaSubTotal = subTotalRow + deltaRows;
    const novaLinhaCondicoes = formaPagRow + deltaRows;

    // ===== LIMPAR LINHAS DE PRODUTO (preserva bordas/estilos) =====
    for (let r = linhaInicial; r < novaLinhaSubTotal; r++) {
      const row = worksheet.getRow(r);
      for (let c = 1; c <= 5; c++) {
        row.getCell(c).value = null;
      }
    }

    // ===== PREENCHER PRODUTOS =====
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    items.forEach((item, index) => {
      const linha = worksheet.getRow(linhaInicial + index);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.getCell(1).value = index + 1;                       // A - ITEM
      linha.getCell(2).value = Math.round(item.quantity ?? 0);  // B - QTDE
      linha.getCell(3).value = item.product ?? '';              // C - DESCRIÇÃO
      linha.getCell(4).value = item.unitPrice;                  // D - VALOR UNIT
      linha.getCell(5).value = totalProduto;                    // E - VALOR TOTAL

      linha.getCell(2).numFmt = '0';
      linha.getCell(4).numFmt = accountingFmt;
      linha.getCell(5).numFmt = accountingFmt;

      linha.commit();
    });

    // ===== TOTAIS =====
    worksheet.getCell(`E${novaLinhaSubTotal}`).value = total;
    worksheet.getCell(`E${novaLinhaSubTotal}`).numFmt = accountingFmt;
    worksheet.getCell(`E${novaLinhaSubTotal + 1}`).value = null; // DESCONTO
    worksheet.getCell(`E${novaLinhaSubTotal + 2}`).value = null; // FRETE
    worksheet.getCell(`E${novaLinhaSubTotal + 3}`).value = total;
    worksheet.getCell(`E${novaLinhaSubTotal + 3}`).numFmt = accountingFmt;

    // ===== FORMA DE PAG. / PRAZO DE ENTREGA =====
    worksheet.getCell(`A${novaLinhaCondicoes}`).value =
      `FORMA DE PAG.: ${paymentMethod}`.trim();
    worksheet.getCell(`B${novaLinhaCondicoes}`).value = null;
    worksheet.getCell(`E${novaLinhaCondicoes}`).value =
      `PRAZO DE ENTREGA: ${deliveryTime}`.trim();
    worksheet.getCell(`D${novaLinhaCondicoes}`).value = null;
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
