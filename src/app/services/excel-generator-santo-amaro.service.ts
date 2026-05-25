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
    worksheet.views = [{ showGridLines: false }];

    // Limpar células com "#VALUE!" deixadas pelo recurso "Picture in Cell"
    // do Excel 365 (ExcelJS não suporta esse formato e converte em erro)
    this.clearValueErrors(worksheet);

    // Inserir a logo programaticamente (substituindo a que estava como
    // "Picture in Cell" no template). twoCellAnchor mantém a imagem
    // ancorada à grade e evita deslocamento ao exportar PDF.
    await this.insertCompanyLogo(workbook, worksheet);

    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== DATA DO ORÇAMENTO =====
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('B4');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // DEBUG: imprime o customer recebido
    console.log('[Santo Amaro] Customer recebido:', customer);
    console.log('[Santo Amaro] Quote info:', {
      label: quote.label,
      source: (quote as any).source,
      paymentMethod: quote.paymentMethod,
      deliveryTime: quote.deliveryTime
    });

    // ===== DADOS DO CLIENTE =====
    if (customer) {
      // C7:P7 mergido - Cliente
      if (customer.client) {
        worksheet.getCell('C7').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      // C8:P8 mergido - Fantasia
      if (customer.fantasyName) {
        worksheet.getCell('C8').value = String(customer.fantasyName).replace(/[\r\n]+/g, ' ').trim();
      } else if (customer.client) {
        worksheet.getCell('C8').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      // C9:L9 mergido - Endereço
      if (customer.address) {
        worksheet.getCell('C9').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      // C10:K10 mergido - Cidade
      if (customer.city) {
        worksheet.getCell('C10').value = customer.city;
      }
      // C11 - Bairro
      if (customer.district) {
        worksheet.getCell('C11').value = customer.district;
      }
      // C12 - CEP
      if (customer.postalCode) {
        worksheet.getCell('C12').value = customer.postalCode;
      }
      // C13:F13 mergido - CNPJ/CPF
      if (customer.document) {
        worksheet.getCell('C13').value = customer.document;
      }
      // C16:L16 mergido - Contato
      if (customer.contact) {
        worksheet.getCell('C16').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
      // C17:P17 mergido - Email
      if (customer.email) {
        worksheet.getCell('C17').value = customer.email;
      }
      // C18:P18 mergido - Vendedor
      if (customer.salesperson) {
        worksheet.getCell('C18').value = customer.salesperson;
      }
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    // A20 = "Forma Pagto. :" -> valor ao lado em B/C
    // A21 = "Prazo de entrega" -> valor ao lado
    if (quote.paymentMethod) {
      worksheet.getCell('C20').value = quote.paymentMethod;
    }
    if (quote.deliveryTime) {
      worksheet.getCell('C21').value = quote.deliveryTime;
    }

    // ===== PRODUTOS =====
    this.fillProducts(worksheet, quote.items);

    const arquivo = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([arquivo]),
      `${quote.label.replace(/\s+/g, '_')}_${this.capitalizeCompanyName(quote.company.name).replace(/\W+/g, '_')}.xlsx`
    );
  }

  private fillProducts(worksheet: ExcelJS.Worksheet, items: QuoteItem[]) {
    const linhaInicial = 24;
    // Encontra dinamicamente a linha do "Total Produtos:" no template
    const linhaTotalOriginal =
      this.findRowContaining(worksheet, 'total produtos', linhaInicial, 200) ?? 31;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos da primeira linha de produto (24) - colunas A (1) até P (16)
    const estiloProduto = Array.from({ length: 16 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(c + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    // Capturar merges da linha modelo (B:C e D:K)
    // Vamos recriar esses merges depois para cada linha de produto

    const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Desfazer merges da região para evitar conflitos
    this.unmergeFrom(worksheet, linhaInicial);

    // Ajustar quantidade de linhas (em chunks para evitar bug do ExcelJS)
    if (linhasParaAdicionar > 0) {
      this.spliceRowsInChunks(
        worksheet,
        linhaTotalOriginal,
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

    const linhaTotal = linhaInicial + quantidadeProdutos;

    // Preencher produtos
    items.forEach((item, index) => {
      const rowNumber = linhaInicial + index;
      const linha = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      linha.height = alturaProduto;

      // Aplicar estilo capturado em todas as colunas A-P
      for (let col = 1; col <= 16; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      // Recriar merge B:L (descrição) - referência e Unid. foram removidas
      this.safeMerge(worksheet, `B${rowNumber}:L${rowNumber}`);

      // Preencher dados
      // A = Código | B(:L) = Descrição | M = Desc% | N = Qtde | O = Vl.Unit | P = Vl.Total
      linha.getCell(1).value = item.code ?? index + 1;     // A - Código
      linha.getCell(2).value = this.capitalize(item.product ?? '');          // B - Descrição
      linha.getCell(13).value = discount > 0 ? discount / 100 : 0; // M - Desc %
      linha.getCell(14).value = Math.round(item.quantity ?? 0); // N - Qtde
      linha.getCell(15).value = item.unitPrice;             // O - Vl Unitário
      linha.getCell(16).value = totalProduto;               // P - Vl Total

      // Formatação numérica
      linha.getCell(13).numFmt = '0.00%';
      linha.getCell(14).numFmt = '0';
      linha.getCell(15).numFmt = accountingFmt;
      linha.getCell(16).numFmt = accountingFmt;

      // Alinhamento consistente em todas as linhas (justificação uniforme)
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(13).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(14).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(15).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
      linha.getCell(16).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };

      linha.commit();
    });

    // Atualizar TOTAL (linha 31 deslocada)
    // Recriar o merge A:O para o rótulo
    this.safeMerge(worksheet, `A${linhaTotal}:O${linhaTotal}`);
    const totalLabel = worksheet.getCell(`A${linhaTotal}`);
    if (!totalLabel.value || (typeof totalLabel.value === 'string' && !totalLabel.value.includes('Total'))) {
      totalLabel.value = 'Total Produtos:';
      totalLabel.alignment = { horizontal: 'right', vertical: 'middle' };
      totalLabel.font = { ...(totalLabel.font || {}), bold: true };
    }

    const totalCell = worksheet.getCell(`P${linhaTotal}`);
    totalCell.value = total;
    totalCell.numFmt = accountingFmt;
    totalCell.font = { ...(totalCell.font || {}), bold: true };
  }

  /**
   * spliceRows do ExcelJS falha silenciosamente quando o count é grande.
   * Workaround: dividir em chunks de 10 linhas.
   */
  private spliceRowsInChunks(
    worksheet: ExcelJS.Worksheet,
    start: number,
    count: number,
    inserts: any[] = []
  ) {
    const CHUNK_SIZE = 10;

    if (count === 0 && inserts.length > 0) {
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

    let toRemove = count;
    while (toRemove > 0) {
      const chunk = Math.min(toRemove, CHUNK_SIZE);
      worksheet.spliceRows(start, chunk);
      toRemove -= chunk;
    }
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
    style: ReturnType<ExcelGeneratorSantoAmaroService['snapshotStyle']>
  ) {
    const newStyle: any = { ...cell.style };
    if (style.font) newStyle.font = JSON.parse(JSON.stringify(style.font));
    if (style.alignment) newStyle.alignment = JSON.parse(JSON.stringify(style.alignment));
    if (style.border) newStyle.border = JSON.parse(JSON.stringify(style.border));
    if (style.fill) newStyle.fill = JSON.parse(JSON.stringify(style.fill));
    if (style.numFmt) newStyle.numFmt = style.numFmt;
    cell.style = newStyle;
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

  /**
   * Limpa células com erro #VALUE! deixadas pelo Excel 365 quando se usa
   * "Picture in Cell" (Insert > Picture > Place in Cell).
   * Esse formato não é suportado pelo ExcelJS e fica como erro ao ler.
   */
  private clearValueErrors(worksheet: ExcelJS.Worksheet) {
    for (let r = 1; r <= 10; r++) {
      for (let c = 1; c <= 18; c++) {
        const cell = worksheet.getRow(r).getCell(c);
        const v = cell.value as any;
        if (v && typeof v === 'object' && v.error === '#VALUE!') {
          cell.value = null;
        }
      }
    }
  }

  /**
   * Insere a logo Santo Amaro como imagem flutuante usando twoCellAnchor.
   * Substitui a "Picture in Cell" do template que não é compatível com ExcelJS.
   */
  private async insertCompanyLogo(
    workbook: ExcelJS.Workbook,
    worksheet: ExcelJS.Worksheet
  ) {
    try {
      const logoResponse = await fetch('/logoSantoAmaroLimpeza.jpeg');
      if (!logoResponse.ok) return;

      const logoBuffer = await logoResponse.arrayBuffer();
      const logoId = workbook.addImage({
        buffer: logoBuffer,
        extension: 'jpeg'
      });

      // Ancorar entre O2 e P5 (cobre as células onde estava a "Picture in Cell")
      worksheet.addImage(logoId, {
        tl: { col: 14, row: 1 } as any,
        br: { col: 16, row: 5 } as any
      });
    } catch {
      // Logo não encontrada, continua sem
    }
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;

    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.min(Math.max(actualRowCount, 40), 200);
    for (let rowNumber = 23; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      (row as any).outlineLevel = 0;
      row.hidden = false;
    }
  }

  private capitalize(str: string): string {
    if (!str) return str;
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  }

  private capitalizeCompanyName(str: string): string {
    if (!str) return str;
    return str
      .split(/\s+/)
      .map(word => this.capitalize(word))
      .join(' ');
  }
}
