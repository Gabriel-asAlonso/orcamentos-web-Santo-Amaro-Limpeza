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
    this.clearValueErrors(worksheet);

    // Inserir a logo programaticamente
    await this.insertCompanyLogo(workbook, worksheet);

    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== DATA DO ORÇAMENTO =====
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('D4');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== NÚMERO DO ORÇAMENTO (Aleatório 2000-3000) =====
    const numeroOrcamento = Math.floor(Math.random() * 1001) + 2000;
    worksheet.getCell('K4').value = numeroOrcamento;

    console.log('[Santo Amaro] Customer recebido:', customer);
    console.log('[Santo Amaro] Quote info:', {
      label: quote.label,
      source: (quote as any).source,
      paymentMethod: quote.paymentMethod,
      deliveryTime: quote.deliveryTime
    });

    // ===== DADOS DO CLIENTE =====
    if (customer) {
      if (customer.client) {
        worksheet.getCell('C7').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.fantasyName) {
        worksheet.getCell('C8').value = String(customer.fantasyName).replace(/[\r\n]+/g, ' ').trim();
      } else if (customer.client) {
        worksheet.getCell('C8').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        worksheet.getCell('C9').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.district) {
        worksheet.getCell('N9').value = customer.district;
      }
      if (customer.city) {
        worksheet.getCell('C10').value = customer.city;
      }
      if (customer.state) {
        worksheet.getCell('M10').value = customer.state;
      }
      if (customer.postalCode) {
        worksheet.getCell('O10').value = customer.postalCode;
      }
      if (customer.document) {
        worksheet.getCell('C11').value = customer.document;
      }
      if (customer.contact) {
        worksheet.getCell('C14').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.email) {
        worksheet.getCell('C15').value = customer.email;
      }
      if (customer.salesperson) {
        worksheet.getCell('C16').value = customer.salesperson;
      }
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    // Forma de Pagamento (N12)
    if (quote.paymentMethod) {
      worksheet.getCell('N12').value = quote.paymentMethod;
    }

    // Prazo de entrega vai para Observações (C60)
    if (quote.deliveryTime) {
      worksheet.getCell('C60').value = `ENTREGA: ${quote.deliveryTime}`;
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
    const linhaInicial = 19;
    // Encontra dinamicamente a linha do "Total Produtos:" no template
    const linhaTotalOriginal =
      this.findRowContaining(worksheet, 'total produtos', linhaInicial, 200) ?? 43;
    const quantidadeProdutos = Math.max(items.length, 1);
    const quantidadeLinhasModelo = Math.max(linhaTotalOriginal - linhaInicial, 1);
    const linhasParaRemover = Math.max(quantidadeLinhasModelo - quantidadeProdutos, 0);
    const linhasParaAdicionar = Math.max(quantidadeProdutos - quantidadeLinhasModelo, 0);

    // Capturar estilos da primeira linha de produto (19)
    const estiloProduto = Array.from({ length: 16 }, (_, c) =>
      this.snapshotStyle(worksheet.getRow(linhaInicial).getCell(c + 1))
    );
    const alturaProduto = worksheet.getRow(linhaInicial).height ?? 16;

    const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    // Calcular total
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    // Calcular quantidade total de itens
    const quantidadeItens = items.reduce((sum, item) => sum + (item.quantity ?? 0), 0);

    // Desfazer merges da região
    this.unmergeFrom(worksheet, linhaInicial);

    // Ajustar quantidade de linhas
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

      // Aplicar estilo capturado
      for (let col = 1; col <= 16; col++) {
        const cell = linha.getCell(col);
        cell.value = null;
        this.applyStyle(cell, estiloProduto[col - 1]);
      }

      // Merge C(:L) para descrição
      this.safeMerge(worksheet, `C${rowNumber}:L${rowNumber}`);

      // Preencher dados - Layout novo:
      // A=Código | C(:L)=Descrição | L=Unid | M=Desc% | N=Qtde | O=Vl.Unit | P=Vl.Total
      linha.getCell(1).value = item.code ?? index + 1;              // A - Código
      linha.getCell(3).value = item.product ?? '';                  // C - Descrição
      linha.getCell(12).value = item.unit ?? '';                    // L - Unidade
      linha.getCell(13).value = discount > 0 ? discount / 100 : 0;  // M - Desc %
      linha.getCell(14).value = Math.round(item.quantity ?? 0);     // N - Qtde
      linha.getCell(15).value = item.unitPrice;                     // O - Vl.Unit
      linha.getCell(16).value = totalProduto;                       // P - Vl.Total

      // Formatação numérica
      linha.getCell(13).numFmt = '0.00%';
      linha.getCell(14).numFmt = '0';
      linha.getCell(15).numFmt = accountingFmt;
      linha.getCell(16).numFmt = accountingFmt;

      // Alinhamento
      linha.getCell(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(3).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      linha.getCell(12).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(13).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(14).alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
      linha.getCell(15).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };
      linha.getCell(16).alignment = { horizontal: 'right', vertical: 'middle', wrapText: false };

      linha.commit();
    });

    // Atualizar totalizações na coluna O
    // K43/O43: Total Produtos
    worksheet.getCell(`O${linhaTotal}`).value = total;
    worksheet.getCell(`O${linhaTotal}`).numFmt = accountingFmt;
    worksheet.getCell(`O${linhaTotal}`).font = { bold: true };

    // K44/O44: Qtde de Itens
    const linhaQtdeItens = linhaTotal + 1;
    worksheet.getCell(`O${linhaQtdeItens}`).value = quantidadeItens;
    worksheet.getCell(`O${linhaQtdeItens}`).numFmt = '0';

    // Zeros em O46 (IPI), O48 (Desconto), O50 (Acréscimo), O52 (Total Subs), O54 (FCP), O56 (Frete)
    const linhasZeros = [46, 48, 50, 52, 54, 56];
    linhasZeros.forEach(linha => {
      worksheet.getCell(`O${linha}`).value = 0;
      worksheet.getCell(`O${linha}`).numFmt = accountingFmt;
    });

    // K58/O58: Total Final
    worksheet.getCell('O58').value = total;
    worksheet.getCell('O58').numFmt = accountingFmt;
    worksheet.getCell('O58').font = { bold: true };
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
      if (!logoResponse.ok) {
        console.warn('[Santo Amaro] Logo não encontrada:', logoResponse.status, logoResponse.statusText);
        return;
      }

      const logoBuffer = await logoResponse.arrayBuffer();
      const logoId = workbook.addImage({
        buffer: logoBuffer,
        extension: 'jpeg'
      });

      worksheet.addImage(logoId, {
        tl: { col: 14, row: 1 } as any,
        br: { col: 16, row: 5 } as any
      });
      console.log('[Santo Amaro] Logo inserida com sucesso');
    } catch (error) {
      console.error('[Santo Amaro] Erro ao inserir logo:', error);
    }
  }

  private clearWorksheetRowGrouping(worksheet: ExcelJS.Worksheet) {
    (worksheet as any).rowOutlineLevel = 0;

    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.min(Math.max(actualRowCount, 40), 200);
    for (let rowNumber = 18; rowNumber <= ultimaLinha; rowNumber++) {
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
