import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorSantoAmaroService {
  private templatePath = 'assets/templates/cotacao-template-santo-amaro.xlsx';
  private readonly LINHA_INICIAL_PRODUTOS = 19;
  private readonly LINHA_TOTAL_PRODUTOS = 43;
  // Última linha de produto no template (linhas 19-39 = 21 produtos modelo)
  private readonly ULTIMA_LINHA_TEMPLATE = 39;
  // A última linha com conteúdo útil é a 61 (linha de Observações).
  private readonly ULTIMA_LINHA_CONTEUDO = 61;

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

    // Remover linhas fantasmas (61-1000) que viram páginas extras no PDF
    this.removePhantomRows(worksheet);

    // Definir área de impressão para garantir que o PDF não tenha páginas vazias
    this.setPrintArea(worksheet);

    // A logo já vem embutida no template - não inserir novamente
    // (evita duplicação da logo no lado direito)

    const customer = quote.customerData;
    const quoteDate = quote.date ? new Date(quote.date) : new Date();

    // ===== DATA DO ORÇAMENTO =====
    // E4 = valor da data (D4 é o label "Data:")
    const dia = String(quoteDate.getDate()).padStart(2, '0');
    const mes = String(quoteDate.getMonth() + 1).padStart(2, '0');
    const ano = quoteDate.getFullYear();
    const dataCell = worksheet.getCell('E4');
    dataCell.numFmt = '@';
    dataCell.value = `${dia}/${mes}/${ano}`;

    // ===== NÚMERO DO ORÇAMENTO (Aleatório 2000-3000) =====
    // J4 = número do orçamento (era K4 no template antigo)
    const numeroOrcamento = Math.floor(Math.random() * 1001) + 2000;
    worksheet.getCell('J4').value = numeroOrcamento;

    console.log('[Santo Amaro] Customer recebido:', customer);
    console.log('[Santo Amaro] Quote info:', {
      label: quote.label,
      source: (quote as any).source,
      paymentMethod: quote.paymentMethod,
      deliveryTime: quote.deliveryTime
    });

    // ===== DADOS DO CLIENTE =====
    // Mapeamento do novo template (Modelo Final Sta Limpeza):
    //   D7  = Cliente       (era C7)
    //   D8  = Fantasia      (era C8)
    //   D9  = Endereço      (era C9)
    //   M9  = Bairro        (era N9)
    //   D10 = Cidade        (era C10)
    //   L10 = UF            (era M10)
    //   N10 = CEP           (era O10)
    //   D11 = CNPJ/CPF      (era C11)
    //   D14 = Contato       (era C14)
    //   D15 = Email         (era C15)
    //   D16 = Vendedor      (era C16)
    if (customer) {
      if (customer.client) {
        worksheet.getCell('D7').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.fantasyName) {
        worksheet.getCell('D8').value = String(customer.fantasyName).replace(/[\r\n]+/g, ' ').trim();
      } else if (customer.client) {
        worksheet.getCell('D8').value = String(customer.client).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.address) {
        worksheet.getCell('D9').value = String(customer.address).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.district) {
        worksheet.getCell('M9').value = customer.district;
      }
      if (customer.city) {
        worksheet.getCell('D10').value = customer.city;
      }
      if (customer.state) {
        worksheet.getCell('L10').value = customer.state;
      }
      if (customer.postalCode) {
        worksheet.getCell('N10').value = customer.postalCode;
      }
      if (customer.document) {
        worksheet.getCell('D11').value = customer.document;
      }
      if (customer.contact) {
        worksheet.getCell('D14').value = String(customer.contact).replace(/[\r\n]+/g, ' ').trim();
      }
      if (customer.email) {
        worksheet.getCell('D15').value = customer.email;
      }
      if (customer.salesperson) {
        worksheet.getCell('D16').value = customer.salesperson;
      }
    }

    // ===== CONDIÇÕES COMERCIAIS =====
    // Forma de Pagamento (M12 - era N12)
    if (quote.paymentMethod) {
      worksheet.getCell('M12').value = quote.paymentMethod;
    }

    // Prazo de entrega vai para Observações (D61 - era C60)
    if (quote.deliveryTime) {
      worksheet.getCell('D61').value = `ENTREGA: ${quote.deliveryTime}`;
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
    // Formato accounting do Excel: R$ alinhado à esquerda e número à direita.
    // Para zeros mostra "R$    -" e negativos viram "-R$ X,XX".
    const accountingFmt = '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-';

    // Cálculos de totais
    const total = items.reduce((sum, item) => {
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      return sum + (subtotal - subtotal * (discount / 100));
    }, 0);

    const quantidadeItens = items.reduce((sum, item) => sum + (item.quantity ?? 0), 0);

    // IMPORTANTE: Apenas LIMPAR valores das colunas relevantes, sem mexer em
    // bordas, fills ou outras propriedades. As bordas vêm do template.
    // Novo template: B=Código, D=Descrição, K=Unid., L=Desc%, M=Qtde, N=Vl.Unit, O=Vl.Total
    const colunasProdutos = [2, 4, 11, 12, 13, 14, 15]; // B, D, K, L, M, N, O
    for (let r = this.LINHA_INICIAL_PRODUTOS; r <= this.ULTIMA_LINHA_TEMPLATE; r++) {
      const row = worksheet.getRow(r);
      colunasProdutos.forEach(c => {
        row.getCell(c).value = null;
      });
    }

    // Preencher produtos - SOMENTE valores, preserva bordas/estilos do template
    items.forEach((item, index) => {
      const rowNumber = this.LINHA_INICIAL_PRODUTOS + index;

      // Se ultrapassar o limite do template, não preencher
      if (rowNumber > this.ULTIMA_LINHA_TEMPLATE) return;

      const row = worksheet.getRow(rowNumber);
      const discount = item.discount ?? 0;
      const subtotal = item.quantity * item.unitPrice;
      const totalProduto = subtotal - subtotal * (discount / 100);

      // Preencher dados conforme NOVO template (apenas valores):
      // B=Código | D=Descrição | K=Unid. | L=Desc% | M=Qtde. | N=Vl.Unit. | O=Vl.Total
      row.getCell(2).value = item.code ?? index + 1;             // B - Código
      row.getCell(4).value = item.product ?? '';                 // D - Descrição
      row.getCell(11).value = item.unit ?? '';                   // K - Unidade
      row.getCell(12).value = discount > 0 ? discount / 100 : 0; // L - Desc %
      row.getCell(13).value = Math.round(item.quantity ?? 0);    // M - Qtde
      row.getCell(14).value = item.unitPrice;                    // N - Vl.Unit
      row.getCell(15).value = totalProduto;                      // O - Vl.Total

      // Aplicar APENAS formatação numérica (sem alterar bordas)
      row.getCell(12).numFmt = '0.00%';
      row.getCell(13).numFmt = '0';
      row.getCell(14).numFmt = accountingFmt;
      row.getCell(15).numFmt = accountingFmt;
    });

    // ===== TOTALIZAÇÕES (mapeamento do novo template) =====
    // Total Produtos: N43 (era O43)
    worksheet.getCell(`N${this.LINHA_TOTAL_PRODUTOS}`).value = total;
    worksheet.getCell(`N${this.LINHA_TOTAL_PRODUTOS}`).numFmt = accountingFmt;

    // Qtde. de Itens: N45 (era O44)
    worksheet.getCell('N45').value = quantidadeItens;
    worksheet.getCell('N45').numFmt = '0';

    // Zeros nos campos de taxas/descontos no novo template:
    // IPI=N47, Desconto=N49, Acréscimo=N51, Total Subs Trib=N53,
    // Total FCP ST=N55, Total Frete=N57
    [47, 49, 51, 53, 55, 57].forEach(row => {
      const cell = worksheet.getCell(`N${row}`);
      cell.value = 0;
      cell.numFmt = accountingFmt;
    });

    // Total final: N59 (era O58)
    worksheet.getCell('N59').value = total;
    worksheet.getCell('N59').numFmt = accountingFmt;
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

  /**
   * Remove "linhas fantasmas" do template - linhas vazias após
   * ULTIMA_LINHA_CONTEUDO que têm altura definida e fazem o PDF
   * gerar páginas em branco extras.
   *
   * No template original existem ~1000 linhas vazias com altura customHeight,
   * que somadas dão várias páginas extras no PDF.
   */
  private removePhantomRows(worksheet: ExcelJS.Worksheet) {
    const actualRowCount = (worksheet as any).actualRowCount ?? worksheet.rowCount;
    const ultimaLinha = Math.max(actualRowCount, this.ULTIMA_LINHA_CONTEUDO + 1);

    for (let rowNumber = this.ULTIMA_LINHA_CONTEUDO + 1; rowNumber <= ultimaLinha; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      // Limpa valores de todas as colunas (A-P, 16 colunas)
      for (let c = 1; c <= 16; c++) {
        const cell = row.getCell(c);
        cell.value = null;
        cell.style = {};
      }
      // Zera altura para a linha não ocupar espaço na impressão
      row.height = 0;
      row.hidden = true;
      (row as any).outlineLevel = 0;
    }
  }

  /**
   * Define a área de impressão para limitar o PDF apenas ao conteúdo útil,
   * evitando páginas extras causadas por linhas/colunas fantasmas.
   * Mantém orientação landscape e ajusta a largura à página (sem comprimir
   * a altura, para não distorcer as linhas de produto).
   */
  private setPrintArea(worksheet: ExcelJS.Worksheet) {
    worksheet.pageSetup = {
      ...(worksheet.pageSetup ?? {}),
      printArea: `A1:P${this.ULTIMA_LINHA_CONTEUDO}`,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0, // 0 = não limita altura (deixa Excel quebrar páginas naturalmente)
      orientation: 'landscape'
    };
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
