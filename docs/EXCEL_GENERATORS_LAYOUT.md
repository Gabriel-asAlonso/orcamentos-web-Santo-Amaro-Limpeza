# Sistema de Layouts Específicos por Empresa

## Visão Geral

O sistema agora suporta layouts personalizados de Excel para cada empresa, permitindo que cada fornecedor tenha seu próprio template e estrutura de apresentação de orçamentos.

## Arquitetura

### Componentes

1. **ExcelGeneratorFactory** (`src/app/services/excel-generator-factory.ts`)
   - Factory que seleciona o service correto baseado no `company.id`
   - Método: `getGenerator(companyId: string)` retorna o service apropriado
   - Método: `generateQuoteExcel(quote: Quote)` encapsula a geração

2. **ExcelGeneratorService** (padrão)
   - Usado para: Laguna, Peg-Limp, GrapaClean, Santa Limpeza
   - Template: `assets/templates/cotacao-template.xlsx`
   - Layout: Original do sistema

3. **ExcelGeneratorSplashService** (novo)
   - Usado para: Splash (id: 'splash')
   - Template: `assets/templates/cotacao-template-Splash.xlsx`
   - Layout: Personalizado com especificações da Splash
     - Logo: A4:B5
     - Empresa: A2:F3
     - Campos de cliente com labels específicos

## Como Adicionar um Novo Layout

### Passo 1: Criar o Template Excel

1. Crie um novo arquivo Excel com o layout desejado
2. Nomeie conforme a convenção: `cotacao-template-[empresa].xlsx`
3. Coloque na pasta: `public/assets/templates/`

### Passo 2: Criar o Service

1. Copie e modifique o `ExcelGeneratorSplashService` como base:

```typescript
// src/app/services/excel-generator-[empresa].service.ts

import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Quote, QuoteItem } from '../model/orcamento.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelGenerator[Empresa]Service {
  private templatePath = 'assets/templates/cotacao-template-[empresa].xlsx';

  async generateQuoteExcel(quote: Quote) {
    // Implementação específica para a empresa
  }

  // Métodos privados para preencher dados
  private fillCustomerData(...) { }
  private fillProducts(...) { }
  // ... etc
}
```

2. Customize os métodos:
   - `fillCustomerData()`: Preencha os campos de cliente conforme layout
   - `fillProducts()`: Preencha a tabela de produtos
   - Outros métodos conforme necessário

### Passo 3: Registrar na Factory

Edite `src/app/services/excel-generator-factory.ts`:

```typescript
export class ExcelGeneratorFactory {
  private defaultGenerator = inject(ExcelGeneratorService);
  private splashGenerator = inject(ExcelGeneratorSplashService);
  private [empresa]Generator = inject(ExcelGenerator[Empresa]Service); // NOVO

  getGenerator(companyId: string) {
    switch (companyId.toLowerCase()) {
      case 'splash':
        return this.splashGenerator;
      case '[empresa]':  // NOVO
        return this.[empresa]Generator;
      default:
        return this.defaultGenerator;
    }
  }
}
```

### Passo 4: Atualizar Company ID (se necessário)

Verifique se o `company.id` em `src/app/data/empresas.ts` corresponde ao case utilizado na factory.

## Estrutura de Exemplo: Splash

```
Implementação do Service:
├── prepareWorksheet()
│   └── Ajusta largura das colunas
├── insertCompanyLogo()
│   └── Insere logo em A4:B5
├── fillCustomerData()
│   └── Preenche campos: cliente, fantasia, endereço, etc.
├── fillProducts()
│   └── Preenche tabela de produtos com totais
└── Utilitários privados
    ├── safeMergeCells()
    ├── clearWorksheetImages()
    └── clearWorksheetRowGrouping()
```

## Melhores Práticas

### 1. Reutilizar Lógica Comum

Se múltiplos services compartilham lógica similar, considere:
- Criar uma classe base abstrata
- Extrair métodos comuns para um utilitário
- Exemplo: `fillProducts()` é similar em todos

### 2. Template Excel

- Mantenha a estrutura de headers consistente
- Use nomes significativos para ranges (colunas nomeadas)
- Preserve formatação visual do cliente
- Deixe espaço para rows dinâmicas de produtos

### 3. Tratamento de Erros

- Use try-catch para fetch de logos (podem não existir)
- Valide merge de cells (podem já estar mescladas)
- Implemente fallbacks para dados ausentes

### 4. Formatação

Padrões de formatação mantidos:
```typescript
// Valores monetários
cell.numFmt = 'R$ #,##0.00';

// Percentuais
cell.numFmt = '0.00%';

// Datas
cell.numFmt = 'dd/mm/yyyy';
```

## Fluxo de Geração

```
1. usuário clica "Baixar Orçamento" em uma empresa
   ↓
2. orcamentos.ts chama downloadQuote(quote)
   ↓
3. excelGeneratorFactory.generateQuoteExcel(quote)
   ↓
4. Factory identifica company.id
   ↓
5. Retorna service apropriado (Splash ou padrão)
   ↓
6. Service carrega template
   ↓
7. Service preenche dados conforme layout específico
   ↓
8. Service exporta e faz download do arquivo
```

## Testando um Novo Layout

1. Crie o template Excel em `public/assets/templates/`
2. Crie o service específico
3. Registre na factory
4. Teste selecionando a empresa e clicando "Baixar"
5. Valide:
   - Logo está na posição correta
   - Dados estão preenchidos corretamente
   - Formatação é preservada
   - Produtos aparecem com espaçamento adequado

## Empresas Configuradas

| Empresa | ID | Service | Template |
|---------|----|----|----------|
| Splash | `splash` | ExcelGeneratorSplashService | `cotacao-template-Splash.xlsx` |
| Laguna | `Laguna` | ExcelGeneratorService (padrão) | `cotacao-template.xlsx` |
| Peg-Limp | `peg-limp` | ExcelGeneratorService (padrão) | `cotacao-template.xlsx` |
| GrapaClean | `grapa-clean` | ExcelGeneratorService (padrão) | `cotacao-template.xlsx` |
| Santa Limpeza | `Santa Limpeza` | ExcelGeneratorService (padrão) | `cotacao-template.xlsx` |

## Próximos Passos

1. Testar layout da Splash com dados reais
2. Ajustar posições de campos e formatação conforme necessário
3. Considerar criar layouts específicos para outras empresas
4. Documentar especificidades de cada layout
