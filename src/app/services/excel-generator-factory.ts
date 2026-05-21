import { Injectable, inject } from '@angular/core';
import { Quote } from '../model/orcamento.model';
import { ExcelGeneratorService } from './excel-generator.service';
import { ExcelGeneratorSplashService } from './excel-generator-splash.service';
import { ExcelGeneratorPegLimpService } from './excel-generator-peglimp.service';
import { ExcelGeneratorGrapaCleanService } from './excel-generator-grapaclean.service';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorFactory {
  private defaultGenerator = inject(ExcelGeneratorService);
  private splashGenerator = inject(ExcelGeneratorSplashService);
  private pegLimpGenerator = inject(ExcelGeneratorPegLimpService);
  private grapaCleanGenerator = inject(ExcelGeneratorGrapaCleanService);

  getGenerator(companyId: string) {
    switch (companyId.toLowerCase()) {
      case 'splash':
        return this.splashGenerator;
      case 'peg-limp':
        return this.pegLimpGenerator;
      case 'grapa-clean':
        return this.grapaCleanGenerator;
      default:
        // Santa Limpeza, Laguna e outras empresas usam o gerador padrão
        return this.defaultGenerator;
    }
  }

  async generateQuoteExcel(quote: Quote) {
    const generator = this.getGenerator(quote.company.id);
    return generator.generateQuoteExcel(quote);
  }
}
