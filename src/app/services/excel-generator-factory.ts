import { Injectable, inject } from '@angular/core';
import { Quote } from '../model/orcamento.model';
import { ExcelGeneratorService } from './excel-generator.service';
import { ExcelGeneratorSplashService } from './excel-generator-splash.service';
import { ExcelGeneratorLagunaService } from './excel-generator-laguna.service';
import { ExcelGeneratorSantoAmaroService } from './excel-generator-santo-amaro.service';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorFactory {
  private defaultGenerator = inject(ExcelGeneratorService);
  private splashGenerator = inject(ExcelGeneratorSplashService);
  private lagunaGenerator = inject(ExcelGeneratorLagunaService);
  private santoAmaroGenerator = inject(ExcelGeneratorSantoAmaroService);

  getGenerator(companyId: string) {
    switch (companyId.toLowerCase()) {
      case 'splash':
        return this.splashGenerator;
      case 'laguna':
        return this.lagunaGenerator;
      case 'santa limpeza':
        return this.santoAmaroGenerator;
      default:
        return this.defaultGenerator;
    }
  }

  async generateQuoteExcel(quote: Quote) {
    const generator = this.getGenerator(quote.company.id);
    return generator.generateQuoteExcel(quote);
  }
}
