import { Injectable, inject } from '@angular/core';
import { Quote } from '../model/orcamento.model';
import { ExcelGeneratorService } from './excel-generator.service';
import { ExcelGeneratorSplashService } from './excel-generator-splash.service';

@Injectable({
  providedIn: 'root'
})
export class ExcelGeneratorFactory {
  private defaultGenerator = inject(ExcelGeneratorService);
  private splashGenerator = inject(ExcelGeneratorSplashService);

  getGenerator(companyId: string) {
    switch (companyId.toLowerCase()) {
      case 'splash':
        return this.splashGenerator;
      default:
        return this.defaultGenerator;
    }
  }

  async generateQuoteExcel(quote: Quote) {
    const generator = this.getGenerator(quote.company.id);
    return generator.generateQuoteExcel(quote);
  }
}
