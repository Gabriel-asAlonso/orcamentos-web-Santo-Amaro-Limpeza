import { TestBed } from '@angular/core/testing';
import { ExcelGeneratorFactory } from './excel-generator-factory';
import { ExcelGeneratorOmaService } from './excel-generator-oma.service';

describe('ExcelGeneratorFactory', () => {
  it('seleciona o gerador específico da OMA', () => {
    TestBed.configureTestingModule({});

    const factory = TestBed.inject(ExcelGeneratorFactory);
    const omaGenerator = TestBed.inject(ExcelGeneratorOmaService);

    expect(factory.getGenerator('oma')).toBe(omaGenerator);
    expect(factory.getGenerator('OMA')).toBe(omaGenerator);
  });
});
