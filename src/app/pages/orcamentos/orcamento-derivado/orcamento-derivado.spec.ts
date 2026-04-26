import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OrcamentoDerivado } from './orcamento-derivado';

describe('OrcamentoDerivado', () => {
  let component: OrcamentoDerivado;
  let fixture: ComponentFixture<OrcamentoDerivado>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [OrcamentoDerivado],
    }).compileComponents();

    fixture = TestBed.createComponent(OrcamentoDerivado);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
