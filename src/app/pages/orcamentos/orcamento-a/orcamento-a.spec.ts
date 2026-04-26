import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OrcamentoA } from './orcamento-a';

describe('OrcamentoA', () => {
  let component: OrcamentoA;
  let fixture: ComponentFixture<OrcamentoA>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [OrcamentoA],
    }).compileComponents();

    fixture = TestBed.createComponent(OrcamentoA);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
