import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OrcamentoGerados } from './orcamento-gerados';

describe('OrcamentoGerados', () => {
  let component: OrcamentoGerados;
  let fixture: ComponentFixture<OrcamentoGerados>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [OrcamentoGerados],
    }).compileComponents();

    fixture = TestBed.createComponent(OrcamentoGerados);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
