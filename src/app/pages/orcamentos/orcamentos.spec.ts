import { ComponentFixture, TestBed } from '@angular/core/testing';

import { orcamentos } from './orcamentos';

describe('Orcamentos', () => {
  let component: orcamentos;
  let fixture: ComponentFixture<orcamentos>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [orcamentos],
    }).compileComponents();

    fixture = TestBed.createComponent(orcamentos);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
