import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Orcamentos } from './orcamentos';

describe('Orcamentos', () => {
  let component: Orcamentos;
  let fixture: ComponentFixture<Orcamentos>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Orcamentos],
    }).compileComponents();

    fixture = TestBed.createComponent(Orcamentos);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
