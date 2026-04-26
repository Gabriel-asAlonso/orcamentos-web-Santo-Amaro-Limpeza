import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Importar } from './importar';

describe('Importar', () => {
  let component: Importar;
  let fixture: ComponentFixture<Importar>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Importar],
    }).compileComponents();

    fixture = TestBed.createComponent(Importar);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
