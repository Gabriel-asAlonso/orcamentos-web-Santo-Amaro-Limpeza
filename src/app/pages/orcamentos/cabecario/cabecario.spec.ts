import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Cabecario } from './cabecario';

describe('Cabecario', () => {
  let component: Cabecario;
  let fixture: ComponentFixture<Cabecario>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Cabecario],
    }).compileComponents();

    fixture = TestBed.createComponent(Cabecario);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
