import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Tetes } from './tetes';

describe('Tetes', () => {
  let component: Tetes;
  let fixture: ComponentFixture<Tetes>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Tetes],
    }).compileComponents();

    fixture = TestBed.createComponent(Tetes);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
