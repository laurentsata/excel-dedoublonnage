import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelDedoublonnageComponent } from './excel-dedoublonnage.component';

describe('ExcelDedoublonnageComponent', () => {
  let component: ExcelDedoublonnageComponent;
  let fixture: ComponentFixture<ExcelDedoublonnageComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExcelDedoublonnageComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExcelDedoublonnageComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
