import { Component } from '@angular/core';
import { ExcelDedoublonnageComponent } from './excel-dedoublonnage/excel-dedoublonnage.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [ExcelDedoublonnageComponent],
  template: `<app-excel-dedoublonnage></app-excel-dedoublonnage>`
})
export class AppComponent {}