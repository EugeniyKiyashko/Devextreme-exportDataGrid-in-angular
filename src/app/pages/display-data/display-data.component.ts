import { Component } from '@angular/core';
import 'devextreme/data/odata/store';
import * as excelExporter from 'devextreme/exporter/exceljs/excelExporter';
import ExcelJS from 'exceljs/dist/exceljs.js';
import { saveAs } from 'file-saver';

@Component({
  templateUrl: 'display-data.component.html'
})

export class DisplayDataComponent {
  dataSource: any;
  priority: any[];

  constructor() {
    this.dataSource = {
      store: {
        type: 'odata',
        key: 'Task_ID',
        url: 'https://js.devexpress.com/Demos/DevAV/odata/Tasks'
      },
      expand: 'ResponsibleEmployee',
      select: [
        'Task_ID',
        'Task_Subject',
        'Task_Start_Date',
        'Task_Due_Date',
        'Task_Status',
        'Task_Priority',
        'Task_Completion',
      ]
    };
    this.priority = [
      { name: 'High', value: 4 },
      { name: 'Urgent', value: 3 },
      { name: 'Normal', value: 2 },
      { name: 'Low', value: 1 }
    ];
  }

  onExporting(e) {
    var workbook = new ExcelJS.Workbook();    
    var worksheet = workbook.addWorksheet('Main sheet');
    var context = this;
    excelExporter.exportDataGrid({
        component: e.component,
        worksheet: worksheet,
        topLeftCell: { row: 4, column: 1 }
    }).then(function() {  
        context.customizeHeader(worksheet);
        return Promise.resolve();
    }).then(function() {
        workbook.xlsx.writeBuffer().then(function(buffer) {
            saveAs(new Blob([buffer], { type: 'application/octet-stream' }), 'DataGrid.xlsx');
        });
    });
    e.cancel = true;
  }

  customizeHeader(worksheet){
    for(var columnIndex = 1; columnIndex < 10; columnIndex++){
        worksheet.getColumn(columnIndex).width = 20;
    }
    worksheet.getColumn(2).width = 60;
    worksheet.mergeCells(1, 1, 3, 9);

    worksheet.getRow(1).getCell(1).value = "Sandra Johnson report";
    worksheet.getRow(1).getCell(1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'ff5722'}, bgColor:{argb:'ff5722'}};
    worksheet.getRow(1).getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  }
}
