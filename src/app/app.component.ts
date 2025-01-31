import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { RouterOutlet } from '@angular/router';
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  WidthType,
  TextRun,
  TableBorders,
  BorderStyle,
  ShadingType,
  AlignmentType,
  Indent,
} from 'docx';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  coded!: string;

  generateDocx() {
    const tableData = [
      [
        'for (var i = 2; primeArray.length < numPrimes; primeArray.length < numPrimes; i++) { ',
      ],
      ['    PrimeCheck(i); //'],
      ['}'],
      ['  for(var i = 2; i < candidate && isPrime; i++){'],
    ];

    const tableRows = tableData.map(
      (row, index) =>
        new TableRow({
          children: row.map(
            (cell) =>
              new TableCell({
                width: { size: 2000, type: WidthType.DXA },
                shading: {
                  fill: 'D3D3D3',
                  type: ShadingType.CLEAR,
                },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cell,
                        font: 'Courier New',
                        size: 24,
                        noProof : true,
                      }),
                    ],
                    alignment: AlignmentType.LEFT,
                  }),
                ],
              })
          ),
        })
    );

    // Create a table with styling (borders and indentation)
    const table = new Table({
      rows: tableRows,
      width: { size: 9000, type: WidthType.DXA }, // Adjust table width
      borders: {
        top: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
        left: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
        right: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
      },
      indent: {
        size: 720,
      },
    });

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Table Report', font: 'Courier New' }),
              ],
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph(' '),
            table,
          ],
        },
      ],
    });

    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, 'StyledTableReport.docx');
    });
  }
}
