import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { RouterOutlet } from '@angular/router';
import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  coded!: string;
  size: number = 12;

  generateDocx() {
    const tableData = this.coded.split('\n').map((e) => [e]);

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
                        size: this.size, // 24 means 12pt
                        noProof: true,
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
