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

  // async generateDocx() {
  //   const response = await fetch('assets/sample.docx');

  //   const template = await response.arrayBuffer();

  //   const zip = new PizZip(template);
  //   const doc = new Docxtemplater().loadZip(zip);

  //   const data = {
  //     name: 'John Doe',
  //     content: 'This is dynamically generated content with custom stylings.',
  //     code: `<pre>${this.coded}</pre>`,
  //   };
  //   doc.setData(data);
  //   const linesArray = this.coded.split('\n'); // Split input into lines
  //   console.log(linesArray); // Log the array of lines

  //   try {
  //     doc.render();
  //   } catch (error) {
  //     console.error('Error rendering document:', error);
  //     throw error;
  //   }

  //   const blob = doc.getZip().generate({ type: 'blob' });
  //   // saveAs(blob, 'generated-document.docx');
  // }

  generateDocx() {
    // Table Data
    const tableData = [
      ['1',
        'for (var i = 2; primeArray.length < numPrimes; primeArray.length < numPrimes; i++) { ',
      ], // Header
      ['2','    PrimeCheck(i); //'],
      ['3','}'],
      ['4','  for(var i = 2; i < candidate && isPrime; i++){'],
    ];

    // Create table rows with styling
    const tableRows = tableData.map(
      (row, index) =>
        new TableRow({
          children: row.map(
            (cell) =>
              new TableCell({
                width: { size: 2000, type: WidthType.DXA },
                shading: {
                  fill: 'D3D3D3', // Light Gray background for header
                  type: ShadingType.CLEAR,
                },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cell,
                        font: 'Courier New',
                        size: 24, // Font size (24 = 12pt)
                        noProof : true
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
        // bottom: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
        left: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
        right: { style: BorderStyle.SINGLE, size: 3, color: '000000' },
        // insideHorizontal: {
        //   style: BorderStyle.SINGLE,
        //   size: 1,
        //   color: '000000',
        // },
        // insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
      },
      indent: {
        size: 720, // Adds left indentation (720 twips = 0.5 inch)
      },
    });

    // Create a document
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
            new Paragraph(' '), // Blank line for spacing
            table,
          ],
        },
      ],
    });

    // Generate and download DOCX file
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, 'StyledTableReport.docx');
    });
  }
}
