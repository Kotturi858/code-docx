import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { BrowserModule } from '@angular/platform-browser';
import { RouterOutlet } from '@angular/router';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  coded!: string;
  title = 'code-docx';

  async generateDocx() {

    const response = await fetch('assets/sample.docx');

    const template = await response.arrayBuffer();

    // Initialize PizZip and Docxtemplater
    const zip = new PizZip(template);
    const doc = new Docxtemplater().loadZip(zip);

    // Replace placeholders with data
    const data = {
      name: 'John Doe',
      content: 'This is dynamically generated content with custom stylings.',
    };
    doc.setData(data);

    try {
      doc.render();
    } catch (error) {
      console.error('Error rendering document:', error);
      throw error;
    }

    // Generate the .docx file
    const blob = doc.getZip().generate({ type: 'blob' });
    saveAs(blob, 'generated-document.docx');
  }
}
