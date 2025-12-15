import { Component, signal, WritableSignal, OnInit, effect } from '@angular/core';
import { CommonModule } from '@angular/common';
import { QuillModule } from 'ngx-quill';
import { FormsModule } from '@angular/forms';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import { HttpClient } from '@angular/common/http';
import Quill from 'quill';

interface CommitmentLanguage {
  code: string;
  name: string;
  content: string;
}

interface DocumentItem {
  id: string;
  name: string;
  path: string;
}

@Component({
  selector: 'app-commitment-collation',
  standalone: true,
  imports: [CommonModule, QuillModule, FormsModule],
  templateUrl: './commitment-collation.component.html',
  styleUrl: './commitment-collation.component.scss'
})
export class CommitmentCollationComponent implements OnInit {
  private readonly STORAGE_KEY = 'COMMITMENT_COLLATION_DATA';
  languages: WritableSignal<CommitmentLanguage[]> = signal([]);

  documents: DocumentItem[] = [
    { id: '1', name: 'Project_Plan.pdf', path: '/docs/Project_Plan.pdf' },
    { id: '2', name: 'Budget_Overview.xlsx', path: '/docs/Budget_Overview.xlsx' },
    { id: '3', name: 'Design_Specs.docx', path: '/docs/Design_Specs.docx' },
    { id: '4', name: 'Meeting_Notes.txt', path: '/docs/Meeting_Notes.txt' },
    { id: '5', name: 'Q1_Report.pdf', path: '/docs/Q1_Report.pdf' },
  ];

  isModalOpen = signal(false);
  searchQuery = signal('');
  editors: Quill[] = [];

  quillModules = {
    toolbar: {
      container: [
        ['bold', 'italic', 'underline'],
        [{ 'list': 'ordered' }, { 'list': 'bullet' }],
        ['link']
      ],
      handlers: {
        'link': (value: any) => {
          this.isModalOpen.set(true);
          setTimeout(() => {
            const input = document.querySelector('.search-bar input') as HTMLElement;
            if (input) input.focus();
          }, 50);
          return false;
        }
      }
    }
  };

  lastActiveEditor: Quill | null = null;

  constructor(private http: HttpClient) {
    const Link = Quill.import('formats/link') as any;
    class MyLink extends Link {
      static create(value: any) {
        const node = super.create(value);
        value = this['sanitize'](value);
        node.setAttribute('href', value);
        node.setAttribute('target', '_blank');
        return node;
      }
    }
    Quill.register(MyLink, true);
  }

  ngOnInit() {
    this.loadFromLocalStorage();
    if (this.languages().length === 0) {
      this.languages.set([{ code: '001', name: 'Commitment Name', content: '' }]);
    }
  }

  saveToLocalStorage() {
    try {
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.languages()));
    } catch (e) {
      console.error('Error saving to localStorage', e);
    }
  }

  private loadFromLocalStorage() {
    try {
      const stored = localStorage.getItem(this.STORAGE_KEY);
      if (stored) {
        this.languages.set(JSON.parse(stored));
      }
    } catch (e) {
      console.error('Error loading from localStorage', e);
    }
  }

  onEditorCreated(quill: Quill) {
    this.editors.push(quill);
    quill.on('selection-change', (range) => {
      if (range) {
        this.lastActiveEditor = quill;
      }
    });
  }

  get filteredDocuments() {
    const query = this.searchQuery().toLowerCase();
    return this.documents.filter(doc => doc.name.toLowerCase().includes(query));
  }

  closeModal() {
    this.isModalOpen.set(false);
    this.searchQuery.set('');
  }

  selectDocument(doc: DocumentItem) {
    const activeEditor = this.lastActiveEditor;
    
    if (activeEditor) {
      activeEditor.focus();
      
      const range = activeEditor.getSelection(); 
      
      if (range) {
        if (range.length > 0) {
          activeEditor.format('link', doc.path);
        } else {
          activeEditor.insertText(range.index, doc.name, 'link', doc.path);
          activeEditor.setSelection(range.index + doc.name.length, 0);
        }
      }
    }
    this.closeModal();
  }

  addCommitment() {
    this.languages.update(langs => [
      ...langs,
      { 
        code: `00${langs.length + 1}`, 
        name: 'New Commitment', 
        content: '' 
      }
    ]);
    this.saveToLocalStorage();
  }

  async previewCommitment() {
    // Ensure we are using the latest from storage as requested, 
    // though in this single-page app state memory is usually sufficient.
    // We will re-read to be safe and strictly follow "refer to local storage".
    this.loadFromLocalStorage(); 

    try {
      const templateData = await this.http.get('commitment_Template.docx', { responseType: 'arraybuffer' }).toPromise();
      
      if (!templateData) {
        throw new Error('Could not load template');
      }

      const zip = new PizZip(templateData);
      
      let docXml = zip.file('word/document.xml')?.asText();
      if (docXml) {
        docXml = docXml.replace(/{[^{}]+}/g, (match: string) => {
           return match.replace(/<[^>]+>/g, '');
        });

        docXml = docXml.replace(/{language}/g, '{@language}');
        
        zip.file('word/document.xml', docXml);
      }

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      const commitmentsData = this.languages().map(lang => {
        const wml = this.convertHtmlToWml(lang.content);
        console.log('HTML Input:', lang.content);
        console.log('WML Output:', wml);
        return { language: wml };
      });

      doc.render({
        commitments: commitmentsData
      });

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      saveAs(out, 'Commitment_Preview.docx');
    } catch (error) {
      console.error('Error generating document:', error);
      alert('Failed to generate preview. Please ensure "commitment_Template.docx" exists in the public folder.');
    }
  }

  private convertHtmlToWml(html: string): string {
    if (!html || html.trim() === '') {
      return '<w:p><w:r><w:t></w:t></w:r></w:p>';
    }

    html = html.replace(/&nbsp;/g, ' ');

    const div = document.createElement('div');
    div.innerHTML = html;

    let wml = '';

    const nodes = Array.from(div.childNodes);
    
    if (!html.includes('<p>') && !html.includes('<P>')) {
      wml += '<w:p>';
      nodes.forEach(node => {
        wml += this.processNodeContent(node);
      });
      wml += '</w:p>';
    } else {
      nodes.forEach(node => {
        if (node.nodeName === 'P') {
          wml += '<w:p>';
          const childWml = this.processNodeContent(node);
          wml += childWml || '<w:r><w:t></w:t></w:r>';
          wml += '</w:p>';
        } else if (node.nodeType === Node.TEXT_NODE && node.textContent?.trim()) {
          wml += '<w:p>';
          wml += this.processNodeContent(node);
          wml += '</w:p>';
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          wml += '<w:p>';
          wml += this.processNodeContent(node);
          wml += '</w:p>';
        }
      });
    }

    return wml || '<w:p><w:r><w:t></w:t></w:r></w:p>';
  }

  private processNodeContent(node: Node): string {
    let wml = '';
    
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || '';
      if (text) {
        wml += `<w:r><w:t xml:space="preserve">${this.escapeXml(text)}</w:t></w:r>`;
      }
    } else if (node.nodeName === 'A') {
      const anchor = node as HTMLAnchorElement;
      const href = anchor.getAttribute('href') || '';
      const text = anchor.textContent || '';
      
      wml += `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>`;
      wml += `<w:r><w:instrText xml:space="preserve"> HYPERLINK "${this.escapeXml(href)}" \\h </w:instrText></w:r>`;
      wml += `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
      wml += `<w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${this.escapeXml(text)}</w:t></w:r>`;
      wml += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
      
    } else if (node.nodeName === 'STRONG' || node.nodeName === 'B') {
      wml += `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${this.escapeXml(node.textContent || '')}</w:t></w:r>`;
    } else if (node.nodeName === 'EM' || node.nodeName === 'I') {
      wml += `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${this.escapeXml(node.textContent || '')}</w:t></w:r>`;
    } else if (node.nodeName === 'U') {
      wml += `<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${this.escapeXml(node.textContent || '')}</w:t></w:r>`;
    } else if (node.nodeName === 'SPAN') {
      const span = node as HTMLElement;
      const style = span.getAttribute('style') || '';
      let highlight = '';
      
      if (style.includes('background-color')) {
        highlight = '<w:highlight w:val="yellow"/>'; 
      }

      if (highlight) {
        wml += `<w:r><w:rPr>${highlight}</w:rPr><w:t xml:space="preserve">${this.escapeXml(node.textContent || '')}</w:t></w:r>`;
      } else if (node.hasChildNodes()) {
        Array.from(node.childNodes).forEach(child => {
          wml += this.processNodeContent(child);
        });
      } else {
        wml += `<w:r><w:t xml:space="preserve">${this.escapeXml(node.textContent || '')}</w:t></w:r>`;
      }
    } else if (node.nodeName === 'OL' || node.nodeName === 'UL') {
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeName === 'LI') {
          wml += '</w:p><w:p>';
          wml += this.processNodeContent(child);
        }
      });
    } else if (node.nodeName === 'BR') {
      wml += '<w:r><w:br/></w:r>';
    } else if (node.hasChildNodes()) {
      Array.from(node.childNodes).forEach(child => {
        wml += this.processNodeContent(child);
      });
    }

    return wml;
  }

  private escapeXml(unsafe: string): string {
    return unsafe.replace(/[<>&'"]/g, (c) => {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
        default: return c;
      }
    });
  }
}