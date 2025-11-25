import { Component, signal, WritableSignal } from '@angular/core';
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
export class CommitmentCollationComponent {
  languages: WritableSignal<CommitmentLanguage[]> = signal([
    { code: '001', name: 'Commitment Name', content: '' }
  ]);

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
      // Restore selection to the last known position to ensure operations work
      // We need to focus it first often
      activeEditor.focus();
      
      const range = activeEditor.getSelection(); 
      
      if (range) {
        if (range.length > 0) {
          // Text is selected, turn it into a link
          activeEditor.format('link', doc.path);
        } else {
          // No text selected, insert document name and link it
          activeEditor.insertText(range.index, doc.name, 'link', doc.path);
          // Move cursor after the inserted link
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
  }

  async previewCommitment() {
    try {
      const templateData = await this.http.get('commitment_Template.docx', { responseType: 'arraybuffer' }).toPromise();
      
      if (!templateData) {
        throw new Error('Could not load template');
      }

      const zip = new PizZip(templateData);
      
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      const data: any = {};
      this.languages().forEach(lang => {
        data[`${lang.code}_content`] = lang.content;
      });
      data.languages = this.languages();

      doc.render(data);

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
}
