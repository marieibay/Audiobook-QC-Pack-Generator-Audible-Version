import { Component, ChangeDetectionStrategy, signal, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Correction, Status } from './models';
import { FileParserService } from './services/file-parser.service';
import { PdfService } from './services/pdf.service';

type UIState = 'idle' | 'parsing' | 'confirm' | 'generating' | 'complete';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule],
})
export class AppComponent {
  // Fix: Explicitly type injected services to resolve type inference issue.
  private fileParserService: FileParserService = inject(FileParserService);
  // Fix: Explicitly type injected services to resolve type inference issue.
  private pdfService: PdfService = inject(PdfService);

  qcFile = signal<File | null>(null);
  scriptFile = signal<File | null>(null);
  status = signal<Status | null>(null);
  generatedPdfBytes = signal<Uint8Array | null>(null);
  generatedPageCount = signal<number>(0);
  pageOffset = signal<number>(0);
  isAudibleProject = signal<boolean>(false);
  isPostQcProject = signal<boolean>(false);
  
  uiState = signal<UIState>('idle');
  parsedCorrections = signal<Correction[]>([]);
  instructionsVisible = signal(false);

  toggleInstructions(): void {
    this.instructionsVisible.update(visible => !visible);
  }

  onQcFileChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0] ?? null;
    this.qcFile.set(file);
    this.resetToIdle();
  }

  onScriptFileChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0] ?? null;
    this.scriptFile.set(file);
    this.resetToIdle();
  }

  onPageOffsetChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const value = input.valueAsNumber;
    this.pageOffset.set(isNaN(value) ? 0 : value);
  }

  onAudibleChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.isAudibleProject.set(input.checked);
  }

  onPostQcChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.isPostQcProject.set(input.checked);
  }
  
  async startParsing(): Promise<void> {
    const currentQcFile = this.qcFile();
    const isAudible = this.isAudibleProject();
    const isPostQc = this.isPostQcProject();

    if (!currentQcFile || !this.scriptFile()) {
      this.status.set({ text: 'Please upload both QC Report and Script files.', type: 'error' });
      return;
    }

    this.uiState.set('parsing');
    this.status.set({ text: 'Parsing QC report...', type: 'info' });

    try {
      const corrections = await this.fileParserService.parseQcFile(currentQcFile, isAudible, isPostQc);
      this.parsedCorrections.set(corrections);

      if (corrections.length === 0) {
        this.status.set({ text: 'Parsing complete. No corrections requiring a pickup were found.', type: 'warning' });
        this.uiState.set('complete');
        this.generatedPdfBytes.set(new Uint8Array()); 
        this.generatedPageCount.set(0);
      } else {
        this.status.set({ text: `Found ${corrections.length} corrections. Please confirm to proceed.`, type: 'info' });
        this.uiState.set('confirm');
      }
    } catch (error) {
      console.error('Error parsing QC file:', error);
      const message = error instanceof Error ? error.message : 'An unknown error occurred.';
      this.status.set({ text: `Failed to parse QC file: ${message}`, type: 'error' });
      this.uiState.set('idle');
    }
  }

  async generateConfirmedQCPack(): Promise<void> {
    const corrections = this.parsedCorrections();
    const currentScriptFile = this.scriptFile();
    const pageOffset = this.pageOffset();
    const isAudible = this.isAudibleProject();

    if (!currentScriptFile || corrections.length === 0) {
      this.status.set({ text: 'Cannot proceed. Script file or corrections are missing.', type: 'error' });
      this.uiState.set('idle');
      return;
    }

    this.uiState.set('generating');
    this.status.set({ text: 'Generating annotated QC Pack PDF...', type: 'info' });

    try {
      const scriptPdfBytes = await currentScriptFile.arrayBuffer();
      const { pdfBytes, pageCount } = await this.pdfService.createQCPack(scriptPdfBytes, corrections, pageOffset, isAudible);
      
      this.generatedPdfBytes.set(pdfBytes);
      this.generatedPageCount.set(pageCount);
      this.status.set({ text: `QC Pack generated successfully with ${pageCount} pages! Ready to download.`, type: 'success' });
      this.uiState.set('complete');

    } catch (error) {
      console.error('Error generating QC Pack:', error);
      const message = error instanceof Error ? error.message : 'An unknown error occurred.';
      this.status.set({ text: `Failed to generate QC Pack: ${message}`, type: 'error' });
      this.uiState.set('idle');
    }
  }

  downloadGeneratedPack(): void {
    const pdfBytes = this.generatedPdfBytes();
    if (pdfBytes && pdfBytes.length > 0) {
      const scriptFileName = this.scriptFile()?.name ?? 'script';
      const downloadName = `${scriptFileName.replace(/\.pdf$/i, '')}_QCPack.pdf`;
      this.downloadFile(pdfBytes, downloadName, 'application/pdf');
    }
  }
  
  resetToIdle(): void {
    this.uiState.set('idle');
    this.status.set(null);
    this.generatedPdfBytes.set(null);
    this.generatedPageCount.set(0);
    this.parsedCorrections.set([]);
  }

  reset(): void {
    this.qcFile.set(null);
    this.scriptFile.set(null);
    
    const qcInput = document.getElementById('qcInput') as HTMLInputElement;
    const scriptInput = document.getElementById('scriptInput') as HTMLInputElement;
    if (qcInput) qcInput.value = '';
    if (scriptInput) scriptInput.value = '';

    this.pageOffset.set(0);
    this.isAudibleProject.set(false);
    this.isPostQcProject.set(false);
    this.resetToIdle();
  }

  private downloadFile(data: Uint8Array, filename: string, mimeType: string): void {
    const blob = new Blob([data], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
}