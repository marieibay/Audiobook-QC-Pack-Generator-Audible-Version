import { Injectable } from '@angular/core';
import { Correction, PageText, PageTextItem } from '../models';

declare var PDFLib: any;
declare var pdfjsLib: any;

@Injectable({ providedIn: 'root' })
export class PdfService {

  async createQCPack(originalPdfBytes: ArrayBuffer, corrections: Correction[], pageOffset: number): Promise<{pdfBytes: Uint8Array, pageCount: number}> {
    const { PDFDocument, rgb, StandardFonts } = PDFLib;
    
    const originalPdfDoc = await PDFDocument.load(originalPdfBytes);
    const qcPackPdfDoc = await PDFDocument.create();
    const pageTexts = await this.extractPdfTextWithItems(originalPdfBytes);

    const correctionsByPage = new Map<number, (Correction & { itemsToUnderline: PageTextItem[] })[]>();

    for (const corr of corrections) {
        const pageNum = corr.Page + pageOffset;
        if (!correctionsByPage.has(pageNum)) {
            correctionsByPage.set(pageNum, []);
        }

        const pageData = pageTexts.find(pt => pt.pageNum === pageNum);
        let itemsToUnderline: PageTextItem[] = [];

        if (pageData) {
            const sortedItems = this.groupItemsIntoLines(pageData.items).flat();
            
            const contextLocation = this.findItemSequence(corr.ContextPhrase, sortedItems);
            
            if (contextLocation) {
                itemsToUnderline = sortedItems.slice(contextLocation.startIndex, contextLocation.endIndex + 1);
            } else {
                 console.warn(`Could not find context for correction on page ${pageNum}. Context Phrase:`, corr.ContextPhrase);
            }
        } else {
             console.warn(`Could not find page ${pageNum} in script PDF for correction:`, corr);
        }

        correctionsByPage.get(pageNum)!.push({ ...corr, itemsToUnderline });
    }

    const pagesToInclude = Array.from(correctionsByPage.keys()).sort((a, b) => a - b);
    const timesRomanFont = await qcPackPdfDoc.embedFont(StandardFonts.TimesRoman);

    for (const pageNum of pagesToInclude) {
        const pageIndex = pageNum - 1;
        if (pageIndex < 0 || pageIndex >= originalPdfDoc.getPageCount()) continue;

        const [copiedPage] = await qcPackPdfDoc.copyPages(originalPdfDoc, [pageIndex]);
        const correctionsForPage = correctionsByPage.get(pageNum)!;

        // Underlining disabled as per user request. The code is commented out but preserved.
        /*
        for (const correction of correctionsForPage) {
            const itemsToUnderline = correction.itemsToUnderline;
            
            if (itemsToUnderline.length > 0) {
                const sentenceLines = this.groupItemsIntoLines(itemsToUnderline);
                for (const line of sentenceLines) {
                    if (line.length === 0) continue;
                    const y = line[0].y;
                    const minX = Math.min(...line.map(i => i.x));
                    const maxX = Math.max(...line.map(i => i.x + i.width));
                    copiedPage.drawLine({
                        start: { x: minX, y: y - 2 },
                        end: { x: maxX, y: y - 2 },
                        thickness: 1, color: rgb(0, 0, 0),
                    });
                }
            }
        }
        */
        
        const notesWithTimestamps = correctionsForPage.map(c => {
            let noteText = c.Notes;
            const formattedTimestamp = this.formatTimestampForNote(c.Timestamp);
            if (c.Track && formattedTimestamp) {
                noteText += `\n${c.Track}/${formattedTimestamp}`;
            }
            return noteText;
        });

        const allNotesForPage = notesWithTimestamps.join('\n\n');
        this.drawNotesBox(copiedPage, allNotesForPage, timesRomanFont, rgb);
        
        qcPackPdfDoc.addPage(copiedPage);
    }
    
    return {
      pdfBytes: await qcPackPdfDoc.save(),
      pageCount: pagesToInclude.length,
    };
  }
  
  private drawNotesBox(page: any, text: string, font: any, rgb: any): void {
      const { width, height } = page.getSize();
      const topMargin = 30; 
      const fontSize = 12;
      const lineHeight = 15;
      const maxWidth = width * 0.6;
      const horizontalMargin = (width - maxWidth) / 2;
      
      const lines = text.split('\n');
      let textBlockHeight = 0;
      lines.forEach(line => {
        const lineCount = Math.ceil(font.widthOfTextAtSize(line, fontSize) / maxWidth);
        textBlockHeight += (lineCount > 0 ? lineCount : 1) * lineHeight;
      });
      textBlockHeight -= (lineHeight - fontSize);

      page.drawRectangle({
          x: horizontalMargin - 5,
          y: height - topMargin - textBlockHeight - 10,
          width: maxWidth + 10,
          height: textBlockHeight + 15,
          color: rgb(1, 1, 1),
          borderColor: rgb(0, 0, 0),
          borderWidth: 1,
      });

      const linesToDraw = text.split('\n');
      let currentY = height - topMargin - fontSize;

      for (const line of linesToDraw) {
          page.drawText(line, {
              x: horizontalMargin,
              y: currentY,
              size: fontSize,
              font: font,
              color: rgb(0, 0, 0),
              lineHeight: lineHeight,
              maxWidth: maxWidth,
          });
          
          const lineCount = Math.ceil(font.widthOfTextAtSize(line, fontSize) / maxWidth);
          const heightOfThisBlock = (lineCount > 0 ? lineCount : 1) * lineHeight;
          currentY -= heightOfThisBlock;
      }
  }

  private async extractPdfTextWithItems(pdfBytes: ArrayBuffer): Promise<PageText[]> {
    const pdfDoc = await pdfjsLib.getDocument({ data: pdfBytes }).promise;
    const pageTexts: PageText[] = [];
    for (let i = 1; i <= pdfDoc.numPages; i++) {
        const page = await pdfDoc.getPage(i);
        const textContent = await page.getTextContent();
        
        const items: PageTextItem[] = textContent.items.map((item: any) => ({
            str: item.str,
            x: item.transform[4],
            y: item.transform[5],
            width: item.width,
            height: item.height,
        }));

        const pageText = textContent.items.map((item: any) => item.str).join(' ');
        pageTexts.push({ pageNum: i, content: pageText, items });
    }
    return pageTexts;
  }

  private findItemSequence(textToFind: string, pageItems: PageTextItem[]): { startIndex: number; endIndex: number } | null {
    const normalizedSearch = this.normalizeForSearch(textToFind);
    if (!normalizedSearch) return null;

    // Try to find exact match first
    for (let startIdx = 0; startIdx < pageItems.length; startIdx++) {
      for (let endIdx = startIdx; endIdx < pageItems.length; endIdx++) {
        const sequence = pageItems.slice(startIdx, endIdx + 1);
        const sequenceText = sequence.map(item => item.str).join(' ');
        const normalized = this.normalizeForSearch(sequenceText);
        
        if (normalized === normalizedSearch) {
          return { startIndex: startIdx, endIndex: endIdx };
        }
        
        if (normalized.length > normalizedSearch.length * 2) {
          break;
        }
      }
    }
    
    // No exact match - we need to handle the case where the text is within a single item
    // or spans items with extra text. Find which single item contains the search text.
    for (let i = 0; i < pageItems.length; i++) {
      const itemText = pageItems[i].str;
      const normalized = this.normalizeForSearch(itemText);
      
      if (normalized.includes(normalizedSearch)) {
        // The entire search phrase is within this single item
        return { startIndex: i, endIndex: i };
      }
    }
    
    // Still no match - search phrase spans multiple items
    // Find the tightest span that contains it
    let bestMatch: { startIndex: number; endIndex: number } | null = null;
    let bestScore = Infinity;
    
    for (let startIdx = 0; startIdx < pageItems.length; startIdx++) {
      for (let endIdx = startIdx; endIdx < pageItems.length; endIdx++) {
        const sequence = pageItems.slice(startIdx, endIdx + 1);
        const sequenceText = sequence.map(item => item.str).join(' ');
        const normalized = this.normalizeForSearch(sequenceText);
        
        if (normalized.includes(normalizedSearch)) {
          const extraChars = normalized.length - normalizedSearch.length;
          
          if (extraChars < bestScore) {
            bestScore = extraChars;
            bestMatch = { startIndex: startIdx, endIndex: endIdx };
          }
        }
        
        if (normalized.length > normalizedSearch.length * 2) {
          break;
        }
      }
    }
    
    if (!bestMatch) {
      console.warn('Could not find match for:', textToFind);
      return null;
    }
    
    // Aggressively trim - remove one item at a time from each end
    let { startIndex, endIndex } = bestMatch;
    
    // Trim from start
    while (startIndex < endIndex) {
      const without = pageItems.slice(startIndex + 1, endIndex + 1);
      const withoutText = without.map(item => item.str).join(' ');
      const normalized = this.normalizeForSearch(withoutText);
      
      if (normalized.includes(normalizedSearch)) {
        startIndex++;
      } else {
        break;
      }
    }
    
    // Trim from end
    while (endIndex > startIndex) {
      const without = pageItems.slice(startIndex, endIndex);
      const withoutText = without.map(item => item.str).join(' ');
      const normalized = this.normalizeForSearch(withoutText);
      
      if (normalized.includes(normalizedSearch)) {
        endIndex--;
      } else {
        break;
      }
    }
    
    return { startIndex, endIndex };
  }

  private groupItemsIntoLines(items: PageTextItem[]): PageTextItem[][] {
    if (!items || items.length === 0) return [];

    const lines = new Map<number, PageTextItem[]>();
    const Y_TOLERANCE = 5; 
    
    const sortedItems = [...items].sort((a, b) => {
        if (Math.abs(a.y - b.y) > Y_TOLERANCE) return b.y - a.y;
        return a.x - b.x;
    });

    sortedItems.forEach(item => {
        if (item.str.trim() === '') return;
        let foundLine = false;
        for (const y of lines.keys()) {
            if (Math.abs(y - item.y) < Y_TOLERANCE) {
                lines.get(y)!.push(item);
                foundLine = true;
                break;
            }
        }
        if (!foundLine) {
            lines.set(item.y, [item]);
        }
    });
    
    return Array.from(lines.entries())
        .sort((a, b) => b[0] - a[0])
        .map(entry => entry[1].sort((a,b) => a.x - b.x));
  }
  
  private normalizeForSearch(s: string): string {
    if (!s) return '';
    return s
      .toLowerCase()
      .replace(/ﬁ/g, 'fi')
      .replace(/ﬂ/g, 'fl')
      .replace(/ﬀ/g, 'ff')
      .replace(/ﬃ/g, 'ffi')
      .replace(/ﬄ/g, 'ffl')
      .replace(/['']/g, "'")
      .replace(/[""]/g, '"')
      // Remove punctuation but keep letters, numbers, apostrophes, and spaces.
      .replace(/[^a-z0-9'\s]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private formatTimestampForNote(timestamp: string | undefined): string {
    if (!timestamp) return '';
    const parts = timestamp.split(':');
    if (parts.length === 3) {
        const minutes = parts[1];
        const seconds = Math.floor(parseFloat(parts[2])).toString().padStart(2, '0');
        return `${minutes}:${seconds}`;
    }
    if (parts.length === 2) {
        const minutes = parts[0];
        const seconds = Math.floor(parseFloat(parts[1])).toString().padStart(2, '0');
        return `${minutes}:${seconds}`;
    }
    return timestamp;
  }
}
