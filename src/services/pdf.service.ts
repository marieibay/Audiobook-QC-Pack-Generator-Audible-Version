import { Injectable } from '@angular/core';
import { Correction, PageText, PageTextItem } from '../models';

declare var PDFLib: any;
declare var pdfjsLib: any;

interface UnderlineSegment {
  item: PageTextItem;
  startFrac: number; // 0–1 within item.width
  endFrac: number;   // 0–1 within item.width
}

@Injectable({ providedIn: 'root' })
export class PdfService {

  async createQCPack(
    originalPdfBytes: ArrayBuffer,
    corrections: Correction[],
    pageOffset: number,
    isAudible: boolean
  ): Promise<{ pdfBytes: Uint8Array; pageCount: number }> {
    const { PDFDocument, rgb, StandardFonts, cmyk } = PDFLib;

    const originalPdfDoc = await PDFDocument.load(originalPdfBytes);
    const qcPackPdfDoc = await PDFDocument.create();
    const pageTexts = await this.extractPdfTextWithItems(originalPdfBytes);

    const correctionsByPage = new Map<number, (Correction & { 
        underlineSegments: UnderlineSegment[], 
        oblongSegments: UnderlineSegment[],
    })[]>();

    for (const corr of corrections) {
      const pageNum = corr.Page + pageOffset;

      if (!correctionsByPage.has(pageNum)) {
        correctionsByPage.set(pageNum, []);
      }

      const pageData = pageTexts.find(pt => pt.pageNum === pageNum);
      let underlineSegments: UnderlineSegment[] = [];
      let oblongSegments: UnderlineSegment[] = [];

      if (!isAudible) {
        if (pageData) {
          const sortedItems = this.groupItemsIntoLines(pageData.items).flat();
          const underlineRanges = this.findItemSegmentsForPhrase(corr.ContextPhrase, sortedItems);
          
          if (underlineRanges) {
            underlineSegments = underlineRanges.map(r => ({
              item: sortedItems[r.itemIndex],
              startFrac: r.startFrac,
              endFrac: r.endFrac,
            }));
          } else {
            console.warn(`Could not precisely match context phrase on page ${pageNum}. Context Phrase:`, corr.ContextPhrase);
          }
          
          if ((corr.correctionType === 'misread' || corr.correctionType === 'missing' || corr.correctionType === 'inserted') && corr.wordsForOblong && corr.wordsForOblong.length > 0 && underlineSegments.length > 0) {
              const sentenceItems = underlineSegments.map(seg => seg.item);
              const oblongRanges = this.findItemSegmentsForPhrase(corr.wordsForOblong.join(' '), sentenceItems);

              if(oblongRanges) {
                   oblongSegments = oblongRanges.map(r => ({
                      item: sentenceItems[r.itemIndex],
                      startFrac: r.startFrac,
                      endFrac: r.endFrac,
                  }));
              } else {
                   console.warn(`Could not find words for oblong WITHIN CONTEXT on page ${pageNum}. Words:`, corr.wordsForOblong.join(' '));
              }
          }

        } else {
          console.warn(`Could not find page ${pageNum} in script PDF for correction:`, corr);
        }
      }

      correctionsByPage.get(pageNum)!.push({ ...corr, underlineSegments, oblongSegments });
    }

    const pagesToInclude = Array.from(correctionsByPage.keys()).sort((a, b) => a - b);
    const timesRomanFont = await qcPackPdfDoc.embedFont(StandardFonts.TimesRoman);

    for (const pageNum of pagesToInclude) {
      const pageIndex = pageNum - 1;
      if (pageIndex < 0 || pageIndex >= originalPdfDoc.getPageCount()) continue;

      const [copiedPage] = await qcPackPdfDoc.copyPages(originalPdfDoc, [pageIndex]);
      const correctionsForPage = correctionsByPage.get(pageNum)!;
      
      for (const correction of correctionsForPage) {
        for (const seg of correction.underlineSegments) {
          const { item, startFrac, endFrac } = seg;
          const clampedStart = Math.max(0, Math.min(1, startFrac));
          const clampedEnd = Math.max(clampedStart, Math.min(1, endFrac));
          if (clampedEnd <= clampedStart) continue;

          copiedPage.drawLine({
            start: { x: item.x + item.width * clampedStart, y: item.y - 2 },
            end: { x: item.x + item.width * clampedEnd, y: item.y - 2 },
            thickness: 1, color: rgb(0, 0, 0),
          });
        }
        
        const oblongsToDraw = this.groupSegmentsIntoLines(correction.oblongSegments);
        for(const lineOfSegments of oblongsToDraw) {
            if (lineOfSegments.length === 0) continue;
            
            const firstItem = lineOfSegments[0].item;
            const lastItem = lineOfSegments[lineOfSegments.length - 1].item;
            
            const x = firstItem.x + firstItem.width * lineOfSegments[0].startFrac;
            const width = (lastItem.x + lastItem.width * lineOfSegments[lineOfSegments.length - 1].endFrac) - x;
            
            const ellipseCenterX = x + width / 2;
            const ellipseCenterY = firstItem.y + (firstItem.height * 0.35) - 1;
            const ellipseXScale = width / 2 + 2;
            const ellipseYScale = (firstItem.height * 0.6) + 1;

            copiedPage.drawEllipse({
                x: ellipseCenterX,
                y: ellipseCenterY,
                xScale: ellipseXScale,
                yScale: ellipseYScale,
                borderColor: rgb(0, 0, 0),
                borderWidth: 1,
            });
        }
      }

      // Notes box
      const notesWithTimestamps = correctionsForPage.map(c => {
        let noteText = c.Notes;
        if (!isAudible) {
            const formattedTimestamp = this.formatTimestampForNote(c.Timestamp);
            if (c.Track && formattedTimestamp) {
              noteText += `\n${c.Track}/${formattedTimestamp}`;
            }
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

  private groupSegmentsIntoLines(segments: UnderlineSegment[]): UnderlineSegment[][] {
    if (!segments || segments.length === 0) return [];

    const lines = new Map<number, UnderlineSegment[]>();
    const Y_TOLERANCE = 5; 
    
    segments.forEach(seg => {
        let foundLine = false;
        for (const y of lines.keys()) {
            if (Math.abs(y - seg.item.y) < Y_TOLERANCE) {
                lines.get(y)!.push(seg);
                foundLine = true;
                break;
            }
        }
        if (!foundLine) {
            lines.set(seg.item.y, [seg]);
        }
    });
    
    return Array.from(lines.values()).map(line => line.sort((a,b) => a.item.x - b.item.x));
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

  private findItemSegmentsForPhrase(
    textToFind: string,
    pageItems: PageTextItem[]
  ): { itemIndex: number; startFrac: number; endFrac: number }[] | null {
    const normalizedSearch = this.normalizeForSearch(textToFind);
    if (!normalizedSearch) return null;

    const searchTokens = normalizedSearch.split(' ').filter(Boolean);
    if (searchTokens.length === 0) return null;

    const itemNorms: string[] = pageItems.map(i => this.normalizeForSearch(i.str));

    interface TokenEntry {
      token: string;
      itemIndex: number;
      itemNormStart: number;
      itemNormEnd: number;
    }

    const tokenList: TokenEntry[] = [];

    itemNorms.forEach((norm, itemIndex) => {
      if (!norm) return;
      const pieces = norm.split(' ').filter(Boolean);
      let offset = 0;

      for (const piece of pieces) {
        const start = norm.indexOf(piece, offset);
        if (start === -1) continue;
        const end = start + piece.length;
        tokenList.push({
          token: piece,
          itemIndex,
          itemNormStart: start,
          itemNormEnd: end,
        });
        offset = end;
      }
    });

    if (tokenList.length === 0) {
      return null;
    }

    for (let startTokIdx = 0; startTokIdx < tokenList.length; startTokIdx++) {
      let tIdx = startTokIdx;
      let sIdx = 0;

      while (sIdx < searchTokens.length && tIdx < tokenList.length) {
        const pageTok = tokenList[tIdx].token;
        const searchTok = searchTokens[sIdx];

        if (pageTok === searchTok) {
          sIdx++;
          tIdx++;
        } else if (
          tIdx + 1 < tokenList.length &&
          pageTok + tokenList[tIdx + 1].token === searchTok
        ) {
          sIdx++;
          tIdx += 2;
        } else {
          break;
        }
      }

      if (sIdx === searchTokens.length) {
        const perItem = new Map<number, { start: number; end: number }>();

        for (let k = startTokIdx; k < tIdx; k++) {
          const entry = tokenList[k];
          const existing = perItem.get(entry.itemIndex);
          if (!existing) {
            perItem.set(entry.itemIndex, {
              start: entry.itemNormStart,
              end: entry.itemNormEnd,
            });
          } else {
            existing.start = Math.min(existing.start, entry.itemNormStart);
            existing.end = Math.max(existing.end, entry.itemNormEnd);
          }
        }

        const ranges = Array.from(perItem.entries())
          .map(([itemIndex, range]) => {
            const norm = itemNorms[itemIndex];
            const totalLen = norm.length || 1;

            let startFrac = range.start / totalLen;
            let endFrac = range.end / totalLen;

            startFrac = Math.max(0, Math.min(1, startFrac));
            endFrac = Math.max(startFrac, Math.min(1, endFrac));

            return { itemIndex, startFrac, endFrac };
          })
          .sort((a, b) => a.itemIndex - b.itemIndex);

        if (ranges.length === 0) return null;
        return ranges;
      }
    }

    console.warn('Phrase not found on page for:', textToFind);
    return null;
  }

  private groupItemsIntoLines(items: PageTextItem[]): PageTextItem[][] {
    if (!items || items.length === 0) return [];

    const lines = new Map<number, PageTextItem[]>();
    const Y_TOLERANCE = 5;

    const sortedItems = [...items].sort((a, b) => {
      if (Math.abs(a.y - b.y) > Y_TOLERANCE) return b.y - a.y; // top to bottom
      return a.x - b.x; // left to right
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
      .map(entry => entry[1].sort((a, b) => a.x - b.x));
  }

  private normalizeForSearch(s: string): string {
    if (!s) return '';
    return s
      .toLowerCase()
      // ligatures
      .replace(/ﬁ/g, 'fi')
      .replace(/ﬂ/g, 'fl')
      .replace(/ﬀ/g, 'ff')
      .replace(/ﬃ/g, 'ffi')
      .replace(/ﬄ/g, 'ffl')
      // quotes
      .replace(/[‘’]/g, "'")
      .replace(/[“”]/g, '"')
      // keep letters, numbers, apostrophes, spaces
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