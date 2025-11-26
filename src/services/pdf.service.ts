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
    pageOffset: number
  ): Promise<{ pdfBytes: Uint8Array; pageCount: number }> {
    const { PDFDocument, rgb, StandardFonts } = PDFLib;

    const originalPdfDoc = await PDFDocument.load(originalPdfBytes);
    const qcPackPdfDoc = await PDFDocument.create();
    const pageTexts = await this.extractPdfTextWithItems(originalPdfBytes);

    // pageNum -> corrections + underline segments
    const correctionsByPage = new Map<number, (Correction & { underlineSegments: UnderlineSegment[] })[]>();

    for (const corr of corrections) {
      const pageNum = corr.Page + pageOffset;

      if (!correctionsByPage.has(pageNum)) {
        correctionsByPage.set(pageNum, []);
      }

      const pageData = pageTexts.find(pt => pt.pageNum === pageNum);
      let underlineSegments: UnderlineSegment[] = [];

      if (pageData) {
        // Put items in reading order first
        const sortedItems = this.groupItemsIntoLines(pageData.items).flat();

        const ranges = this.findItemSegmentsForPhrase(corr.ContextPhrase, sortedItems);

        if (ranges) {
          underlineSegments = ranges.map(r => ({
            item: sortedItems[r.itemIndex],
            startFrac: r.startFrac,
            endFrac: r.endFrac,
          }));
        } else {
          console.warn(
            `Could not precisely match context phrase on page ${pageNum}. Context Phrase:`,
            corr.ContextPhrase
          );
        }
      } else {
        console.warn(`Could not find page ${pageNum} in script PDF for correction:`, corr);
      }

      correctionsByPage.get(pageNum)!.push({ ...corr, underlineSegments });
    }

    const pagesToInclude = Array.from(correctionsByPage.keys()).sort((a, b) => a - b);
    const timesRomanFont = await qcPackPdfDoc.embedFont(StandardFonts.TimesRoman);

    for (const pageNum of pagesToInclude) {
      const pageIndex = pageNum - 1;
      if (pageIndex < 0 || pageIndex >= originalPdfDoc.getPageCount()) continue;

      const [copiedPage] = await qcPackPdfDoc.copyPages(originalPdfDoc, [pageIndex]);
      const correctionsForPage = correctionsByPage.get(pageNum)!;

      // 🔽 Draw underlines: ONLY the matched phrase segments
      for (const correction of correctionsForPage) {
        for (const seg of correction.underlineSegments) {
          const { item, startFrac, endFrac } = seg;

          // safety clamp
          const clampedStart = Math.max(0, Math.min(1, startFrac));
          const clampedEnd = Math.max(clampedStart + 0.01, Math.min(1, endFrac)); // ensure at least a tiny line

          const x1 = item.x + item.width * clampedStart;
          const x2 = item.x + item.width * clampedEnd;
          const y = item.y;

          copiedPage.drawLine({
            start: { x: x1, y: y - 2 },
            end: { x: x2, y: y - 2 },
            thickness: 1,
            color: rgb(0, 0, 0),
          });
        }
      }

      // Notes box
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

  // ---------------------------------------------------------------------------
  // Notes box
  // ---------------------------------------------------------------------------

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

  // ---------------------------------------------------------------------------
  // pdf.js text extraction
  // ---------------------------------------------------------------------------

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

  // ---------------------------------------------------------------------------
  // Robust phrase matcher: returns per-item fractional coverage
  // ---------------------------------------------------------------------------

  private findItemSegmentsForPhrase(
    textToFind: string,
    pageItems: PageTextItem[]
  ): { itemIndex: number; startFrac: number; endFrac: number }[] | null {
    const normalizedSearch = this.normalizeForSearch(textToFind);
    if (!normalizedSearch) return null;

    const searchTokens = normalizedSearch.split(' ').filter(Boolean);
    if (searchTokens.length === 0) return null;

    // Precompute normalized strings for items
    const itemNorms: string[] = pageItems.map(i => this.normalizeForSearch(i.str));

    interface TokenEntry {
      token: string;
      itemIndex: number;
      itemNormStart: number;
      itemNormEnd: number;
    }

    const tokenList: TokenEntry[] = [];

    // Build token list with positions inside each itemNorm
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
      console.warn('No tokens on page when searching for:', textToFind);
      return null;
    }

    // Try to match searchTokens against tokenList in order.
    for (let startTokIdx = 0; startTokIdx < tokenList.length; startTokIdx++) {
      let tIdx = startTokIdx;
      let sIdx = 0;

      while (sIdx < searchTokens.length && tIdx < tokenList.length) {
        const pageTok = tokenList[tIdx].token;
        const searchTok = searchTokens[sIdx];

        if (pageTok === searchTok) {
          // Direct match
          sIdx++;
          tIdx++;
        } else if (
          tIdx + 1 < tokenList.length &&
          pageTok + tokenList[tIdx + 1].token === searchTok
        ) {
          // Handle split words: "par" + "ticular" vs "particular"
          sIdx++;
          tIdx += 2;
        } else {
          break;
        }
      }

      if (sIdx === searchTokens.length) {
        // Match found from tokenList[startTokIdx .. tIdx-1]
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

  // ---------------------------------------------------------------------------
  // Group items into lines (for reading order)
  // ---------------------------------------------------------------------------

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

  // ---------------------------------------------------------------------------
  // Normalization & timestamp helper
  // ---------------------------------------------------------------------------

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
