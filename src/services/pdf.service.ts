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
        const mainPageNum = corr.Page + pageOffset;
        let matchFound = false;

        const processMatch = (pageNum: number, segments: UnderlineSegment[]) => {
            let oblongSegments: UnderlineSegment[] = [];
            if (corr.wordsForOblong && corr.wordsForOblong.length > 0 && segments.length > 0) {
                const sentenceItems = segments.map(seg => seg.item);
                
                if (corr.correctionType === 'inserted' && corr.wordsForOblong.length === 2) {
                    // Robust logic: Find each boundary word individually within the sentence context.
                    const word1Segments = this.findItemSegmentsForPhrase(corr.wordsForOblong[0], sentenceItems);
                    const word2Segments = this.findItemSegmentsForPhrase(corr.wordsForOblong[1], sentenceItems);

                    if (word1Segments && word1Segments.length > 0 && word2Segments && word2Segments.length > 0) {
                        const firstWord1Range = word1Segments[0];
                        const firstWord2RangeAfterWord1 = word2Segments.find(seg => 
                            seg.itemIndex > firstWord1Range.itemIndex || 
                            (seg.itemIndex === firstWord1Range.itemIndex && seg.startFrac >= firstWord1Range.endFrac)
                        );

                        if (firstWord2RangeAfterWord1) {
                            oblongSegments = [
                                { item: sentenceItems[firstWord1Range.itemIndex], startFrac: firstWord1Range.startFrac, endFrac: firstWord1Range.endFrac },
                                { item: sentenceItems[firstWord2RangeAfterWord1.itemIndex], startFrac: firstWord2RangeAfterWord1.startFrac, endFrac: firstWord2RangeAfterWord1.endFrac }
                            ];
                        } else {
                            console.warn(`Could not find boundary word 2 after word 1 for INSERTION on page ${pageNum}.`, corr);
                        }
                    } else {
                        console.warn(`Could not find one or both boundary words for INSERTION on page ${pageNum}.`, corr);
                    }
                } else if (corr.correctionType === 'misread' || corr.correctionType === 'missing') {
                    const finalOblongRanges = this.findItemSegmentsForPhrase(corr.wordsForOblong.join(' '), sentenceItems);
                     if (finalOblongRanges) {
                        oblongSegments = finalOblongRanges.map(r => ({
                            item: sentenceItems[r.itemIndex],
                            startFrac: r.startFrac,
                            endFrac: r.endFrac,
                        }));
                    } else {
                        console.warn(`Could not find words for oblong WITHIN CONTEXT on page ${pageNum}. Words:`, corr.wordsForOblong.join(' '));
                    }
                }
            }
            if (!correctionsByPage.has(pageNum)) {
                correctionsByPage.set(pageNum, []);
            }
            correctionsByPage.get(pageNum)!.push({ ...corr, underlineSegments: segments, oblongSegments });
        };

        // For non-Audible projects, we perform text search and highlighting
        if (!isAudible) {
            const pageData = pageTexts.find(pt => pt.pageNum === mainPageNum);
            
            // Attempt 1: Single page search
            if (pageData) {
                const sortedItems = this.groupItemsIntoLines(pageData.items).flat();
                const underlineRanges = this.findItemSegmentsForPhrase(corr.ContextPhrase, sortedItems);
                if (underlineRanges) {
                    matchFound = true;
                    const segments = underlineRanges.map(r => ({ item: sortedItems[r.itemIndex], startFrac: r.startFrac, endFrac: r.endFrac }));
                    processMatch(mainPageNum, segments);
                }
            }

            // Attempt 2: Forward cross-page search
            if (!matchFound) {
                const nextPageNum = mainPageNum + 1;
                const nextPageData = pageTexts.find(pt => pt.pageNum === nextPageNum);
                if (pageData && nextPageData) {
                    const items1 = this.groupItemsIntoLines(pageData.items).flat();
                    const items2 = this.groupItemsIntoLines(nextPageData.items).flat();
                    const combinedItems = [...items1, ...items2];
                    const underlineRanges = this.findItemSegmentsForPhrase(corr.ContextPhrase, combinedItems);
                    if (underlineRanges) {
                        matchFound = true;
                        const segments1 = underlineRanges
                            .filter(r => r.itemIndex < items1.length)
                            .map(r => ({ item: combinedItems[r.itemIndex], startFrac: r.startFrac, endFrac: r.endFrac }));
                        const segments2 = underlineRanges
                            .filter(r => r.itemIndex >= items1.length)
                            .map(r => ({ item: combinedItems[r.itemIndex], startFrac: r.startFrac, endFrac: r.endFrac }));
                        
                        if (segments1.length > 0) processMatch(mainPageNum, segments1);
                        if (segments2.length > 0) processMatch(nextPageNum, segments2);
                    }
                }
            }

            // Attempt 3: Backward cross-page search
            if (!matchFound) {
                const prevPageNum = mainPageNum - 1;
                const prevPageData = pageTexts.find(pt => pt.pageNum === prevPageNum);
                if (prevPageData && pageData) {
                    const items1 = this.groupItemsIntoLines(prevPageData.items).flat();
                    const items2 = this.groupItemsIntoLines(pageData.items).flat();
                    const combinedItems = [...items1, ...items2];
                    const underlineRanges = this.findItemSegmentsForPhrase(corr.ContextPhrase, combinedItems);
                    if (underlineRanges) {
                        matchFound = true;
                        const segments1 = underlineRanges
                            .filter(r => r.itemIndex < items1.length)
                            .map(r => ({ item: combinedItems[r.itemIndex], startFrac: r.startFrac, endFrac: r.endFrac }));
                        const segments2 = underlineRanges
                            .filter(r => r.itemIndex >= items1.length)
                            .map(r => ({ item: combinedItems[r.itemIndex], startFrac: r.startFrac, endFrac: r.endFrac }));
                        
                        if (segments1.length > 0) processMatch(prevPageNum, segments1);
                        if (segments2.length > 0) processMatch(mainPageNum, segments2);
                    }
                }
            }
        }

        // If no match found (or if Audible), add correction to main page without any segments.
        if (!matchFound) {
            if (!isAudible) {
                console.warn(`Could not find context phrase for correction on page ${mainPageNum}:`, corr.ContextPhrase);
            }
            if (!correctionsByPage.has(mainPageNum)) {
                correctionsByPage.set(mainPageNum, []);
            }
            correctionsByPage.get(mainPageNum)!.push({ ...corr, underlineSegments: [], oblongSegments: [] });
        }
    }


    const pagesToInclude = Array.from(correctionsByPage.keys()).sort((a, b) => a - b);
    const notesFont = await qcPackPdfDoc.embedFont(StandardFonts.Helvetica);

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
      this.drawNotesBox(copiedPage, allNotesForPage, notesFont, rgb);

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
    const topMargin = 25;
    const leftMargin = 25;
    const fontSize = 10;
    const lineHeight = 13;
    const maxWidth = width * 0.4;
    const padding = 8;

    // Calculate the total height required for the text block
    const lines = text.split('\n');
    let textBlockHeight = 0;
    lines.forEach(line => {
      const wrappedLineCount = Math.ceil(font.widthOfTextAtSize(line, fontSize) / maxWidth) || 1;
      textBlockHeight += wrappedLineCount * lineHeight;
    });
    // Adjust for the last line's leading to make padding even
    if (text.length > 0) {
      textBlockHeight -= (lineHeight - fontSize);
    }
    
    const boxWidth = maxWidth + (padding * 2);
    const boxHeight = textBlockHeight + (padding * 2);

    const boxX = leftMargin;
    const boxY = height - topMargin - boxHeight;

    // Draw the box
    page.drawRectangle({
      x: boxX,
      y: boxY,
      width: boxWidth,
      height: boxHeight,
      color: rgb(1, 1, 1),
      borderColor: rgb(0, 0, 0),
      borderWidth: 1,
    });

    // Draw the text line by line inside the box
    let currentY = boxY + boxHeight - padding - fontSize;
    for (const line of lines) {
      page.drawText(line, {
        x: boxX + padding,
        y: currentY,
        size: fontSize,
        font: font,
        color: rgb(0, 0, 0),
        lineHeight: lineHeight,
        maxWidth: maxWidth,
      });
      const wrappedLineCount = Math.ceil(font.widthOfTextAtSize(line, fontSize) / maxWidth) || 1;
      currentY -= wrappedLineCount * lineHeight;
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
    if (!textToFind || !pageItems || pageItems.length === 0) {
      return null;
    }
  
    // 1. Build a continuous string representation (corpus) and a map from its characters back to the original items.
    let corpus = '';
    const charMap: { itemIndex: number; charIndexInItem: number }[] = [];
    pageItems.forEach((item, itemIndex) => {
      for (let i = 0; i < item.str.length; i++) {
        corpus += item.str[i];
        charMap.push({ itemIndex, charIndexInItem: i });
      }
    });
  
    // 2. Normalize both the corpus and the search phrase. Normalization involves collapsing whitespace
    // and removing characters that interfere with matching, while keeping track of indices.
    const normalizedSearch = this.normalizeForSearch(textToFind);
    const { normalizedText: normalizedCorpus, originalIndices } = this.normalizeAndMap(corpus);
  
    if (!normalizedSearch) return null;
  
    // 3. Find the normalized search phrase within the normalized corpus.
    const matchIndex = normalizedCorpus.indexOf(normalizedSearch);
  
    if (matchIndex === -1) {
        // console.warn('Phrase not found on page for:', textToFind); // This can be noisy, disable for now
        return null;
    }
  
    // 4. Map the start and end of the match in the normalized corpus back to indices in the original corpus.
    const matchEndIndex = matchIndex + normalizedSearch.length - 1;
    const originalStartIndex = originalIndices[matchIndex];
    const originalEndIndex = originalIndices[matchEndIndex];
  
    // 5. Use the character map to find all the original items and character ranges that are part of the match.
    const perItem = new Map<number, { start: number; end: number }>();
    for (let i = originalStartIndex; i <= originalEndIndex; i++) {
      if (i >= charMap.length) continue;
      const { itemIndex, charIndexInItem } = charMap[i];
      if (itemIndex === -1) continue; // Skip characters that were not in an original item
  
      const existing = perItem.get(itemIndex);
      if (!existing) {
        perItem.set(itemIndex, { start: charIndexInItem, end: charIndexInItem });
      } else {
        existing.start = Math.min(existing.start, charIndexInItem);
        existing.end = Math.max(existing.end, charIndexInItem);
      }
    }
  
    // 6. Convert the character ranges into fractional start/end positions for each item.
    const ranges = Array.from(perItem.entries()).map(([itemIndex, range]) => {
      const item = pageItems[itemIndex];
      const itemLen = item.str.length || 1;
      
      // The end character is inclusive, so add 1 to get the exclusive end position for slicing.
      const startFrac = range.start / itemLen;
      const endFrac = (range.end + 1) / itemLen;
  
      return {
        itemIndex,
        startFrac: Math.max(0, Math.min(1, startFrac)),
        endFrac: Math.max(0, Math.min(1, endFrac)),
      };
    }).sort((a, b) => a.itemIndex - b.itemIndex);
  
    return ranges.length > 0 ? ranges : null;
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

  private normalizeAndMap(s: string): { normalizedText: string, originalIndices: number[] } {
    if (!s) return { normalizedText: '', originalIndices: [] };
    
    let normalized = '';
    const indices: number[] = [];
    let lastCharWasSpace = true;

    for (let i = 0; i < s.length; i++) {
        let char = s[i].toLowerCase();

        // Handle ligatures
        if (char === 'ﬁ') { char = 'fi'; }
        else if (char === 'ﬂ') { char = 'fl'; }
        // ... add other ligatures if needed

        const isSpace = /\s/.test(char);
        
        if (isSpace) {
            if (!lastCharWasSpace) {
                normalized += ' ';
                indices.push(i);
                lastCharWasSpace = true;
            }
        } else if (char === '-') {
            // Heuristic for soft hyphens used for line/page breaks.
            // If the very next character in the raw corpus is a letter, assume it's a soft hyphen
            // and skip it to join the word parts.
            if (i + 1 < s.length && /[a-zA-Z]/.test(s[i + 1])) {
                // Don't add a space or the hyphen. The next char will append directly.
                lastCharWasSpace = false; 
            } else {
                // It's a hard hyphen (e.g., state-of-the-art) or at the end of text. Treat as a space.
                if (!lastCharWasSpace) {
                    normalized += ' ';
                    indices.push(i);
                    lastCharWasSpace = true;
                }
            }
        } else if (/[a-z0-9']/.test(char) || /[‘’]/.test(char) || /[“”]/.test(char)) {
            // Standardize quotes
            if (/[‘’]/.test(char)) char = "'";
            if (/[“”]/.test(char)) char = '"';

            normalized += char;
            indices.push(i);
            lastCharWasSpace = false;
        } else {
             // Treat other punctuation as a potential space break
             if (!lastCharWasSpace) {
                normalized += ' ';
                indices.push(i);
                lastCharWasSpace = true;
            }
        }
    }
    return { normalizedText: normalized.trim(), originalIndices: indices };
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