import { Injectable } from '@angular/core';
import { Correction, PageText, PageTextItem } from '../models';

declare var PDFLib: any;
declare var pdfjsLib: any;

interface UnderlineSegment {
  item: PageTextItem;
  itemIndex: number;
  startFrac: number;
  endFrac: number;
  startChar?: number; // Added for exact font measurement
  endChar?: number;   // Added for exact font measurement
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

      // For both Audible and non-Audible, we try to find the text first
      let searchPages = [mainPageNum];
      // Basic cross-page support (current, next, prev)
      searchPages.push(mainPageNum + 1);
      if (mainPageNum > 1) searchPages.push(mainPageNum - 1);

      // We need to access pageTexts for these pages
      const pagesData = searchPages.map(p => ({ pageNum: p, data: pageTexts.find(pt => pt.pageNum === p) })).filter(x => x.data);

      let foundInPageNum = -1;
      let matchIndices: { start: number; end: number } | null = null;
      let corpusInfo: { corpus: string; charMap: any[] } | null = null;
      let pageItems: PageTextItem[] = [];

      // Try to find the phrase in the pages
      for (const { pageNum, data } of pagesData) {
        if (!data) continue;

        // Flatten items
        const items = this.groupItemsIntoLines(data.items).flat();

        // Build corpus
        let corpus = '';
        const charMap: { itemIndex: number; charIndexInItem: number }[] = [];
        items.forEach((item, itemIndex) => {
          for (let i = 0; i < item.str.length; i++) {
            corpus += item.str[i];
            charMap.push({ itemIndex, charIndexInItem: i });
          }
        });

        const normalizedSearch = this.normalizeForSearch(corr.ContextPhrase);
        const { normalizedText: normalizedCorpus, originalIndices } = this.normalizeAndMap(corpus);

        // 1. Strict
        let matchIndex = normalizedCorpus.indexOf(normalizedSearch);
        let matchedLength = normalizedSearch.length;
        let indicesToMap = originalIndices;

        // 2. Fuzzy
        if (matchIndex === -1) {
          const aggressiveNormalize = (s: string) => {
            let n = '';
            const idxs: number[] = [];
            for (let i = 0; i < s.length; i++) {
              if (/[a-z0-9]/i.test(s[i])) {
                n += s[i].toLowerCase();
                idxs.push(i);
              }
            }
            return { text: n, indices: idxs };
          };
          const aggSearch = aggressiveNormalize(corr.ContextPhrase);
          const aggCorpus = aggressiveNormalize(corpus); // Use raw corpus
          const aggMatchIndex = aggCorpus.text.indexOf(aggSearch.text);
          if (aggMatchIndex !== -1) {
            matchIndex = aggCorpus.indices[aggMatchIndex];
            // end index calculation
            const aggMatchEnd = aggMatchIndex + aggSearch.text.length - 1;
            const bufEnd = aggCorpus.indices[aggMatchEnd];

            foundInPageNum = pageNum;
            matchIndices = { start: matchIndex, end: bufEnd };
            corpusInfo = { corpus, charMap };
            pageItems = items;
            break;
          }
        } else {
          // Strict match found
          foundInPageNum = pageNum;
          const startI = indicesToMap[matchIndex];
          const endI = indicesToMap[matchIndex + matchedLength - 1];
          matchIndices = { start: startI, end: endI };
          corpusInfo = { corpus, charMap };
          pageItems = items;
          break;
        }
      }

      if (foundInPageNum !== -1 && matchIndices && corpusInfo) {
        matchFound = true;
        const { start, end } = matchIndices!;
        const { corpus, charMap } = corpusInfo!;

        if (isAudible) {
          // --- Audible Logic ---

          // We have matchIndices (Context Phrase) and corpus for that match.
          // Search for the specific correction word(s) inside the Context Phrase to be precise.
          let specificStart = start;
          let specificEnd = end;

          const phrasesToFind = corr.correctionType === 'inserted' ? corr.wordsForOblong : [corr.wordsForOblong.join(' ')];
          const targetPhrase = phrasesToFind[0];

          if (targetPhrase) {
            const contextStr = corpus.substring(start, end + 1);

            // Normalize context segment and target phrase
            const normTarget = this.normalizeForSearch(targetPhrase);
            const { normalizedText: normContext, originalIndices: contextIndices } = this.normalizeAndMap(contextStr);

            const localMatchIndex = normContext.indexOf(normTarget);

            if (localMatchIndex !== -1) {
              // Found specific word inside context!
              const localStart = contextIndices[localMatchIndex];
              const localEnd = contextIndices[localMatchIndex + normTarget.length - 1]; // inclusive of last char

              specificStart = start + localStart;
              specificEnd = start + localEnd;
            }
          }

          // 1. Expand (up to) 3 words: [Prev] [Target] [Next] (STAY WITHIN PUNCTUATION BOUNDARIES)

          // Helper to check if a character is a word character
          const isWordChar = (char: string) => /[a-zA-Z0-9']/.test(char);
          // Characters that attached to words (should be underlined)
          const isAttachedPunctuation = (char: string) => /[.,!?;:"“”‘’()\[\]]/.test(char);
          // Characters that act as boundaries (do NOT cross these for the 3-word expansion)
          const isUnderlineBoundary = (char: string) => /[.?!,;:—()\[\]"“”]/.test(char);

          const findBoundaryStart = (fromIndex: number): number => {
            let i = fromIndex;
            while (i >= 0) {
              if (isUnderlineBoundary(corpus[i])) return i + 1;
              i--;
            }
            return 0;
          };

          const findBoundaryEnd = (fromIndex: number): number => {
            let i = fromIndex;
            while (i < corpus.length) {
              if (isUnderlineBoundary(corpus[i])) return i;
              i++;
            }
            return corpus.length - 1;
          };

          // Underline boundaries for the target word
          const underlineLimitStart = findBoundaryStart(specificStart - 1);
          const underlineLimitEnd = findBoundaryEnd(specificEnd);

          // Helper to expand a word range to include attached punctuation
          const expandWord = (startIdx: number, endIdx: number, limitStart: number, limitEnd: number) => {
            let s = startIdx;
            let e = endIdx;
            // Expand core word
            while (s > limitStart && isWordChar(corpus[s - 1])) s--;
            while (e < limitEnd && isWordChar(corpus[e + 1])) e++;
            // Expand to include leading/trailing punctuation
            while (s > limitStart && isAttachedPunctuation(corpus[s - 1])) s--;
            while (e < limitEnd && isAttachedPunctuation(corpus[e + 1])) e++;
            return { s, e };
          };

          // Find boundaries for current word
          const currentWord = expandWord(specificStart, specificEnd, underlineLimitStart, underlineLimitEnd);
          let wordStart = currentWord.s;
          let wordEnd = currentWord.e;

          // Now find the previous word (MUST BE WITHIN SAME PUNCTUATION BOUNDARY)
          if (currentWord.s > underlineLimitStart) {
            let p = currentWord.s - 1;
            // Skip spacing (within boundary)
            while (p >= underlineLimitStart && /\s/.test(corpus[p])) p--;
            // If we found a character (still in same boundary), treat as prev word
            if (p >= underlineLimitStart && !isUnderlineBoundary(corpus[p])) {
              const prevWord = expandWord(p, p, underlineLimitStart, underlineLimitEnd);
              wordStart = prevWord.s;
            }
          }

          // Now find the next word (MUST BE WITHIN SAME PUNCTUATION BOUNDARY)
          if (currentWord.e < underlineLimitEnd) {
            let n = currentWord.e + 1;
            // Skip spacing (within boundary)
            while (n <= underlineLimitEnd && /\s/.test(corpus[n])) n++;
            // If we found a character (still in same boundary), treat as next word
            if (n <= underlineLimitEnd && !isUnderlineBoundary(corpus[n])) {
              const nextWord = expandWord(n, n, underlineLimitStart, underlineLimitEnd);
              wordEnd = nextWord.e;
            }
          }

          const threeWordSegments = this.mapRangeToItems(wordStart, wordEnd, charMap, pageItems);

          // 2. Expand to 3 sentences: [Prev] [Target] [Next] (STAY WITHIN .?! BOUNDARIES)
          const isSentenceEnd = (char: string) => /[.?!]/.test(char);
          const findSentenceStart = (fromIndex: number): number => {
            let i = fromIndex;
            while (i >= 0) {
              if (isSentenceEnd(corpus[i])) return i + 1;
              i--;
            }
            return 0;
          };
          const findSentenceEnd = (fromIndex: number): number => {
            let i = fromIndex;
            while (i < corpus.length) {
              if (isSentenceEnd(corpus[i])) return i;
              i++;
            }
            return corpus.length - 1;
          };
          // Start of target phrase is inside "Target Sentence"
          // Scan back to find start of Target Sentence
          let targetSentStart = findSentenceStart(start - 1);
          // Scan back to find start of Prev Sentence
          let prevSentStart = findSentenceStart(targetSentStart - 2);

          // End of target phrase is inside "Target Sentence"
          // Scan forward to find end of Target Sentence
          let targetSentEnd = findSentenceEnd(end);
          // Scan forward to find end of Next Sentence
          let nextSentEnd = findSentenceEnd(targetSentEnd + 1);

          // Handle edge cases where punctuation might be spaces away or missing
          // For simplicity, we define the range [prevSentStart, nextSentEnd]
          const sentenceSegments = this.mapRangeToItems(prevSentStart, nextSentEnd, charMap, pageItems);

          if (!correctionsByPage.has(foundInPageNum)) correctionsByPage.set(foundInPageNum, []);

          // We'll store:
          // underlineSegments -> Used for Red Underline (3 words)
          // oblongSegments -> Used for Yellow Highlight (3 sentences) (abusing the field, but we'll check isAudible in drawing loop)
          correctionsByPage.get(foundInPageNum)!.push({
            ...corr,
            underlineSegments: threeWordSegments,
            oblongSegments: sentenceSegments
          });

        } else {
          // --- Standard Logic ---
          const segments = this.mapRangeToItems(start, end, charMap, pageItems);

          let oblongSegments: UnderlineSegment[] = [];
          // ... (Copy existing oblong logic for standard) ...
          if (corr.wordsForOblong && corr.wordsForOblong.length > 0 && segments.length > 0) {
            const sentenceItems = segments.map(seg => seg.item);
            let phrasesToFind: string[] = [];
            if (corr.correctionType === 'inserted') {
              phrasesToFind = corr.wordsForOblong;
            } else {
              phrasesToFind = [corr.wordsForOblong.join(' ')];
            }
            const allRanges = phrasesToFind.flatMap(phrase => this.findItemSegmentsForPhrase(phrase, sentenceItems) || []);
            if (allRanges.length > 0) {
              oblongSegments = allRanges.map(r => ({
                item: sentenceItems[r.itemIndex], itemIndex: r.itemIndex, startFrac: r.startFrac, endFrac: r.endFrac,
              }));
            }
          }

          if (!correctionsByPage.has(foundInPageNum)) correctionsByPage.set(foundInPageNum, []);
          correctionsByPage.get(foundInPageNum)!.push({ ...corr, underlineSegments: segments, oblongSegments });
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

        if (isAudible) {
          // --- DRAW AUDIBLE STYLE ---

          // 1. Draw Yellow Highlight for Sentences (oblongSegments)
          // We need to merge them into lines first
          const sentenceLines = this.groupSegmentsIntoLines(correction.oblongSegments);
          for (const line of sentenceLines) {
            if (line.length === 0) continue;
            const first = line[0];
            const last = line[line.length - 1];
            const startX = first.item.x + first.item.width * first.startFrac;
            const endX = last.item.x + last.item.width * last.endFrac;
            const y = first.item.y;
            const height = first.item.height || 10;

            copiedPage.drawRectangle({
              x: startX,
              y: y - 2, // slight padding
              width: endX - startX,
              height: height + 4,
              color: rgb(1, 1, 0), // Yellow
              opacity: 0.75,
              blendMode: PDFLib.BlendMode.Multiply,
            });
          }

          // 2. Draw Red Underline for Words (underlineSegments)
          for (const seg of correction.underlineSegments) {
            const { item, startChar, endChar } = seg;

            // Calculate precise offsets using font metrics if indices available
            let startX = item.x;
            let endX = item.x + item.width;

            if (startChar !== undefined && endChar !== undefined) {
              const fontSize = item.height || 12;
              const textBefore = item.str.substring(0, startChar);
              const textWithin = item.str.substring(0, endChar + 1);

              // Measure exact widths using the notes font (Helvetica)
              // Note: If the actual PDF font is very different, this is still an approximation,
              // but it's much better than linear interpolation.
              const offsetStart = notesFont.widthOfTextAtSize(textBefore, fontSize);
              const offsetEnd = notesFont.widthOfTextAtSize(textWithin, fontSize);

              startX = item.x + offsetStart;
              endX = item.x + offsetEnd;
            } else {
              // Fallback to fractional if indices missing
              startX = item.x + item.width * seg.startFrac;
              endX = item.x + item.width * seg.endFrac;
            }

            const padding = 2.0; // points of bleed for visibility

            copiedPage.drawLine({
              start: { x: startX - padding, y: item.y - 2 },
              end: { x: endX + padding, y: item.y - 2 },
              thickness: 1.3, color: rgb(1, 0, 0), // Red
            });
          }

        } else {
          // --- DRAW STANDARD STYLE ---
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
          for (const lineOfSegments of oblongsToDraw) {
            if (lineOfSegments.length === 0) continue;

            const firstItem = lineOfSegments[0].item;
            const lastItem = lineOfSegments[lineOfSegments.length - 1].item;

            const x = firstItem.x + firstItem.width * lineOfSegments[0].startFrac;
            const width = (lastItem.x + lastItem.width * lineOfSegments[lineOfSegments.length - 1].endFrac) - x;

            const ellipseCenterX = x + width / 2;
            const ellipseCenterY = firstItem.y + (firstItem.height * 0.45);
            const ellipseXScale = width / 2 + 10;
            const ellipseYScale = (firstItem.height * 0.5) + 2;

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
      }

      // Notes box logic - Group identical notes
      const groupedCorrections = new Map<string, (Correction & {
        underlineSegments: UnderlineSegment[],
        oblongSegments: UnderlineSegment[],
      })[]>();

      for (const c of correctionsForPage) {
        const noteKey = c.Notes.trim();
        if (!groupedCorrections.has(noteKey)) {
          groupedCorrections.set(noteKey, []);
        }
        groupedCorrections.get(noteKey)!.push(c);
      }

      const noteBlocks: string[] = [];

      for (const [noteText, group] of groupedCorrections) {
        let block = noteText;

        if (!isAudible) {
          const timestamps = group
            .map(c => {
              const ts = this.formatTimestampForNote(c.Timestamp);
              if (c.Track && ts) {
                return `${c.Track}/${ts}`;
              } else if (ts) {
                return ts;
              }
              return '';
            })
            .filter(s => s !== '');

          if (timestamps.length > 0) {
            block += `\n${timestamps.join(' - ')}`;
          }
        }
        noteBlocks.push(block);
      }

      const allNotesForPage = noteBlocks.join('\n\n');
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

    return Array.from(lines.values()).map(line => line.sort((a, b) => a.item.x - b.item.x));
  }

  private drawNotesBox(page: any, text: string, font: any, rgb: any): void {
    const { width, height } = page.getSize();
    const topMargin = 25;
    const leftMargin = 25;
    const fontSize = 10;
    const lineHeight = 12;
    const maxWidth = width * 0.6;
    const padding = 8;

    // Helper to wrap text manually to avoid pdf-lib wrapping mismatch
    const wrapText = (paragraph: string): string[] => {
      const words = paragraph.split(' ');
      let lines: string[] = [];
      let currentLine = words[0];

      for (let i = 1; i < words.length; i++) {
        const word = words[i];
        const testLine = currentLine + ' ' + word;
        const width = font.widthOfTextAtSize(testLine, fontSize);
        if (width <= maxWidth) {
          currentLine = testLine;
        } else {
          lines.push(currentLine);
          currentLine = word;
        }
      }
      lines.push(currentLine);
      return lines;
    };

    const paragraphs = text.split('\n');
    const linesToDraw: string[] = [];

    paragraphs.forEach(paragraph => {
      if (paragraph === '') {
        linesToDraw.push('');
      } else {
        const wrapped = wrapText(paragraph);
        linesToDraw.push(...wrapped);
      }
    });

    let textBlockHeight = linesToDraw.length * lineHeight;

    // Adjust for the last line's leading to make padding even
    if (linesToDraw.length > 0) {
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

    // Draw the text line by line
    let currentY = boxY + boxHeight - padding - fontSize;
    for (const line of linesToDraw) {
      page.drawText(line, {
        x: boxX + padding,
        y: currentY,
        size: fontSize,
        font: font,
        color: rgb(0, 0, 0),
      });
      currentY -= lineHeight;
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

    // 3. Find the normalized search phrase within the normalized corpus (Attempt 1: Standard Normalization)
    let matchIndex = normalizedCorpus.indexOf(normalizedSearch);
    let matchedLength = normalizedSearch.length;
    let indicesToUse = originalIndices;

    // Attempt 2: Aggressive "Fuzzy" Search (Strip all non-alphanumeric)
    if (matchIndex === -1) {
      const aggressiveNormalize = (s: string) => {
        let n = '';
        const idxs: number[] = [];
        for (let i = 0; i < s.length; i++) {
          if (/[a-z0-9]/i.test(s[i])) {
            n += s[i].toLowerCase();
            idxs.push(i);
          }
        }
        return { text: n, indices: idxs };
      };

      const aggSearch = aggressiveNormalize(textToFind);
      const aggCorpus = aggressiveNormalize(corpus);

      const aggMatchIndex = aggCorpus.text.indexOf(aggSearch.text);
      if (aggMatchIndex !== -1) {
        // Found it with aggressive search!
        // We need to map the aggCorpus indices back to the original corpus indices.
        // aggCorpus.indices[x] gives the index in 'corpus' where the x-th char of aggCorpus came from.

        matchIndex = aggCorpus.indices[aggMatchIndex]; // Start index in original corpus

        // For end index, we look at the last character of the match in aggCorpus
        const aggMatchEndIndex = aggMatchIndex + aggSearch.text.length - 1;
        const corpusEndIndex = aggCorpus.indices[aggMatchEndIndex];

        // We can't use 'originalIndices' map from Attempt 1 directly because that one included spaces/punctuation.
        // But we have mapped directly to 'corpus' indices now.
        // So we can construct a fake 'indicesToUse' that is just a 1:1 map for the range we care about, 
        // OR simpler: just operate on corpus indices directly.

        // Let's adapt the variables to flow into step 4.
        // Step 4 expects 'matchIndex' to be an index into 'normalizedCorpus' array of 'originalIndices'.
        // That is complicated to shim. Let's rewrite Step 4 & 5 slightly to accept raw corpus indices if possible, 
        // or just perform the mapping here.

        // Simpler approach: Calculate originalStart/End directly here.
        const originalStartIndex = matchIndex;
        const originalEndIndex = corpusEndIndex;

        return this.mapRangeToItems(originalStartIndex, originalEndIndex, charMap, pageItems);
      }
    }

    if (matchIndex === -1) {
      console.warn(`Could not find phrase on page (even with fuzzy search): "${textToFind}"`);
      // console.log('Page text snippet:', normalizedCorpus.substring(0, 200) + '...');
      return null;
    }

    // 4. Map the start and end of the match in the normalized corpus back to indices in the original corpus.
    // (This path is for Attempt 1)
    const matchEndIndex = matchIndex + matchedLength - 1;
    const originalStartIndex = indicesToUse[matchIndex];
    const originalEndIndex = indicesToUse[matchEndIndex];

    return this.mapRangeToItems(originalStartIndex, originalEndIndex, charMap, pageItems);
  }

  private mapRangeToItems(
    originalStartIndex: number,
    originalEndIndex: number,
    charMap: { itemIndex: number; charIndexInItem: number }[],
    pageItems: PageTextItem[]
  ): UnderlineSegment[] {
    // 5. Use the character map to find all the original items and character ranges that are part of the match.
    const perItem = new Map<number, { start: number; end: number }>();
    for (let i = originalStartIndex; i <= originalEndIndex; i++) {
      if (i < 0 || i >= charMap.length) continue;
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

      const startChar = range.start;
      const endChar = range.end;

      return {
        item,
        itemIndex,
        startFrac: Math.max(0, Math.min(1, startFrac)),
        endFrac: Math.max(0, Math.min(1, endFrac)),
        startChar,
        endChar
      };
    }).sort((a, b) => a.itemIndex - b.itemIndex);

    return ranges;
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
        // treat hyphen as a joining char (no space) logic
        // only if it's strictly within a word (e.g. state-of-the-art) but in PDF extraction
        // a hyphen at end of line often means word wrap.
        // Simplified heuristic: Treat hyphen as nothing (join) if followed by letter, 
        // OR treat as space if surrounded by spaces.

        // For now, let's treat it as a space unless it looks like a line-break hyphen
        // If we assume the PDF extraction already outputted the hyphen character, it's safer to 
        // generally treat it as a valid character OR a space, but sticking to "replace with space"
        // is often safer for "sentence search" unless it's a specific compound word.

        // The previous logic was: if next is letter, join (soft hyphen).
        // Let's refine: If current is hyphen, check next char.
        if (i + 1 < s.length && /[a-z]/i.test(s[i + 1])) {
          // hyphen then letter -> likely soft hyphen or compound word.
          // We'll strip it to match "stateoftheart" style OR "disconnected" -> "disconnected"
          // This helps if the search phrase is "disconnected" but PDF has "dis-connected"
          lastCharWasSpace = false;
        } else {
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