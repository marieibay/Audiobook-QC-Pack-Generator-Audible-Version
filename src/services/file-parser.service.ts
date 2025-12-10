import { Injectable } from '@angular/core';
import { Correction } from '../models';

declare var Papa: any;
declare var XLSX: any;

@Injectable({ providedIn: 'root' })
export class FileParserService {

  async parseQcFile(file: File, isAudible: boolean): Promise<Correction[]> {
    if (file.name.endsWith('.csv')) {
      return new Promise((resolve, reject) => {
        Papa.parse(file, {
          skipEmptyLines: true,
          complete: (results: any) => {
            if (results.errors && results.errors.length > 0) {
              console.error('CSV parsing errors:', results.errors);
              return reject(new Error(`CSV Parsing Error: ${results.errors[0].message}`));
            }
            try {
              const corrections = this.parseRows(results.data, isAudible);
              resolve(corrections);
            } catch (error) {
              reject(error);
            }
          },
          error: (err: any) => {
            reject(new Error('Failed to read or parse CSV file.'));
          }
        });
      });
    }

    if (file.name.endsWith('.xlsx')) {
      const fileContent = await file.arrayBuffer();
      const workbook = XLSX.read(fileContent, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      return this.parseRows(rows, isAudible);
    }

    throw new Error("Unsupported file type. Please upload a .csv or .xlsx file.");
  }

  private hasHeaders(rows: any[][], requiredCols: string[]): boolean {
    for (const row of rows) {
        if (!Array.isArray(row)) continue;
        const upperRow = row.map(h => h ? h.toString().trim().toUpperCase() : '');
        if (requiredCols.every(col => upperRow.includes(col))) {
            return true;
        }
    }
    return false;
  }

  private parseRows(rows: any[][], isAudible: boolean): Correction[] {
    // Attempt to auto-detect format
    const hasPostQcHeaders = this.hasHeaders(rows, ['CD-TRK', 'TIME', 'TEXT', 'EDITOR COMMENTS']);
    const hasStandardHeaders = this.hasHeaders(rows, ['ID', 'PAGE', 'CONTEXT', 'NOTES']);
    
    if (hasPostQcHeaders && !hasStandardHeaders) {
      // Looks like a Post QC file, parse it as such
      return this.parsePostQcRows(rows, isAudible);
    }
    
    // Default to standard parsing (which will throw a specific error if headers are not found)
    return this.parseStandardQcRows(rows, isAudible);
  }

  private parsePostQcRows(rows: any[][], isAudible: boolean): Correction[] {
    const { header, data } = this.findPostQcHeaderAndData(rows);
    const headerUpper = header.map(h => h ? h.toString().trim().toUpperCase() : '');

    const trackIndex = headerUpper.indexOf('CD-TRK');
    const timeIndex = headerUpper.indexOf('TIME');
    const pageIndex = headerUpper.findIndex(h => h.startsWith('PAGE'));
    const textIndex = headerUpper.indexOf('TEXT');
    const problemIndex = headerUpper.indexOf('PROBLEM DESCRIPTION');
    const editorCommentsIndex = headerUpper.indexOf('EDITOR COMMENTS');
    
    if ([trackIndex, timeIndex, pageIndex, textIndex, problemIndex, editorCommentsIndex].includes(-1)) {
        throw new Error('Could not find all required headers for Post QC format.');
    }

    const corrections: Correction[] = [];
    let pickupId = 1;

    for (const row of data) {
        if (!row || row.length === 0) continue;

        const editorComment = row[editorCommentsIndex] ? row[editorCommentsIndex].toString().trim().toLowerCase() : '';
        if (editorComment !== 'fix not possible without pickup' && editorComment !== 'please fix.' && editorComment !== 'please add to pu.') {
            continue;
        }

        const pageStr = row[pageIndex] ? row[pageIndex].toString() : '0';
        const pageMatch = pageStr.match(/\d+/);
        const page = pageMatch ? parseInt(pageMatch[0], 10) : 0;
        
        const fullText = row[textIndex] ? row[textIndex].toString() : '';
        const problemDescription = row[problemIndex] ? row[problemIndex].toString() : '';

        const rawContext = fullText.replace(/\[|\]/g, '');
        
        const { formattedNote, wordsForOblong, correctionType, searchableContext } = this.processNotes(problemDescription, rawContext, isAudible);

        // For Post QC, if wordsForOblong is empty, try to get it from the brackets in the text
        let finalWordsForOblong = wordsForOblong;
        if (finalWordsForOblong.length === 0) {
            const oblongMatch = fullText.match(/\[(.*?)\]/);
            if (oblongMatch && oblongMatch[1]) {
                finalWordsForOblong = oblongMatch[1].split(' ').filter(Boolean);
            }
        }
        
        corrections.push({
            Id: String(pickupId++),
            Page: page,
            ContextPhrase: searchableContext,
            Notes: formattedNote,
            Track: row[trackIndex] ? row[trackIndex].toString() : '',
            Timestamp: row[timeIndex] ? row[timeIndex].toString() : '',
            correctionType: correctionType,
            wordsForOblong: finalWordsForOblong,
        });
    }
    return corrections;
  }

  private parseStandardQcRows(rows: any[][], isAudible: boolean): Correction[] {
    const { header, data } = this.findStandardHeaderAndData(rows);
    const headerUpper = header.map(h => h ? h.toString().trim().toUpperCase() : '');

    const idIndex = headerUpper.indexOf('ID');
    const pageIndex = headerUpper.indexOf('PAGE');
    const contextIndex = headerUpper.indexOf('CONTEXT');
    const notesIndex = headerUpper.indexOf('NOTES');
    const timeCodeIndex = headerUpper.indexOf('TIME CODE');
    
    const corrections: Correction[] = [];
    let currentTrack = '';

    for (const row of data) {
      if (!row || row.length === 0) continue;

      let trackFilename = '';
      for (const cell of row) {
        const cellStr = cell ? cell.toString() : '';
        if (cellStr.toLowerCase().endsWith('.wav')) {
          trackFilename = cellStr;
          break;
        }
      }

      if (trackFilename) {
        const trackMatch = trackFilename.match(/^(\d+)/);
        if (trackMatch) {
          currentTrack = trackMatch[1];
        }
        continue;
      }
      
      if (!row[idIndex] || isNaN(Number(row[idIndex]))) {
        continue;
      }

      let status = '';
      if (contextIndex !== -1) {
          for (let i = contextIndex + 1; i < row.length; i++) {
              const cellValue = row[i] ? row[i].toString().trim().toLowerCase() : '';
              if (cellValue === 'fix not possible without pickup') {
                  status = cellValue;
                  break;
              }
          }
      }

      if (status === 'fix not possible without pickup') {
        const notes = row[notesIndex] ? row[notesIndex].toString() : '';
        const rawContext = row[contextIndex] ? row[contextIndex].toString() : '';
        const timestamp = timeCodeIndex !== -1 && row[timeCodeIndex] ? row[timeCodeIndex].toString() : '';

        const processedNote = this.processNotes(notes, rawContext, isAudible);

        if (!processedNote.searchableContext) {
            console.warn(`Correction on page ${row[pageIndex]} skipped because it has no context phrase.`);
            continue;
        }

        corrections.push({
          Id: row[idIndex].toString(),
          Page: Number(row[pageIndex]),
          ContextPhrase: processedNote.searchableContext,
          Notes: processedNote.formattedNote,
          Track: currentTrack,
          Timestamp: timestamp,
          correctionType: processedNote.correctionType,
          wordsForOblong: processedNote.wordsForOblong,
        });
      }
    }
    return corrections;
  }

  private findPostQcHeaderAndData(rows: any[][]): { header: string[], data: any[][] } {
    let headerIndex = -1;
    const requiredCols = ['CD-TRK', 'TIME', 'TEXT', 'EDITOR COMMENTS'];

    for (let i = 0; i < rows.length; i++) {
        if (!Array.isArray(rows[i])) continue;
        const row = rows[i].map(h => h ? h.toString().trim().toUpperCase() : '');
        if (requiredCols.every(col => row.includes(col))) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex === -1) {
        throw new Error(`Could not find required header columns for Post QC format (${requiredCols.join(', ')}) in the QC file.`);
    }

    const header = rows[headerIndex].map(h => h ? h.toString() : '');
    const data = rows.slice(headerIndex + 1);
    return { header, data };
  }
  
  private findStandardHeaderAndData(rows: any[][]): { header: string[], data: any[][] } {
    let headerIndex = -1;
    const requiredCols = ['ID', 'PAGE', 'CONTEXT', 'NOTES'];

    for (let i = 0; i < rows.length; i++) {
        if (!Array.isArray(rows[i])) continue;
        const row = rows[i].map(h => h ? h.toString().trim().toUpperCase() : '');
        if (requiredCols.every(col => row.includes(col))) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex === -1) {
        throw new Error(`Could not find required header columns for Standard format (${requiredCols.join(', ')}) in the QC file.`);
    }

    const header = rows[headerIndex].map(h => h ? h.toString() : '');
    const data = rows.slice(headerIndex + 1);
    return { header, data };
  }
  
  private normalizeText(s: string): string {
    if (!s) return '';
    return s.replace(/[﹏_]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  
  private processNotes(
    notes: string,
    rawContext: string,
    isAudible: boolean
  ): {
      formattedNote: string,
      wordsForOblong: string[],
      correctionType: 'misread' | 'missing' | 'inserted',
      searchableContext: string,
  } {
      if (!notes) {
          return { formattedNote: '', wordsForOblong: [], correctionType: 'misread', searchableContext: this.normalizeText(rawContext) };
      }
      notes = notes.trim();
      let originalPrefix = '';

      const prefixMatch = notes.match(/^(MR|MW):\s*/i);
      if (prefixMatch) {
          originalPrefix = prefixMatch[1].toUpperCase();
          notes = notes.substring(prefixMatch[0].length);
      }

      const sbMatch = notes.match(/^(.+?)\s+S\/B\s+(.+)$/i);
      if (sbMatch) {
          const wrong = sbMatch[1].trim().replace(/^["']|["']$/g, '');
          const right = sbMatch[2].trim().replace(/^["']|["']$/g, '');
          let formattedNote = `read as "${wrong}" should be read as "${right}"`;
          if (isAudible) {
              const prefix = originalPrefix === 'MW' ? 'MW: ' : 'MR: ';
              formattedNote = prefix + formattedNote;
          }
          return {
              formattedNote,
              wordsForOblong: right.split(' ').filter(Boolean),
              correctionType: 'misread',
              searchableContext: this.normalizeText(rawContext)
          };
      }

      const readAsMatch = notes.match(/^(.+?)\s+read as\s+(.+)$/i);
      if (readAsMatch) {
          // This format is "<correct_word> read as <incorrect_word>"
          const correctWord = readAsMatch[1].trim().replace(/^["']|["']$/g, '');
          const incorrectWord = readAsMatch[2].trim().replace(/^["']|["']$/g, '');
          
          let formattedNote = `read as "${incorrectWord}" should be read as "${correctWord}"`;
          
          if (isAudible) {
              // This is always a misread, so we enforce MR prefix for Audible projects.
              formattedNote = `MR: ${formattedNote}`;
          }

          return {
              formattedNote,
              wordsForOblong: correctWord.split(' ').filter(Boolean), // Encircle the correct word from the script
              correctionType: 'misread',
              searchableContext: this.normalizeText(rawContext)
          };
      }

      const missingMatch = notes.match(/^(?:Word(?:s)?\s+)?Missing:\s+(.+)$/i);
      if (missingMatch) {
          const missing = missingMatch[1].trim().replace(/^["']|["']$/g, '');
          let formattedNote = `"${missing}" is missing and should be read.`;
          if (isAudible) {
              const wordCount = missing.split(/\s+/).filter(Boolean).length;
              const prefix = wordCount > 3 ? 'ML:' : 'MW:'; // Differentiate between missing word and missing line
              formattedNote = `${prefix} ${formattedNote}`;
          }
          return {
              formattedNote,
              wordsForOblong: missing.split(' ').filter(Boolean),
              correctionType: 'missing',
              searchableContext: this.normalizeText(rawContext)
          };
      }

      const omittedMatch = notes.match(/^omitted\s+(.+)$/i);
      if (omittedMatch) {
          const omitted = omittedMatch[1].trim().replace(/^["']|["']$/g, '');
          let formattedNote = `"${omitted}" was omitted and should be read.`;
          if (isAudible) {
            const wordCount = omitted.split(/\s+/).filter(Boolean).length;
            const prefix = wordCount > 3 ? 'ML:' : 'MW:'; // Differentiate between missing word and missing line
            formattedNote = `${prefix} ${formattedNote}`;
          }
          return {
              formattedNote,
              wordsForOblong: omitted.split(' ').filter(Boolean),
              correctionType: 'missing',
              searchableContext: this.normalizeText(rawContext)
          };
      }
      
      const insertedMatch = notes.match(/^(?:Word(?:s)?\s+)?Inserted:\s+(.+)$/i);
      if (insertedMatch) {
          const inserted = insertedMatch[1].trim().replace(/^["']|["']$/g, '');
          let formattedNote = `"${inserted}" was inserted and should be omitted.`;
          if (isAudible) {
            const wordCount = inserted.split(/\s+/).filter(Boolean).length;
            const prefix = wordCount > 3 ? 'ML:' : 'MW:'; // Differentiate between inserted word and inserted line
            formattedNote = `${prefix} ${formattedNote}`;
          }
          
          let wordsForOblong: string[] = [];
          const placeholderMatch = rawContext.match(/(.*?)[\s_﹏]+(.*?)/s);
          if (placeholderMatch) {
              const beforeText = placeholderMatch[1].trim();
              const afterText = placeholderMatch[2].trim();
              
              const beforeWords = beforeText.split(/\s+/);
              const afterWords = afterText.split(/\s+/);
      
              if (beforeWords.length > 0 && afterWords.length > 0) {
                  wordsForOblong = [beforeWords[beforeWords.length - 1], afterWords[0]];
              }
          }

          return {
              formattedNote,
              wordsForOblong: wordsForOblong,
              correctionType: 'inserted',
              searchableContext: this.normalizeText(rawContext)
          };
      }

      let formattedNote = notes;
      if (isAudible) {
        if (!originalPrefix) {
          formattedNote = `MW: ${notes}`;
        } else {
          formattedNote = `${originalPrefix}: ${notes}`;
        }
      }
      return {
          formattedNote,
          wordsForOblong: [],
          correctionType: 'misread',
          searchableContext: this.normalizeText(rawContext)
      };
  }
}