import { Injectable } from '@angular/core';
import { Correction } from '../models';

declare var Papa: any;
declare var XLSX: any;

@Injectable({ providedIn: 'root' })
export class FileParserService {

  async parseQcFile(file: File, isAudible: boolean, isPostQc: boolean): Promise<Correction[]> {
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
              const corrections = this.parseRows(results.data, isAudible, isPostQc);
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
      return this.parseRows(rows, isAudible, isPostQc);
    }

    throw new Error("Unsupported file type. Please upload a .csv or .xlsx file.");
  }

  private parseRows(rows: any[][], isAudible: boolean, isPostQc: boolean): Correction[] {
    if (isPostQc) {
      return this.parsePostQcRows(rows, isAudible);
    }
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

        const editorComment = row[editorCommentsIndex] ? row[editorCommentsIndex].toString().trim() : '';
        if (editorComment !== 'Fixed Not Possible Without Pickup') {
            continue;
        }

        const pageStr = row[pageIndex] ? row[pageIndex].toString() : '0';
        const pageMatch = pageStr.match(/\d+/);
        const page = pageMatch ? parseInt(pageMatch[0], 10) : 0;
        
        const fullText = row[textIndex] ? row[textIndex].toString().replace(/\u200b/g, '') : '';
        const problemDescription = row[problemIndex] ? row[problemIndex].toString().replace(/\u200b/g, '') : '';

        let wordsForOblong: string[] = [];
        // Context is always the full text with brackets removed.
        const contextPhrase = fullText.replace(/\[\[|\]\]/g, '');
        
        // First, prioritize [[...]] from the Text column for what to encircle.
        const oblongMatch = fullText.match(/\[\[(.*?)\]\]/);
        if (oblongMatch && oblongMatch[1]) {
            wordsForOblong = oblongMatch[1].split(' ').filter(Boolean);
        } else {
            // As a fallback, check for 'noise on "..."' in the Problem Description.
            const noiseMatch = problemDescription.match(/noise on ["'](.*?)["']/i);
            if (noiseMatch && noiseMatch[1]) {
                wordsForOblong = noiseMatch[1].split(' ').filter(Boolean);
            }
        }
        
        corrections.push({
            Id: String(pickupId++),
            Page: page,
            ContextPhrase: contextPhrase,
            Notes: problemDescription,
            Track: row[trackIndex] ? row[trackIndex].toString() : '',
            Timestamp: row[timeIndex] ? row[timeIndex].toString() : '',
            correctionType: 'misread', // Treat all as misread for highlighting
            wordsForOblong: wordsForOblong,
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
    const statusIndex = contextIndex + 1; // Assume status is in the column after context

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

      const status = row[statusIndex] ? row[statusIndex].toString().trim() : '';

      if (status === 'Fix Not Possible Without Pickup') {
        const notes = row[notesIndex] ? row[notesIndex].toString() : '';
        const originalContext = this.normalizeText(row[contextIndex] ? row[contextIndex].toString() : '');
        const timestamp = timeCodeIndex !== -1 && row[timeCodeIndex] ? row[timeCodeIndex].toString() : '';

        const processedNote = this.processNotes(notes, originalContext, isAudible);

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
    return s.replace(/\u200b/g, '').replace(/[﹏_]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  
  private processNotes(
    notes: string, 
    originalContext: string, 
    isAudible: boolean
  ): { 
      formattedNote: string, 
      wordsForOblong: string[], 
      correctionType: 'misread' | 'missing' | 'inserted',
      searchableContext: string,
  } {
      if (!notes) return { formattedNote: '', wordsForOblong: [], correctionType: 'misread', searchableContext: originalContext };
      notes = notes.trim().replace(/\u200b/g, '');

      const sbMatch = notes.match(/^(?:(MR|MW):\s*)?(.+?)\s+S\/B\s+(.+)$/i);
      if (sbMatch) {
          const originalPrefix = sbMatch[1];
          const wrong = sbMatch[2].trim().replace(/^["']|["']$/g, '');
          const right = sbMatch[3].trim().replace(/^["']|["']$/g, '');
          
          let formattedNote = `read as "${wrong}" should be read as "${right}"`;

          if (isAudible) {
            const prefix = (originalPrefix && originalPrefix.toUpperCase() === 'MW') ? 'MW: ' : 'MR: ';
            formattedNote = prefix + formattedNote;
          }

          return { 
            formattedNote, 
            wordsForOblong: right.split(' ').filter(w => w.length > 0),
            correctionType: 'misread',
            searchableContext: originalContext
          };
      }
      
      const soundsLikeMatch = notes.match(/^(.*)\s+sounds like\s+(.*)$/i);
      if (soundsLikeMatch) {
          const right = soundsLikeMatch[1].trim().replace(/^["']|["']$/g, '');
          const wrong = soundsLikeMatch[2].trim().replace(/^["']|["']$/g, '');
          
          let formattedNote = `read as "${wrong}" should be read as "${right}"`;
          
          if (isAudible) {
              formattedNote = 'MR: ' + formattedNote;
          }

          return {
              formattedNote,
              wordsForOblong: right.split(' ').filter(w => w.length > 0),
              correctionType: 'misread',
              searchableContext: originalContext
          };
      }

      const omittedLineMatch = notes.match(/^omitted line:\s*(.+)$/i);
      if (omittedLineMatch) {
          const omitted = omittedLineMatch[1].trim().replace(/^["']|["']$/g, '');
          const omittedWords = omitted.split(' ').filter(w => w.length > 0);
          let formattedNote = `"${omitted}" is missing and should be read.`;
          
          if (isAudible) {
              formattedNote = 'ML: ' + formattedNote;
          }
          
          return {
              formattedNote,
              wordsForOblong: omittedWords,
              correctionType: 'missing',
              searchableContext: originalContext
          };
      }

      const missingMatch = notes.match(/^(?:Word(?:s)?\s+)?Missing:\s+(.+)$/i);
      if (missingMatch) {
          const missing = missingMatch[1].trim().replace(/^["']|["']$/g, '');
          const missingWords = missing.split(' ').filter(w => w.length > 0);
          let formattedNote = `"${missing}" is missing and should be read.`;
          
          if (isAudible) {
              const prefix = missingWords.length >= 2 ? 'ML: ' : 'MW: ';
              formattedNote = prefix + formattedNote;
          }
          
          return {
              formattedNote,
              wordsForOblong: missingWords,
              correctionType: 'missing',
              searchableContext: originalContext
          };
      }

      const omittedMatch = notes.match(/^omitted:\s*(.+)$/i);
      if (omittedMatch) {
          const omitted = omittedMatch[1].trim().replace(/^["']|["']$/g, '');
          const omittedWords = omitted.split(' ').filter(w => w.length > 0);
          let formattedNote = `"${omitted}" is missing and should be read.`;
          
          if (isAudible) {
              const prefix = omittedWords.length >= 2 ? 'ML: ' : 'MW: ';
              formattedNote = prefix + formattedNote;
          }
          
          return {
              formattedNote,
              wordsForOblong: omittedWords,
              correctionType: 'missing',
              searchableContext: originalContext
          };
      }

      const omittedPlainMatch = notes.match(/^omitted\s+(.+)$/i);
      if (omittedPlainMatch) {
          const omitted = omittedPlainMatch[1].trim().replace(/^["']|["']$/g, '');
          const omittedWords = omitted.split(' ').filter(w => w.length > 0);
          let formattedNote = `"${omitted}" is missing and should be read.`;
          
          if (isAudible) {
              const prefix = omittedWords.length >= 2 ? 'ML: ' : 'MW: ';
              formattedNote = prefix + formattedNote;
          }
          
          return {
              formattedNote,
              wordsForOblong: omittedWords,
              correctionType: 'missing',
              searchableContext: originalContext
          };
      }

      const insertedMatch = notes.match(/^(?:Word(?:s)?\s+)?Inserted:\s+(.+)$/i);
      if (insertedMatch) {
          const inserted = insertedMatch[1].trim().replace(/^["']|["']$/g, '');
          const insertedWords = inserted.split(' ').filter(w => w.length > 0);
          let formattedNote = `"${inserted}" was inserted and should be omitted.`;

          if (isAudible) {
              const prefix = insertedWords.length >= 2 ? 'ML: ' : 'MW: ';
              formattedNote = prefix + formattedNote;
          }

          return {
              formattedNote,
              wordsForOblong: insertedWords,
              correctionType: 'inserted',
              searchableContext: originalContext
          };
      }
      
      return { 
        formattedNote: notes, 
        wordsForOblong: [], 
        correctionType: 'misread', 
        searchableContext: originalContext 
      };
  }
}