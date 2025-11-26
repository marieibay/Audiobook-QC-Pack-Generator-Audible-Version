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

  private parseRows(rows: any[][], isAudible: boolean): Correction[] {
    const { header, data } = this.findHeaderAndData(rows);
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

      // Check for a track filename row (e.g., '003_Chapter_One.wav') by searching all cells.
      let trackFilename = '';
      for (const cell of row) {
        const cellStr = cell ? cell.toString() : '';
        if (cellStr.toLowerCase().endsWith('.wav')) {
          trackFilename = cellStr;
          break;
        }
      }

      if (trackFilename) {
        // Found a track filename, extract the track number and move to the next row.
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
        const originalContext = row[contextIndex] ? row[contextIndex].toString().replace(/[﹏_]+/g, ' ').replace(/\s+/g, ' ').trim() : '';
        const timestamp = timeCodeIndex !== -1 && row[timeCodeIndex] ? row[timeCodeIndex].toString() : '';

        const { formattedNote, wordsToCorrect } = this.processNotes(notes, isAudible);

        if (!originalContext) {
            console.warn(`Correction on page ${row[pageIndex]} skipped because it has no context phrase.`);
            continue;
        }

        corrections.push({
          Id: row[idIndex].toString(),
          Page: Number(row[pageIndex]),
          ContextPhrase: originalContext,
          Notes: formattedNote,
          WordsToCorrect: wordsToCorrect,
          Track: currentTrack,
          Timestamp: timestamp,
        });
      }
    }
    return corrections;
  }

  private findHeaderAndData(rows: any[][]): { header: string[], data: any[][] } {
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
        throw new Error(`Could not find required header columns (${requiredCols.join(', ')}) in the QC file.`);
    }

    const header = rows[headerIndex].map(h => h ? h.toString() : '');
    const data = rows.slice(headerIndex + 1);
    return { header, data };
  }
  
  private processNotes(notes: string, isAudible: boolean): { formattedNote: string, wordsToCorrect: string[] } {
      if (!notes) return { formattedNote: '', wordsToCorrect: [] };
      notes = notes.trim();

      const sbMatch = notes.match(/^(?:(MR|MW):\s*)?(.+?)\s+S\/B\s+(.+)$/i);
      if (sbMatch) {
          const originalPrefix = sbMatch[1];
          const wrong = sbMatch[2].trim().replace(/^"|"$/g, '');
          const right = sbMatch[3].trim().replace(/^"|"$/g, '');
          
          // Use the 'right' word(s) for searching, as this is what's present in the script PDF.
          const wordsToCorrect = right.split(' ').filter(w => w.length > 0);
          let formattedNote = `read as "${wrong}" should be read as "${right}"`;

          if (isAudible) {
            const prefix = (originalPrefix && originalPrefix.toUpperCase() === 'MW') ? 'MW: ' : 'MR: ';
            formattedNote = prefix + formattedNote;
          }

          return { formattedNote, wordsToCorrect };
      }

      const missingMatch = notes.match(/^(?:Word(?:s)?\s+)?Missing:\s+(.+)$/i);
      if (missingMatch) {
          const missing = missingMatch[1].trim().replace(/^"|"$/g, '');
          let formattedNote = `"${missing}" is missing and should be read.`;
          if (isAudible) {
              formattedNote = `MW: ${formattedNote}`;
          }
          return {
              formattedNote,
              // The words are missing, so they cannot be found in the PDF.
              // We must rely on the ContextPhrase for positioning.
              wordsToCorrect: []
          };
      }

      const insertedMatch = notes.match(/^(?:Word(?:s)?\s+)?Inserted:\s+(.+)$/i);
      if (insertedMatch) {
          const inserted = insertedMatch[1].trim().replace(/^"|"$/g, '');
          return {
              formattedNote: `"${inserted}" was inserted and should be omitted.`,
              wordsToCorrect: []
          };
      }
      
      return { formattedNote: notes, wordsToCorrect: [] };
  }
}