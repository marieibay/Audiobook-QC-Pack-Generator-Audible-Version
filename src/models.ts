export interface Correction {
  Id: string;
  Page: number;
  ContextPhrase: string;
  Notes: string;
  WordsToCorrect?: string[];
  Track?: string;
  Timestamp?: string;
}

export type Status = {
  text: string;
  type: 'info' | 'success' | 'error' | 'warning';
};

export interface PageTextItem {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PageText {
  pageNum: number;
  content: string; // The raw text for easy searching
  items: PageTextItem[]; // The structured text for accurate positioning
}

export interface BoundingBox {
  x: number;
  y: number;
  width: number;
  height: number;
}