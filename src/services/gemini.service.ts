import { Injectable } from '@angular/core';
import { GoogleGenAI } from '@google/genai';

@Injectable({ providedIn: 'root' })
export class GeminiService {
  private ai: GoogleGenAI | null = null;

  constructor() {
    // API key is expected to be available in the execution environment
    // as per the app's operational requirements.
    const apiKey = (process.env as any).API_KEY;
    if (apiKey) {
      this.ai = new GoogleGenAI({ apiKey });
    } else {
      console.warn('Gemini API key not found. AI-powered sentence matching will be disabled.');
    }
  }

  async findClosestSentence(pageText: string, contextPhrase: string): Promise<string | null> {
    if (!this.ai) {
      return null;
    }

    const model = 'gemini-2.5-flash';
    const prompt = `
      You are a text processing expert. Your task is to find a SEARCH PHRASE within a larger TEXT BLOCK.
      The text may have small differences due to formatting (like ligatures ﬁ vs fi, or smart quotes ’ vs ').
      Your goal is to find the exact text in the TEXT BLOCK that corresponds to the SEARCH PHRASE.

      1. Read the SEARCH PHRASE.
      2. Read the TEXT BLOCK.
      3. Find the sentence in the TEXT BLOCK that is the best semantic and character-level match for the SEARCH PHRASE.
      4. Respond with ONLY the verbatim text of the matching sentence from the TEXT BLOCK. Do not alter it. Do not add explanations or quotes.

      SEARCH PHRASE:
      ---
      ${contextPhrase}
      ---

      TEXT BLOCK:
      ---
      ${pageText}
      ---
    `;

    try {
      const response = await this.ai.models.generateContent({
        model: model,
        contents: prompt,
      });

      const matchedSentence = response.text.trim();

      // Validate that Gemini's response is a substring of the original page text to prevent hallucinations.
      if (matchedSentence && pageText.includes(matchedSentence)) {
        return matchedSentence;
      }
      
      console.warn('Gemini returned a sentence not present in the original text:', matchedSentence);
      return null;
    } catch (error) {
      console.error('Error calling Gemini API:', error);
      return null;
    }
  }
}