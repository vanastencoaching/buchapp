/* ============================================
   Claude API Integration
   ============================================ */

const ClaudeAPI = {
  /**
   * Returns the currently selected model from the dropdown
   */
  getModel() {
    const select = document.getElementById('model-select');
    return select ? select.value : 'claude-sonnet-4-20250514';
  },

  /**
   * Returns style instructions based on selected writing style
   */
  getStyleInstructions() {
    const select = document.getElementById('style-select');
    const style = select ? select.value : 'warm';

    switch (style) {
      case 'warm':
        return `SCHREIBSTIL: Herzlich-ermutigend
- Schreibe in einem warmen, professionellen, motivierenden Ton
- Verwende die Du-Ansprache (kein Sie)
- F\u00fcge wo passend Coaching-Impulse, Reflexionsfragen oder \u00dcbungen ein
- Verwende Zitate f\u00fcr besonders pr\u00e4gnante Aussagen
- Positiv formuliert, keine Verneinungen
- Wie eine gute Freundin die coacht`;

      case 'sachlich':
        return `SCHREIBSTIL: Sachlich-strukturiert
- Schreibe in einem klaren, pr\u00e4zisen, faktenorientierten Ton
- Verwende die Du-Ansprache
- Strukturiere mit klaren Zwischen\u00fcberschriften und Aufz\u00e4hlungen
- Fakten und Methoden stehen im Vordergrund
- Kompakt und auf den Punkt
- Wie ein gutes Fachbuch`;

      case 'story':
        return `SCHREIBSTIL: Story-basiert
- Schreibe erz\u00e4hlerisch, mit Geschichten und Metaphern
- Verwende die Du-Ansprache
- Beginne Abschnitte mit kleinen Anekdoten oder Beispielen
- Nutze bildhafte Sprache und Vergleiche
- Baue einen Spannungsbogen auf
- Wie ein Erz\u00e4hler am Lagerfeuer der Wissen teilt`;

      default:
        return '';
    }
  },

  /**
   * Sends a transcript to Claude API and receives structured book chapter content.
   * @param {string} apiKey - Anthropic API key
   * @param {string} transcript - The raw transcript text
   * @param {string} chapterTitle - The chapter title
   * @param {number} chapterNumber - The chapter number
   * @param {string} bookTitle - The overall book title
   * @param {string} customInstructions - Additional user instructions
   * @param {object} chapterContext - Context about neighboring chapters for transitions
   * @returns {Promise<object>} Structured chapter content
   */
  async generateChapter(apiKey, transcript, chapterTitle, chapterNumber, bookTitle, customInstructions, chapterContext) {
    const styleInstructions = ClaudeAPI.getStyleInstructions();
    const systemPrompt = `Du bist ein professioneller Buchautor und Ghostwriter für Coaching- und Persönlichkeitsentwicklungs-Bücher.
Du arbeitest für Van Asten Coaching - Bewusstseinswerkstatt.

Deine Aufgabe: Verwandle das gegebene Transkript in ein professionelles Buchkapitel.

${styleInstructions}

WICHTIGE REGELN:
- Strukturiere den Text klar mit Abschnitten und visueller Vielfalt
- Entferne F\u00fcllw\u00f6rter, Wiederholungen und typische Sprachmuster aus dem Transkript
- Behalte die Kernaussagen und Inhalte bei
- Das Kapitel soll sich wie ein professionelles, visuell ansprechendes Sachbuch lesen
- Schreibe IMMER mit korrekten deutschen Umlauten (\u00e4, \u00f6, \u00fc, \u00df) - NIEMALS ae, oe, ue verwenden

DESIGN-REGELN F\u00dcR VISUELLE VIELFALT (SEHR WICHTIG):
- Nutze MINDESTENS 5 verschiedene Content-Typen pro Kapitel (nicht nur paragraphs!)
- Setze nach 2-3 Abs\u00e4tzen einen visuellen Akzent (quote, tip, goldbox, exercise oder list)
- Verwende subheading_red f\u00fcr wichtige Kernpunkte und subheading_green f\u00fcr Wachstumsthemen
- Jedes Kapitel sollte mindestens enthalten: 1 quote, 1 tip ODER goldbox, 1 exercise, 1 list
- Wechsle zwischen Flie\u00dftext und visuellen Elementen ab - das Buch soll NICHT wie eine Textwand aussehen

AUSGABEFORMAT: Antworte NUR mit einem validen JSON-Objekt (kein Markdown, kein Kommentar). Struktur:
{
  "chapter_title": "Kapitel-\u00dcberschrift",
  "content": [
    {"type": "paragraph", "text": "Einleitender Absatz..."},
    {"type": "subheading_red", "text": "Wichtiger Kernpunkt"},
    {"type": "paragraph", "text": "Erkl\u00e4render Text..."},
    {"type": "quote", "text": "Ein pr\u00e4gnantes Zitat oder Kernaussage"},
    {"type": "paragraph", "text": "Weiterer Text..."},
    {"type": "subheading_green", "text": "Dein Wachstumsschritt"},
    {"type": "tip", "text": "Ein Coaching-Impuls oder praktischer Tipp"},
    {"type": "list", "items": ["Punkt 1", "Punkt 2", "Punkt 3"]},
    {"type": "goldbox", "text": "Eine Schl\u00fcsselerkenntnis zum Mitnehmen"},
    {"type": "exercise", "text": "Eine Reflexionsfrage oder \u00dcbung"}
  ]
}

Erlaubte Content-Typen:
- "paragraph" - Normaler Fließtext
- "heading" - Zwischenüberschrift (Überschrift 2 Ebene)
- "subheading_red" - Rote Akzent-Überschrift für wichtige Punkte
- "subheading_green" - Grüne Akzent-Überschrift für Wachstumsthemen
- "quote" - Prägnantes Zitat
- "tip" - Coaching-Impuls mit grünem Balken
- "goldbox" - Besonders wichtige Erkenntnis in Gold-Box
- "exercise" - Reflexionsfrage oder Übung
- "list" - Aufzählungsliste (mit "items" Array)

${customInstructions ? '\nZUS\u00c4TZLICHE ANWEISUNGEN DES AUTORS:\n' + customInstructions : ''}`;

    // Build chapter context for transitions
    let contextInfo = '';
    if (chapterContext) {
      if (chapterContext.prevTitle) {
        contextInfo += '\nVORHERIGES KAPITEL: "' + chapterContext.prevTitle + '"';
        if (chapterContext.prevSummary) {
          contextInfo += ' - ' + chapterContext.prevSummary;
        }
        contextInfo += '\n-> Beginne dieses Kapitel mit einem sanften \u00dcbergang, der an das vorherige Kapitel ankn\u00fcpft.';
      }
      if (chapterContext.nextTitle) {
        contextInfo += '\nN\u00c4CHSTES KAPITEL: "' + chapterContext.nextTitle + '"';
        if (chapterContext.nextSummary) {
          contextInfo += ' - ' + chapterContext.nextSummary;
        }
        contextInfo += '\n-> Schlie\u00dfe dieses Kapitel mit einem Ausblick ab, der neugierig auf das n\u00e4chste Kapitel macht.';
      }
      if (chapterContext.totalChapters) {
        contextInfo += '\nDas Buch hat insgesamt ' + chapterContext.totalChapters + ' Kapitel. Dies ist Kapitel ' + chapterNumber + '.';
      }
    }

    const userMessage = `Buch: "${bookTitle}"
Kapitel ${chapterNumber}: "${chapterTitle}"
${contextInfo ? '\nKONTEXT F\u00dcR KAPITEL\u00dcBERG\u00c4NGE:' + contextInfo + '\n' : ''}
Hier ist das Transkript, das du in ein Buchkapitel verwandeln sollst:

---
${transcript}
---

Bitte verwandle dieses Transkript in ein professionelles Buchkapitel. Antworte NUR mit dem JSON-Objekt.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify({
        model: ClaudeAPI.getModel(),
        max_tokens: 8000,
        system: systemPrompt,
        messages: [
          { role: 'user', content: userMessage }
        ]
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      if (response.status === 401) {
        throw new Error('Ungueltiger API-Schluessel. Bitte pruefe deinen Schluessel und versuche es erneut.');
      }
      if (response.status === 429) {
        throw new Error('Zu viele Anfragen. Bitte warte einen Moment und versuche es erneut.');
      }
      throw new Error(`API-Fehler (${response.status}): ${errorData.error?.message || 'Unbekannter Fehler'}`);
    }

    const data = await response.json();
    const textContent = data.content.find(c => c.type === 'text')?.text;

    if (!textContent) {
      throw new Error('Keine Antwort von der KI erhalten.');
    }

    // Parse JSON from response - handle potential markdown code blocks
    let jsonStr = textContent.trim();
    if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/^```(?:json)?\s*/, '').replace(/\s*```$/, '');
    }

    try {
      const parsed = JSON.parse(jsonStr);
      if (!parsed.content || !Array.isArray(parsed.content)) {
        throw new Error('Unerwartetes Antwortformat');
      }
      return parsed;
    } catch (e) {
      // If JSON parsing fails, try to extract JSON from the response
      const jsonMatch = textContent.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const parsed = JSON.parse(jsonMatch[0]);
          if (parsed.content && Array.isArray(parsed.content)) {
            return parsed;
          }
        } catch (e2) {
          // Fall through to error
        }
      }
      throw new Error('Die KI-Antwort konnte nicht verarbeitet werden. Bitte versuche es erneut.');
    }
  },

  /**
   * Generate a book introduction
   */
  async generateIntroduction(apiKey, bookTitle, subtitle, chapterTitles, customInstructions) {
    const systemPrompt = `Du bist ein professioneller Buchautor für Coaching- und Persönlichkeitsentwicklungs-Bücher.
Du arbeitest für Van Asten Coaching - Bewusstseinswerkstatt.
Schreibe IMMER mit korrekten deutschen Umlauten (ä, ö, ü, ß) - NIEMALS ae, oe, ue verwenden.

Erstelle eine warmherzige, einladende Einleitung für ein Buch. Die Einleitung soll:
- Den Leser persönlich mit Du ansprechen
- Erklären, warum es dieses Buch gibt
- Einen kurzen Überblick über die Kapitel geben
- Motivieren, das Buch durchzuarbeiten
- Positiv formuliert sein, keine Verneinungen
- Ca. 300-500 Wörter lang sein

${customInstructions ? '\nZUSÄTZLICHE ANWEISUNGEN:\n' + customInstructions : ''}

Antworte NUR mit einem validen JSON-Objekt:
{
  "content": [
    {"type": "paragraph", "text": "..."},
    {"type": "quote", "text": "..."},
    {"type": "paragraph", "text": "..."}
  ]
}`;

    const userMessage = `Buch: "${bookTitle}"${subtitle ? '\nUntertitel: "' + subtitle + '"' : ''}
Kapitel: ${chapterTitles.join(', ')}

Schreibe eine Einleitung für dieses Buch.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify({
        model: ClaudeAPI.getModel(),
        max_tokens: 3000,
        system: systemPrompt,
        messages: [{ role: 'user', content: userMessage }]
      })
    });

    if (!response.ok) throw new Error('Einleitung konnte nicht generiert werden.');
    const data = await response.json();
    const text = data.content.find(c => c.type === 'text')?.text || '';
    let jsonStr = text.trim().replace(/^```(?:json)?\s*/, '').replace(/\s*```$/, '');
    const match = jsonStr.match(/\{[\s\S]*\}/);
    return match ? JSON.parse(match[0]) : { content: [{ type: 'paragraph', text: jsonStr }] };
  },

  /**
   * Generate a book closing/outro
   */
  async generateOutro(apiKey, bookTitle, subtitle, chapterTitles, customInstructions) {
    const systemPrompt = `Du bist ein professioneller Buchautor für Coaching- und Persönlichkeitsentwicklungs-Bücher.
Du arbeitest für Van Asten Coaching - Bewusstseinswerkstatt.
Schreibe IMMER mit korrekten deutschen Umlauten (ä, ö, ü, ß) - NIEMALS ae, oe, ue verwenden.

Erstelle ein warmherziges Abschlusswort für ein Buch. Das Abschlusswort soll:
- Den Leser persönlich mit Du ansprechen
- Die wichtigsten Erkenntnisse zusammenfassen
- Zum Handeln und Umsetzen motivieren
- Sich beim Leser bedanken
- Positiv und ermutigend enden
- Ca. 200-400 Wörter lang sein
- Mit einem motivierenden Schlusssatz enden

${customInstructions ? '\nZUSÄTZLICHE ANWEISUNGEN:\n' + customInstructions : ''}

Antworte NUR mit einem validen JSON-Objekt:
{
  "content": [
    {"type": "paragraph", "text": "..."},
    {"type": "goldbox", "text": "..."},
    {"type": "paragraph", "text": "..."}
  ]
}`;

    const userMessage = `Buch: "${bookTitle}"${subtitle ? '\nUntertitel: "' + subtitle + '"' : ''}
Kapitel: ${chapterTitles.join(', ')}

Schreibe ein Abschlusswort für dieses Buch.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify({
        model: ClaudeAPI.getModel(),
        max_tokens: 3000,
        system: systemPrompt,
        messages: [{ role: 'user', content: userMessage }]
      })
    });

    if (!response.ok) throw new Error('Abschlusswort konnte nicht generiert werden.');
    const data = await response.json();
    const text = data.content.find(c => c.type === 'text')?.text || '';
    let jsonStr = text.trim().replace(/^```(?:json)?\s*/, '').replace(/\s*```$/, '');
    const match = jsonStr.match(/\{[\s\S]*\}/);
    return match ? JSON.parse(match[0]) : { content: [{ type: 'paragraph', text: jsonStr }] };
  }
};
