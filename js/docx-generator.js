/* ============================================
   DOCX Generator
   Manipulates the Van Asten template.docx
   using JSZip to create finished book documents.
   ============================================ */

const DocxGenerator = {
  templateZip: null,

  /**
   * Load the template.docx file
   */
  async loadTemplate() {
    const response = await fetch('template.docx');
    if (!response.ok) {
      throw new Error('Template-Datei konnte nicht geladen werden. Stelle sicher, dass template.docx im gleichen Ordner liegt.');
    }
    const blob = await response.blob();
    this.templateZip = await JSZip.loadAsync(blob);
    return true;
  },

  /**
   * Escape special XML characters
   */
  escapeXml(text) {
    if (!text) return '';
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  },

  /**
   * Convert smart quotes for German text
   */
  smartQuotes(text) {
    if (!text) return '';
    return text
      .replace(/'/g, '\u2019')      // Apostroph
      .replace(/"([^"]*?)"/g, '\u201E$1\u201C')  // Deutsche Anfuehrungszeichen
      .replace(/--/g, '\u2013');     // Gedankenstrich
  },

  /**
   * Generate XML for a normal paragraph
   */
  xmlParagraph(text) {
    const escaped = this.escapeXml(this.smartQuotes(text));
    // Split by bold markers **text** and italic markers *text*
    const parts = this.parseInlineFormatting(escaped);
    const runs = parts.map(p => {
      if (p.bold && p.italic) {
        return `<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">${p.text}</w:t></w:r>`;
      } else if (p.bold) {
        return `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${p.text}</w:t></w:r>`;
      } else if (p.italic) {
        return `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${p.text}</w:t></w:r>`;
      }
      return `<w:r><w:t xml:space="preserve">${p.text}</w:t></w:r>`;
    }).join('');
    return `<w:p>${runs}</w:p>`;
  },

  /**
   * Parse inline formatting markers (bold/italic)
   */
  parseInlineFormatting(text) {
    const parts = [];
    // Simple regex-based approach for **bold** and *italic*
    const regex = /(\*\*(.+?)\*\*|\*(.+?)\*|[^*]+)/g;
    let match;
    while ((match = regex.exec(text)) !== null) {
      if (match[2]) {
        // Bold
        parts.push({ text: match[2], bold: true, italic: false });
      } else if (match[3]) {
        // Italic
        parts.push({ text: match[3], bold: false, italic: true });
      } else {
        parts.push({ text: match[0], bold: false, italic: false });
      }
    }
    return parts.length > 0 ? parts : [{ text, bold: false, italic: false }];
  },

  /**
   * Generate XML for styled paragraph
   */
  xmlStyled(styleId, text, underline) {
    const escaped = this.escapeXml(this.smartQuotes(text));
    const rPr = underline ? '<w:rPr><w:u w:val="single"/></w:rPr>' : '';
    return `<w:p><w:pPr><w:pStyle w:val="${styleId}"/></w:pPr><w:r>${rPr}<w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
  },

  /**
   * Generate XML for a list
   */
  xmlList(items) {
    return items.map(item => {
      const escaped = this.escapeXml(this.smartQuotes(item));
      return `<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
    }).join('');
  },

  /**
   * Generate XML for a page break
   */
  xmlPageBreak() {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  },

  /**
   * Generate the closing page XML (with logo)
   */
  xmlClosingPage() {
    return this.xmlPageBreak() +
      `<w:p><w:pPr><w:spacing w:before="4000" w:after="0"/><w:jc w:val="center"/></w:pPr></w:p>` +
      `<w:p><w:pPr><w:spacing w:after="600"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:noProof/></w:rPr>` +
      `<w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="3200000" cy="1194000"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="201" name="Closing Logo" descr="Van Asten Coaching Logo"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="201" name="logo_truly_transparent.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rId13"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3200000" cy="1194000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
      `<w:p><w:pPr><w:spacing w:before="800"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:color w:val="D4B565"/><w:sz w:val="24"/></w:rPr>` +
      `<w:t>Van Asten Coaching \u2013 Bewusstseinswerkstatt</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:spacing w:before="400"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:color w:val="6DB89A"/><w:sz w:val="20"/></w:rPr>` +
      `<w:t>www.vanasten-coaching.de</w:t></w:r></w:p>`;
  },

  /**
   * Generate the cover page XML
   */
  xmlCoverPage(title, subtitle) {
    const escapedTitle = this.escapeXml(this.smartQuotes(title));
    const escapedSubtitle = this.escapeXml(this.smartQuotes(subtitle));

    return `<w:p><w:pPr><w:spacing w:before="2000" w:after="0"/><w:jc w:val="center"/></w:pPr></w:p>` +
      `<w:p><w:pPr><w:spacing w:after="600"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:noProof/></w:rPr>` +
      `<w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="3200000" cy="1194000"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="200" name="Cover Logo" descr="Van Asten Coaching Logo"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="200" name="logo_truly_transparent.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rId13"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3200000" cy="1194000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
      `<w:p><w:pPr><w:spacing w:before="1600"/></w:pPr></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Titel"/><w:jc w:val="center"/></w:pPr><w:r><w:t>${escapedTitle}</w:t></w:r></w:p>` +
      (escapedSubtitle ? `<w:p><w:pPr><w:pStyle w:val="Untertitel"/><w:jc w:val="center"/></w:pPr><w:r><w:t>${escapedSubtitle}</w:t></w:r></w:p>` : '') +
      `<w:p><w:pPr><w:spacing w:before="1200"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:color w:val="2F3230"/></w:rPr><w:t>Van Asten Coaching \u2013 Bewusstseinswerkstatt</w:t></w:r></w:p>` +
      this.xmlPageBreak();
  },

  /**
   * Convert a structured chapter object to OOXML
   */
  /**
   * Generate a gold decorative line for visual separation
   */
  xmlGoldLine() {
    return '<w:p><w:pPr><w:spacing w:before="200" w:after="200"/><w:pBdr>' +
      '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="D4B565"/>' +
      '</w:pBdr></w:pPr></w:p>';
  },

  /**
   * Generate spacing paragraph
   */
  xmlSpacing(before, after) {
    return `<w:p><w:pPr><w:spacing w:before="${before || 0}" w:after="${after || 0}"/></w:pPr></w:p>`;
  },

  chapterToXml(chapter, chapterNumber) {
    let xml = '';

    // Chapter title as Heading 1
    xml += this.xmlStyled('berschrift1', `Kapitel ${chapterNumber}: ${chapter.chapter_title}`);

    // Gold line under chapter title
    xml += this.xmlGoldLine();

    // Process each content block with smart spacing
    let prevType = '';
    for (const block of chapter.content) {
      // Add extra spacing before headings (visual breathing room)
      if ((block.type === 'heading' || block.type === 'subheading_red' || block.type === 'subheading_green') && prevType === 'paragraph') {
        xml += this.xmlSpacing(200, 0);
      }

      switch (block.type) {
        case 'paragraph':
          xml += this.xmlParagraph(block.text);
          break;

        case 'heading':
          xml += this.xmlSpacing(300, 0);
          xml += this.xmlStyled('berschrift1', block.text);
          xml += this.xmlGoldLine();
          break;

        case 'subheading_red':
          xml += this.xmlStyled('berschrift4', block.text, true);
          break;

        case 'subheading_green':
          xml += this.xmlStyled('berschrift5', block.text, true);
          break;

        case 'quote':
          xml += this.xmlSpacing(100, 0);
          xml += this.xmlStyled('Zitat', block.text);
          xml += this.xmlSpacing(0, 100);
          break;

        case 'tip':
          xml += this.xmlSpacing(100, 0);
          xml += this.xmlStyled('Hinweis', block.text);
          xml += this.xmlSpacing(0, 100);
          break;

        case 'goldbox':
          xml += this.xmlSpacing(100, 0);
          xml += this.xmlStyled('GoldBox', block.text);
          xml += this.xmlSpacing(0, 100);
          break;

        case 'exercise':
          xml += this.xmlSpacing(200, 0);
          xml += this.xmlStyled('berschrift5', 'Reflexion und \u00dcbung', true);
          xml += this.xmlStyled('Hinweis', block.text);
          xml += this.xmlSpacing(0, 200);
          break;

        case 'list':
          if (block.items && Array.isArray(block.items)) {
            xml += this.xmlSpacing(80, 0);
            xml += this.xmlList(block.items);
            xml += this.xmlSpacing(0, 80);
          }
          break;

        default:
          // Unknown type, treat as paragraph
          if (block.text) {
            xml += this.xmlParagraph(block.text);
          }
          break;
      }
      prevType = block.type;
    }

    return xml;
  },

  /**
   * Generate the complete book document
   * @param {string} title - Book title
   * @param {string} subtitle - Book subtitle
   * @param {Array} chapters - Array of structured chapter objects
   * @param {object|null} introData - Introduction content (optional)
   * @param {object|null} outroData - Outro/closing content (optional)
   * @returns {Promise<Blob>} The generated .docx as a Blob
   */
  async generateBook(title, subtitle, chapters, introData, outroData) {
    if (!this.templateZip) {
      await this.loadTemplate();
    }

    // Clone the template ZIP
    const newZip = new JSZip();
    const templateFiles = this.templateZip;

    // Copy all files from template
    for (const [path, file] of Object.entries(templateFiles.files)) {
      if (!file.dir) {
        const content = await file.async('arraybuffer');
        newZip.file(path, content);
      }
    }

    // Read the original document.xml to extract sectPr
    const originalDocXml = await templateFiles.file('word/document.xml').async('string');

    // Extract the sectPr (must be preserved exactly)
    const sectPrMatch = originalDocXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    if (!sectPrMatch) {
      throw new Error('Template-Fehler: sectPr nicht gefunden');
    }
    const sectPr = sectPrMatch[0];

    // Extract the XML header and w:body opening tag
    const xmlHeader = originalDocXml.match(/^[\s\S]*?<w:body>/)?.[0];
    if (!xmlHeader) {
      throw new Error('Template-Fehler: w:body nicht gefunden');
    }

    // Build the new document content
    let bodyContent = '';

    // Cover page
    bodyContent += this.xmlCoverPage(title, subtitle);

    // Table of Contents header
    bodyContent += this.xmlStyled('berschrift1', 'Inhaltsverzeichnis');
    bodyContent += '<w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>';

    // TOC: Einleitung
    if (introData) {
      bodyContent += this.xmlParagraph('Einleitung');
    }

    // TOC: Chapters
    for (let i = 0; i < chapters.length; i++) {
      const chTitle = chapters[i].chapter_title || `Kapitel ${i + 1}`;
      bodyContent += this.xmlParagraph(`Kapitel ${i + 1}: ${chTitle}`);
    }

    // TOC: Abschlusswort
    if (outroData) {
      bodyContent += this.xmlParagraph('Abschlusswort');
    }

    bodyContent += this.xmlPageBreak();

    // Introduction
    if (introData && introData.content) {
      bodyContent += this.xmlStyled('berschrift1', 'Einleitung');
      bodyContent += '<w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>';
      for (const block of introData.content) {
        switch (block.type) {
          case 'paragraph': bodyContent += this.xmlParagraph(block.text); break;
          case 'quote': bodyContent += this.xmlStyled('Zitat', block.text); break;
          case 'tip': bodyContent += this.xmlStyled('Hinweis', block.text); break;
          case 'goldbox': bodyContent += this.xmlStyled('GoldBox', block.text); break;
          default: if (block.text) bodyContent += this.xmlParagraph(block.text); break;
        }
      }
      bodyContent += this.xmlPageBreak();
    }

    // Generate each chapter
    for (let i = 0; i < chapters.length; i++) {
      bodyContent += this.chapterToXml(chapters[i], i + 1);

      // Page break between chapters (but not after last)
      if (i < chapters.length - 1) {
        bodyContent += this.xmlPageBreak();
      }
    }

    // Outro / Abschlusswort
    if (outroData && outroData.content) {
      bodyContent += this.xmlPageBreak();
      bodyContent += this.xmlStyled('berschrift1', 'Abschlusswort');
      bodyContent += '<w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>';
      for (const block of outroData.content) {
        switch (block.type) {
          case 'paragraph': bodyContent += this.xmlParagraph(block.text); break;
          case 'quote': bodyContent += this.xmlStyled('Zitat', block.text); break;
          case 'tip': bodyContent += this.xmlStyled('Hinweis', block.text); break;
          case 'goldbox': bodyContent += this.xmlStyled('GoldBox', block.text); break;
          default: if (block.text) bodyContent += this.xmlParagraph(block.text); break;
        }
      }
    }

    // Closing page with logo
    bodyContent += this.xmlClosingPage();

    // Assemble complete document.xml
    const newDocXml = xmlHeader + bodyContent + sectPr + '</w:body></w:document>';

    // Replace document.xml in the zip
    newZip.file('word/document.xml', newDocXml);

    // Generate the new .docx file
    const blob = await newZip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    return blob;
  }
};
