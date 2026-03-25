/* ============================================
   Buch-Generator - Main Application
   ============================================ */

(function () {
  'use strict';

  // -- State --
  const state = {
    chapters: [],           // { id, fileName, title, text }
    generatedChapters: [],  // AI-generated chapter data (JSON objects)
    introData: null,
    outroData: null,
    generatedBook: null,
    isGenerating: false,
    modalChapterIndex: -1,
    modalRegenMode: false
  };

  let draggedItem = null;

  // -- DOM Elements --
  const $ = (sel) => document.querySelector(sel);
  const apiKeyInput = $('#api-key-input');
  const toggleApiKeyBtn = $('#toggle-api-key');
  const bookTitleInput = $('#book-title');
  const bookSubtitleInput = $('#book-subtitle');
  const bookInstructionsInput = $('#book-instructions');
  const uploadZone = $('#upload-zone');
  const fileInput = $('#file-input');
  const chapterList = $('#chapter-list');
  const chapterCount = $('#chapter-count');
  const estimatedTime = $('#estimated-time');
  const generateBtn = $('#generate-btn');
  const progressArea = $('#progress-area');
  const progressBar = $('#progress-bar');
  const progressText = $('#progress-text');
  const downloadBtn = $('#download-btn');
  const resetBtn = $('#reset-btn');
  const stepPreview = $('#step-preview');
  const previewCards = $('#preview-cards');
  const chapterModal = $('#chapter-modal');
  const modalCloseBtn = $('#modal-close-btn');
  const modalTitle = $('#modal-title');
  const modalBadge = $('#modal-badge');
  const modalBody = $('#modal-body');
  const regenSection = $('#regen-section');
  const regenInstructions = $('#regen-instructions');
  const regenBtn = $('#regen-btn');

  // -- Init --
  function init() {
    // Try to load saved API key from localStorage (persists across sessions)
    const savedKey = localStorage.getItem('buchapp-api-key');
    if (savedKey) {
      apiKeyInput.value = savedKey;
    }

    setupEventListeners();
    updateUI();
    updateStepper();
  }

  // -- Event Listeners --
  function setupEventListeners() {
    // API Key toggle
    toggleApiKeyBtn.addEventListener('click', () => {
      const isPassword = apiKeyInput.type === 'password';
      apiKeyInput.type = isPassword ? 'text' : 'password';
      toggleApiKeyBtn.textContent = isPassword ? 'Verbergen' : 'Anzeigen';
    });

    // Save API key to localStorage on change (persists across sessions)
    apiKeyInput.addEventListener('input', () => {
      localStorage.setItem('buchapp-api-key', apiKeyInput.value);
      updateUI();
    });

    // File upload
    uploadZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

    // Drag & Drop for files
    uploadZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadZone.classList.add('drag-over');
    });
    uploadZone.addEventListener('dragleave', () => {
      uploadZone.classList.remove('drag-over');
    });
    uploadZone.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadZone.classList.remove('drag-over');
      handleFiles(e.dataTransfer.files);
    });

    // Generate button
    generateBtn.addEventListener('click', generateBook);

    // Download button
    downloadBtn.addEventListener('click', finalizeBook);

    // Reset button (new book)
    resetBtn.addEventListener('click', resetForNewBook);

    // Input changes update UI
    bookTitleInput.addEventListener('input', updateUI);
    bookSubtitleInput.addEventListener('input', updateUI);

    // Stepper click navigation
    document.querySelectorAll('.stepper-step').forEach(step => {
      step.addEventListener('click', () => {
        const stepNum = step.dataset.step;
        const sections = ['step-api', 'step-meta', 'step-transcripts', 'step-generate', 'step-preview'];
        const target = document.getElementById(sections[stepNum - 1]);
        if (target) {
          target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      });
    });

    // Scroll-based stepper update
    window.addEventListener('scroll', updateStepper, { passive: true });

    // Modal close
    modalCloseBtn.addEventListener('click', closeModal);
    chapterModal.addEventListener('click', (e) => {
      if (e.target === chapterModal) closeModal();
    });

    // Escape key closes modal
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !chapterModal.hidden) {
        closeModal();
      }
    });

    // Regenerate button inside modal
    regenBtn.addEventListener('click', () => {
      const idx = state.modalChapterIndex;
      const instructions = regenInstructions.value.trim();
      if (idx >= 0) {
        regenerateChapter(idx, instructions);
      }
    });
  }

  // -- Stepper Logic --
  function updateStepper() {
    const sections = [
      { el: document.getElementById('step-api'), step: 1 },
      { el: document.getElementById('step-meta'), step: 2 },
      { el: document.getElementById('step-transcripts'), step: 3 },
      { el: document.getElementById('step-generate'), step: 4 },
      { el: document.getElementById('step-preview'), step: 5 }
    ];

    const scrollPos = window.scrollY + 200;
    let activeStep = 1;

    for (const section of sections) {
      if (section.el && !section.el.hidden && section.el.offsetTop <= scrollPos) {
        activeStep = section.step;
      }
    }

    // If preview is visible, step 5 can be active via scroll
    // If still generating or generated but no preview, stay on step 4
    if (state.isGenerating) {
      activeStep = Math.min(activeStep, 4);
    }

    const steps = document.querySelectorAll('.stepper-step');
    const lines = document.querySelectorAll('.stepper-line');

    steps.forEach((step, index) => {
      const stepNum = index + 1;
      step.classList.remove('active', 'completed');
      if (stepNum === activeStep) {
        step.classList.add('active');
      } else if (stepNum < activeStep) {
        step.classList.add('completed');
      }
    });

    lines.forEach((line, index) => {
      line.classList.remove('completed');
      if (index < activeStep - 1) {
        line.classList.add('completed');
      }
    });
  }

  // -- File Handling --
  function handleFileSelect(e) {
    handleFiles(e.target.files);
    fileInput.value = '';
  }

  async function handleFiles(files) {
    for (const file of files) {
      try {
        const text = await readFileAsText(file);
        const chapterNum = state.chapters.length + 1;
        const title = 'Kapitel ' + chapterNum;

        state.chapters.push({
          id: Date.now() + Math.random(),
          fileName: file.name,
          title: title,
          text: text
        });
      } catch (err) {
        showError('Datei "' + file.name + '" konnte nicht gelesen werden: ' + err.message);
      }
    }
    renderChapterList();
    updateUI();
  }

  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      // For text/markdown files
      if (file.type === 'text/plain' || file.name.endsWith('.md') || file.name.endsWith('.txt')) {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => reject(new Error('Lesefehler'));
        reader.readAsText(file, 'UTF-8');
        return;
      }

      // For Word documents (.docx and .doc)
      if (file.name.endsWith('.docx') || file.name.endsWith('.doc') || file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || file.type === 'application/msword') {
        const reader = new FileReader();
        reader.onload = async (e) => {
          try {
            const zip = await JSZip.loadAsync(e.target.result);
            const docXml = await zip.file('word/document.xml').async('string');

            // Parse XML and extract text from w:t elements
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(docXml, 'application/xml');

            // Get all paragraphs
            const paragraphs = xmlDoc.getElementsByTagNameNS(
              'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p'
            );

            const lines = [];
            for (const p of paragraphs) {
              const texts = p.getElementsByTagNameNS(
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't'
              );
              let line = '';
              for (const t of texts) {
                line += t.textContent || '';
              }
              lines.push(line);
            }

            const text = lines.join('\n').trim();
            if (!text) {
              reject(new Error('Das Word-Dokument enth\u00e4lt keinen lesbaren Text.'));
            } else {
              resolve(text);
            }
          } catch (err) {
            reject(new Error('Word-Dokument konnte nicht gelesen werden: ' + err.message));
          }
        };
        reader.onerror = () => reject(new Error('Lesefehler'));
        reader.readAsArrayBuffer(file);
        return;
      }

      // For other files, try as text
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error('Dateiformat wird nicht unterst\u00fctzt. Bitte verwende .txt, .md oder .docx Dateien.'));
      reader.readAsText(file, 'UTF-8');
    });
  }

  // -- Chapter List Rendering --
  function renderChapterList() {
    chapterList.innerHTML = '';

    state.chapters.forEach((chapter, index) => {
      const item = document.createElement('div');
      item.className = 'chapter-item';
      item.draggable = true;
      item.dataset.id = chapter.id;

      item.innerHTML =
        '<span class="chapter-drag-handle" title="Ziehen zum Umsortieren">&#9776;</span>' +
        '<span class="chapter-number">' + (index + 1) + '</span>' +
        '<div class="chapter-info">' +
          '<input class="chapter-name-input" type="text" value="' + escapeHtml(chapter.title) + '"' +
          ' data-id="' + chapter.id + '" placeholder="Kapitelname">' +
          '<div class="chapter-file-name">' + escapeHtml(chapter.fileName) + ' &mdash; ' + formatSize(chapter.text.length) + ' Zeichen</div>' +
        '</div>' +
        '<div class="chapter-actions">' +
          '<button class="btn btn-danger" data-delete="' + chapter.id + '" title="Entfernen">&times;</button>' +
        '</div>';

      // Chapter name edit
      const nameInput = item.querySelector('.chapter-name-input');
      nameInput.addEventListener('input', (e) => {
        const ch = state.chapters.find(c => c.id == e.target.dataset.id);
        if (ch) ch.title = e.target.value;
      });

      // Delete button
      const deleteBtn = item.querySelector('[data-delete]');
      deleteBtn.addEventListener('click', () => {
        state.chapters = state.chapters.filter(c => c.id != chapter.id);
        renderChapterList();
        updateUI();
      });

      // Drag & Drop for reordering
      item.addEventListener('dragstart', (e) => {
        draggedItem = chapter.id;
        item.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
      });

      item.addEventListener('dragend', () => {
        item.classList.remove('dragging');
        document.querySelectorAll('.drag-over-item').forEach(el => el.classList.remove('drag-over-item'));
        draggedItem = null;
      });

      item.addEventListener('dragover', (e) => {
        e.preventDefault();
        if (draggedItem && draggedItem !== chapter.id) {
          item.classList.add('drag-over-item');
        }
      });

      item.addEventListener('dragleave', () => {
        item.classList.remove('drag-over-item');
      });

      item.addEventListener('drop', (e) => {
        e.preventDefault();
        item.classList.remove('drag-over-item');
        if (draggedItem && draggedItem !== chapter.id) {
          reorderChapters(draggedItem, chapter.id);
        }
      });

      chapterList.appendChild(item);
    });
  }

  function reorderChapters(fromId, toId) {
    const fromIndex = state.chapters.findIndex(c => c.id == fromId);
    const toIndex = state.chapters.findIndex(c => c.id == toId);
    if (fromIndex === -1 || toIndex === -1) return;

    const [moved] = state.chapters.splice(fromIndex, 1);
    state.chapters.splice(toIndex, 0, moved);
    renderChapterList();
  }

  // -- UI Updates --
  function updateUI() {
    const count = state.chapters.length;
    chapterCount.textContent = count;

    // Estimate ~1-2 minutes per chapter
    if (count > 0) {
      const minTime = count;
      const maxTime = count * 2;
      estimatedTime.textContent = 'ca. ' + minTime + '-' + maxTime + ' Minuten';
    } else {
      estimatedTime.textContent = '-';
    }

    // Enable/disable generate button (API key is optional - demo mode without)
    const hasApiKey = apiKeyInput.value.trim().length > 10;
    const hasTitle = bookTitleInput.value.trim().length > 0;
    const hasChapters = count > 0;

    generateBtn.disabled = !(hasTitle && hasChapters) || state.isGenerating;

    // Update button text based on mode
    if (!state.isGenerating) {
      generateBtn.textContent = hasApiKey ? 'Buch erstellen (mit KI)' : 'Buch erstellen (Demo-Modus)';
    }

    // Update time estimate based on mode
    if (count > 0) {
      if (hasApiKey) {
        const minTime = count + 1;
        const maxTime = (count * 2) + 2;
        estimatedTime.textContent = 'ca. ' + minTime + '-' + maxTime + ' Minuten (inkl. Einleitung & Abschluss)';
      } else {
        estimatedTime.textContent = 'wenige Sekunden (Demo)';
      }
    }
  }

  // -- Reset for New Book --
  function resetForNewBook() {
    state.chapters = [];
    state.generatedChapters = [];
    state.introData = null;
    state.outroData = null;
    state.generatedBook = null;
    state.isGenerating = false;

    bookTitleInput.value = '';
    bookSubtitleInput.value = '';
    bookInstructionsInput.value = '';

    chapterList.innerHTML = '';
    progressArea.hidden = true;
    progressBar.style.width = '0%';

    // Hide preview section
    stepPreview.hidden = true;
    previewCards.innerHTML = '';

    renderChapterList();
    updateUI();
    updateStepper();

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // -- Book Generation --
  async function generateBook() {
    if (state.isGenerating) return;
    state.isGenerating = true;
    state.generatedBook = null;
    state.generatedChapters = [];
    state.introData = null;
    state.outroData = null;

    const apiKey = apiKeyInput.value.trim();
    const title = bookTitleInput.value.trim();
    const subtitle = bookSubtitleInput.value.trim();
    const instructions = bookInstructionsInput.value.trim();
    const isDemoMode = apiKey.length <= 10;

    // Show progress, hide preview
    progressArea.hidden = false;
    stepPreview.hidden = true;
    generateBtn.disabled = true;
    updateProgress(0, 'Starte...');

    try {
      const generatedChapters = [];
      const totalChapters = state.chapters.length;
      let introData = null;
      let outroData = null;

      if (isDemoMode) {
        // DEMO MODE: Use transcript text directly without AI
        updateProgress(10, 'Demo-Modus: Transkripte werden direkt verwendet...');

        for (let i = 0; i < totalChapters; i++) {
          const chapter = state.chapters[i];
          updateProgress(10 + ((i / totalChapters) * 75), 'Kapitel ' + (i + 1) + ' von ' + totalChapters + ': "' + chapter.title + '" wird vorbereitet...');

          // Convert raw text into structured chapter format
          const paragraphs = chapter.text.split(/\n\s*\n/).filter(p => p.trim());
          const content = paragraphs.map(p => ({
            type: 'paragraph',
            text: p.trim()
          }));

          generatedChapters.push({
            chapter_title: chapter.title,
            content: content
          });
        }
      } else {
        // AI MODE: Send to Claude API
        const chapterTitles = state.chapters.map(c => c.title);

        // Step 1: Generate Introduction
        updateProgress(3, 'Erstelle Einleitung...');
        try {
          introData = await ClaudeAPI.generateIntroduction(apiKey, title, subtitle, chapterTitles, instructions);
        } catch (e) {
          console.warn('Einleitung konnte nicht generiert werden:', e);
        }

        // Step 2: Generate each chapter with context for transitions
        updateProgress(10, 'Starte Kapitel-Verarbeitung...');
        for (let i = 0; i < totalChapters; i++) {
          const chapter = state.chapters[i];
          const progressPercent = 10 + ((i / totalChapters) * 65);
          updateProgress(progressPercent, 'Kapitel ' + (i + 1) + ' von ' + totalChapters + ': "' + chapter.title + '" wird verarbeitet...');

          // Build chapter context for smooth transitions
          var chapterContext = {
            totalChapters: totalChapters
          };
          if (i > 0) {
            chapterContext.prevTitle = state.chapters[i - 1].title;
            // If previous chapter was already generated, use its title
            if (generatedChapters[i - 1]) {
              chapterContext.prevTitle = generatedChapters[i - 1].chapter_title || state.chapters[i - 1].title;
              // Create a short summary from the previous chapter's first paragraphs
              var prevContent = generatedChapters[i - 1].content || [];
              var prevTexts = [];
              for (var p = 0; p < prevContent.length && prevTexts.length < 2; p++) {
                if (prevContent[p].text) prevTexts.push(prevContent[p].text);
              }
              chapterContext.prevSummary = prevTexts.join(' ').substring(0, 200);
            }
          }
          if (i < totalChapters - 1) {
            chapterContext.nextTitle = state.chapters[i + 1].title;
          }

          const result = await ClaudeAPI.generateChapter(
            apiKey,
            chapter.text,
            chapter.title,
            i + 1,
            title,
            instructions,
            chapterContext
          );

          generatedChapters.push(result);
          updateProgress(10 + (((i + 1) / totalChapters) * 65), 'Kapitel ' + (i + 1) + ' von ' + totalChapters + ' fertig.');
        }

        // Step 3: Generate Outro/Closing
        updateProgress(80, 'Erstelle Abschlusswort...');
        try {
          outroData = await ClaudeAPI.generateOutro(apiKey, title, subtitle, chapterTitles, instructions);
        } catch (e) {
          console.warn('Abschlusswort konnte nicht generiert werden:', e);
        }
      }

      // Store generated data in state
      state.generatedChapters = generatedChapters;
      state.introData = introData;
      state.outroData = outroData;

      updateProgress(100, 'Generierung abgeschlossen! Vorschau wird geladen...');

      // Show preview instead of immediately building DOCX
      setTimeout(() => {
        progressArea.hidden = true;
        showPreview();
      }, 600);

    } catch (err) {
      showError(err.message);
      progressArea.hidden = true;
    } finally {
      state.isGenerating = false;
      generateBtn.disabled = false;
      updateUI();
    }
  }

  function updateProgress(percent, text) {
    progressBar.style.width = percent + '%';
    progressText.textContent = text;
  }


  // ============================================
  //  Preview Cards
  // ============================================

  function showPreview() {
    stepPreview.hidden = false;
    renderPreviewCards();
    updateStepper();

    // Scroll to preview section
    setTimeout(() => {
      stepPreview.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
  }

  function renderPreviewCards() {
    previewCards.innerHTML = '';

    // Intro card (if exists)
    if (state.introData && state.introData.content) {
      const card = createPreviewCard(-1, 'Einleitung', state.introData.content, 'intro');
      previewCards.appendChild(card);
    }

    // Chapter cards
    state.generatedChapters.forEach((chapter, index) => {
      const title = chapter.chapter_title || ('Kapitel ' + (index + 1));
      const card = createPreviewCard(index, title, chapter.content, 'chapter');
      previewCards.appendChild(card);
    });

    // Outro card (if exists)
    if (state.outroData && state.outroData.content) {
      const card = createPreviewCard(-2, 'Abschlusswort', state.outroData.content, 'outro');
      previewCards.appendChild(card);
    }
  }

  function createPreviewCard(index, title, content, cardType) {
    const card = document.createElement('div');
    card.className = 'preview-card';
    card.dataset.index = index;

    if (cardType === 'intro') card.classList.add('intro-card');
    if (cardType === 'outro') card.classList.add('outro-card');

    // Determine number label
    let numberLabel = '';
    if (cardType === 'intro') {
      numberLabel = 'E';
    } else if (cardType === 'outro') {
      numberLabel = 'A';
    } else {
      numberLabel = '' + (index + 1);
    }

    // Extract plain text excerpt from content array
    const excerpt = extractExcerpt(content, 200);

    // Check if API key present for regen availability
    const hasApiKey = apiKeyInput.value.trim().length > 10;
    const canRegen = hasApiKey && cardType === 'chapter';

    card.innerHTML =
      '<div class="preview-card-header">' +
        '<span class="preview-card-number">' + escapeHtml(numberLabel) + '</span>' +
        '<span class="preview-card-title">' + escapeHtml(title) + '</span>' +
      '</div>' +
      '<div class="preview-card-badge">' +
        '<span class="badge badge-done">Fertig</span>' +
      '</div>' +
      '<div class="preview-card-excerpt">' + escapeHtml(excerpt) + '</div>' +
      '<div class="preview-card-actions">' +
        '<button class="btn btn-preview" data-action="preview">Vorschau</button>' +
        (canRegen
          ? '<button class="btn btn-regen" data-action="regen">Neu generieren</button>'
          : '') +
      '</div>';

    // Event: Preview button
    card.querySelector('[data-action="preview"]').addEventListener('click', () => {
      openChapterModal(index, false, cardType);
    });

    // Event: Regen button
    const regenButton = card.querySelector('[data-action="regen"]');
    if (regenButton) {
      regenButton.addEventListener('click', () => {
        openChapterModal(index, true, cardType);
      });
    }

    return card;
  }

  function extractExcerpt(content, maxLen) {
    if (!content || !Array.isArray(content)) return '';
    let text = '';
    for (const block of content) {
      if (block.text) {
        text += block.text + ' ';
      } else if (block.items && Array.isArray(block.items)) {
        text += block.items.join(', ') + ' ';
      }
      if (text.length >= maxLen) break;
    }
    text = text.trim();
    if (text.length > maxLen) {
      text = text.substring(0, maxLen).replace(/\s+\S*$/, '') + '...';
    }
    return text;
  }


  // ============================================
  //  Chapter Modal
  // ============================================

  function openChapterModal(index, regenMode, cardType) {
    state.modalChapterIndex = index;
    state.modalRegenMode = regenMode;

    // Determine which content to show
    let title = '';
    let content = [];

    if (index === -1 && cardType === 'intro') {
      title = 'Einleitung';
      content = state.introData ? state.introData.content : [];
    } else if (index === -2 && cardType === 'outro') {
      title = 'Abschlusswort';
      content = state.outroData ? state.outroData.content : [];
    } else {
      const chapter = state.generatedChapters[index];
      if (chapter) {
        title = chapter.chapter_title || ('Kapitel ' + (index + 1));
        content = chapter.content || [];
      }
    }

    modalTitle.textContent = title;
    modalBadge.innerHTML = '<span class="badge badge-done">Fertig</span>';
    modalBody.innerHTML = renderFormattedContent(content);

    // Show/hide regen section
    if (regenMode && index >= 0) {
      regenSection.hidden = false;
      regenInstructions.value = '';
      regenBtn.disabled = false;
      regenBtn.textContent = 'Neu generieren';
    } else {
      regenSection.hidden = true;
    }

    chapterModal.hidden = false;
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    chapterModal.hidden = true;
    document.body.style.overflow = '';
    state.modalChapterIndex = -1;
    state.modalRegenMode = false;
  }

  function renderFormattedContent(content) {
    if (!content || !Array.isArray(content)) return '<p class="preview-paragraph">Kein Inhalt vorhanden.</p>';

    let html = '';
    for (const block of content) {
      switch (block.type) {
        case 'paragraph':
          html += '<p class="preview-paragraph">' + escapeHtml(block.text) + '</p>';
          break;
        case 'heading':
          html += '<h3 class="preview-heading">' + escapeHtml(block.text) + '</h3>';
          break;
        case 'subheading_red':
          html += '<h4 class="preview-subheading-red">' + escapeHtml(block.text) + '</h4>';
          break;
        case 'subheading_green':
          html += '<h4 class="preview-subheading-green">' + escapeHtml(block.text) + '</h4>';
          break;
        case 'quote':
          html += '<blockquote class="preview-quote">' + escapeHtml(block.text) + '</blockquote>';
          break;
        case 'tip':
          html += '<div class="preview-tip">' + escapeHtml(block.text) + '</div>';
          break;
        case 'goldbox':
          html += '<div class="preview-goldbox">' + escapeHtml(block.text) + '</div>';
          break;
        case 'exercise':
          html += '<div class="preview-exercise">' + escapeHtml(block.text) + '</div>';
          break;
        case 'list':
          if (block.items && Array.isArray(block.items)) {
            html += '<ul class="preview-list">';
            for (const item of block.items) {
              html += '<li>' + escapeHtml(item) + '</li>';
            }
            html += '</ul>';
          }
          break;
        default:
          if (block.text) {
            html += '<p class="preview-paragraph">' + escapeHtml(block.text) + '</p>';
          }
          break;
      }
    }
    return html;
  }


  // ============================================
  //  Chapter Regeneration
  // ============================================

  async function regenerateChapter(index, additionalInstructions) {
    if (index < 0 || index >= state.generatedChapters.length) return;

    const apiKey = apiKeyInput.value.trim();
    if (apiKey.length <= 10) {
      showError('Zum Neu-Generieren wird ein API-Schl\u00fcssel ben\u00f6tigt.');
      return;
    }

    const chapter = state.chapters[index];
    if (!chapter) return;

    const title = bookTitleInput.value.trim();
    const baseInstructions = bookInstructionsInput.value.trim();

    // Combine base instructions with additional per-chapter instructions
    let combinedInstructions = baseInstructions;
    if (additionalInstructions) {
      combinedInstructions = combinedInstructions
        ? combinedInstructions + '\n\nZUS\u00c4TZLICH F\u00dcR DIESES KAPITEL:\n' + additionalInstructions
        : additionalInstructions;
    }

    // Update UI to show regenerating state
    regenBtn.disabled = true;
    regenBtn.textContent = 'Wird generiert...';
    modalBadge.innerHTML = '<span class="badge badge-regenerating">Wird generiert...</span>';

    // Also update the card badge
    updateCardBadge(index, 'regenerating');

    try {
      // Build context from neighboring chapters for transitions
      var regenContext = { totalChapters: state.generatedChapters.length };
      if (index > 0 && state.generatedChapters[index - 1]) {
        regenContext.prevTitle = state.generatedChapters[index - 1].chapter_title;
      }
      if (index < state.generatedChapters.length - 1 && state.generatedChapters[index + 1]) {
        regenContext.nextTitle = state.generatedChapters[index + 1].chapter_title;
      }

      const result = await ClaudeAPI.generateChapter(
        apiKey,
        chapter.text,
        chapter.title,
        index + 1,
        title,
        combinedInstructions,
        regenContext
      );

      // Replace the chapter data
      state.generatedChapters[index] = result;

      // Re-render the preview card for this chapter
      renderPreviewCards();

      // Update modal if still open on the same chapter
      if (!chapterModal.hidden && state.modalChapterIndex === index) {
        const newTitle = result.chapter_title || ('Kapitel ' + (index + 1));
        modalTitle.textContent = newTitle;
        modalBody.innerHTML = renderFormattedContent(result.content);
        modalBadge.innerHTML = '<span class="badge badge-done">Fertig</span>';
        regenBtn.disabled = false;
        regenBtn.textContent = 'Neu generieren';
        regenInstructions.value = '';
      }

    } catch (err) {
      showError('Kapitel konnte nicht neu generiert werden: ' + err.message);
      // Reset UI
      modalBadge.innerHTML = '<span class="badge badge-done">Fertig</span>';
      regenBtn.disabled = false;
      regenBtn.textContent = 'Neu generieren';
      updateCardBadge(index, 'done');
    }
  }

  function updateCardBadge(index, status) {
    const cards = previewCards.querySelectorAll('.preview-card');
    // Find the card matching this index. Cards include intro (index -1), chapters (0..n), outro (-2)
    for (const card of cards) {
      if (parseInt(card.dataset.index) === index) {
        const badgeContainer = card.querySelector('.preview-card-badge');
        if (status === 'regenerating') {
          badgeContainer.innerHTML = '<span class="badge badge-regenerating">Wird generiert...</span>';
        } else {
          badgeContainer.innerHTML = '<span class="badge badge-done">Fertig</span>';
        }
        break;
      }
    }
  }


  // ============================================
  //  Finalize & Download
  // ============================================

  async function finalizeBook() {
    if (state.generatedChapters.length === 0) {
      showError('Keine generierten Kapitel vorhanden.');
      return;
    }

    const title = bookTitleInput.value.trim();
    const subtitle = bookSubtitleInput.value.trim();

    // Disable download button while generating
    downloadBtn.disabled = true;
    downloadBtn.textContent = 'Erstelle Word-Dokument...';

    try {
      // Load template if not already loaded
      await DocxGenerator.loadTemplate();

      // Generate the DOCX
      const bookBlob = await DocxGenerator.generateBook(
        title,
        subtitle,
        state.generatedChapters,
        state.introData,
        state.outroData
      );

      state.generatedBook = {
        blob: bookBlob,
        fileName: title.replace(/[^a-zA-Z0-9\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df\s-]/g, '').replace(/\s+/g, '_') + '.docx'
      };

      // Trigger download
      saveAs(state.generatedBook.blob, state.generatedBook.fileName);

    } catch (err) {
      showError('Fehler beim Erstellen des Word-Dokuments: ' + err.message);
    } finally {
      downloadBtn.disabled = false;
      downloadBtn.textContent = 'Buch herunterladen (.docx)';
    }
  }


  // -- Error Handling --
  function showError(message) {
    const toast = document.createElement('div');
    toast.className = 'error-toast';
    toast.innerHTML =
      '<button class="close-toast">&times;</button>' +
      escapeHtml(message);
    document.body.appendChild(toast);

    toast.querySelector('.close-toast').addEventListener('click', () => toast.remove());
    setTimeout(() => toast.remove(), 8000);
  }

  // -- Utilities --
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function formatSize(chars) {
    if (chars < 1000) return chars;
    return (chars / 1000).toFixed(1) + 'k';
  }

  // -- Start --
  init();
})();
