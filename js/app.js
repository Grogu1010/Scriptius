const editor = document.getElementById('scriptEditor');
const titleInput = document.getElementById('scriptTitle');
const beatNotes = document.getElementById('beatNotes');
const autosaveStatus = document.getElementById('autosaveStatus');
const wordStat = document.getElementById('wordStat');
const sceneStat = document.getElementById('sceneStat');
const runtimeStat = document.getElementById('runtimeStat');
const pageIndicator = document.getElementById('pageIndicator');
const sceneList = document.getElementById('sceneList');
const sceneCountLabel = document.getElementById('sceneCount');
const insightsList = document.getElementById('insightsList');
const previewContainer = document.getElementById('scriptPreview');
const snackbar = document.getElementById('snackbar');
const exportMenu = document.getElementById('exportMenu');
const exportMenuBtn = document.getElementById('exportMenuBtn');
const newDocumentBtn = document.getElementById('newDocumentBtn');
const importBtn = document.getElementById('importDocumentBtn');
const filePicker = document.getElementById('filePicker');
const downloadFountainBtn = document.getElementById('downloadFountainBtn');
const downloadTxtBtn = document.getElementById('downloadTxtBtn');
const printBtn = document.getElementById('printBtn');
const shareBtn = document.getElementById('shareBtn');
const shareLinkInput = document.getElementById('shareLink');
const copyShareLinkBtn = document.getElementById('copyShareLinkBtn');
const shareModal = document.getElementById('shareModal');
const shareModalLink = document.getElementById('shareModalLink');
const shareModalCopy = document.getElementById('shareModalCopy');
const closeShareModalBtn = document.getElementById('closeShareModal');
const nativeShareBtn = document.getElementById('nativeShareBtn');
const focusModeToggle = document.getElementById('focusModeToggle');
const statsToggle = document.getElementById('statsToggle');
const themeToggle = document.getElementById('themeToggle');
const toolbar = document.querySelector('.toolbar');
const autosaveIndicator = document.querySelector('[data-autosave-indicator]');
const progressTrack = document.getElementById('scriptProgressTrack');
const progressValue = document.getElementById('scriptProgressValue');
const progressLabel = document.getElementById('progressLabel');

const libraryBtn = document.getElementById('libraryBtn');
const libraryDrawer = document.getElementById('libraryDrawer');
const closeLibraryBtn = document.getElementById('closeLibraryBtn');
const librarySearch = document.getElementById('librarySearch');
const createDocumentBtn = document.getElementById('createDocumentBtn');
const documentList = document.getElementById('documentList');
const emptyLibraryMessage = document.getElementById('emptyLibraryMessage');
const viewMenu = document.getElementById('viewMenu');
const viewMenuBtn = document.getElementById('viewMenuBtn');
const toggleOutlineBtn = document.getElementById('toggleOutlineBtn');
const togglePreviewBtn = document.getElementById('togglePreviewBtn');

const DOCUMENTS_KEY = 'scriptius-documents';
const ACTIVE_DOCUMENT_KEY = 'scriptius-active-document';
const LEGACY_STORAGE_KEY = 'scriptius-document';
const THEME_KEY = 'scriptius-theme';
const LAYOUT_KEY = 'scriptius-layout';
const AUTOSAVE_DELAY = 800;
const LINES_PER_PAGE = 55;
const PROGRESS_PAGE_GOAL = 110;
const NUMBER_FORMATTER = new Intl.NumberFormat();

const INDENTS = Object.freeze({
  action: '',
  scene: '',
  dialogue: '          ',
  parenthetical: '               ',
  character: '                    ',
  transition: '                              ',
  lyric: '          ',
  section: '',
  empty: '',
});

const CHARACTER_INDENT = INDENTS.character.length;
const DIALOGUE_INDENT = INDENTS.dialogue.length;
const PARENTHETICAL_INDENT = INDENTS.parenthetical.length;

const SCENE_PATTERN = /^(INT\.?|EXT\.?|EST\.?|INT\.?\s*\/\s*EXT\.?|I\/E\.?|INT\.?\s*&\s*EXT\.?)\b/i;
const TRANSITION_PATTERN = /(CUT TO:?|SMASH TO:?|MATCH CUT TO:?|DISSOLVE TO:?|WIPE TO:?|FADE OUT:?|FADE IN:?|BACK TO:?|TO BLACK:?|END (?:OF|ON)|IRIS OUT:?|IRIS IN:?|.*TO:)$/i;

let formattingResult = { formattedText: '', lineInfos: [] };
let isFormatting = false;
let autosaveTimer = null;
let lastSavedAt = null;
let isDirty = false;
let documents = [];
let currentDocumentId = null;
let libraryPreviouslyFocused = null;
const previewState = {
  lineToPage: [],
  pageCount: 1,
};

init();

function init() {
  applyStoredTheme();
  applyStoredLayout();
  attachEventListeners();
  bootstrapDocuments();
  loadFromShareLink();
  updateViewMenuLabels();
  editor.focus();
  updateAutosaveStatus();
  setInterval(updateAutosaveStatus, 30000);
}

function attachEventListeners() {
  editor.addEventListener('input', () => {
    if (isFormatting) return;
    applyFormattingAndRefresh();
  });

  editor.addEventListener('keydown', handleEditorKeydown);
  editor.addEventListener('click', handleCaretChange);
  editor.addEventListener('keyup', handleCaretChange);

  document.addEventListener('selectionchange', () => {
    if (document.activeElement === editor) {
      handleCaretChange();
    }
  });

  titleInput.addEventListener('input', () => {
    markDirty();
    scheduleAutosave();
  });

  beatNotes.addEventListener('input', () => {
    markDirty();
    scheduleAutosave();
  });

  libraryBtn.addEventListener('click', openLibrary);
  closeLibraryBtn.addEventListener('click', closeLibrary);
  libraryDrawer.addEventListener('click', (event) => {
    if (event.target === libraryDrawer) {
      closeLibrary();
    }
  });
  librarySearch.addEventListener('input', renderDocumentLibrary);
  createDocumentBtn.addEventListener('click', () => {
    if (isDirty && editor.value.trim()) {
      const proceed = window.confirm('Start a new script? Unsaved changes will be lost.');
      if (!proceed) return;
    }
    const doc = createDocument();
    loadDocument(doc, { preserveCaret: false });
    closeLibrary();
    editor.focus();
    showSnackbar('Started a new script');
  });
  documentList.addEventListener('click', handleLibraryClick);

  newDocumentBtn.addEventListener('click', handleNewDocument);
  importBtn.addEventListener('click', () => filePicker.click());
  filePicker.addEventListener('change', handleImportFile);

  downloadFountainBtn.addEventListener('click', () => {
    const filename = createFilename('.fountain');
    downloadFile(filename, formattingResult.formattedText);
    showSnackbar('Fountain file downloaded');
  });

  downloadTxtBtn.addEventListener('click', () => {
    const filename = createFilename('.txt');
    downloadFile(filename, formattingResult.formattedText);
    showSnackbar('Text file downloaded');
  });

  printBtn.addEventListener('click', () => window.print());

  exportMenuBtn.addEventListener('click', toggleExportMenu);
  document.addEventListener('click', (event) => {
    if (!exportMenu.contains(event.target) && event.target !== exportMenuBtn) {
      exportMenu.classList.remove('open');
      exportMenuBtn.setAttribute('aria-expanded', 'false');
    }
    if (!viewMenu.contains(event.target) && event.target !== viewMenuBtn) {
      closeViewMenu();
    }
  });

  viewMenuBtn.addEventListener('click', toggleViewMenu);
  toggleOutlineBtn.addEventListener('click', () => toggleWorkspacePanel('outline'));
  togglePreviewBtn.addEventListener('click', () => toggleWorkspacePanel('preview'));

  shareBtn.addEventListener('click', openShareModal);
  copyShareLinkBtn.addEventListener('click', handleCopyShareLink);
  shareModalCopy.addEventListener('click', handleCopyShareLink);
  closeShareModalBtn.addEventListener('click', closeShareModal);
  shareModal.addEventListener('click', (event) => {
    if (event.target === shareModal) closeShareModal();
  });

  nativeShareBtn.addEventListener('click', async () => {
    const link = ensureShareLink();
    if (!navigator.share) {
      showSnackbar('Native share is not supported here');
      return;
    }
    try {
      await navigator.share({
        title: titleInput.value || 'Scriptius draft',
        text: 'Take a look at my Scriptius draft.',
        url: link,
      });
    } catch (error) {
      if (error?.name !== 'AbortError') {
        console.error(error);
        showSnackbar('Unable to share right now');
      }
    }
  });

  focusModeToggle.addEventListener('click', () => {
    document.body.classList.toggle('focus-mode');
    focusModeToggle.textContent = document.body.classList.contains('focus-mode')
      ? 'Exit focus mode'
      : 'Focus mode';
  });

  statsToggle.addEventListener('click', () => {
    document.body.classList.toggle('stats-visible');
    statsToggle.textContent = document.body.classList.contains('stats-visible') ? 'Hide stats' : 'Show stats';
  });

  themeToggle.addEventListener('click', toggleTheme);

  toolbar.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-snippet]');
    if (!button) return;
    insertSnippet(button.dataset.snippet);
  });

  sceneList.addEventListener('click', (event) => {
    const button = event.target.closest('.scene-link');
    if (!button) return;
    const lineIndex = Number(button.dataset.lineIndex);
    focusLine(lineIndex);
  });

  previewContainer.addEventListener('click', (event) => {
    const lineEl = event.target.closest('.script-line');
    if (!lineEl) return;
    const lineIndex = Number(lineEl.dataset.lineIndex);
    focusLine(lineIndex);
  });

  document.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
      event.preventDefault();
      const filename = createFilename('.fountain');
      downloadFile(filename, formattingResult.formattedText);
      showSnackbar('Saved a Fountain copy');
    }
    if (event.key === 'Escape') {
      if (shareModal.classList.contains('open')) {
        closeShareModal();
      } else if (libraryDrawer.classList.contains('open')) {
        closeLibrary();
      } else {
        closeViewMenu();
      }
    }
  });
}

function handleEditorKeydown(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    insertLineBreak(event.shiftKey);
    return;
  }

  if (event.key === 'Tab') {
    event.preventDefault();
    insertSnippet('    ');
  }
}

function handleCaretChange() {
  const caret = editor.selectionStart;
  highlightFromCaret(caret, { scrollPreview: false });
}

function insertSnippet(snippet) {
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  const text = editor.value;
  const before = text.slice(0, start);
  const after = text.slice(end);
  const nextValue = before + snippet + after;
  const caret = start + snippet.length;
  editor.value = nextValue;
  editor.setSelectionRange(caret, caret);
  applyFormattingAndRefresh();
}

function insertLineBreak(forcePlain) {
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  const text = editor.value;
  const before = text.slice(0, start);
  const after = text.slice(end);

  let indent = '';
  if (!forcePlain) {
    const currentLineInfo = findLineInfoForCaret(start);
    const nextType = suggestNextType(currentLineInfo?.type);
    indent = INDENTS[nextType] || '';
  }

  const insertion = `\n${indent}`;
  editor.value = `${before}${insertion}${after}`;
  const caret = start + insertion.length;
  editor.setSelectionRange(caret, caret);
  applyFormattingAndRefresh();
}

function suggestNextType(currentType) {
  switch (currentType) {
    case 'character':
    case 'parenthetical':
      return 'dialogue';
    case 'dialogue':
      return 'dialogue';
    case 'lyric':
      return 'lyric';
    case 'scene':
    case 'transition':
    case 'section':
      return 'action';
    default:
      return 'action';
  }
}

function findLineInfoForCaret(position) {
  const lineInfos = formattingResult.lineInfos;
  if (!lineInfos.length) return null;
  const lineIndex = getLineIndexFromCaret(position, lineInfos);
  return lineInfos[lineIndex] || null;
}

function toggleExportMenu() {
  const isOpen = exportMenu.classList.toggle('open');
  exportMenuBtn.setAttribute('aria-expanded', String(isOpen));
}

function toggleViewMenu() {
  const isOpen = viewMenu.classList.toggle('open');
  viewMenuBtn.setAttribute('aria-expanded', String(isOpen));
}

function closeViewMenu() {
  if (!viewMenu.classList.contains('open')) return;
  viewMenu.classList.remove('open');
  viewMenuBtn.setAttribute('aria-expanded', 'false');
}

function toggleWorkspacePanel(panel) {
  if (panel === 'outline') {
    document.body.classList.toggle('hide-outline');
  } else if (panel === 'preview') {
    document.body.classList.toggle('hide-preview');
  }
  persistLayout();
  updateViewMenuLabels();
}

function applyStoredLayout() {
  try {
    const stored = localStorage.getItem(LAYOUT_KEY);
    if (!stored) return;
    const payload = JSON.parse(stored);
    if (payload?.hideOutline) {
      document.body.classList.add('hide-outline');
    }
    if (payload?.hidePreview) {
      document.body.classList.add('hide-preview');
    }
  } catch (error) {
    console.warn('Failed to load layout preferences', error);
  }
}

function persistLayout() {
  const payload = {
    hideOutline: document.body.classList.contains('hide-outline'),
    hidePreview: document.body.classList.contains('hide-preview'),
  };
  try {
    localStorage.setItem(LAYOUT_KEY, JSON.stringify(payload));
  } catch (error) {
    console.warn('Failed to store layout preferences', error);
  }
}

function updateViewMenuLabels() {
  const outlineHidden = document.body.classList.contains('hide-outline');
  const previewHidden = document.body.classList.contains('hide-preview');
  toggleOutlineBtn.dataset.state = '✓';
  togglePreviewBtn.dataset.state = '✓';
  toggleOutlineBtn.textContent = outlineHidden ? 'Show outline' : 'Hide outline';
  togglePreviewBtn.textContent = previewHidden ? 'Show preview' : 'Hide preview';
  toggleOutlineBtn.setAttribute('aria-pressed', String(!outlineHidden));
  togglePreviewBtn.setAttribute('aria-pressed', String(!previewHidden));
}

function openLibrary() {
  closeViewMenu();
  exportMenu.classList.remove('open');
  exportMenuBtn.setAttribute('aria-expanded', 'false');
  renderDocumentLibrary();
  libraryDrawer.classList.add('open');
  libraryDrawer.setAttribute('aria-hidden', 'false');
  document.body.classList.add('drawer-open');
  libraryPreviouslyFocused = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  requestAnimationFrame(() => {
    librarySearch.focus();
    librarySearch.select();
  });
}

function closeLibrary() {
  libraryDrawer.classList.remove('open');
  libraryDrawer.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('drawer-open');
  closeViewMenu();
  if (libraryPreviouslyFocused?.focus) {
    libraryPreviouslyFocused.focus();
  }
  libraryPreviouslyFocused = null;
}

function handleLibraryClick(event) {
  const deleteBtn = event.target.closest('.document-card__delete');
  const bodyBtn = event.target.closest('.document-card__body');
  const item = event.target.closest('.document-card');
  if (!item) return;
  const docId = item.dataset.documentId;
  if (!docId) return;

  if (deleteBtn) {
    event.stopPropagation();
    const doc = documents.find((entry) => entry.id === docId);
    const title = doc?.title?.trim() || 'Untitled script';
    const proceed = window.confirm(`Delete "${title}"? This cannot be undone.`);
    if (!proceed) return;
    deleteDocument(docId);
    renderDocumentLibrary();
    return;
  }

  if (bodyBtn) {
    if (docId === currentDocumentId) {
      closeLibrary();
      return;
    }
    if (isDirty) {
      const proceed = window.confirm('Switch scripts without saving recent changes?');
      if (!proceed) return;
    }
    const doc = documents.find((entry) => entry.id === docId);
    if (!doc) return;
    loadDocument(doc, { preserveCaret: false });
    closeLibrary();
    editor.focus();
    showSnackbar(`Switched to ${doc.title?.trim() || 'Untitled script'}`);
  }
}

function deleteDocument(id) {
  const index = documents.findIndex((doc) => doc.id === id);
  if (index === -1) return;
  documents.splice(index, 1);
  persistDocuments();
  if (!documents.length) {
    const doc = createDocument();
    loadDocument(doc, { preserveCaret: false });
    showSnackbar('Script deleted');
    renderDocumentLibrary();
    return;
  }
  if (currentDocumentId === id) {
    const fallback = documents[index] || documents[index - 1] || documents[0];
    loadDocument(fallback, { preserveCaret: false });
  }
  showSnackbar('Script deleted');
  renderDocumentLibrary();
}

function renderDocumentLibrary() {
  const queryRaw = librarySearch.value.trim();
  const query = queryRaw.toLowerCase();
  const sorted = [...documents].sort((a, b) => getDocumentSortValue(b) - getDocumentSortValue(a));
  const filtered = query
    ? sorted.filter((doc) => {
        const haystack = `${doc.title} ${doc.content} ${doc.beats}`.toLowerCase();
        return haystack.includes(query);
      })
    : sorted;

  documentList.innerHTML = '';
  const fragment = document.createDocumentFragment();

  filtered.forEach((doc) => {
    const item = document.createElement('li');
    item.className = 'document-card';
    item.dataset.documentId = doc.id;
    if (doc.id === currentDocumentId) {
      item.classList.add('active');
    }

    const bodyButton = document.createElement('button');
    bodyButton.type = 'button';
    bodyButton.className = 'document-card__body';
    const title = document.createElement('span');
    title.className = 'document-card__title';
    title.textContent = doc.title?.trim() || 'Untitled script';
    const meta = document.createElement('span');
    meta.className = 'document-card__meta';
    meta.textContent = formatDocumentTimestamp(doc.updatedAt);
    const excerpt = document.createElement('span');
    excerpt.className = 'document-card__excerpt';
    excerpt.textContent = getDocumentExcerpt(doc);
    bodyButton.append(title, meta, excerpt);

    const actions = document.createElement('div');
    actions.className = 'document-card__actions';
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'document-card__delete';
    deleteButton.textContent = 'Delete';
    actions.appendChild(deleteButton);

    item.append(bodyButton, actions);
    fragment.appendChild(item);
  });

  documentList.appendChild(fragment);
  if (filtered.length === 0) {
    if (queryRaw) {
      emptyLibraryMessage.textContent = `No scripts match "${queryRaw}"`;
    } else {
      emptyLibraryMessage.textContent = 'Your autosaved scripts will appear here.';
    }
  } else {
    emptyLibraryMessage.textContent = 'Your autosaved scripts will appear here.';
  }
  emptyLibraryMessage.hidden = filtered.length > 0;
}

function getDocumentSortValue(doc) {
  const updated = doc.updatedAt ? Date.parse(doc.updatedAt) : 0;
  if (Number.isFinite(updated) && updated) {
    return updated;
  }
  const created = doc.createdAt ? Date.parse(doc.createdAt) : 0;
  return Number.isFinite(created) ? created : 0;
}

function formatDocumentTimestamp(isoString) {
  if (!isoString) return 'Never saved';
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return 'Never saved';
  const diff = Date.now() - date.getTime();
  if (diff < 45000) return 'Updated moments ago';
  if (diff < 3600000) return `Updated ${Math.max(1, Math.round(diff / 60000))}m ago`;
  if (diff < 86400000) return `Updated ${Math.round(diff / 3600000)}h ago`;
  return `Updated ${date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}`;
}

function getDocumentExcerpt(doc) {
  const content = doc.content || '';
  const beats = doc.beats || '';
  const lines = content.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (lines.length) {
    return lines[0].slice(0, 120);
  }
  const beatLines = beats.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (beatLines.length) {
    return beatLines[0].slice(0, 120);
  }
  return 'No pages yet — start writing!';
}

function applyFormattingAndRefresh(options = {}) {
  const { preserveCaret = true } = options;
  const rawText = editor.value.replace(/\r\n?/g, '\n');
  const caret = preserveCaret ? editor.selectionStart : rawText.length;
  const { formattedText, lineInfos } = formatScript(rawText);

  let newCaret = caret;
  if (formattedText !== rawText) {
    newCaret = preserveCaret ? mapCaretPosition(caret, lineInfos, formattedText) : formattedText.length;
    isFormatting = true;
    editor.value = formattedText;
    isFormatting = false;
  }

  if (preserveCaret) {
    const clamped = Math.min(newCaret, formattedText.length);
    editor.setSelectionRange(clamped, clamped);
  }

  formattingResult = { formattedText, lineInfos };
  updatePreview(lineInfos);
  updateOutline(lineInfos);
  updateStats(lineInfos);
  updateInsights(lineInfos);
  highlightFromCaret(editor.selectionStart, { scrollPreview: false });
  scheduleAutosave();
}

function formatScript(text) {
  const lines = text.split('\n');
  const lineInfos = [];
  const formattedLines = [];
  let previousType = 'action';
  let offsetOriginal = 0;
  let offsetFormatted = 0;

  for (let i = 0; i < lines.length; i += 1) {
    const raw = lines[i];
    const info = analyseLine(raw, previousType);
    const exportLine = info.type === 'empty' ? '' : info.indent + info.content + info.trailingWhitespace;
    const startOriginal = offsetOriginal;
    const endOriginal = startOriginal + raw.length;
    const startFormatted = offsetFormatted;
    const endFormatted = startFormatted + exportLine.length;

    lineInfos.push({
      index: i,
      raw,
      type: info.type,
      content: info.content,
      indent: info.indent,
      indentLength: info.indent.length,
      formattedContentLength: info.content.length,
      leadingWhitespace: info.leadingWhitespace,
      trailingWhitespaceLength: info.trailingWhitespace.length,
      startOriginal,
      endOriginal,
      startFormatted,
      endFormatted,
      endFormattedWithBreak: endFormatted + (i < lines.length - 1 ? 1 : 0),
      hasBreak: i < lines.length - 1,
      location: info.location,
      timeOfDay: info.timeOfDay,
    });

    formattedLines.push(exportLine);

    if (info.type !== 'empty') {
      previousType = info.type;
    }

    offsetOriginal = endOriginal + (i < lines.length - 1 ? 1 : 0);
    offsetFormatted = endFormatted + (i < lines.length - 1 ? 1 : 0);
  }

  const formattedText = formattedLines.join('\n');
  return { formattedText, lineInfos };
}

function analyseLine(rawLine, previousType) {
  const normalisedTabs = rawLine.replace(/\t/g, '    ');
  const trimmedRight = normalisedTabs.replace(/\s+$/u, '');
  const trailingWhitespaceLength = normalisedTabs.length - trimmedRight.length;
  const trailingWhitespace = trailingWhitespaceLength > 0 ? normalisedTabs.slice(-trailingWhitespaceLength) : '';
  const trimmedLeft = trimmedRight.replace(/^\s+/u, '');
  const leadingWhitespace = trimmedRight.length - trimmedLeft.length;
  const trimmed = trimmedLeft;

  if (!trimmed) {
    return {
      type: 'empty',
      content: '',
      indent: INDENTS.empty,
      leadingWhitespace,
      trailingWhitespace: '',
      location: null,
      timeOfDay: null,
    };
  }

  if (isSceneHeading(trimmed)) {
    const content = normaliseSceneHeading(trimmed);
    const sceneBits = parseSceneParts(content);
    return {
      type: 'scene',
      content,
      indent: INDENTS.scene,
      leadingWhitespace,
      trailingWhitespace,
      location: sceneBits.location,
      timeOfDay: sceneBits.timeOfDay,
    };
  }

  if (isTransition(trimmed)) {
    return {
      type: 'transition',
      content: normaliseTransition(trimmed),
      indent: INDENTS.transition,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (isSectionHeading(trimmed)) {
    return {
      type: 'section',
      content: trimmed.replace(/^#+\s*/, '').toUpperCase(),
      indent: INDENTS.section,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (leadingWhitespace >= CHARACTER_INDENT - 2 && isCharacter(trimmed)) {
    return {
      type: 'character',
      content: normaliseCharacter(trimmed),
      indent: INDENTS.character,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (leadingWhitespace >= PARENTHETICAL_INDENT - 2 && isParenthetical(trimmed)) {
    return {
      type: 'parenthetical',
      content: normaliseParenthetical(trimmed),
      indent: INDENTS.parenthetical,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (leadingWhitespace >= DIALOGUE_INDENT - 2 && !isParenthetical(trimmed)) {
    return {
      type: 'dialogue',
      content: trimmed,
      indent: INDENTS.dialogue,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (isCharacter(trimmed)) {
    return {
      type: 'character',
      content: normaliseCharacter(trimmed),
      indent: INDENTS.character,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (trimmed.startsWith('~')) {
    return {
      type: 'lyric',
      content: trimmed.replace(/^~\s*/, ''),
      indent: INDENTS.lyric,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  if (previousType === 'character' || previousType === 'dialogue' || previousType === 'parenthetical') {
    if (isParenthetical(trimmed)) {
      return {
        type: 'parenthetical',
        content: normaliseParenthetical(trimmed),
        indent: INDENTS.parenthetical,
        leadingWhitespace,
        trailingWhitespace,
        location: null,
        timeOfDay: null,
      };
    }
  }

  if (previousType === 'character' || previousType === 'parenthetical') {
    return {
      type: 'dialogue',
      content: trimmed,
      indent: INDENTS.dialogue,
      leadingWhitespace,
      trailingWhitespace,
      location: null,
      timeOfDay: null,
    };
  }

  return {
    type: 'action',
    content: trimmed,
    indent: INDENTS.action,
    leadingWhitespace,
    trailingWhitespace,
    location: null,
    timeOfDay: null,
  };
}

function isSceneHeading(line) {
  return SCENE_PATTERN.test(line.trim()) || /^(INT|EXT)\b/.test(line.trim());
}

function normaliseSceneHeading(line) {
  return line.replace(/\s+/g, ' ').trim().toUpperCase();
}

function parseSceneParts(heading) {
  const withoutPrefix = heading.replace(SCENE_PATTERN, '').trim();
  const parts = withoutPrefix.split(' - ');
  const location = parts[0]?.trim() || heading;
  const timeOfDay = parts.slice(1).join(' - ').trim();
  return { location, timeOfDay };
}

function isTransition(line) {
  return TRANSITION_PATTERN.test(line.trim());
}

function normaliseTransition(line) {
  const upper = line.trim().toUpperCase();
  return upper.endsWith(':') ? upper : `${upper}:`;
}

function isSectionHeading(line) {
  return /^#/.test(line.trim());
}

function isCharacter(line) {
  if (!line) return false;
  if (line.length > 35) return false;
  if (line.includes(':')) return false;
  if (isTransition(line) || isSceneHeading(line)) return false;
  const withoutParenthetical = line.replace(/\(.*?\)/g, '');
  if (/[a-z]/.test(withoutParenthetical)) return false;
  if (!/[A-Z]/.test(withoutParenthetical)) return false;
  if (/^[0-9\W]+$/.test(withoutParenthetical)) return false;
  return true;
}

function normaliseCharacter(line) {
  return line.replace(/\s+/g, ' ').trim().toUpperCase();
}

function isParenthetical(line) {
  return line.startsWith('(');
}

function normaliseParenthetical(line) {
  let result = line.trim();
  if (!result.startsWith('(')) result = `(${result}`;
  if (!result.endsWith(')')) result = `${result})`;
  return result;
}

function mapCaretPosition(originalCaret, lineInfos, formattedText) {
  if (!lineInfos.length) return 0;
  for (let i = 0; i < lineInfos.length; i += 1) {
    const info = lineInfos[i];
    if (originalCaret <= info.endOriginal) {
      const offsetInLine = Math.max(0, originalCaret - info.startOriginal);
      if (info.type === 'empty') {
        return info.startFormatted;
      }
      if (offsetInLine <= info.leadingWhitespace) {
        return info.startFormatted + Math.min(offsetInLine, info.indentLength);
      }
      const afterIndent = offsetInLine - info.leadingWhitespace;
      const contentLength = info.formattedContentLength;
      const trailingLength = info.trailingWhitespaceLength || 0;
      if (afterIndent <= contentLength) {
        return info.startFormatted + info.indentLength + Math.min(afterIndent, contentLength);
      }
      const trailingOffset = Math.min(afterIndent - contentLength, trailingLength);
      return info.startFormatted + info.indentLength + contentLength + trailingOffset;
    }
    if (originalCaret === info.endOriginal + 1 && info.hasBreak) {
      return info.endFormatted + 1;
    }
  }
  return formattedText.length;
}

function updatePreview(lineInfos) {
  previewContainer.innerHTML = '';
  previewState.lineToPage = new Array(lineInfos.length);
  previewState.pageCount = 1;

  const fragment = document.createDocumentFragment();
  let currentPageNumber = 1;
  let currentPage = createPage(currentPageNumber);
  fragment.appendChild(currentPage);
  let countableLines = 0;

  if (!lineInfos.length) {
    const placeholder = document.createElement('div');
    placeholder.className = 'script-line action';
    placeholder.textContent = 'Start writing to see formatted pages here.';
    currentPage.appendChild(placeholder);
    previewState.pageCount = currentPageNumber;
    previewContainer.appendChild(fragment);
    updatePageIndicator(null);
    return;
  }

  lineInfos.forEach((info) => {
    const shouldCount = info.type !== 'empty';
    if (shouldCount && countableLines >= LINES_PER_PAGE) {
      currentPageNumber += 1;
      currentPage = createPage(currentPageNumber);
      fragment.appendChild(currentPage);
      countableLines = 0;
    }

    const lineEl = document.createElement('div');
    lineEl.className = `script-line ${info.type}`;
    lineEl.dataset.lineIndex = String(info.index);
    if (info.content) {
      lineEl.textContent = info.content;
    } else {
      lineEl.innerHTML = '&nbsp;';
    }
    currentPage.appendChild(lineEl);
    previewState.lineToPage[info.index] = currentPageNumber;

    if (shouldCount) {
      countableLines += 1;
    }
  });

  previewState.pageCount = currentPageNumber;
  previewContainer.appendChild(fragment);
}

function createPage(pageNumber) {
  const page = document.createElement('div');
  page.className = 'script-page';
  page.dataset.pageNumber = String(pageNumber);
  page.setAttribute('aria-label', `Page ${pageNumber}`);
  return page;
}

function updateOutline(lineInfos) {
  sceneList.innerHTML = '';
  const template = document.getElementById('sceneListItemTemplate');
  const fragment = document.createDocumentFragment();
  const scenes = lineInfos.filter((line) => line.type === 'scene');

  scenes.forEach((scene, index) => {
    const clone = template.content.cloneNode(true);
    const button = clone.querySelector('.scene-link');
    const numberEl = clone.querySelector('.scene-number');
    const titleEl = clone.querySelector('.scene-title');
    const metaEl = clone.querySelector('.scene-meta');

    numberEl.textContent = `#${index + 1}`;
    titleEl.textContent = scene.location || scene.content;
    const time = scene.timeOfDay ? ` · ${scene.timeOfDay}` : '';
    const page = getPageForLine(scene.index);
    metaEl.textContent = `Page ${NUMBER_FORMATTER.format(page)}${time}`;

    button.dataset.lineIndex = String(scene.index);
    fragment.appendChild(clone);
  });

  sceneList.appendChild(fragment);
  const label = formatCount(scenes.length, 'scene', 'scenes');
  sceneCountLabel.textContent = label;
}

function updateStats(lineInfos) {
  let wordCount = 0;
  let nonEmptyLines = 0;
  lineInfos.forEach((info) => {
    if (info.type !== 'empty') {
      nonEmptyLines += 1;
    }
    wordCount += countWords(info.content);
  });

  const scenes = lineInfos.filter((line) => line.type === 'scene').length;
  const fallbackPages = Math.max(1, Math.ceil(nonEmptyLines / LINES_PER_PAGE));
  const pageCount = Math.max(1, previewState.pageCount || fallbackPages);

  wordStat.textContent = formatCount(wordCount, 'word', 'words');
  sceneStat.textContent = formatCount(scenes, 'scene', 'scenes');
  runtimeStat.textContent = `Runtime: ${formatCount(pageCount, 'min', 'mins')}`;
  updateProgressUI(wordCount, pageCount);
}

function updateProgressUI(wordCount, pageCount) {
  if (progressTrack) {
    const value = wordCount ? Math.min(pageCount, PROGRESS_PAGE_GOAL) : 0;
    progressTrack.setAttribute('aria-valuenow', String(value));
  }
  if (progressValue) {
    const ratioBase = wordCount ? Math.min(pageCount / PROGRESS_PAGE_GOAL, 1) : 0;
    const ratio = Math.max(0, ratioBase);
    progressValue.style.setProperty('--progress', ratio.toFixed(4));
  }
  if (progressLabel) {
    progressLabel.textContent = buildProgressMessage(wordCount, pageCount);
  }
}

function buildProgressMessage(wordCount, pageCount) {
  if (!wordCount) {
    return 'Start charting your opening act.';
  }
  let stage;
  if (pageCount < 15) {
    stage = 'Act I foundations are taking shape';
  } else if (pageCount < 40) {
    stage = 'Build momentum toward the midpoint';
  } else if (pageCount < 70) {
    stage = 'Drive the second act forward';
  } else if (pageCount < 100) {
    stage = 'Line up the finale beats and payoff';
  } else {
    stage = 'Feature length achieved — tighten and refine';
  }
  return `${formatCount(wordCount, 'word', 'words')} captured — ${stage}.`;
}

function updateInsights(lineInfos) {
  insightsList.innerHTML = '';
  const template = document.getElementById('insightItemTemplate');
  const fragment = document.createDocumentFragment();

  const characters = new Map();
  const locations = new Map();
  let totalLines = 0;
  let dialogueLines = 0;
  let actionLines = 0;
  let currentScene = null;
  let currentSceneLength = 0;
  let longestScene = { title: '', length: 0 };

  lineInfos.forEach((info) => {
    if (info.type !== 'empty') {
      totalLines += 1;
    }
    if (info.type === 'dialogue') {
      dialogueLines += 1;
    }
    if (info.type === 'action') {
      actionLines += 1;
    }
    if (info.type === 'character') {
      const key = info.content;
      characters.set(key, (characters.get(key) || 0) + 1);
    }
    if (info.type === 'scene') {
      if (currentScene && currentSceneLength > longestScene.length) {
        longestScene = { title: currentScene, length: currentSceneLength };
      }
      currentScene = info.content;
      currentSceneLength = 0;
      if (info.location) {
        locations.set(info.location, (locations.get(info.location) || 0) + 1);
      }
    } else if (info.type !== 'empty' && currentScene) {
      currentSceneLength += 1;
    }
  });

  if (currentScene && currentSceneLength > longestScene.length) {
    longestScene = { title: currentScene, length: currentSceneLength };
  }

  const dialoguePercentage = totalLines ? Math.round((dialogueLines / totalLines) * 100) : 0;
  const actionPercentage = totalLines ? Math.round((actionLines / totalLines) * 100) : 0;
  const topLocation = Array.from(locations.entries()).sort((a, b) => b[1] - a[1])[0];

  const insights = [
    {
      label: 'Dialogue vs action',
      value: `${dialoguePercentage}% dialogue / ${actionPercentage}% action`,
    },
    {
      label: 'Unique characters',
      value: characters.size ? NUMBER_FORMATTER.format(characters.size) : '—',
    },
    {
      label: 'Most used location',
      value: topLocation ? `${topLocation[0]} (${formatCount(topLocation[1], 'scene', 'scenes')})` : '—',
    },
  ];

  if (longestScene.title) {
    insights.push({
      label: 'Longest scene',
      value: `${shortenSceneTitle(longestScene.title)} (${formatCount(longestScene.length, 'line', 'lines')})`,
    });
  }

  insights.slice(0, 4).forEach((insight) => {
    const clone = template.content.cloneNode(true);
    clone.querySelector('.insight-label').textContent = insight.label;
    clone.querySelector('.insight-value').textContent = insight.value;
    fragment.appendChild(clone);
  });

  insightsList.appendChild(fragment);
}

function shortenSceneTitle(title) {
  return title.replace(/\s+-\s+.*$/, '');
}

function highlightFromCaret(caret, options = {}) {
  const { scrollPreview = false } = options;
  const lineInfos = formattingResult.lineInfos;
  if (!lineInfos.length) {
    updatePageIndicator(null);
    return;
  }
  const lineIndex = getLineIndexFromCaret(caret, lineInfos);
  updatePageIndicator(lineIndex);
  highlightPreviewLine(lineIndex, { scroll: scrollPreview });
  highlightOutlineScene(lineIndex, lineInfos);
}

function getLineIndexFromCaret(caret, lineInfos) {
  for (let i = 0; i < lineInfos.length; i += 1) {
    const info = lineInfos[i];
    if (caret <= info.endFormattedWithBreak || i === lineInfos.length - 1) {
      return i;
    }
  }
  return lineInfos.length - 1;
}

function highlightPreviewLine(lineIndex, options = {}) {
  const { scroll = false } = options;
  previewContainer.querySelectorAll('.script-line.active').forEach((node) => node.classList.remove('active'));
  const target = previewContainer.querySelector(`.script-line[data-line-index="${lineIndex}"]`);
  if (target) {
    target.classList.add('active');
    if (scroll) {
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  }
}

function highlightOutlineScene(lineIndex, lineInfos) {
  const sceneButtons = sceneList.querySelectorAll('.scene-link');
  sceneButtons.forEach((button) => button.classList.remove('active'));

  const sceneIndex = findSceneIndexForLine(lineIndex, lineInfos);
  if (sceneIndex === -1) return;
  const button = sceneButtons[sceneIndex];
  if (button) {
    button.classList.add('active');
  }
}

function updatePageIndicator(lineIndex) {
  const totalPages = Math.max(1, previewState.pageCount || 1);
  const safeIndex = Number.isInteger(lineIndex) && lineIndex >= 0 ? lineIndex : 0;
  const currentPage = getPageForLine(safeIndex);
  pageIndicator.textContent = `Page ${NUMBER_FORMATTER.format(currentPage)} of ${NUMBER_FORMATTER.format(totalPages)}`;
}

function getPageForLine(lineIndex) {
  if (!Number.isInteger(lineIndex) || lineIndex < 0) {
    return 1;
  }
  const mapped = previewState.lineToPage[lineIndex];
  const fallback = Math.max(1, Math.ceil((lineIndex + 1) / LINES_PER_PAGE));
  const page = Number.isInteger(mapped) && mapped > 0 ? mapped : fallback;
  if (Number.isInteger(previewState.pageCount) && previewState.pageCount > 0) {
    return Math.min(previewState.pageCount, page);
  }
  return page;
}

function findSceneIndexForLine(lineIndex, lineInfos) {
  let sceneCounter = -1;
  for (let i = 0; i <= lineIndex && i < lineInfos.length; i += 1) {
    if (lineInfos[i].type === 'scene') {
      sceneCounter += 1;
    }
  }
  return sceneCounter;
}

function countWords(text) {
  if (!text) return 0;
  const matches = text.trim().match(/\b\w+\b/g);
  return matches ? matches.length : 0;
}

function formatCount(count, singular, plural) {
  const unit = count === 1 ? singular : plural;
  return `${NUMBER_FORMATTER.format(count)} ${unit}`;
}

function scheduleAutosave() {
  markDirty();
  if (!currentDocumentId) return;
  if (autosaveTimer) {
    clearTimeout(autosaveTimer);
  }
  autosaveTimer = setTimeout(performAutosave, AUTOSAVE_DELAY);
}

function markDirty() {
  if (isDirty) return;
  isDirty = true;
  updateAutosaveStatus();
}

function performAutosave() {
  if (!currentDocumentId) return;
  const now = new Date();
  const updates = {
    title: titleInput.value,
    content: formattingResult.formattedText,
    beats: beatNotes.value,
    updatedAt: now.toISOString(),
  };
  try {
    upsertDocument(currentDocumentId, updates);
    lastSavedAt = now;
    isDirty = false;
    updateAutosaveStatus();
    renderDocumentLibrary();
  } catch (error) {
    console.warn('Autosave failed', error);
  }
}

function updateAutosaveStatus() {
  let message = '';
  let state = 'idle';

  if (!lastSavedAt) {
    if (isDirty) {
      message = 'Unsaved changes';
      state = 'dirty';
    } else {
      message = 'Draft not saved';
      state = 'idle';
    }
  } else if (isDirty) {
    message = 'Unsaved changes';
    state = 'dirty';
  } else {
    const diffMs = Date.now() - lastSavedAt.getTime();
    if (diffMs < 2000) {
      message = 'Saved just now';
      state = 'fresh';
    } else if (diffMs < 60000) {
      message = `Saved ${Math.round(diffMs / 1000)}s ago`;
      state = 'fresh';
    } else if (diffMs < 3600000) {
      message = `Saved ${Math.round(diffMs / 60000)}m ago`;
      state = 'stale';
    } else {
      message = `Saved ${lastSavedAt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
      state = 'stale';
    }
  }

  autosaveStatus.textContent = message;
  if (autosaveIndicator) {
    autosaveIndicator.dataset.state = state;
  }
}

function bootstrapDocuments() {
  documents = loadDocumentsFromStorage();
  if (!documents.length) {
    createDocument();
    documents = loadDocumentsFromStorage();
  }
  const storedActiveId = localStorage.getItem(ACTIVE_DOCUMENT_KEY);
  const target = documents.find((doc) => doc.id === storedActiveId) || documents[0];
  loadDocument(target, { preserveCaret: false });
}

function loadDocumentsFromStorage() {
  try {
    const stored = localStorage.getItem(DOCUMENTS_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed)) {
        return parsed
          .filter((doc) => doc && (doc.id || doc.createdAt))
          .map((doc) => ({
            id: doc.id || generateDocumentId(),
            title: doc.title || '',
            content: (doc.content || '').replace(/\r\n?/g, '\n'),
            beats: doc.beats || '',
            createdAt: doc.createdAt || doc.updatedAt || new Date().toISOString(),
            updatedAt: doc.updatedAt || doc.createdAt || new Date().toISOString(),
          }));
      }
    }
  } catch (error) {
    console.warn('Failed to load document library', error);
  }

  try {
    const legacy = localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!legacy) return [];
    const payload = JSON.parse(legacy);
    const timestamp = payload.updatedAt || new Date().toISOString();
    const migrated = {
      id: generateDocumentId(),
      title: payload.title || '',
      content: (payload.content || '').replace(/\r\n?/g, '\n'),
      beats: payload.beats || '',
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    localStorage.removeItem(LEGACY_STORAGE_KEY);
    localStorage.setItem(DOCUMENTS_KEY, JSON.stringify([migrated]));
    return [migrated];
  } catch (error) {
    console.warn('Failed to migrate legacy draft', error);
  }

  return [];
}

function persistDocuments() {
  try {
    localStorage.setItem(DOCUMENTS_KEY, JSON.stringify(documents));
  } catch (error) {
    console.warn('Failed to store documents', error);
  }
}

function upsertDocument(id, updates) {
  const index = documents.findIndex((doc) => doc.id === id);
  if (index === -1) return;
  documents[index] = {
    ...documents[index],
    ...updates,
  };
  persistDocuments();
  localStorage.setItem(ACTIVE_DOCUMENT_KEY, id);
}

function createDocument(initial = {}) {
  const nowIso = new Date().toISOString();
  const doc = {
    id: generateDocumentId(),
    title: initial.title || '',
    content: (initial.content || '').replace(/\r\n?/g, '\n'),
    beats: initial.beats || '',
    createdAt: initial.createdAt || nowIso,
    updatedAt: initial.updatedAt || nowIso,
  };
  documents.push(doc);
  persistDocuments();
  renderDocumentLibrary();
  return doc;
}

function loadDocument(doc, options = {}) {
  if (!doc) return;
  const { preserveCaret = false } = options;
  currentDocumentId = doc.id;
  localStorage.setItem(ACTIVE_DOCUMENT_KEY, currentDocumentId);
  titleInput.value = doc.title || '';
  editor.value = (doc.content || '').replace(/\r\n?/g, '\n');
  beatNotes.value = doc.beats || '';
  lastSavedAt = doc.updatedAt ? new Date(doc.updatedAt) : null;
  isDirty = false;
  resetShareLink();
  updateAutosaveStatus();
  const shouldSeed = !doc.content && documents.length === 1 && doc.createdAt === doc.updatedAt;
  if (shouldSeed) {
    seedWelcomePage();
  }
  applyFormattingAndRefresh({ preserveCaret });
  renderDocumentLibrary();
}

function resetShareLink() {
  shareLinkInput.value = '';
  shareModalLink.value = '';
}

function seedWelcomePage() {
  const template = `INT. DREAMSCAPE - NIGHT

A swirling cosmic void. Possibilities everywhere.

NARRATOR
Welcome to Scriptius. Start writing your story...
`;
  editor.value = template;
}

function loadFromShareLink() {
  const hash = window.location.hash || '';
  if (!hash.startsWith('#share=')) return;
  const encoded = hash.slice('#share='.length);
  try {
    const payload = decodeSharePayload(encoded);
    if (payload?.content) {
      const doc = createDocument({
        title: payload.title || 'Shared draft',
        content: payload.content,
        beats: payload.beats || '',
        updatedAt: new Date().toISOString(),
      });
      loadDocument(doc, { preserveCaret: false });
      showSnackbar('Loaded shared draft');
      history.replaceState(null, document.title, window.location.pathname + window.location.search);
    }
  } catch (error) {
    console.error('Failed to load shared draft', error);
    showSnackbar('Unable to load shared script');
  }
}

function handleNewDocument() {
  if (isDirty && editor.value.trim()) {
    const proceed = window.confirm('Start a new script? Unsaved changes will be lost.');
    if (!proceed) return;
  }
  const doc = createDocument();
  loadDocument(doc, { preserveCaret: false });
  editor.focus();
  showSnackbar('Started a new script');
}

function handleImportFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    editor.value = String(reader.result).replace(/\r\n?/g, '\n');
    titleInput.value = file.name.replace(/\.[^.]+$/, '');
    applyFormattingAndRefresh({ preserveCaret: false });
    showSnackbar('Imported draft');
    filePicker.value = '';
  };
  reader.onerror = () => {
    console.error('Import failed', reader.error);
    showSnackbar('Failed to import file');
  };
  reader.readAsText(file);
}

function createFilename(extension) {
  const base = titleInput.value.trim() || 'Untitled Script';
  return `${base.replace(/[^a-z0-9]+/gi, '_')}${extension}`;
}

function downloadFile(filename, content) {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function openShareModal() {
  const link = ensureShareLink();
  shareModalLink.value = link;
  shareLinkInput.value = link;
  shareModal.classList.add('open');
  shareModalLink.select();
}

function closeShareModal() {
  shareModal.classList.remove('open');
}

function ensureShareLink() {
  const link = shareLinkInput.value || generateShareLink();
  shareLinkInput.value = link;
  shareModalLink.value = link;
  return link;
}

function generateShareLink() {
  const payload = {
    title: titleInput.value,
    content: formattingResult.formattedText,
    beats: beatNotes.value,
  };
  const encoded = encodeSharePayload(payload);
  const base = window.location.href.split('#')[0];
  return `${base}#share=${encoded}`;
}

function encodeSharePayload(payload) {
  const json = JSON.stringify(payload);
  const bytes = new TextEncoder().encode(json);
  let binary = '';
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function decodeSharePayload(encoded) {
  const binary = atob(encoded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const json = new TextDecoder().decode(bytes);
  return JSON.parse(json);
}

function generateDocumentId() {
  const cryptoObj = globalThis.crypto;
  if (cryptoObj?.randomUUID) {
    return cryptoObj.randomUUID();
  }
  return `doc-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

async function handleCopyShareLink() {
  const link = ensureShareLink();
  try {
    await navigator.clipboard.writeText(link);
    showSnackbar('Share link copied');
  } catch (error) {
    console.warn('Clipboard failed', error);
    showSnackbar('Press Ctrl+C to copy');
    shareModalLink.select();
  }
}

function showSnackbar(message) {
  snackbar.textContent = message;
  snackbar.classList.add('show');
  setTimeout(() => snackbar.classList.remove('show'), 2400);
}

function toggleTheme() {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  setTheme(next);
}

function applyStoredTheme() {
  const stored = localStorage.getItem(THEME_KEY);
  if (stored) {
    setTheme(stored);
  }
}

function setTheme(theme) {
  document.documentElement.dataset.theme = theme;
  localStorage.setItem(THEME_KEY, theme);
  themeToggle.textContent = theme === 'dark' ? '☀️' : '🌗';
}

function focusLine(lineIndex) {
  const info = formattingResult.lineInfos[lineIndex];
  if (!info) return;
  const caret = info.startFormatted;
  editor.focus();
  editor.setSelectionRange(caret, caret);
  highlightFromCaret(caret, { scrollPreview: true });
}

window.addEventListener('beforeunload', (event) => {
  if (!isDirty) return;
  event.preventDefault();
  event.returnValue = '';
});
