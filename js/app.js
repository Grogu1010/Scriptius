const editor = document.getElementById('scriptEditor');
const writingStage = document.querySelector('.writing-stage');
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

const suggestionPanel = document.getElementById('suggestionPanel');
const suggestionLabel = document.getElementById('suggestionLabel');
const suggestionList = document.getElementById('suggestionList');

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
const writingPanel = document.querySelector('.writing-panel');
const viewModeButtons = document.querySelectorAll('.view-switch__btn[data-view-mode]');
const storyDirectionBtn = document.getElementById('storyDirectionBtn');
const storyDirectionOutput = document.getElementById('storyDirectionOutput');
const storyDirectionStatus = document.getElementById('storyDirectionStatus');
const storyDirectionText = document.getElementById('storyDirectionText');
const storyDirectionActions = document.getElementById('storyDirectionActions');
const storyDirectionImplementBtn = document.getElementById('storyDirectionImplementBtn');
const storyDirectionApproveBtn = document.getElementById('storyDirectionApproveBtn');
const storyDirectionRejectBtn = document.getElementById('storyDirectionRejectBtn');
const storyDirectionFeedbackStatus = document.getElementById('storyDirectionFeedbackStatus');
const STORY_DIRECTION_DEFAULT_LABEL = storyDirectionBtn?.textContent?.trim() || 'Story direction';

const DOCUMENTS_KEY = 'scriptius-documents';
const ACTIVE_DOCUMENT_KEY = 'scriptius-active-document';
const LEGACY_STORAGE_KEY = 'scriptius-document';
const THEME_KEY = 'scriptius-theme';
const LAYOUT_KEY = 'scriptius-layout';
const VIEW_MODE_KEY = 'scriptius-view-mode';
const AUTOSAVE_DELAY = 800;
const LINES_PER_PAGE = 55;
const PROGRESS_PAGE_GOAL = 110;
const NUMBER_FORMATTER = new Intl.NumberFormat();

const SCENE_PREFIX_SUGGESTIONS = Object.freeze([
  'INT.',
  'EXT.',
  'INT./EXT.',
  'EXT./INT.',
  'I/E.',
  'EST.',
  'INT & EXT.',
]);

const CHARACTER_QUALIFIER_SUGGESTIONS = Object.freeze([
  "CONT'D",
  'V.O.',
  'O.S.',
  'O.C.',
  'INTO PHONE',
  'ON RADIO',
  'FILTERED',
  'SOTTO',
]);

const TIME_OF_DAY_SUGGESTIONS = Object.freeze([
  'DAY',
  'NIGHT',
  'MORNING',
  'AFTERNOON',
  'EVENING',
  'LATER',
  'CONTINUOUS',
  'SAME',
  'DAWN',
  'DUSK',
]);

const MAX_SUGGESTIONS = 6;

const suggestionData = {
  characterCounts: new Map(),
  qualifierCounts: new Map(),
  locationData: new Map(),
  timeCounts: new Map(),
  scenePrefixes: new Map(),
};

const suggestionState = {
  items: [],
  activeIndex: -1,
  context: null,
};

const CARET_MEASURE_PROPERTIES = Object.freeze([
  'direction',
  'boxSizing',
  'width',
  'height',
  'overflowX',
  'overflowY',
  'borderTopWidth',
  'borderRightWidth',
  'borderBottomWidth',
  'borderLeftWidth',
  'borderStyle',
  'paddingTop',
  'paddingRight',
  'paddingBottom',
  'paddingLeft',
  'fontStyle',
  'fontVariant',
  'fontWeight',
  'fontStretch',
  'fontSize',
  'fontSizeAdjust',
  'lineHeight',
  'fontFamily',
  'textAlign',
  'textTransform',
  'textIndent',
  'textDecoration',
  'letterSpacing',
  'wordSpacing',
  'tabSize',
  'MozTabSize',
]);

let caretMirrorDiv = null;
let caretMirrorSpan = null;
let suggestionPositionFrame = null;

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
let currentViewMode = writingPanel?.dataset.viewMode || 'editor';
const previewState = {
  lineToPage: [],
  pageCount: 1,
};
let storageError = false;
let storageErrorNotified = false;
let storyDirectionAbortController = null;
let lastStoryDirectionIdea = '';
let storyDirectionFeedbackSelection = null;

init();

function init() {
  applyStoredTheme();
  applyStoredLayout();
  applyStoredViewMode();
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
  editor.addEventListener('blur', closeSuggestionPanel);
  editor.addEventListener('focus', () => {
    requestAnimationFrame(updateSuggestionPanel);
  });
  editor.addEventListener('scroll', handleEditorScroll);

  document.addEventListener('selectionchange', () => {
    if (document.activeElement === editor) {
      handleCaretChange();
    }
  });

  window.addEventListener('resize', handleWindowResize);

  if (suggestionList) {
    suggestionList.addEventListener('mousedown', handleSuggestionPointerDown);
    suggestionList.addEventListener('mouseover', handleSuggestionHover);
  }

  titleInput.addEventListener('input', () => {
    markDirty();
    scheduleAutosave();
  });

  beatNotes.addEventListener('input', () => {
    markDirty();
    scheduleAutosave();
  });

  if (storyDirectionBtn) {
    storyDirectionBtn.addEventListener('click', handleStoryDirectionRequest);
  }

  if (storyDirectionImplementBtn) {
    storyDirectionImplementBtn.addEventListener('click', handleStoryDirectionImplement);
  }

  if (storyDirectionApproveBtn) {
    storyDirectionApproveBtn.addEventListener('click', handleStoryDirectionFeedback);
  }

  if (storyDirectionRejectBtn) {
    storyDirectionRejectBtn.addEventListener('click', handleStoryDirectionFeedback);
  }

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

  viewModeButtons.forEach((button) => {
    button.addEventListener('click', () => {
      setViewMode(button.dataset.viewMode);
    });
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
  if (handleSuggestionKeydown(event)) {
    return;
  }

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
  updateSuggestionPanel();
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
      return 'action';
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

function updateSuggestionPanel() {
  if (!suggestionPanel || !suggestionList) return;
  if (currentViewMode !== 'editor') {
    closeSuggestionPanel();
    return;
  }
  if (document.activeElement !== editor) {
    closeSuggestionPanel();
    return;
  }
  const caret = editor.selectionStart;
  const lineInfo = findLineInfoForCaret(caret);
  if (!lineInfo) {
    closeSuggestionPanel();
    return;
  }
  const context = determineSuggestionContext(lineInfo, caret);
  if (!context) {
    closeSuggestionPanel();
    return;
  }
  const items = buildSuggestionsForContext(context);
  if (!items.length) {
    closeSuggestionPanel();
    return;
  }

  const previous = suggestionState.context;
  const contextChanged =
    !previous ||
    previous.type !== context.type ||
    previous.mode !== context.mode ||
    previous.lineIndex !== context.lineIndex;

  suggestionState.context = context;
  suggestionState.items = items;
  if (contextChanged || suggestionState.activeIndex < 0 || suggestionState.activeIndex >= items.length) {
    suggestionState.activeIndex = 0;
  }

  renderSuggestionList(items, context);
}

function buildSuggestionsForContext(context) {
  if (!context) return [];
  if (context.type === 'character') {
    if (context.mode === 'qualifier') {
      return getQualifierSuggestions(context.prefix, context.currentValue);
    }
    return getCharacterNameSuggestions(context.prefix, context.currentValue);
  }
  if (context.type === 'scene') {
    if (context.mode === 'prefix') {
      return getScenePrefixSuggestions(context.prefix, context.currentValue);
    }
    if (context.mode === 'location') {
      return getSceneLocationSuggestions(context.prefix, context.currentValue);
    }
    if (context.mode === 'time') {
      return getSceneTimeSuggestions(context.prefix, context.location, context.currentValue);
    }
  }
  return [];
}

function getSuggestionLabel(context) {
  if (!context) return '';
  if (context.type === 'character') {
    return context.mode === 'qualifier' ? 'Character notations' : 'Character names';
  }
  if (context.type === 'scene') {
    if (context.mode === 'prefix') return 'Scene prefixes';
    if (context.mode === 'location') return 'Scene locations';
    if (context.mode === 'time') return 'Time of day';
  }
  return 'Suggestions';
}

function renderSuggestionList(items, context) {
  if (!suggestionPanel || !suggestionList || !suggestionLabel) return;
  const label = getSuggestionLabel(context);
  if (label) {
    suggestionLabel.textContent = label;
    suggestionLabel.hidden = false;
    suggestionList.setAttribute('aria-label', label);
  } else {
    suggestionLabel.textContent = '';
    suggestionLabel.hidden = true;
    suggestionList.removeAttribute('aria-label');
  }

  suggestionPanel.hidden = false;
  suggestionList.innerHTML = '';
  const fragment = document.createDocumentFragment();
  items.forEach((item, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'suggestion-option';
    button.dataset.index = String(index);
    button.dataset.kind = item.kind;
    button.id = `suggestion-option-${index}`;
    button.setAttribute('role', 'option');
    button.textContent = item.label;
    fragment.appendChild(button);
  });
  suggestionList.appendChild(fragment);
  setActiveSuggestion(suggestionState.activeIndex);
  scheduleSuggestionPositionUpdate();
}

function setActiveSuggestion(index) {
  if (!suggestionList) return;
  const options = suggestionList.querySelectorAll('.suggestion-option');
  if (!options.length) {
    suggestionState.activeIndex = -1;
    suggestionList.removeAttribute('aria-activedescendant');
    return;
  }

  let nextIndex = Number(index);
  if (Number.isNaN(nextIndex) || nextIndex < 0) {
    nextIndex = 0;
  }
  if (nextIndex >= options.length) {
    nextIndex = nextIndex % options.length;
  }

  suggestionState.activeIndex = nextIndex;
  options.forEach((option, idx) => {
    const isActive = idx === nextIndex;
    option.classList.toggle('is-active', isActive);
    option.setAttribute('aria-selected', String(isActive));
    if (isActive) {
      suggestionList.setAttribute('aria-activedescendant', option.id);
      option.scrollIntoView({ block: 'nearest' });
    }
  });
}

function scheduleSuggestionPositionUpdate() {
  if (!suggestionPanel || suggestionPanel.hidden) return;
  if (suggestionPositionFrame) {
    cancelAnimationFrame(suggestionPositionFrame);
  }
  suggestionPositionFrame = requestAnimationFrame(() => {
    suggestionPositionFrame = null;
    positionSuggestionPanel();
  });
}

function positionSuggestionPanel() {
  if (!suggestionPanel || suggestionPanel.hidden) return;
  const anchor = getSuggestionAnchorRect();
  if (!anchor) return;

  const { top, left, height, containerRect } = anchor;
  const edgePadding = 12;
  const containerHeight = containerRect?.height ?? 0;
  const containerWidth = containerRect?.width ?? 0;
  const panelRect = suggestionPanel.getBoundingClientRect();
  const panelWidth = panelRect.width || suggestionPanel.offsetWidth || 0;
  const panelHeight = panelRect.height || suggestionPanel.offsetHeight || 0;

  let computedLeft = left;
  if (containerWidth > 0 && panelWidth > 0) {
    const maxLeft = Math.max(edgePadding, containerWidth - edgePadding - panelWidth);
    const minLeft = edgePadding;
    computedLeft = Math.min(Math.max(computedLeft, minLeft), maxLeft);
  } else {
    computedLeft = Math.max(edgePadding, computedLeft);
  }

  let computedTop = top + height + 8;
  if (containerHeight > 0 && panelHeight > 0) {
    const maxTop = Math.max(edgePadding, containerHeight - edgePadding - panelHeight);
    if (computedTop > maxTop && top - panelHeight - 8 >= edgePadding) {
      computedTop = top - panelHeight - 8;
    }
    computedTop = Math.min(Math.max(computedTop, edgePadding), maxTop);
  } else {
    computedTop = Math.max(edgePadding, computedTop);
  }

  suggestionPanel.style.top = `${computedTop}px`;
  suggestionPanel.style.left = `${computedLeft}px`;
}

function getSuggestionAnchorRect() {
  if (!editor) return null;
  const caret = editor.selectionStart;
  if (!Number.isInteger(caret)) return null;
  const metrics = measureEditorCaret(caret);
  if (!metrics) return null;

  const container =
    suggestionPanel?.offsetParent || writingStage || editor.parentElement || document.body;
  if (!container) return null;

  const containerRect = container.getBoundingClientRect();
  const editorRect = editor.getBoundingClientRect();
  const relativeTop = editorRect.top - containerRect.top + metrics.top - editor.scrollTop;
  const relativeLeft = editorRect.left - containerRect.left + metrics.left - editor.scrollLeft;

  return {
    top: relativeTop,
    left: relativeLeft,
    height: metrics.height,
    containerRect,
  };
}

function measureEditorCaret(position) {
  if (!editor || !Number.isInteger(position)) return null;
  const doc = editor.ownerDocument;
  if (!doc) return null;
  const view = doc.defaultView || window;
  const computed = view.getComputedStyle(editor);

  if (!caretMirrorDiv) {
    caretMirrorDiv = doc.createElement('div');
    caretMirrorDiv.setAttribute('aria-hidden', 'true');
  }
  if (!caretMirrorSpan) {
    caretMirrorSpan = doc.createElement('span');
  }

  const div = caretMirrorDiv;
  const span = caretMirrorSpan;
  const style = div.style;

  style.whiteSpace = 'pre-wrap';
  style.wordWrap = 'break-word';
  style.position = 'absolute';
  style.visibility = 'hidden';
  style.top = '0';
  style.left = '-9999px';

  CARET_MEASURE_PROPERTIES.forEach((prop) => {
    if (prop === 'lineHeight' && editor.nodeName === 'INPUT') {
      style.lineHeight = computed.height;
    } else {
      style[prop] = computed[prop];
    }
  });

  const value = editor.value || '';
  div.textContent = value.slice(0, position);
  span.textContent = value.slice(position) || '.';
  div.appendChild(span);
  doc.body.appendChild(div);

  const borderTop = parseInt(computed.borderTopWidth, 10) || 0;
  const borderLeft = parseInt(computed.borderLeftWidth, 10) || 0;
  const caretTop = span.offsetTop + borderTop;
  const caretLeft = span.offsetLeft + borderLeft;
  let caretHeight = parseInt(computed.lineHeight, 10);
  if (!Number.isFinite(caretHeight) || caretHeight <= 0) {
    const fontSize = parseInt(computed.fontSize, 10) || 0;
    caretHeight = span.offsetHeight || fontSize || 0;
  }

  doc.body.removeChild(div);

  return {
    top: caretTop,
    left: caretLeft,
    height: caretHeight,
  };
}

function moveActiveSuggestion(delta) {
  if (!suggestionState.items.length) return;
  const length = suggestionState.items.length;
  const nextIndex = (suggestionState.activeIndex + delta + length) % length;
  setActiveSuggestion(nextIndex);
}

function applyActiveSuggestion() {
  const index = suggestionState.activeIndex;
  if (index < 0 || index >= suggestionState.items.length) return;
  const item = suggestionState.items[index];
  applySuggestion(item);
}

function applySuggestion(item) {
  if (!item) return;
  const caret = editor.selectionStart;
  const lineInfo = findLineInfoForCaret(caret);
  if (!lineInfo) return;
  const context = determineSuggestionContext(lineInfo, caret);
  if (!context || !suggestionState.context) return;
  if (context.type !== suggestionState.context.type || context.mode !== suggestionState.context.mode) {
    return;
  }

  const lineStart = lineInfo.startFormatted;
  const indentLength = lineInfo.indentLength || 0;
  const lineText = editor.value.slice(lineStart, lineInfo.endFormatted);
  const content = lineText.slice(indentLength);
  const caretInContent = Math.max(0, Math.min(caret - lineStart - indentLength, content.length));

  let rangeStart = null;
  let rangeEnd = null;
  let replacement = '';

  if (context.type === 'character') {
    if (context.mode === 'name') {
      const baseEndRaw = context.baseEnd ?? content.indexOf('(');
      const baseEnd = baseEndRaw === -1 ? content.length : baseEndRaw;
      rangeStart = lineStart + indentLength;
      rangeEnd = lineStart + indentLength + baseEnd;
      replacement = item.value;
      const nextChar = content.charAt(baseEnd);
      if (nextChar === '(') {
        replacement = `${replacement} `;
      }
    } else if (context.mode === 'qualifier') {
      const openIndex = context.openParenOffset ?? content.lastIndexOf('(', caretInContent - 1);
      if (openIndex === -1) return;
      const closingIndex = content.indexOf(')', openIndex);
      rangeStart = lineStart + indentLength + openIndex;
      if (closingIndex !== -1) {
        rangeEnd = lineStart + indentLength + closingIndex + 1;
      } else {
        rangeEnd = lineStart + indentLength + caretInContent;
      }
      replacement = `(${item.value})`;
    }
  } else if (context.type === 'scene') {
    if (context.mode === 'prefix') {
      let start = lineStart + indentLength;
      while (start < lineInfo.endFormatted && editor.value[start] === ' ') {
        start += 1;
      }
      let end = start;
      while (end < lineInfo.endFormatted) {
        const ch = editor.value[end];
        if (ch === ' ' || ch === '-') {
          break;
        }
        end += 1;
      }
      while (end < lineInfo.endFormatted && editor.value[end] === ' ') {
        end += 1;
      }
      rangeStart = start;
      rangeEnd = end;
      replacement = `${item.value} `;
    } else if (context.mode === 'location') {
      const prefixEnd = context.prefixEnd || 0;
      const locationStart = context.locationStart ?? prefixEnd;
      const locationEnd = context.locationEnd ?? content.length;
      rangeStart = lineStart + indentLength + locationStart;
      rangeEnd = lineStart + indentLength + locationEnd;
      const remainder = content.slice(locationEnd);
      replacement = item.value;
      if (locationStart === prefixEnd && prefixEnd > 0) {
        replacement = ` ${replacement}`;
      }
      if (remainder.startsWith('-')) {
        replacement = `${replacement} `;
      }
    } else if (context.mode === 'time') {
      if (!Number.isInteger(context.timeStart) || context.timeStart < 0) return;
      rangeStart = lineStart + indentLength + context.timeStart;
      rangeEnd = lineStart + indentLength + content.length;
      replacement = item.value;
    }
  }

  if (rangeStart === null || rangeEnd === null) return;
  replaceEditorRange(rangeStart, rangeEnd, replacement);
}

function replaceEditorRange(start, end, replacement) {
  const text = editor.value;
  const safeStart = Math.max(0, Math.min(start, text.length));
  const safeEnd = Math.max(safeStart, Math.min(end, text.length));
  const before = text.slice(0, safeStart);
  const after = text.slice(safeEnd);
  editor.value = `${before}${replacement}${after}`;
  const caretPosition = safeStart + replacement.length;
  editor.setSelectionRange(caretPosition, caretPosition);
  closeSuggestionPanel();
  applyFormattingAndRefresh();
}

function handleSuggestionKeydown(event) {
  if (!isSuggestionOpen()) return false;
  if (event.key === 'ArrowDown') {
    event.preventDefault();
    moveActiveSuggestion(1);
    return true;
  }
  if (event.key === 'ArrowUp') {
    event.preventDefault();
    moveActiveSuggestion(-1);
    return true;
  }
  if (event.key === 'Enter') {
    if (event.shiftKey) {
      return false;
    }
    event.preventDefault();
    applyActiveSuggestion();
    return true;
  }
  if (event.key === 'Tab') {
    event.preventDefault();
    applyActiveSuggestion();
    return true;
  }
  if (event.key === 'Escape') {
    event.preventDefault();
    closeSuggestionPanel();
    return true;
  }
  return false;
}

function isSuggestionOpen() {
  return Boolean(suggestionState.items.length && suggestionPanel && !suggestionPanel.hidden);
}

function handleEditorScroll() {
  if (!isSuggestionOpen()) return;
  scheduleSuggestionPositionUpdate();
}

function handleWindowResize() {
  if (!isSuggestionOpen()) return;
  scheduleSuggestionPositionUpdate();
}

function handleSuggestionPointerDown(event) {
  const button = event.target.closest('.suggestion-option');
  if (!button) return;
  event.preventDefault();
  const index = Number(button.dataset.index);
  if (Number.isNaN(index)) return;
  setActiveSuggestion(index);
  applyActiveSuggestion();
}

function handleSuggestionHover(event) {
  const button = event.target.closest('.suggestion-option');
  if (!button) return;
  const index = Number(button.dataset.index);
  if (Number.isNaN(index) || index === suggestionState.activeIndex) return;
  setActiveSuggestion(index);
}

function closeSuggestionPanel() {
  if (!suggestionPanel || !suggestionList || !suggestionLabel) return;
  if (suggestionPositionFrame) {
    cancelAnimationFrame(suggestionPositionFrame);
    suggestionPositionFrame = null;
  }
  suggestionPanel.hidden = true;
  suggestionLabel.textContent = '';
  suggestionLabel.hidden = true;
  suggestionList.innerHTML = '';
  suggestionList.removeAttribute('aria-activedescendant');
  suggestionPanel.style.removeProperty('top');
  suggestionPanel.style.removeProperty('left');
  suggestionPanel.style.removeProperty('width');
  suggestionState.items = [];
  suggestionState.activeIndex = -1;
  suggestionState.context = null;
}

function determineSuggestionContext(lineInfo, caret) {
  if (!lineInfo) return null;
  const lineIndex = Number.isInteger(lineInfo.index) ? lineInfo.index : 0;
  const lineStart = lineInfo.startFormatted;
  const indentLength = lineInfo.indentLength || 0;
  const lineText = editor.value.slice(lineStart, lineInfo.endFormatted);
  const content = lineText.slice(indentLength);
  const uppercaseContent = content.toUpperCase();
  const caretOffset = Math.max(0, Math.min(caret - lineStart, lineText.length));
  const caretInContent = Math.max(0, Math.min(caretOffset - indentLength, uppercaseContent.length));
  const beforeCaret = content.slice(0, caretInContent);
  const uppercaseBeforeCaret = beforeCaret.toUpperCase();
  const trimmedBeforeCaret = uppercaseBeforeCaret.trim();

  const isCharacterLine = lineInfo.type === 'character' || isCharacter(content);
  if (isCharacterLine && (trimmedBeforeCaret.length > 0 || uppercaseContent.includes('('))) {
    const lastOpenParen = uppercaseBeforeCaret.lastIndexOf('(');
    const lastCloseParen = uppercaseBeforeCaret.lastIndexOf(')');
    const insideParenthesis = lastOpenParen !== -1 && (lastCloseParen === -1 || lastCloseParen < lastOpenParen);
    if (insideParenthesis) {
      const qualifierPrefix = uppercaseBeforeCaret.slice(lastOpenParen + 1).trim();
      const afterOpen = uppercaseContent.slice(lastOpenParen + 1);
      const closingIndex = afterOpen.indexOf(')');
      const currentQualifier = closingIndex !== -1 ? afterOpen.slice(0, closingIndex).trim() : qualifierPrefix;
      return {
        type: 'character',
        mode: 'qualifier',
        prefix: qualifierPrefix,
        currentValue: currentQualifier,
        lineIndex,
        indentLength,
        openParenOffset: lastOpenParen,
      };
    }
    const baseEndRaw = uppercaseContent.indexOf('(');
    const baseEnd = baseEndRaw === -1 ? uppercaseContent.length : baseEndRaw;
    const baseName = uppercaseContent.slice(0, baseEnd).trim();
    return {
      type: 'character',
      mode: 'name',
      prefix: trimmedBeforeCaret,
      currentValue: baseName,
      lineIndex,
      indentLength,
      baseEnd,
    };
  }

  const prefixMatch = uppercaseContent.match(SCENE_PATTERN);
  const hasScenePrefix = prefixMatch && prefixMatch.index === 0;
  const hyphenInfo = findSceneHyphen(uppercaseContent);
  const hyphenIndex = hyphenInfo ? hyphenInfo.index : -1;
  const hyphenLength = hyphenInfo ? hyphenInfo.length : 0;
  const timeStart = hyphenIndex !== -1 ? hyphenIndex + hyphenLength : -1;

  if (hasScenePrefix) {
    const prefixRaw = prefixMatch[0];
    const prefixEnd = prefixRaw.length;
    const prefixComplete = prefixRaw.includes('.') || prefixRaw.includes('/') || prefixRaw.includes('&');
    const caretInPrefix = caretInContent < prefixEnd;
    const caretAtEndOfPrefix = caretInContent === prefixEnd;

    let locationStart = prefixEnd;
    if (uppercaseContent.charAt(locationStart) === ' ') {
      locationStart += 1;
    }
    const locationEnd = hyphenIndex !== -1 ? hyphenIndex : uppercaseContent.length;
    const locationFull = uppercaseContent.slice(locationStart, locationEnd).trim();
    const timeFull = timeStart !== -1 ? uppercaseContent.slice(timeStart).trim() : '';

    if (caretInPrefix || (!prefixComplete && caretAtEndOfPrefix)) {
      const prefixBeforeCaret = uppercaseBeforeCaret.slice(0, Math.min(caretInContent, prefixEnd)).trim();
      const currentPrefix = normaliseScenePrefix(prefixRaw);
      return {
        type: 'scene',
        mode: 'prefix',
        prefix: prefixBeforeCaret,
        currentValue: currentPrefix,
        lineIndex,
        indentLength,
        prefixEnd,
        prefixComplete,
      };
    }

    if (timeStart !== -1 && caretInContent >= timeStart) {
      const timePrefix = uppercaseBeforeCaret.slice(timeStart, caretInContent).trim();
      return {
        type: 'scene',
        mode: 'time',
        prefix: timePrefix,
        currentValue: timeFull,
        location: locationFull,
        lineIndex,
        indentLength,
        prefixEnd,
        locationStart,
        locationEnd,
        timeStart,
        hyphenIndex,
      };
    }

    const locationPrefix = uppercaseBeforeCaret.slice(locationStart, Math.min(caretInContent, locationEnd)).trim();
    return {
      type: 'scene',
      mode: 'location',
      prefix: locationPrefix,
      currentValue: locationFull,
      lineIndex,
      indentLength,
      prefixEnd,
      locationStart,
      locationEnd,
      hyphenIndex,
    };
  }

  if (trimmedBeforeCaret.length > 0 && !trimmedBeforeCaret.includes(' ')) {
    const matchesPrefix = SCENE_PREFIX_SUGGESTIONS.some((entry) =>
      normaliseScenePrefix(entry).startsWith(trimmedBeforeCaret)
    );
    if (matchesPrefix) {
      return {
        type: 'scene',
        mode: 'prefix',
        prefix: trimmedBeforeCaret,
        currentValue: trimmedBeforeCaret,
        lineIndex,
        indentLength,
        prefixEnd: trimmedBeforeCaret.length,
        prefixComplete: false,
      };
    }
  }

  return null;
}

function refreshSuggestionData(lineInfos) {
  suggestionData.characterCounts.clear();
  suggestionData.qualifierCounts.clear();
  suggestionData.locationData.clear();
  suggestionData.timeCounts.clear();
  suggestionData.scenePrefixes.clear();

  lineInfos.forEach((info) => {
    if (info.type === 'character') {
      const parts = splitCharacterName(info.content || '');
      if (parts.baseName) {
        incrementCount(suggestionData.characterCounts, parts.baseName);
      }
      parts.qualifiers.forEach((qualifier) => {
        if (qualifier && qualifier.length <= 40) {
          incrementCount(suggestionData.qualifierCounts, qualifier);
        }
      });
    } else if (info.type === 'parenthetical') {
      const trimmed = (info.content || '').trim();
      if (trimmed) {
        const normalised = trimmed.toUpperCase();
        if (normalised.length <= 60) {
          incrementCount(suggestionData.qualifierCounts, normalised);
        }
      }
    }

    if (info.type === 'scene') {
      const prefixMatch = info.content ? info.content.match(SCENE_PATTERN) : null;
      if (prefixMatch) {
        const prefix = normaliseScenePrefix(prefixMatch[0]);
        incrementCount(suggestionData.scenePrefixes, prefix);
      }
      if (info.location) {
        const location = info.location.toUpperCase();
        let entry = suggestionData.locationData.get(location);
        if (!entry) {
          entry = { count: 0, times: new Map() };
          suggestionData.locationData.set(location, entry);
        }
        entry.count += 1;
        if (info.timeOfDay) {
          const time = info.timeOfDay.toUpperCase();
          entry.times.set(time, (entry.times.get(time) || 0) + 1);
          incrementCount(suggestionData.timeCounts, time);
        }
      } else if (info.timeOfDay) {
        const time = info.timeOfDay.toUpperCase();
        incrementCount(suggestionData.timeCounts, time);
      }
    }
  });
}

function splitCharacterName(text = '') {
  const upper = text.toUpperCase();
  const qualifiers = [];
  upper.replace(/\(([^)]+)\)/g, (match, inner) => {
    const cleaned = inner.trim();
    if (cleaned) {
      qualifiers.push(cleaned);
    }
    return '';
  });
  const baseName = upper.replace(/\([^)]*\)/g, '').trim();
  return { baseName, qualifiers };
}

function normaliseScenePrefix(prefix) {
  if (!prefix) return '';
  return prefix
    .toUpperCase()
    .replace(/\s*\/\s*/g, '/')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findSceneHyphen(text) {
  if (!text) return null;
  let index = text.indexOf(' - ');
  if (index !== -1) {
    return { index, length: 3 };
  }
  index = text.indexOf(' -');
  if (index !== -1) {
    return { index, length: 2 };
  }
  index = text.indexOf('- ');
  if (index !== -1) {
    return { index, length: 2 };
  }
  return null;
}

function incrementCount(map, key) {
  if (!key) return;
  map.set(key, (map.get(key) || 0) + 1);
}

function getCharacterNameSuggestions(prefix, currentValue) {
  const upperPrefix = (prefix || '').toUpperCase();
  const current = (currentValue || '').toUpperCase();
  const entries = Array.from(suggestionData.characterCounts.entries()).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });
  const suggestions = [];
  const seen = new Set();
  const filtered = entries.filter(([name]) => !upperPrefix || name.startsWith(upperPrefix));
  const source = filtered.length ? filtered : entries;
  source.forEach(([name]) => {
    if (name === current) return;
    addSuggestionOption(suggestions, seen, name, 'characterName');
  });
  return suggestions.slice(0, MAX_SUGGESTIONS);
}

function getQualifierSuggestions(prefix, currentValue) {
  const upperPrefix = (prefix || '').toUpperCase();
  const current = (currentValue || '').toUpperCase();
  const suggestions = [];
  const seen = new Set();
  const entries = Array.from(suggestionData.qualifierCounts.entries()).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });
  entries.forEach(([qualifier]) => {
    if (qualifier === current) return;
    if (upperPrefix && !qualifier.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, qualifier, 'characterQualifier');
  });
  CHARACTER_QUALIFIER_SUGGESTIONS.forEach((qualifier) => {
    const value = qualifier.toUpperCase();
    if (value === current) return;
    if (upperPrefix && !value.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, value, 'characterQualifier');
  });
  return suggestions.slice(0, MAX_SUGGESTIONS);
}

function getScenePrefixSuggestions(prefix, currentValue) {
  const upperPrefix = (prefix || '').toUpperCase();
  const current = normaliseScenePrefix(currentValue || '');
  const suggestions = [];
  const seen = new Set();
  const entries = Array.from(suggestionData.scenePrefixes.entries()).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });
  entries.forEach(([value]) => {
    const normalised = normaliseScenePrefix(value);
    if (normalised === current) return;
    if (upperPrefix && !normalised.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, normalised, 'scenePrefix');
  });
  SCENE_PREFIX_SUGGESTIONS.forEach((value) => {
    const normalised = normaliseScenePrefix(value);
    if (normalised === current) return;
    if (upperPrefix && !normalised.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, normalised, 'scenePrefix');
  });
  return suggestions.slice(0, MAX_SUGGESTIONS);
}

function getSceneLocationSuggestions(prefix, currentValue) {
  const upperPrefix = (prefix || '').toUpperCase();
  const current = (currentValue || '').toUpperCase();
  const suggestions = [];
  const seen = new Set();
  const entries = Array.from(suggestionData.locationData.entries()).sort((a, b) => {
    if (b[1].count !== a[1].count) return b[1].count - a[1].count;
    return a[0].localeCompare(b[0]);
  });
  const filtered = entries.filter(([location]) => !upperPrefix || location.startsWith(upperPrefix));
  const source = filtered.length ? filtered : entries;
  source.forEach(([location]) => {
    if (location === current) return;
    addSuggestionOption(suggestions, seen, location, 'sceneLocation');
  });
  return suggestions.slice(0, MAX_SUGGESTIONS);
}

function getSceneTimeSuggestions(prefix, location, currentValue) {
  const upperPrefix = (prefix || '').toUpperCase();
  const current = (currentValue || '').toUpperCase();
  const suggestions = [];
  const seen = new Set();

  if (location) {
    const entry = suggestionData.locationData.get(location.toUpperCase());
    if (entry) {
      const locationTimes = Array.from(entry.times.entries()).sort((a, b) => {
        if (b[1] !== a[1]) return b[1] - a[1];
        return a[0].localeCompare(b[0]);
      });
      locationTimes.forEach(([time]) => {
        if (time === current) return;
        if (upperPrefix && !time.startsWith(upperPrefix)) return;
        addSuggestionOption(suggestions, seen, time, 'sceneTime');
      });
    }
  }

  const globalTimes = Array.from(suggestionData.timeCounts.entries()).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });
  globalTimes.forEach(([time]) => {
    if (time === current) return;
    if (upperPrefix && !time.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, time, 'sceneTime');
  });

  TIME_OF_DAY_SUGGESTIONS.forEach((time) => {
    const value = time.toUpperCase();
    if (value === current) return;
    if (upperPrefix && !value.startsWith(upperPrefix)) return;
    addSuggestionOption(suggestions, seen, value, 'sceneTime');
  });

  return suggestions.slice(0, MAX_SUGGESTIONS);
}

function addSuggestionOption(target, seen, value, kind) {
  if (!value || seen.has(value)) return;
  target.push({
    label: kind === 'characterQualifier' ? `(${value})` : value,
    value,
    kind,
  });
  seen.add(value);
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
  togglePreviewBtn.textContent = previewHidden ? 'Show right column' : 'Hide right column';
  toggleOutlineBtn.setAttribute('aria-pressed', String(!outlineHidden));
  togglePreviewBtn.setAttribute('aria-pressed', String(!previewHidden));
}

function setViewMode(mode, { persist = true } = {}) {
  if (!writingPanel) return;
  const nextMode = mode === 'formatted' ? 'formatted' : 'editor';
  currentViewMode = nextMode;
  writingPanel.dataset.viewMode = nextMode;
  viewModeButtons.forEach((button) => {
    const isActive = button.dataset.viewMode === nextMode;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', String(isActive));
  });

  if (nextMode === 'formatted') {
    closeSuggestionPanel();
    requestAnimationFrame(() => {
      editor.focus({ preventScroll: true });
      highlightFromCaret(editor.selectionStart, { scrollPreview: true });
    });
  } else {
    requestAnimationFrame(() => {
      editor.focus();
      updateSuggestionPanel();
    });
  }

  if (persist) {
    persistViewMode();
  }
}

function applyStoredViewMode() {
  if (!writingPanel) return;
  try {
    const stored = localStorage.getItem(VIEW_MODE_KEY);
    if (stored === 'formatted' || stored === 'editor') {
      setViewMode(stored, { persist: false });
    } else {
      setViewMode(currentViewMode, { persist: false });
    }
  } catch (error) {
    console.warn('Failed to load view mode', error);
    setViewMode(currentViewMode, { persist: false });
  }
}

function persistViewMode() {
  if (!writingPanel) return;
  try {
    localStorage.setItem(VIEW_MODE_KEY, currentViewMode);
  } catch (error) {
    console.warn('Failed to store view mode', error);
  }
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
  refreshSuggestionData(lineInfos);
  updatePreview(lineInfos);
  updateOutline(lineInfos);
  updateStats(lineInfos);
  updateInsights(lineInfos);
  highlightFromCaret(editor.selectionStart, { scrollPreview: false });
  updateSuggestionPanel();
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

  if (/^[A-Z]$/.test(trimmed) && (previousType !== 'dialogue' || leadingWhitespace >= CHARACTER_INDENT - 2)) {
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
  const significant = withoutParenthetical.replace(/[^A-Z0-9]/g, '');
  if (significant.length < 2) return false;
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
  autosaveTimer = null;
  const now = new Date();
  const updates = {
    title: titleInput.value,
    content: formattingResult.formattedText,
    beats: beatNotes.value,
    updatedAt: now.toISOString(),
  };
  const success = upsertDocument(currentDocumentId, updates);
  if (!success) {
    return;
  }
  lastSavedAt = now;
  isDirty = false;
  updateAutosaveStatus();
  renderDocumentLibrary();
}

function updateAutosaveStatus() {
  if (!autosaveStatus) return;
  let message = '';
  let state = 'idle';

  if (storageError) {
    message = 'Storage unavailable — changes may not be saved';
    state = 'error';
    autosaveStatus.textContent = message;
    if (autosaveIndicator) {
      autosaveIndicator.dataset.state = state;
    }
    return;
  }

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
    if (storageError) {
      storageError = false;
      storageErrorNotified = false;
      updateAutosaveStatus();
    }
    return true;
  } catch (error) {
    console.warn('Failed to store documents', error);
    storageError = true;
    if (!storageErrorNotified) {
      showSnackbar('Browser storage is unavailable. Changes may not be saved.');
      storageErrorNotified = true;
    }
    updateAutosaveStatus();
    return false;
  }
}

function upsertDocument(id, updates) {
  const index = documents.findIndex((doc) => doc.id === id);
  if (index === -1) return false;
  documents[index] = {
    ...documents[index],
    ...updates,
  };
  const persisted = persistDocuments();
  if (persisted) {
    localStorage.setItem(ACTIVE_DOCUMENT_KEY, id);
  }
  return persisted;
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
  const base64 = btoa(binary);
  return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function decodeSharePayload(encoded) {
  const sanitized = encoded.replace(/\s/g, '');
  let base64 = sanitized.replace(/-/g, '+').replace(/_/g, '/');
  const padding = base64.length % 4;
  if (padding > 0) {
    base64 = base64.padEnd(base64.length + (4 - padding), '=');
  }
  const binary = atob(base64);
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
  let stored = null;
  try {
    stored = localStorage.getItem(THEME_KEY);
  } catch (error) {
    console.warn('Unable to read stored theme preference', error);
  }
  if (stored === 'light' || stored === 'dark') {
    setTheme(stored);
    return;
  }

  const mediaQuery = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  const prefersDark = mediaQuery?.matches;
  setTheme(prefersDark ? 'dark' : 'light', { persist: false });

  if (mediaQuery) {
    const handleChange = (event) => {
      let hasStoredPreference = false;
      try {
        hasStoredPreference = Boolean(localStorage.getItem(THEME_KEY));
      } catch (error) {
        console.warn('Unable to read stored theme preference', error);
      }
      if (hasStoredPreference) {
        if (mediaQuery.removeEventListener) {
          mediaQuery.removeEventListener('change', handleChange);
        } else if (mediaQuery.removeListener) {
          mediaQuery.removeListener(handleChange);
        }
        return;
      }
      setTheme(event.matches ? 'dark' : 'light', { persist: false });
    };

    if (mediaQuery.addEventListener) {
      mediaQuery.addEventListener('change', handleChange);
    } else if (mediaQuery.addListener) {
      mediaQuery.addListener(handleChange);
    }
  }
}

function setTheme(theme, options = {}) {
  const { persist = true } = options;
  document.documentElement.dataset.theme = theme;
  if (persist) {
    try {
      localStorage.setItem(THEME_KEY, theme);
    } catch (error) {
      console.warn('Unable to store theme preference', error);
    }
  } else {
    try {
      localStorage.removeItem(THEME_KEY);
    } catch (error) {
      console.warn('Unable to clear stored theme preference', error);
    }
  }
  if (themeToggle) {
    const isDark = theme === 'dark';
    const label = isDark ? 'Switch to light theme' : 'Switch to dark theme';
    themeToggle.textContent = isDark ? '☀️' : '🌗';
    themeToggle.setAttribute('aria-label', label);
    themeToggle.setAttribute('title', label);
  }
}

function focusLine(lineIndex) {
  const info = formattingResult.lineInfos[lineIndex];
  if (!info) return;
  const caret = info.startFormatted;
  editor.focus();
  editor.setSelectionRange(caret, caret);
  highlightFromCaret(caret, { scrollPreview: true });
}

async function handleStoryDirectionRequest(event) {
  event?.preventDefault?.();
  if (!storyDirectionBtn || !storyDirectionOutput || !storyDirectionStatus) {
    return;
  }

  const scriptExcerpt = createExcerpt(editor?.value || '', 3500);
  const beatsExcerpt = createExcerpt(beatNotes?.value || '', 1200);

  if (!scriptExcerpt && !beatsExcerpt) {
    setStoryDirectionMessage(
      'Add some pages or jot ideas in Quick beats so Story Direction has context to work with.',
      { storeIdea: false }
    );
    storyDirectionStatus.textContent = '';
    return;
  }

  if (storyDirectionAbortController) {
    storyDirectionAbortController.abort();
  }

  storyDirectionAbortController = new AbortController();

  setStoryDirectionLoading(true);

  const userPrompt = [
    scriptExcerpt ? `SCRIPT SO FAR:\n${scriptExcerpt}` : 'SCRIPT SO FAR:\n(No script text yet.)',
    beatsExcerpt ? `QUICK BEATS:\n${beatsExcerpt}` : 'QUICK BEATS:\n(No quick beats entered.)',
    'TASK: Suggest a specific next story development in two or three sentences. Reference characters or threads already mentioned and explain why the move keeps momentum.'
  ].join('\n\n');

  try {
    const response = await fetch('https://histage.netlify.app/.netlify/functions/posts', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'HiStageReasoning',
        messages: [
          {
            role: 'system',
            content:
              'You are a seasoned screenwriting partner. Read the draft and quick beats carefully and craft one concise suggestion for what should happen next in the story.'
          },
          { role: 'user', content: userPrompt }
        ],
        temperature: 0.25,
        max_tokens: 700
      }),
      signal: storyDirectionAbortController.signal
    });

    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}`);
    }

    const data = await response.json();
    const content = data?.choices?.[0]?.message?.content?.trim();

    if (content) {
      setStoryDirectionMessage(content);
      storyDirectionStatus.textContent = 'Idea ready.';
    } else {
      setStoryDirectionMessage('The AI did not return a suggestion. Please try again.', { isError: true });
      storyDirectionStatus.textContent = '';
    }
  } catch (error) {
    if (error?.name === 'AbortError') {
      return;
    }
    console.error('Story direction request failed', error);
    setStoryDirectionMessage('Story Direction ran into a problem. Try again in a moment.', { isError: true });
    storyDirectionStatus.textContent = '';
  } finally {
    setStoryDirectionLoading(false);
    storyDirectionAbortController = null;
  }
}

function setStoryDirectionLoading(isLoading) {
  if (!storyDirectionBtn || !storyDirectionStatus) return;
  storyDirectionBtn.disabled = isLoading;
  storyDirectionBtn.textContent = isLoading ? 'Thinking…' : STORY_DIRECTION_DEFAULT_LABEL;
  storyDirectionStatus.textContent = isLoading ? 'Gathering a next beat…' : storyDirectionStatus.textContent;
  if (storyDirectionOutput) {
    if (isLoading) {
      storyDirectionOutput.setAttribute('aria-busy', 'true');
    } else {
      storyDirectionOutput.removeAttribute('aria-busy');
    }
  }
  updateStoryDirectionControls({ isLoading });
}

function setStoryDirectionMessage(message, options = {}) {
  if (!storyDirectionOutput) return;
  const { isError = false, storeIdea = !isError } = options;

  storyDirectionOutput.hidden = false;
  storyDirectionOutput.classList.toggle('is-error', Boolean(isError));

  if (storyDirectionText) {
    storyDirectionText.textContent = message;
  } else {
    storyDirectionOutput.textContent = message;
  }

  lastStoryDirectionIdea = storeIdea ? message : '';
  resetStoryDirectionFeedback();
  updateStoryDirectionControls();
}

function handleStoryDirectionImplement(event) {
  event?.preventDefault?.();
  if (!editor || !lastStoryDirectionIdea) return;

  const idea = lastStoryDirectionIdea.trim();
  if (!idea) return;

  const currentValue = editor.value || '';
  const trimmedCurrent = currentValue.replace(/\s+$/, '');
  const separator = trimmedCurrent ? '\n\n' : '';
  const nextValue = `${trimmedCurrent}${separator}${idea}`;

  editor.value = nextValue;
  editor.dispatchEvent(new Event('input', { bubbles: true }));
  editor.focus();
  const caret = editor.value.length;
  editor.setSelectionRange(caret, caret);

  if (storyDirectionFeedbackStatus) {
    storyDirectionFeedbackStatus.hidden = false;
    storyDirectionFeedbackStatus.textContent = 'Added the idea to your script.';
    storyDirectionFeedbackStatus.classList.remove('is-positive', 'is-negative');
    storyDirectionFeedbackStatus.classList.add('is-positive');
  }
}

function handleStoryDirectionFeedback(event) {
  event?.preventDefault?.();
  if (!lastStoryDirectionIdea) return;

  const target = event?.currentTarget;
  if (!target || target.disabled) return;

  const feedback = target.dataset?.feedback;
  if (!feedback) return;

  if (storyDirectionFeedbackSelection === feedback) {
    storyDirectionFeedbackSelection = null;
    if (storyDirectionApproveBtn) storyDirectionApproveBtn.classList.remove('is-active');
    if (storyDirectionRejectBtn) storyDirectionRejectBtn.classList.remove('is-active');
    if (storyDirectionFeedbackStatus) {
      storyDirectionFeedbackStatus.textContent = '';
      storyDirectionFeedbackStatus.classList.remove('is-positive', 'is-negative');
      storyDirectionFeedbackStatus.hidden = true;
    }
    return;
  }

  storyDirectionFeedbackSelection = feedback;

  if (storyDirectionApproveBtn) {
    storyDirectionApproveBtn.classList.toggle('is-active', feedback === 'approve');
  }
  if (storyDirectionRejectBtn) {
    storyDirectionRejectBtn.classList.toggle('is-active', feedback === 'reject');
  }

  if (storyDirectionFeedbackStatus) {
    storyDirectionFeedbackStatus.hidden = false;
    storyDirectionFeedbackStatus.classList.remove('is-positive', 'is-negative');
    if (feedback === 'approve') {
      storyDirectionFeedbackStatus.textContent = 'Marked this idea as helpful.';
      storyDirectionFeedbackStatus.classList.add('is-positive');
    } else {
      storyDirectionFeedbackStatus.textContent = "Marked this idea as not quite right.";
      storyDirectionFeedbackStatus.classList.add('is-negative');
    }
  }
}

function resetStoryDirectionFeedback() {
  storyDirectionFeedbackSelection = null;
  if (storyDirectionApproveBtn) {
    storyDirectionApproveBtn.classList.remove('is-active');
  }
  if (storyDirectionRejectBtn) {
    storyDirectionRejectBtn.classList.remove('is-active');
  }
  if (storyDirectionFeedbackStatus) {
    storyDirectionFeedbackStatus.textContent = '';
    storyDirectionFeedbackStatus.classList.remove('is-positive', 'is-negative');
    storyDirectionFeedbackStatus.hidden = true;
  }
}

function updateStoryDirectionControls(options = {}) {
  const { isLoading = false } = options;
  const hasIdea = Boolean(lastStoryDirectionIdea);

  if (storyDirectionActions) {
    storyDirectionActions.hidden = !hasIdea;
  }

  const shouldDisable = !hasIdea || isLoading;
  if (storyDirectionImplementBtn) {
    storyDirectionImplementBtn.disabled = shouldDisable;
  }
  if (storyDirectionApproveBtn) {
    storyDirectionApproveBtn.disabled = shouldDisable;
  }
  if (storyDirectionRejectBtn) {
    storyDirectionRejectBtn.disabled = shouldDisable;
  }
}

function createExcerpt(text, limit = 3500) {
  if (!text) return '';
  const trimmed = text.trim();
  if (trimmed.length <= limit) return trimmed;
  return `…${trimmed.slice(trimmed.length - limit)}`;
}

window.addEventListener('beforeunload', (event) => {
  if (!isDirty) return;
  event.preventDefault();
  event.returnValue = '';
});
