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

const STORAGE_KEY = 'scriptius-document';
const THEME_KEY = 'scriptius-theme';
const AUTOSAVE_DELAY = 800;

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

init();

function init() {
  applyStoredTheme();
  attachEventListeners();
  loadFromStorage();
  loadFromShareLink();
  if (!editor.value.trim()) {
    seedWelcomePage();
  }
  applyFormattingAndRefresh({ preserveCaret: false });
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
  });

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
  highlightFromCaret(caret);
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
  highlightFromCaret(editor.selectionStart);
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
    const exportLine = info.type === 'empty' ? '' : info.indent + info.content;
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
  const trimmedLeft = trimmedRight.replace(/^\s+/u, '');
  const leadingWhitespace = trimmedRight.length - trimmedLeft.length;
  const trimmed = trimmedLeft;

  if (!trimmed) {
    return { type: 'empty', content: '', indent: INDENTS.empty, leadingWhitespace, location: null, timeOfDay: null };
  }

  if (isSceneHeading(trimmed)) {
    const content = normaliseSceneHeading(trimmed);
    const sceneBits = parseSceneParts(content);
    return {
      type: 'scene',
      content,
      indent: INDENTS.scene,
      leadingWhitespace,
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
      location: null,
      timeOfDay: null,
    };
  }

  return {
    type: 'action',
    content: trimmed,
    indent: INDENTS.action,
    leadingWhitespace,
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
      const trimmedOffset = Math.min(offsetInLine - info.leadingWhitespace, info.formattedContentLength);
      return info.startFormatted + info.indentLength + trimmedOffset;
    }
    if (originalCaret === info.endOriginal + 1 && info.hasBreak) {
      return info.endFormatted + 1;
    }
  }
  return formattedText.length;
}

function updatePreview(lineInfos) {
  previewContainer.innerHTML = '';
  const fragment = document.createDocumentFragment();
  let currentPage = createPage();
  fragment.appendChild(currentPage);
  let countableLines = 0;

  if (!lineInfos.length) {
    const placeholder = document.createElement('div');
    placeholder.className = 'script-line action';
    placeholder.textContent = 'Start writing to see formatted pages here.';
    currentPage.appendChild(placeholder);
    previewContainer.appendChild(fragment);
    return;
  }

  lineInfos.forEach((info) => {
    const shouldCount = info.type !== 'empty';
    if (shouldCount && countableLines >= 55) {
      currentPage = createPage();
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

    if (shouldCount) {
      countableLines += 1;
    }
  });

  previewContainer.appendChild(fragment);
}

function createPage() {
  const page = document.createElement('div');
  page.className = 'script-page';
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
    const page = Math.max(1, Math.round((scene.index + 1) / 55));
    metaEl.textContent = `Page ${page}${time}`;

    button.dataset.lineIndex = String(scene.index);
    fragment.appendChild(clone);
  });

  sceneList.appendChild(fragment);
  const label = `${scenes.length} ${scenes.length === 1 ? 'scene' : 'scenes'}`;
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
  const estimatedPages = Math.max(1, Math.round(nonEmptyLines / 55));

  wordStat.textContent = `${wordCount} ${wordCount === 1 ? 'word' : 'words'}`;
  sceneStat.textContent = `${scenes} ${scenes === 1 ? 'scene' : 'scenes'}`;
  runtimeStat.textContent = `Runtime: ${estimatedPages} ${estimatedPages === 1 ? 'min' : 'mins'}`;
  pageIndicator.textContent = `Est. ${estimatedPages} ${estimatedPages === 1 ? 'page' : 'pages'}`;
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
      value: characters.size ? `${characters.size}` : '—',
    },
    {
      label: 'Most used location',
      value: topLocation ? `${topLocation[0]} (${topLocation[1]})` : '—',
    },
  ];

  if (longestScene.title) {
    insights.push({
      label: 'Longest scene',
      value: `${shortenSceneTitle(longestScene.title)} (${longestScene.length} lines)`,
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

function highlightFromCaret(caret) {
  const lineInfos = formattingResult.lineInfos;
  if (!lineInfos.length) return;
  const lineIndex = getLineIndexFromCaret(caret, lineInfos);
  highlightPreviewLine(lineIndex);
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

function highlightPreviewLine(lineIndex) {
  previewContainer.querySelectorAll('.script-line.active').forEach((node) => node.classList.remove('active'));
  const target = previewContainer.querySelector(`.script-line[data-line-index="${lineIndex}"]`);
  if (target) {
    target.classList.add('active');
    target.scrollIntoView({ block: 'center', behavior: 'smooth' });
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

function scheduleAutosave() {
  markDirty();
  if (autosaveTimer) {
    clearTimeout(autosaveTimer);
  }
  autosaveTimer = setTimeout(performAutosave, AUTOSAVE_DELAY);
}

function markDirty() {
  isDirty = true;
}

function performAutosave() {
  const payload = {
    title: titleInput.value,
    content: formattingResult.formattedText,
    beats: beatNotes.value,
    updatedAt: new Date().toISOString(),
  };
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    lastSavedAt = new Date(payload.updatedAt);
    isDirty = false;
    updateAutosaveStatus();
  } catch (error) {
    console.warn('Autosave failed', error);
  }
}

function updateAutosaveStatus() {
  if (!lastSavedAt) {
    autosaveStatus.textContent = isDirty ? 'Unsaved changes' : 'Draft not saved';
    return;
  }
  const diffMs = Date.now() - lastSavedAt.getTime();
  if (diffMs < 2000) {
    autosaveStatus.textContent = 'Saved just now';
  } else if (diffMs < 60000) {
    autosaveStatus.textContent = `Saved ${Math.round(diffMs / 1000)}s ago`;
  } else if (diffMs < 3600000) {
    autosaveStatus.textContent = `Saved ${Math.round(diffMs / 60000)}m ago`;
  } else {
    autosaveStatus.textContent = `Saved ${lastSavedAt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
  }
}

function loadFromStorage() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return;
    const payload = JSON.parse(stored);
    titleInput.value = payload.title || '';
    editor.value = (payload.content || '').replace(/\r\n?/g, '\n');
    beatNotes.value = payload.beats || '';
    lastSavedAt = payload.updatedAt ? new Date(payload.updatedAt) : null;
  } catch (error) {
    console.warn('Failed to load stored draft', error);
  }
}

function seedWelcomePage() {
  const template = `INT. DREAMSCAPE - NIGHT\n\nA swirling cosmic void. Possibilities everywhere.\n\nNARRATOR\nWelcome to Scriptius. Start writing your story...\n`;
  editor.value = template;
}

function loadFromShareLink() {
  const hash = window.location.hash || '';
  if (!hash.startsWith('#share=')) return;
  const encoded = hash.slice('#share='.length);
  try {
    const payload = decodeSharePayload(encoded);
    if (payload?.content) {
      titleInput.value = payload.title || 'Shared draft';
      editor.value = payload.content.replace(/\r\n?/g, '\n');
      beatNotes.value = payload.beats || '';
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
  titleInput.value = '';
  editor.value = '';
  beatNotes.value = '';
  lastSavedAt = null;
  isDirty = false;
  updateAutosaveStatus();
  applyFormattingAndRefresh({ preserveCaret: false });
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
  highlightFromCaret(caret);
}

window.addEventListener('beforeunload', (event) => {
  if (!isDirty) return;
  event.preventDefault();
  event.returnValue = '';
});

