const state = {
  groups: [],
  selectedId: null,
  promptMenuOpen: false,
  apiKeyPromptDismissedForProvider: null,
  contextSourceMode: 'selection',
  selectedSourceMode: 'selection',
};

const els = {};

const STORAGE_KEY = 'scopedEditAssistantSettingsV4';
const LEGACY_STORAGE_KEYS = ['scopedEditAssistantSettingsV3', 'scopedEditAssistantSettingsV2', 'scopedEditAssistantSettingsV1', 'tightReplaceSettingsV4', 'tightReplaceSettingsV3', 'tightReplaceSettingsV2'];

const PRESET_GOOGLE_SLOW = 'google-slow';
const PRESET_GOOGLE_FAST = 'google-fast';
const PRESET_OPENAI = 'openai';
const PRESET_INTERNAL = 'internal';
const PRESET_CUSTOM = 'custom';
const PRESET_MANUAL = 'manual';
const DEFAULT_PROVIDER = PRESET_GOOGLE_SLOW;

const GOOGLE_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions';
const GOOGLE_SLOW_MODEL = 'gemini-3.1-pro-preview';
const GOOGLE_FAST_MODEL = 'gemini-3-flash-preview';

const OPENAI_ENDPOINT = 'https://api.openai.com/v1/chat/completions';
const OPENAI_MODEL = 'gpt-5.4';

const INTERNAL_ENDPOINT = 'http://127.0.0.1:1234/v1/chat/completions';
const INTERNAL_MODEL = 'qwen/qwen3.5-35b-a3b';

const CUSTOM_DEFAULT_ENDPOINT = INTERNAL_ENDPOINT;
const CUSTOM_DEFAULT_MODEL = INTERNAL_MODEL;

const DEFAULT_TEMPERATURE = '0.1';
const DEFAULT_MAX_TOKENS = '30000';
const DEFAULT_STYLE_GUIDE = '';

const PRESET_META = {
  [PRESET_GOOGLE_SLOW]: {
    badge: 'Google smart preset',
    badgeTone: 'success',
    endpoint: GOOGLE_ENDPOINT,
    model: GOOGLE_SLOW_MODEL,
    apiKeyRequired: true,
    apiKeyPlaceholder: 'Enter your Google API key',
    apiKeyHint: 'Google and ChatGPT / OpenAI presets require a key. Internal and custom can leave this blank if the endpoint does not require auth.',
    statusMessage: 'Google smart preset selected. Default model: gemini-3.1-pro-preview.',
  },
  [PRESET_GOOGLE_FAST]: {
    badge: 'Google fast preset',
    badgeTone: 'success',
    endpoint: GOOGLE_ENDPOINT,
    model: GOOGLE_FAST_MODEL,
    apiKeyRequired: true,
    apiKeyPlaceholder: 'Enter your Google API key',
    apiKeyHint: 'Google and ChatGPT / OpenAI presets require a key. Internal and custom can leave this blank if the endpoint does not require auth.',
    statusMessage: 'Google fast preset selected. Default model: gemini-3-flash-preview.',
  },
  [PRESET_OPENAI]: {
    badge: 'ChatGPT preset',
    badgeTone: 'success',
    endpoint: OPENAI_ENDPOINT,
    model: OPENAI_MODEL,
    apiKeyRequired: true,
    apiKeyPlaceholder: 'Enter your OpenAI API key',
    apiKeyHint: 'Google and ChatGPT / OpenAI presets require a key. Internal and custom can leave this blank if the endpoint does not require auth.',
    statusMessage: 'ChatGPT / OpenAI preset selected. Default endpoint: api.openai.com with gpt-5.4.',
  },
  [PRESET_INTERNAL]: {
    badge: 'Internal preset',
    badgeTone: 'neutral',
    endpoint: INTERNAL_ENDPOINT,
    model: INTERNAL_MODEL,
    apiKeyRequired: false,
    apiKeyPlaceholder: 'Optional for your local endpoint',
    apiKeyHint: 'Internal uses your localhost OpenAI-compatible endpoint. Leave the key blank if your local server does not require auth.',
    statusMessage: 'Internal preset selected. Default endpoint: localhost with qwen/qwen3.5-35b-a3b.',
  },
  [PRESET_CUSTOM]: {
    badge: 'Custom preset',
    badgeTone: 'neutral',
    endpoint: CUSTOM_DEFAULT_ENDPOINT,
    model: CUSTOM_DEFAULT_MODEL,
    apiKeyRequired: false,
    apiKeyPlaceholder: 'Optional, depending on your endpoint',
    apiKeyHint: 'Custom leaves the endpoint and model fully manual. Enter a key only if your endpoint requires authorization.',
    statusMessage: 'Custom preset selected. Your endpoint and model remain manual.',
  },
  [PRESET_MANUAL]: {
    badge: 'Manual / Copilot',
    badgeTone: 'neutral',
    endpoint: '',
    model: '',
    apiKeyRequired: false,
    apiKeyPlaceholder: '',
    apiKeyHint: 'Manual / Copilot mode copies the full prompt to your clipboard so you can paste it into Copilot or any other AI assistant.',
    statusMessage: 'Manual / Copilot mode. Click Send to copy the full prompt to your clipboard.',
  },
};

const OUTPUT_INSTRUCTIONS_TEMPLATE = `You are generating edit suggestions for a Word add-in.

Return exactly one top-level JSON array and nothing else.
Do not wrap the JSON in markdown fences.
Do not add commentary before or after the JSON.
If there are no valid edits, return [] .

Each array item must be an object with this exact shape:
[
  {
    "group": 1,
    "original": "Exact sentence or paragraph from the source text.",
    "replacement": "Full revised sentence or paragraph.",
    "original_span": "Smallest unique substring inside original that should be replaced.",
    "replacement_span": "Exact text that replaces original_span.",
    "reasoning_summary": "1 to 3 sentences explaining the editorial reason for the change."
  }
]

Field rules:
- group: integer, starting at 1 and increasing by 1 for each suggestion.
- original: must be copied exactly from the source text, preserving punctuation, capitalization, apostrophes, and spacing.
- replacement: the full revised version of original, not just the changed fragment.
- original_span: the narrowest possible substring inside original that can be replaced safely. It must appear exactly once within original. If the smallest span is ambiguous or the change is an insertion, widen the span just enough to make it unique and anchor the edit.
- replacement_span: the exact text that replaces original_span. Leave this empty only for a pure deletion.
- reasoning_summary: short, user-facing editorial rationale focused on clarity, tone, concision, grammar, consistency, or correctness.

Output rules:
- Use only text that is supported by the source provided to you.
- Do not invent facts, figures, citations, or new content beyond the requested edit.
- Do not change text outside the requested edit.
- Keep original_span and replacement_span as tight as possible so tracked changes stay minimal.
- Escape all JSON correctly.

Example:
[
  {
    "group": 1,
    "original": "This report is divided into the following chapters:",
    "replacement": "This report is organized as follows:",
    "original_span": "divided into the following chapters",
    "replacement_span": "organized as follows",
    "reasoning_summary": "This change makes the wording more concise and formal. It also uses a cleaner phrase that reads more naturally in report introductions."
  },
  {
    "group": 2,
    "original": "It is absolutely essential that we must review the budget before the final deadline.",
    "replacement": "We must review the budget before the final deadline.",
    "original_span": "It is absolutely essential that we must",
    "replacement_span": "We must",
    "reasoning_summary": "Removes redundant phrasing to make the sentence more direct and concise."
  }
]`;

const DEFAULT_SYSTEM_PROMPT = `You are editing text for a Word add-in that applies tightly scoped tracked changes.

Output only valid JSON as a top-level array. Do not add markdown fences. Do not add commentary before or after the JSON.

For each requested edit, output one object with these fields:

group: integer, starting at 1.
original: the exact original sentence or paragraph from the document.
replacement: the exact full replacement sentence or paragraph.
original_span: the smallest exact substring inside original that should be replaced.
replacement_span: the exact replacement text for original_span.
reasoning_summary: a brief plain-language explanation of the edit, in 1 to 3 sentences.

Rules:

original must match the document exactly, including punctuation, apostrophes, capitalization, and spacing.
replacement must be the full rewritten sentence or paragraph.
original_span and replacement_span must be as tight as possible so tracked changes show the narrowest useful edit.
If the smallest possible span would be ambiguous because it appears more than once inside original, widen original_span just enough to make it unique within original.
Do not use an empty original_span. If the edit is an insertion, widen the span slightly so there is still a concrete replaceable substring.
Do not use an empty replacement_span unless the edit is a deletion.
reasoning_summary must be concise and user-facing. It should explain the editorial rationale, not hidden internal reasoning.
Do not change text outside the requested edit.
Escape JSON correctly.

Tone and style instructions:
10. Preserve a formal, neutral, and professional tone.
11. Prefer clear, direct, analytical wording over conversational, promotional, emotional, or dramatic phrasing.
12. Keep the prose concise, but do not make it abrupt. Maintain smooth, polished sentence flow appropriate for a public report.
13. Avoid slang, idioms, rhetorical questions, intensifiers, and casual transitions.
14. Prefer precise and literal wording. Do not add flourish, personality, or persuasive framing unless the source text already requires it.
15. Maintain an objective, non-partisan, and evidence-oriented style. Do not overstate certainty or exaggerate findings.
16. Prefer terminology that sounds institutional, policy-focused, and technically accurate when that matches the source.
17. Keep meaning stable. Do not introduce new interpretation, emphasis, or implied judgment.
18. When tightening text, favor edits that improve clarity, readability, and consistency without making the prose sound informal.
19. Preserve a measured cadence: short enough to be clear, but not choppy; detailed enough to be precise, but not wordy.
20. Keep terminology, capitalization, and naming conventions consistent within the edited sentence or paragraph and with the surrounding document when possible.

Good reasoning_summary examples:

"Removes unnecessary wording and makes the sentence more concise."
"Replaces repetitive phrasing with a cleaner formulation while preserving meaning."
"Improves clarity by naming the subject more directly and tightening the sentence structure."

Example:
[
{
"group": 1,
"original": "This report is divided into the following chapters:",
"replacement": "This report is organized as follows:",
"original_span": "divided into the following chapters",
"replacement_span": "organized as follows",
"reasoning_summary": "This change makes the wording more concise and formal. It also uses a cleaner phrase that reads more naturally in report introductions."
}
]

A shorter add-on block you could append instead is:

Tone and style instructions:
10. Preserve a formal, neutral, and professional tone.
11. Prefer clear, direct, analytical wording over conversational or promotional phrasing.
12. Keep edits concise while maintaining smooth sentence flow and public-report formality.
13. Avoid slang, idioms, rhetorical flourish, and exaggerated emphasis.
14. Use precise, institutional, and policy-appropriate language where relevant.
15. Maintain an objective, evidence-oriented style and do not overstate certainty.
16. Do not introduce new interpretation, judgment, or emphasis beyond the original meaning.
17. Improve clarity and consistency without making the prose sound casual or simplified.`;

const COMPREHENSIVE_CHAT_PROMPT = `Review the provided document text and return tightly scoped edits suitable for tracked changes. Take a comprehensive pass: improve clarity, consistency, tone, grammar, spelling, capitalization, terminology, and internal alignment of numbers and text where the provided document supports that review. Check for inconsistencies between tables and prose, and across sections or chapters, but only make corrections that are supported by the provided text.`;

const CLARITY_CHAT_PROMPT = `Review the provided document text and return tightly scoped edits focused on clarity. Improve sentence structure, remove unnecessary wording, resolve awkward or indirect phrasing, and make the logic easier to follow on a first read while preserving meaning and keeping the tone formal and professional.`;

const CONSISTENCY_CHAT_PROMPT = `Review the provided document text and return tightly scoped edits focused on consistency of numbers and text. Check for alignment between tables and running text, and across sections or chapters, including figures, percentages, dates, units, capitalization, labels, terminology, and repeated references. Only make consistency edits that are clearly supported by the provided text.`;

const TONE_CHAT_PROMPT = `Review the provided document text and return tightly scoped edits focused on tone. Make the prose read as formal, neutral, professional, and publication-ready without changing meaning, level of certainty, or factual posture. Remove conversational, promotional, emotional, or overly emphatic wording where present.`;

const PROOFREADING_CHAT_PROMPT = `Review the provided document text and return tightly scoped proofreading edits only. Correct grammar, spelling, punctuation, capitalization, and obvious typographical errors while preserving wording, structure, meaning, and tone as much as possible. Do not make broader stylistic or substantive revisions unless needed to fix a clear language error.`;

const CHAT_PROMPT_PRESETS = {
  comprehensive: {
    label: 'Comprehensive',
    badge: 'Comprehensive',
    tone: 'success',
    prompt: COMPREHENSIVE_CHAT_PROMPT,
  },
  clarity: {
    label: 'Clarity',
    badge: 'Clarity',
    tone: 'neutral',
    prompt: CLARITY_CHAT_PROMPT,
  },
  consistency: {
    label: 'Consistency of numbers and text',
    badge: 'Consistency',
    tone: 'neutral',
    prompt: CONSISTENCY_CHAT_PROMPT,
  },
  tone: {
    label: 'Tone',
    badge: 'Tone',
    tone: 'neutral',
    prompt: TONE_CHAT_PROMPT,
  },
  proofreading: {
    label: 'Proofreading',
    badge: 'Proofreading',
    tone: 'neutral',
    prompt: PROOFREADING_CHAT_PROMPT,
  },
  custom: {
    label: 'Custom / current text',
    badge: 'Custom / current text',
    tone: 'neutral',
    prompt: '',
  },
};

const DEFAULT_CHAT_PROMPT_PRESET = 'comprehensive';

Office.onReady((info) => {
  bindElements();
  bindEvents();
  hydrateLlmSettings();
  initializeToggleButtons();
  updateProviderUi();
  updateResponseAutoLoadBadge('Awaiting response', 'neutral');
  updateRequestFeedback('Ready', 'neutral', 'Nothing sent yet.');
  updateSourceModeCards();
  updateSourceStatusSummary();

  if (info.host !== Office.HostType.Word) {
    setGlobalStatus('This add-in only works in Microsoft Word.', 'error');
    setLlmStatus('This add-in only works in Microsoft Word.', 'error');
    updateRequestFeedback('Unavailable', 'error', 'This add-in only works in Microsoft Word.');
    return;
  }

  setGlobalStatus('Ready. Send a request normally, or use the advanced manual JSON tools when needed.', 'neutral');
  setLlmStatus('Ready. Choose a source, set instructions, and send when you are ready.', 'neutral');
  updateSourceStatusSummary();

  if (els.documentContextInput.value.trim()) {
    if (state.contextSourceMode === 'selection' || state.contextSourceMode === 'body') {
      void maybeRestoreSavedContextFromWord().finally(() => updateSourceStatusSummary());
    } else {
      setLlmStatus('Restored your last source text from local storage.', 'neutral');
      updateSourceStatusSummary();
    }
  }
});

function bindEvents() {
  if (els.toggleLlmSectionBtn) {
    els.toggleLlmSectionBtn.addEventListener('click', () => toggleSection(els.toggleLlmSectionBtn, els.llmSectionBody));
  }
  if (els.toggleAdvancedPanelBtn) {
    els.toggleAdvancedPanelBtn.addEventListener('click', () => toggleSection(els.toggleAdvancedPanelBtn, els.advancedPanelBody));
  }
  if (els.toggleJsonSectionBtn) {
    els.toggleJsonSectionBtn.addEventListener('click', () => toggleSection(els.toggleJsonSectionBtn, els.jsonSectionBody));
  }

  if (els.loadBtn) els.loadBtn.addEventListener('click', onLoadSuggestions);
  if (els.copyOutputInstructionsBtn) els.copyOutputInstructionsBtn.addEventListener('click', onCopyOutputInstructions);
  if (els.trackingBtn) els.trackingBtn.addEventListener('click', onCheckTracking);
  if (els.findBtn) els.findBtn.addEventListener('click', onFindSentence);
  if (els.prevMatchBtn) els.prevMatchBtn.addEventListener('click', () => onCycleMatch(-1));
  if (els.nextMatchBtn) els.nextMatchBtn.addEventListener('click', () => onCycleMatch(1));
  if (els.applyBtn) els.applyBtn.addEventListener('click', onApplySelected);
  if (els.prevGroupBtn) els.prevGroupBtn.addEventListener('click', () => onCycleGroup(-1));
  if (els.nextGroupBtn) els.nextGroupBtn.addEventListener('click', () => onCycleGroup(1));
  if (els.applyAllBtn) els.applyAllBtn.addEventListener('click', onApplyAllExactSingles);
  if (els.copyJsonExampleBtn) els.copyJsonExampleBtn.addEventListener('click', insertSampleJson);

  if (els.resetSystemPromptBtn) els.resetSystemPromptBtn.addEventListener('click', onResetSystemPrompt);
  if (els.loadSelectionBtn) els.loadSelectionBtn.addEventListener('click', onLoadSelectionIntoContext);
  if (els.loadBodyBtn) els.loadBodyBtn.addEventListener('click', onLoadBodyIntoContext);
  if (els.refreshSourceBtn) els.refreshSourceBtn.addEventListener('click', onRefreshSourcePreview);
  if (els.sourceModeSelection) els.sourceModeSelection.addEventListener('change', onSourceModeChanged);
  if (els.sourceModeBody) els.sourceModeBody.addEventListener('change', onSourceModeChanged);
  if (els.clearManualSourceBtn) els.clearManualSourceBtn.addEventListener('click', onClearManualSource);
  if (els.clearDocumentContextBtn) els.clearDocumentContextBtn.addEventListener('click', onClearManualSource);

  if (els.documentContextInput) {
    els.documentContextInput.addEventListener('input', () => {
      state.contextSourceMode = els.documentContextInput.value.trim() ? 'manual' : getSelectedSourceMode();
      updateSourceStatusSummary();
      persistLlmSettings();
    });
  }

  if (els.styleGuideInput) {
    els.styleGuideInput.addEventListener('input', () => {
      updateStyleGuideBadge();
      persistLlmSettings();
    });
  }
  if (els.clearStyleGuideBtn) els.clearStyleGuideBtn.addEventListener('click', onClearStyleGuide);

  if (els.chatPromptPreset) els.chatPromptPreset.addEventListener('change', onChatPromptPresetChanged);
  if (els.chatPromptMenuBtn) els.chatPromptMenuBtn.addEventListener('click', onToggleChatPromptMenu);
  if (els.chatPromptMenuItems?.length) {
    els.chatPromptMenuItems.forEach((item) => {
      item.addEventListener('click', () => onChatPromptMenuItemSelected(item.dataset.promptPreset || 'custom'));
    });
  }

  document.addEventListener('click', onDocumentClick);
  document.addEventListener('keydown', onDocumentKeydown);

  if (els.sendChatBtn) els.sendChatBtn.addEventListener('click', () => onSendChatToLlm({ focusSuggestions: false }));
  if (els.sendAndLoadBtn) els.sendAndLoadBtn.addEventListener('click', () => onSendChatToLlm({ focusSuggestions: true }));
  if (els.copyResponseToJsonBtn) els.copyResponseToJsonBtn.addEventListener('click', onCopyResponseToJson);
  if (els.loadResponseBtn) els.loadResponseBtn.addEventListener('click', onLoadResponseAsSuggestions);

  if (els.llmProvider) els.llmProvider.addEventListener('change', onProviderChanged);
  if (els.applyProviderPresetBtn) els.applyProviderPresetBtn.addEventListener('click', onApplyProviderPreset);
  if (els.saveAsCustomBtn) els.saveAsCustomBtn.addEventListener('click', onSaveAsCustom);

  if (els.parseCopilotResponseBtn) els.parseCopilotResponseBtn.addEventListener('click', onParseCopilotResponse);

  if (els.saveApiKeyModalBtn) els.saveApiKeyModalBtn.addEventListener('click', onSaveApiKeyFromModal);
  if (els.closeApiKeyModalBtn) els.closeApiKeyModalBtn.addEventListener('click', closeApiKeyModal);
  if (els.openSettingsFromApiKeyModalBtn) {
    els.openSettingsFromApiKeyModalBtn.addEventListener('click', () => {
      setSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody, true);
      closeApiKeyModal();
      els.llmApiKey?.focus();
    });
  }
  if (els.apiKeyModalInput) {
    els.apiKeyModalInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        onSaveApiKeyFromModal();
      }
    });
  }
  if (els.apiKeyModal) {
    els.apiKeyModal.addEventListener('click', (event) => {
      if (event.target === els.apiKeyModal) {
        closeApiKeyModal();
      }
    });
  }
}
function bindElements() {
  Object.assign(els, {
    // Step 1 / source
    toggleLlmSectionBtn: document.getElementById('toggleLlmSectionBtn'),
    llmSectionBody: document.getElementById('llmSectionBody'),
    sourceSelectionCard: document.getElementById('sourceSelectionCard'),
    sourceBodyCard: document.getElementById('sourceBodyCard'),
    sourceModeSelection: document.getElementById('sourceModeSelection'),
    sourceModeBody: document.getElementById('sourceModeBody'),
    llmStatus: document.getElementById('llmStatus'),
    advancedPanelBody: document.getElementById('advancedPanelBody'),
    documentContextInput: document.getElementById('documentContextInput'),
    loadSelectionBtn: document.getElementById('loadSelectionBtn'),
    loadBodyBtn: document.getElementById('loadBodyBtn'),
    clearManualSourceBtn: document.getElementById('clearManualSourceBtn'),

    // Settings / provider
    llmProvider: document.getElementById('llmProvider'),
    applyProviderPresetBtn: document.getElementById('applyProviderPresetBtn'),
    saveAsCustomBtn: document.getElementById('saveAsCustomBtn'),
    toggleAdvancedPanelBtn: document.getElementById('toggleAdvancedPanelBtn'),
    llmApiKey: document.getElementById('llmApiKey'),
    apiKeyHint: document.getElementById('apiKeyHint'),
    apiKeyRequirementBadge: document.getElementById('apiKeyRequirementBadge'),
    llmEndpoint: document.getElementById('llmEndpoint'),
    llmModel: document.getElementById('llmModel'),
    llmTemperature: document.getElementById('llmTemperature'),
    llmMaxTokens: document.getElementById('llmMaxTokens'),
    systemPromptInput: document.getElementById('systemPromptInput'),
    resetSystemPromptBtn: document.getElementById('resetSystemPromptBtn'),

    // Prompt / style guide
    chatInput: document.getElementById('chatInput'),
    chatPromptPreset: document.getElementById('chatPromptPreset'),
    chatPromptMenuBtn: document.getElementById('chatPromptMenuBtn'),
    chatPromptMenu: document.getElementById('chatPromptMenu'),
    chatPromptMenuItems: Array.from(document.querySelectorAll('.preset-menu-item')),
    styleGuideInput: document.getElementById('styleGuideInput'),
    styleGuideSavedBadge: document.getElementById('styleGuideSavedBadge'),
    clearStyleGuideBtn: document.getElementById('clearStyleGuideBtn'),

    // Send / request feedback
    sendAndLoadBtn: document.getElementById('sendAndLoadBtn'),
    requestSection: document.getElementById('requestSection'),
    requestStateBadge: document.getElementById('requestStateBadge'),
    requestFeedback: document.getElementById('requestFeedback'),
    requestProgressText: document.getElementById('requestProgressText'),

    // Response / JSON tools
    llmResponse: document.getElementById('llmResponse'),
    responseAutoLoadBadge: document.getElementById('responseAutoLoadBadge'),
    copyResponseToJsonBtn: document.getElementById('copyResponseToJsonBtn'),
    loadResponseBtn: document.getElementById('loadResponseBtn'),
    toggleJsonSectionBtn: document.getElementById('toggleJsonSectionBtn'),
    jsonSectionBody: document.getElementById('jsonSectionBody'),
    jsonInput: document.getElementById('jsonInput'),
    loadBtn: document.getElementById('loadBtn'),
    copyOutputInstructionsBtn: document.getElementById('copyOutputInstructionsBtn'),
    trackingBtn: document.getElementById('trackingBtn'),
    globalStatus: document.getElementById('globalStatus'),

    // Suggestions / detail pane
    groupList: document.getElementById('groupList'),
    detailEmpty: document.getElementById('detailEmpty'),
    detailPane: document.getElementById('detailPane'),
    detailGroup: document.getElementById('detailGroup'),
    detailMatchCount: document.getElementById('detailMatchCount'),
    detailStatus: document.getElementById('detailStatus'),
    detailOriginal: document.getElementById('detailOriginal'),
    detailReplacement: document.getElementById('detailReplacement'),
    detailSpanPreview: document.getElementById('detailSpanPreview'),
    detailReasoning: document.getElementById('detailReasoning'),
    findBtn: document.getElementById('findBtn'),
    prevMatchBtn: document.getElementById('prevMatchBtn'),
    nextMatchBtn: document.getElementById('nextMatchBtn'),
    applyBtn: document.getElementById('applyBtn'),
    prevGroupBtn: document.getElementById('prevGroupBtn'),
    nextGroupBtn: document.getElementById('nextGroupBtn'),
    groupStatus: document.getElementById('groupStatus'),

    // API key modal
    apiKeyModal: document.getElementById('apiKeyModal'),
    apiKeyModalInput: document.getElementById('apiKeyModalInput'),
    apiKeyModalStatus: document.getElementById('apiKeyModalStatus'),
    saveApiKeyModalBtn: document.getElementById('saveApiKeyModalBtn'),
    closeApiKeyModalBtn: document.getElementById('closeApiKeyModalBtn'),
    openSettingsFromApiKeyModalBtn: document.getElementById('openSettingsFromApiKeyModalBtn'),

    // Optional/missing-safe elements referenced elsewhere
    providerPresetBadge: document.getElementById('providerPresetBadge'),
    sourceStatusText: document.getElementById('sourceStatusText'),
    sourceStatusBadge: document.getElementById('sourceStatusBadge'),
    clearDocumentContextBtn: document.getElementById('clearDocumentContextBtn'),
    refreshSourceBtn: document.getElementById('refreshSourceBtn'),
    sendChatBtn: document.getElementById('sendChatBtn'),
    applyAllBtn: document.getElementById('applyAllBtn'),
    copyJsonExampleBtn: document.getElementById('copyJsonExampleBtn'),

    // Manual / Copilot paste panel
    copilotPastePanel: document.getElementById('copilotPastePanel'),
    copilotResponseInput: document.getElementById('copilotResponseInput'),
    parseCopilotResponseBtn: document.getElementById('parseCopilotResponseBtn'),
    copilotParseStatus: document.getElementById('copilotParseStatus'),

    // Fallback manual-copy prompt panel
    manualPromptPanel: document.getElementById('manualPromptPanel'),
    manualPromptOutput: document.getElementById('manualPromptOutput'),
  });
}

function initializeToggleButtons() {
  [
    [els.toggleLlmSectionBtn, els.llmSectionBody],
    [els.toggleAdvancedPanelBtn, els.advancedPanelBody],
    [els.toggleJsonSectionBtn, els.jsonSectionBody],
  ].forEach(([button, body]) => {
    if (button && body) {
      setToggleButtonState(button, !body.classList.contains('hidden'));
    }
  });
}

function setToggleButtonState(toggleButton, isExpanded) {
  if (!toggleButton) {
    return;
  }

  const expandedIcon = toggleButton.dataset.expandedIcon || '▾';
  const collapsedIcon = toggleButton.dataset.collapsedIcon || '▸';
  const expandedLabel = toggleButton.dataset.expandedLabel || 'Collapse section';
  const collapsedLabel = toggleButton.dataset.collapsedLabel || 'Expand section';

  toggleButton.setAttribute('aria-expanded', String(isExpanded));
  toggleButton.textContent = isExpanded ? expandedIcon : collapsedIcon;
  toggleButton.setAttribute('aria-label', isExpanded ? expandedLabel : collapsedLabel);
  toggleButton.title = isExpanded ? expandedLabel : collapsedLabel;
}

function toggleSection(toggleButton, sectionBody) {
  const isHidden = sectionBody.classList.contains('hidden');
  setSectionExpanded(toggleButton, sectionBody, isHidden);
}

function setSectionExpanded(toggleButton, sectionBody, shouldExpand) {
  if (!toggleButton || !sectionBody) {
    return;
  }

  sectionBody.classList.toggle('hidden', !shouldExpand);
  setToggleButtonState(toggleButton, shouldExpand);
}

function getDefaultSettings() {
  return {
    llmProvider: DEFAULT_PROVIDER,
    llmEndpoint: GOOGLE_ENDPOINT,
    llmModel: GOOGLE_SLOW_MODEL,
    llmApiKey: '',
    llmTemperature: DEFAULT_TEMPERATURE,
    llmMaxTokens: DEFAULT_MAX_TOKENS,
    systemPromptInput: DEFAULT_SYSTEM_PROMPT,
    chatInput: COMPREHENSIVE_CHAT_PROMPT,
    chatPromptPreset: DEFAULT_CHAT_PROMPT_PRESET,
    styleGuideInput: DEFAULT_STYLE_GUIDE,
    documentContextInput: '',
    lastLoadedSourceType: 'selection',
    selectedSourceMode: 'selection',

    googleSlowEndpoint: GOOGLE_ENDPOINT,
    googleSlowModel: GOOGLE_SLOW_MODEL,
    googleSlowApiKey: '',

    googleFastEndpoint: GOOGLE_ENDPOINT,
    googleFastModel: GOOGLE_FAST_MODEL,
    googleFastApiKey: '',

    openaiEndpoint: OPENAI_ENDPOINT,
    openaiModel: OPENAI_MODEL,
    openaiApiKey: '',

    internalEndpoint: INTERNAL_ENDPOINT,
    internalModel: INTERNAL_MODEL,
    internalApiKey: '',

    customEndpoint: CUSTOM_DEFAULT_ENDPOINT,
    customModel: CUSTOM_DEFAULT_MODEL,
    customApiKey: '',
  };
}

function hydrateLlmSettings() {
  const defaults = getDefaultSettings();
  let settings = {};

  const candidateKeys = [STORAGE_KEY, ...LEGACY_STORAGE_KEYS];
  for (const key of candidateKeys) {
    try {
      const raw = localStorage.getItem(key);
      if (raw) {
        settings = JSON.parse(raw);
        break;
      }
    } catch (error) {
      settings = {};
    }
  }

  const merged = migrateStoredSettings({
    ...defaults,
    ...settings,
  });

  const provider = PRESET_META[merged.llmProvider] ? merged.llmProvider : DEFAULT_PROVIDER;
  els.llmProvider.value = provider;
  els.llmTemperature.value = merged.llmTemperature || DEFAULT_TEMPERATURE;
  els.llmMaxTokens.value = merged.llmMaxTokens || DEFAULT_MAX_TOKENS;
  els.systemPromptInput.value = merged.systemPromptInput || DEFAULT_SYSTEM_PROMPT;
  els.chatInput.value = typeof merged.chatInput === 'string' ? merged.chatInput : COMPREHENSIVE_CHAT_PROMPT;
  els.styleGuideInput.value = typeof merged.styleGuideInput === 'string' ? merged.styleGuideInput : DEFAULT_STYLE_GUIDE;
  els.documentContextInput.value = typeof merged.documentContextInput === 'string' ? merged.documentContextInput : '';
  state.selectedSourceMode = merged.selectedSourceMode || 'selection';
  state.contextSourceMode = merged.lastLoadedSourceType || state.selectedSourceMode || 'selection';
  setSelectedSourceMode(state.selectedSourceMode, { persist: false });
  updateStyleGuideBadge();

  applyPresetFieldsToUi(provider, merged);
  applyDefaultApiKeySectionVisibility();

  const detectedChatPresetId = findMatchingChatPromptPresetId(els.chatInput.value);
  els.chatPromptPreset.value = detectedChatPresetId || merged.chatPromptPreset || DEFAULT_CHAT_PROMPT_PRESET;

  syncChatPromptPresetSelectionFromText();
  persistLlmSettings();
}

function applyDefaultApiKeySectionVisibility() {
  if (els.apiKeyModalInput && els.llmApiKey) {
    els.apiKeyModalInput.value = els.llmApiKey.value;
  }

  maybePromptForApiKey();
}


function migrateStoredSettings(settings) {
  const migrated = { ...settings };

  if (migrated.llmProvider === 'gemini') {
    migrated.llmProvider = PRESET_GOOGLE_SLOW;
  }

  const legacyGeminiEndpoint = migrated.geminiEndpoint || GOOGLE_ENDPOINT;
  const legacyGeminiModel = migrated.geminiModel || GOOGLE_SLOW_MODEL;
  const legacyGeminiApiKey = migrated.geminiApiKey || migrated.llmApiKey || '';

  migrated.googleSlowEndpoint = migrated.googleSlowEndpoint || legacyGeminiEndpoint;
  migrated.googleSlowModel = migrated.googleSlowModel || (legacyGeminiModel === GOOGLE_FAST_MODEL ? GOOGLE_SLOW_MODEL : legacyGeminiModel);
  migrated.googleSlowApiKey = migrated.googleSlowApiKey || legacyGeminiApiKey;

  migrated.googleFastEndpoint = migrated.googleFastEndpoint || legacyGeminiEndpoint;
  migrated.googleFastModel = migrated.googleFastModel || GOOGLE_FAST_MODEL;
  migrated.googleFastApiKey = migrated.googleFastApiKey || legacyGeminiApiKey || migrated.googleSlowApiKey || '';

  migrated.openaiEndpoint =
    migrated.openaiEndpoint ||
    (migrated.llmProvider === PRESET_OPENAI ? migrated.llmEndpoint : '') ||
    OPENAI_ENDPOINT;
  migrated.openaiModel =
    migrated.openaiModel ||
    (migrated.llmProvider === PRESET_OPENAI ? migrated.llmModel : '') ||
    OPENAI_MODEL;
  migrated.openaiApiKey =
    migrated.openaiApiKey ||
    (migrated.llmProvider === PRESET_OPENAI ? migrated.llmApiKey : '') ||
    '';

  const legacyCustomEndpoint = migrated.customEndpoint || (migrated.llmProvider === PRESET_CUSTOM ? migrated.llmEndpoint : '') || CUSTOM_DEFAULT_ENDPOINT;
  const legacyCustomModel = migrated.customModel || (migrated.llmProvider === PRESET_CUSTOM ? migrated.llmModel : '') || CUSTOM_DEFAULT_MODEL;
  const legacyCustomApiKey = migrated.customApiKey || (migrated.llmProvider === PRESET_CUSTOM ? migrated.llmApiKey : '') || '';

  migrated.customEndpoint = legacyCustomEndpoint;
  migrated.customModel = legacyCustomModel;
  migrated.customApiKey = legacyCustomApiKey;

  migrated.internalEndpoint = migrated.internalEndpoint || INTERNAL_ENDPOINT;
  migrated.internalModel = migrated.internalModel || INTERNAL_MODEL;
  migrated.internalApiKey = migrated.internalApiKey || '';

  migrated.documentContextInput = typeof migrated.documentContextInput === 'string' ? migrated.documentContextInput : '';
  migrated.styleGuideInput = typeof migrated.styleGuideInput === 'string' ? migrated.styleGuideInput : DEFAULT_STYLE_GUIDE;
  migrated.selectedSourceMode = migrated.selectedSourceMode || (migrated.lastLoadedSourceType === 'body' ? 'body' : 'selection');
  migrated.lastLoadedSourceType = migrated.lastLoadedSourceType || migrated.selectedSourceMode || 'selection';
  if (!['selection', 'body', 'manual'].includes(migrated.lastLoadedSourceType)) {
    migrated.lastLoadedSourceType = migrated.selectedSourceMode || 'selection';
  }

  return migrated;
}

function persistLlmSettings() {
  const existing = migrateStoredSettings({
    ...getDefaultSettings(),
    ...readStoredSettings(),
  });

  const provider = PRESET_META[els.llmProvider.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  const detectedChatPresetId = findMatchingChatPromptPresetId(els.chatInput.value);

  const settings = {
    ...getDefaultSettings(),
    ...existing,
    llmProvider: provider,
    llmEndpoint: els.llmEndpoint.value.trim(),
    llmModel: els.llmModel.value.trim(),
    llmApiKey: els.llmApiKey.value,
    llmTemperature: els.llmTemperature.value.trim(),
    llmMaxTokens: els.llmMaxTokens.value.trim(),
    systemPromptInput: els.systemPromptInput.value,
    chatInput: els.chatInput.value,
    chatPromptPreset: detectedChatPresetId || 'custom',
    styleGuideInput: els.styleGuideInput?.value || DEFAULT_STYLE_GUIDE,
    documentContextInput: els.documentContextInput.value,
    lastLoadedSourceType: state.contextSourceMode || getSelectedSourceMode(),
    selectedSourceMode: getSelectedSourceMode(),
  };

  const endpoint = els.llmEndpoint.value.trim();
  const model = els.llmModel.value.trim();
  const apiKey = els.llmApiKey.value;

  if (provider === PRESET_GOOGLE_SLOW) {
    settings.googleSlowEndpoint = endpoint || GOOGLE_ENDPOINT;
    settings.googleSlowModel = model || GOOGLE_SLOW_MODEL;
    settings.googleSlowApiKey = apiKey;
    settings.googleFastApiKey = apiKey;
  } else if (provider === PRESET_GOOGLE_FAST) {
    settings.googleFastEndpoint = endpoint || GOOGLE_ENDPOINT;
    settings.googleFastModel = model || GOOGLE_FAST_MODEL;
    settings.googleFastApiKey = apiKey;
    settings.googleSlowApiKey = apiKey;
  } else if (provider === PRESET_OPENAI) {
    settings.openaiEndpoint = endpoint || OPENAI_ENDPOINT;
    settings.openaiModel = model || OPENAI_MODEL;
    settings.openaiApiKey = apiKey;
  } else if (provider === PRESET_INTERNAL) {
    settings.internalEndpoint = endpoint || INTERNAL_ENDPOINT;
    settings.internalModel = model || INTERNAL_MODEL;
    settings.internalApiKey = apiKey;
  } else if (provider !== PRESET_MANUAL) {
    settings.customEndpoint = endpoint || CUSTOM_DEFAULT_ENDPOINT;
    settings.customModel = model || CUSTOM_DEFAULT_MODEL;
    settings.customApiKey = apiKey;
  }

  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  syncChatPromptPresetSelectionFromText();
  updateProviderUi();
}

function readStoredSettings() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
  } catch (error) {
    return {};
  }
}

function getSelectedSourceMode() {
  return els.sourceModeBody?.checked ? 'body' : 'selection';
}

function setSelectedSourceMode(mode, { persist = true } = {}) {
  const safeMode = mode === 'body' ? 'body' : 'selection';
  state.selectedSourceMode = safeMode;

  if (els.sourceModeSelection) {
    els.sourceModeSelection.checked = safeMode === 'selection';
  }
  if (els.sourceModeBody) {
    els.sourceModeBody.checked = safeMode === 'body';
  }

  updateSourceModeCards();
  updateSourceStatusSummary();

  if (persist) {
    persistLlmSettings();
  }
}

function updateSourceModeCards() {
  const selectedMode = getSelectedSourceMode();
  els.sourceSelectionCard?.classList.toggle('is-selected', selectedMode === 'selection');
  els.sourceBodyCard?.classList.toggle('is-selected', selectedMode === 'body');
}

function formatCharacterCount(text) {
  return `${text.length.toLocaleString()} char${text.length === 1 ? '' : 's'}`;
}

function updateSourceStatusSummary(customMessage = '') {
  if (!els.sourceStatusText || !els.sourceStatusBadge) {
    return;
  }

  const cachedText = String(els.documentContextInput?.value || '').trim();
  const manualOverrideActive = state.contextSourceMode === 'manual' && Boolean(cachedText);
  const selectedMode = getSelectedSourceMode();

  let badgeText = selectedMode === 'body' ? 'Whole document' : 'Highlighted text';
  let badgeTone = 'neutral';
  let message = customMessage;

  if (!message) {
    if (manualOverrideActive) {
      badgeText = 'Advanced override';
      badgeTone = 'warning';
      message = `Using the pasted advanced source text (${formatCharacterCount(cachedText)}) until you clear it.`;
    } else if (state.contextSourceMode === 'body' && cachedText) {
      badgeTone = 'success';
      message = `Ready to send the whole document. Latest cached preview: ${formatCharacterCount(cachedText)}.`;
    } else if (state.contextSourceMode === 'selection' && cachedText) {
      badgeTone = 'success';
      message = `Ready to send the current highlighted text. Latest cached preview: ${formatCharacterCount(cachedText)}.`;
    } else if (selectedMode === 'body') {
      message = 'The add-in will read the full Word document when you send the request.';
    } else {
      message = 'The add-in will read the current Word selection when you send the request.';
    }
  }

  els.sourceStatusText.textContent = message;
  els.sourceStatusBadge.textContent = badgeText;
  els.sourceStatusBadge.className = `mini-badge ${badgeTone}`;

  if (els.clearDocumentContextBtn) {
    els.clearDocumentContextBtn.classList.toggle('hidden', !manualOverrideActive);
  }
}

function updateStyleGuideBadge() {
  if (!els.styleGuideSavedBadge) {
    return;
  }

  const hasStyleGuide = Boolean(String(els.styleGuideInput?.value || '').trim());
  els.styleGuideSavedBadge.textContent = hasStyleGuide ? 'Saved locally' : 'Optional';
  els.styleGuideSavedBadge.className = `mini-badge ${hasStyleGuide ? 'success' : 'neutral'}`;
}

async function getDocumentTextForCurrentMode({ refresh = true } = {}) {
  if (!refresh && state.contextSourceMode === 'manual' && els.documentContextInput?.value.trim()) {
    return els.documentContextInput.value.trim();
  }

  const mode = getSelectedSourceMode();
  const text = mode === 'body' ? await getWordBodyText() : await getWordSelectionText();

  if (!text.trim()) {
    throw new Error(mode === 'body' ? 'The document body is empty.' : 'Highlight some text in Word first.');
  }

  els.documentContextInput.value = text;
  state.contextSourceMode = mode;
  updateSourceStatusSummary(`Ready to send ${mode === 'body' ? 'the whole document' : 'the highlighted text'} (${formatCharacterCount(text)}).`);
  persistLlmSettings();
  return text;
}

function onSourceModeChanged(event) {
  const nextMode = event?.target?.value === 'body' ? 'body' : 'selection';
  state.selectedSourceMode = nextMode;
  updateSourceModeCards();
  updateSourceStatusSummary();
  persistLlmSettings();
}

async function onRefreshSourcePreview() {
  try {
    const text = await getDocumentTextForCurrentMode({ refresh: true });
    setLlmStatus(`Refreshed ${getSelectedSourceMode() === 'body' ? 'the full document' : 'the current selection'} preview (${formatCharacterCount(text)}).`, 'success');
  } catch (error) {
    setLlmStatus(`Could not refresh the source preview: ${error.message}`, 'error');
    updateSourceStatusSummary();
  }
}

function onClearManualSource() {
  els.documentContextInput.value = '';
  state.contextSourceMode = getSelectedSourceMode();
  updateSourceStatusSummary();
  persistLlmSettings();
  setLlmStatus('Cleared the advanced source override.', 'success');
}

function onClearStyleGuide() {
  if (!els.styleGuideInput) {
    return;
  }

  els.styleGuideInput.value = DEFAULT_STYLE_GUIDE;
  updateStyleGuideBadge();
  persistLlmSettings();
  setLlmStatus('Cleared the saved style guide.', 'success');
}

function onProviderChanged() {
  const provider = PRESET_META[els.llmProvider.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  state.apiKeyPromptDismissedForProvider = null;
  const stored = migrateStoredSettings({
    ...getDefaultSettings(),
    ...readStoredSettings(),
  });

  applyPresetFieldsToUi(provider, stored);
  applyDefaultApiKeySectionVisibility();
  persistLlmSettings();
  setLlmStatus(PRESET_META[provider].statusMessage, 'success');
}

function onApplyProviderPreset() {
  const provider = PRESET_META[els.llmProvider.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  const defaults = getDefaultSettings();

  if (provider === PRESET_GOOGLE_SLOW) {
    els.llmEndpoint.value = GOOGLE_ENDPOINT;
    els.llmModel.value = GOOGLE_SLOW_MODEL;
    setLlmStatus('Applied the Google smart preset values. Enter your Google API key before sending.', 'success');
  } else if (provider === PRESET_GOOGLE_FAST) {
    els.llmEndpoint.value = GOOGLE_ENDPOINT;
    els.llmModel.value = GOOGLE_FAST_MODEL;
    setLlmStatus('Applied the Google fast preset values. Enter your Google API key before sending.', 'success');
  } else if (provider === PRESET_OPENAI) {
    els.llmEndpoint.value = OPENAI_ENDPOINT;
    els.llmModel.value = OPENAI_MODEL;
    setLlmStatus('Applied the ChatGPT / OpenAI preset values. Enter your OpenAI API key before sending.', 'success');
  } else if (provider === PRESET_INTERNAL) {
    els.llmEndpoint.value = INTERNAL_ENDPOINT;
    els.llmModel.value = INTERNAL_MODEL;
    setLlmStatus('Applied the internal preset values. The API key is optional for your local endpoint.', 'success');
  } else if (provider === PRESET_MANUAL) {
    els.llmEndpoint.value = '';
    els.llmModel.value = '';
    setLlmStatus('Manual / Copilot mode. Click Send to copy the full prompt to your clipboard.', 'success');
  } else {
    els.llmEndpoint.value = defaults.customEndpoint || CUSTOM_DEFAULT_ENDPOINT;
    els.llmModel.value = defaults.customModel || CUSTOM_DEFAULT_MODEL;
    setLlmStatus('Applied the default custom/OpenAI-compatible preset values.', 'success');
  }

  applyDefaultApiKeySectionVisibility();
  persistLlmSettings();
}

function onSaveAsCustom() {
  const endpoint = els.llmEndpoint.value.trim();
  const model = els.llmModel.value.trim();
  const apiKey = els.llmApiKey.value;

  els.llmProvider.value = PRESET_CUSTOM;
  els.llmEndpoint.value = endpoint;
  els.llmModel.value = model;
  els.llmApiKey.value = apiKey;

  applyDefaultApiKeySectionVisibility();
  persistLlmSettings();
  setLlmStatus('Saved the current endpoint, model, and API key as your custom preset.', 'success');
}

function applyPresetFieldsToUi(provider, settings) {
  const safeProvider = PRESET_META[provider] ? provider : DEFAULT_PROVIDER;

  if (safeProvider === PRESET_GOOGLE_SLOW) {
    els.llmEndpoint.value = settings.googleSlowEndpoint || GOOGLE_ENDPOINT;
    els.llmModel.value = settings.googleSlowModel || GOOGLE_SLOW_MODEL;
    els.llmApiKey.value = settings.googleSlowApiKey || '';
  } else if (safeProvider === PRESET_GOOGLE_FAST) {
    els.llmEndpoint.value = settings.googleFastEndpoint || GOOGLE_ENDPOINT;
    els.llmModel.value = settings.googleFastModel || GOOGLE_FAST_MODEL;
    els.llmApiKey.value = settings.googleFastApiKey || settings.googleSlowApiKey || '';
  } else if (safeProvider === PRESET_OPENAI) {
    els.llmEndpoint.value = settings.openaiEndpoint || OPENAI_ENDPOINT;
    els.llmModel.value = settings.openaiModel || OPENAI_MODEL;
    els.llmApiKey.value = settings.openaiApiKey || '';
  } else if (safeProvider === PRESET_INTERNAL) {
    els.llmEndpoint.value = settings.internalEndpoint || INTERNAL_ENDPOINT;
    els.llmModel.value = settings.internalModel || INTERNAL_MODEL;
    els.llmApiKey.value = settings.internalApiKey || '';
  } else if (safeProvider === PRESET_MANUAL) {
    els.llmEndpoint.value = '';
    els.llmModel.value = '';
    els.llmApiKey.value = '';
  } else {
    els.llmEndpoint.value = settings.customEndpoint || CUSTOM_DEFAULT_ENDPOINT;
    els.llmModel.value = settings.customModel || CUSTOM_DEFAULT_MODEL;
    els.llmApiKey.value = settings.customApiKey || '';
  }
}

function updateProviderUi() {
  if (!els.providerPresetBadge || !els.llmProvider) {
    return;
  }

  const provider = PRESET_META[els.llmProvider.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  const meta = PRESET_META[provider];

  els.providerPresetBadge.textContent = meta.badge;
  els.providerPresetBadge.className = `mini-badge ${meta.badgeTone || 'neutral'}`;
  els.llmEndpoint.placeholder = meta.endpoint;
  els.llmModel.placeholder = meta.model;

  if (els.llmApiKey) {
    els.llmApiKey.placeholder = meta.apiKeyPlaceholder;
  }

  if (els.apiKeyRequirementBadge) {
    els.apiKeyRequirementBadge.textContent = meta.apiKeyRequired ? 'Required' : 'Optional';
    els.apiKeyRequirementBadge.className = `mini-badge ${meta.apiKeyRequired ? 'warning' : 'neutral'}`;
  }

  if (els.apiKeyHint) {
    els.apiKeyHint.textContent = meta.apiKeyHint;
  }

  const isManual = provider === PRESET_MANUAL;
  if (els.sendAndLoadBtn) {
    els.sendAndLoadBtn.textContent = isManual ? 'Copy full prompt' : 'Send';
    els.sendAndLoadBtn.dataset.defaultLabel = els.sendAndLoadBtn.textContent;
  }
  if (!isManual && els.copilotPastePanel) {
    els.copilotPastePanel.classList.add('hidden');
  }
}

function onChatPromptPresetChanged() {
  const presetId = els.chatPromptPreset.value;
  updateChatPromptPresetUi(presetId);

  if (presetId === 'custom') {
    persistLlmSettings();
    setLlmStatus('Custom prompt selected. The current Chat / edit request text will be kept as-is.', 'neutral');
    return;
  }

  const preset = getChatPromptPreset(presetId);
  els.chatInput.value = preset.prompt;
  persistLlmSettings();
  setLlmStatus(`Loaded the ${preset.label.toLowerCase()} request into the Chat / edit request box.`, 'success');
}

function syncChatPromptPresetSelectionFromText() {
  if (!els.chatPromptPreset) {
    return;
  }

  const matchedPresetId = findMatchingChatPromptPresetId(els.chatInput.value) || 'custom';
  if (els.chatPromptPreset.value !== matchedPresetId) {
    els.chatPromptPreset.value = matchedPresetId;
  }

  updateChatPromptPresetUi(matchedPresetId);
}

function updateChatPromptPresetUi(presetId) {
  const preset = getChatPromptPreset(presetId);

  if (els.chatPromptMenuBtn) {
    els.chatPromptMenuBtn.textContent = 'Presets ▾';
    els.chatPromptMenuBtn.title = `Choose a ready-made edit request preset. Current: ${preset.label}.`;
    els.chatPromptMenuBtn.setAttribute('aria-label', `Prompt presets. Current: ${preset.label}.`);
  }
}

function getChatPromptPreset(presetId) {
  return CHAT_PROMPT_PRESETS[presetId] || CHAT_PROMPT_PRESETS.custom;
}

function findMatchingChatPromptPresetId(promptText) {
  const text = String(promptText || '');
  const presetIds = Object.keys(CHAT_PROMPT_PRESETS).filter((key) => key !== 'custom');

  for (const presetId of presetIds) {
    if (CHAT_PROMPT_PRESETS[presetId].prompt === text) {
      return presetId;
    }
  }

  return '';
}

function onToggleChatPromptMenu(event) {
  event.preventDefault();
  event.stopPropagation();

  const shouldOpen = !state.promptMenuOpen;
  setChatPromptMenuOpen(shouldOpen);
}

function setChatPromptMenuOpen(isOpen) {
  state.promptMenuOpen = Boolean(isOpen);

  if (els.chatPromptMenu) {
    els.chatPromptMenu.classList.toggle('hidden', !state.promptMenuOpen);
  }
  if (els.chatPromptMenuBtn) {
    els.chatPromptMenuBtn.setAttribute('aria-expanded', String(state.promptMenuOpen));
  }
}

function onChatPromptMenuItemSelected(presetId) {
  if (!els.chatPromptPreset) {
    return;
  }

  els.chatPromptPreset.value = presetId;
  setChatPromptMenuOpen(false);
  onChatPromptPresetChanged();
}

function onDocumentClick(event) {
  if (!state.promptMenuOpen) {
    return;
  }

  const withinMenu = els.chatPromptMenu?.contains(event.target);
  const withinButton = els.chatPromptMenuBtn?.contains(event.target);
  if (!withinMenu && !withinButton) {
    setChatPromptMenuOpen(false);
  }
}

function onDocumentKeydown(event) {
  if (event.key === 'Escape') {
    if (state.promptMenuOpen) {
      setChatPromptMenuOpen(false);
    }

    if (els.apiKeyModal && !els.apiKeyModal.classList.contains('hidden')) {
      closeApiKeyModal();
    }
  }
}

function maybePromptForApiKey({ force = false } = {}) {
  const provider = PRESET_META[els.llmProvider?.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  const requiresApiKey = isApiKeyRequiredForPreset(provider);
  const hasApiKey = Boolean(els.llmApiKey?.value.trim());

  if (!requiresApiKey || hasApiKey) {
    return false;
  }

  if (!force && state.apiKeyPromptDismissedForProvider === provider) {
    return false;
  }

  openApiKeyModal();
  return true;
}

function openApiKeyModal() {
  if (!els.apiKeyModal) {
    return;
  }

  els.apiKeyModal.classList.remove('hidden');
  els.apiKeyModal.setAttribute('aria-hidden', 'false');
  if (els.apiKeyModalInput && els.llmApiKey) {
    els.apiKeyModalInput.value = els.llmApiKey.value;
  }
  if (els.apiKeyModalStatus) {
    els.apiKeyModalStatus.textContent = '';
    els.apiKeyModalStatus.className = 'hint modal-status';
  }

  setTimeout(() => {
    els.apiKeyModalInput?.focus();
  }, 0);
}

function closeApiKeyModal() {
  if (!els.apiKeyModal) {
    return;
  }

  const provider = PRESET_META[els.llmProvider?.value] ? els.llmProvider.value : DEFAULT_PROVIDER;
  if (!els.llmApiKey?.value.trim()) {
    state.apiKeyPromptDismissedForProvider = provider;
  }

  els.apiKeyModal.classList.add('hidden');
  els.apiKeyModal.setAttribute('aria-hidden', 'true');
}

function onSaveApiKeyFromModal() {
  const apiKey = String(els.apiKeyModalInput?.value || '').trim();

  if (!apiKey) {
    if (els.apiKeyModalStatus) {
      els.apiKeyModalStatus.textContent = 'Enter an API key to continue.';
      els.apiKeyModalStatus.className = 'hint modal-status error';
    }
    return;
  }

  els.llmApiKey.value = apiKey;
  state.apiKeyPromptDismissedForProvider = null;
  persistLlmSettings();
  closeApiKeyModal();
  setLlmStatus('API key saved locally in settings.', 'success');
}

function onResetSystemPrompt() {
  els.systemPromptInput.value = DEFAULT_SYSTEM_PROMPT;
  persistLlmSettings();
  setLlmStatus('System instruction reset to the default comprehensive tracked-change JSON prompt.', 'success');
}

function insertSampleJson() {
  els.jsonInput.value = JSON.stringify(
    [
      {
        group: 1,
        original: 'This report is divided into the following chapters:',
        replacement: 'This report is organized as follows:',
        original_span: 'divided into the following chapters',
        replacement_span: 'organized as follows',
        reasoning_summary:
          'This change makes the wording more concise and formal. It also uses a cleaner phrase that reads more naturally in report introductions.',
      },
    ],
    null,
    2
  );
}

function onLoadSuggestions() {
  try {
    const count = loadSuggestionsFromRaw(els.jsonInput.value.trim());
    setGlobalStatus(`Loaded ${count} suggestion${count === 1 ? '' : 's'}.`, 'success');
  } catch (error) {
    setGlobalStatus(error.message, 'error');
  }
}

async function onCopyOutputInstructions() {
  try {
    await copyTextToClipboard(OUTPUT_INSTRUCTIONS_TEMPLATE);
    setGlobalStatus('Copied output instructions to your clipboard.', 'success');
  } catch (error) {
    setGlobalStatus(`Could not copy output instructions: ${error.message}`, 'error');
  }
}

function loadSuggestionsFromRaw(raw) {
  if (!raw) {
    throw new Error('Paste a JSON array first.');
  }

  const parsed = parseSuggestionJson(raw);
  state.groups = parsed.map((item, index) => buildGroup(item, index));
  state.selectedId = state.groups.length ? state.groups[0].id : null;

  renderGroupList();
  renderSelectedGroup();

  return state.groups.length;
}

async function onCheckTracking() {
  try {
    const mode = await getTrackingMode();
    const label = friendlyTrackingMode(mode);
    const cls = mode === 'Off' ? 'warning' : 'success';
    setGlobalStatus(`Track Changes: ${label}.`, cls);
  } catch (error) {
    setGlobalStatus(`Could not read Track Changes mode: ${error.message}`, 'error');
  }
}

async function onLoadSelectionIntoContext() {
  try {
    const text = await getWordSelectionText();
    if (!text.trim()) {
      throw new Error('The current selection is empty.');
    }

    els.documentContextInput.value = text;
    state.contextSourceMode = 'manual';
    updateSourceStatusSummary(`Using the current selection as an advanced override (${formatCharacterCount(text)}).`);
    persistLlmSettings();
    setLlmStatus('Loaded the current Word selection into the advanced source box.', 'success');
  } catch (error) {
    setLlmStatus(`Could not load the current selection: ${error.message}`, 'error');
  }
}

async function onLoadBodyIntoContext() {
  try {
    const text = await getWordBodyText();
    if (!text.trim()) {
      throw new Error('The document body is empty.');
    }

    els.documentContextInput.value = text;
    state.contextSourceMode = 'manual';
    updateSourceStatusSummary(`Using the full document as an advanced override (${formatCharacterCount(text)}).`);
    persistLlmSettings();
    setLlmStatus('Loaded the full document text into the advanced source box.', 'success');
  } catch (error) {
    setLlmStatus(`Could not load the document text: ${error.message}`, 'error');
  }
}

function buildManualFullPrompt({ systemPrompt, documentText, userPrompt, styleGuide }) {
  const parts = [];

  if (systemPrompt.trim()) {
    parts.push(`=== INSTRUCTIONS ===\n${systemPrompt.trim()}`);
  }

  if (documentText.trim()) {
    parts.push(`=== DOCUMENT ===\n${documentText.trim()}`);
  }

  if (styleGuide && styleGuide.trim()) {
    parts.push(`=== STYLE GUIDE ===\n${styleGuide.trim()}`);
  }

  if (userPrompt && userPrompt.trim()) {
    parts.push(`=== EDIT REQUEST ===\n${userPrompt.trim()}`);
  }

  parts.push(`=== OUTPUT FORMAT ===\n${OUTPUT_INSTRUCTIONS_TEMPLATE}`);

  return parts.join('\n\n');
}

async function onCopyManualPrompt() {
  try {
    const systemPrompt = els.systemPromptInput.value.trim();
    const styleGuide = els.styleGuideInput?.value.trim() || '';
    const userPrompt = els.chatInput.value.trim();

    const hasManualSourceOverride = state.contextSourceMode === 'manual' && Boolean(els.documentContextInput.value.trim());
    const documentText = hasManualSourceOverride
      ? els.documentContextInput.value.trim()
      : await getDocumentTextForCurrentMode({ refresh: true });

    const fullPrompt = buildManualFullPrompt({ systemPrompt, documentText, userPrompt, styleGuide });

    let clipboardOk = false;
    try {
      await copyTextToClipboard(fullPrompt);
      clipboardOk = true;
    } catch (_) {
      // Clipboard blocked — fall through to manual panel.
    }

    if (clipboardOk) {
      if (els.copilotPastePanel) {
        els.copilotPastePanel.classList.remove('hidden');
        els.copilotResponseInput?.focus();
      }
      setLlmStatus('Full prompt copied to clipboard. Paste it into Copilot, then paste the response in the box below.', 'success');
      updateRequestFeedback('Copied', 'success', 'Full prompt copied to clipboard. Paste the AI response into the box that appeared below.');
    } else {
      if (els.manualPromptPanel && els.manualPromptOutput) {
        els.manualPromptOutput.value = fullPrompt;
        els.manualPromptPanel.classList.remove('hidden');
        els.manualPromptPanel.open = true;
        els.manualPromptOutput.focus();
        els.manualPromptOutput.select();
      }
      if (els.copilotPastePanel) {
        els.copilotPastePanel.classList.remove('hidden');
      }
      setLlmStatus('Clipboard blocked — copy the prompt from the box below manually.', 'warning');
      updateRequestFeedback('Manual copy needed', 'warning', 'Clipboard access was blocked. The prompt is shown below — copy it manually, then paste the AI response.');
    }
  } catch (error) {
    setLlmStatus(`Could not build prompt: ${error.message}`, 'error');
    updateRequestFeedback('Failed', 'error', `Could not build prompt: ${error.message}`);
  }
}

function onParseCopilotResponse() {
  const responseText = els.copilotResponseInput?.value.trim();
  if (!responseText) {
    setCopilotParseStatus('Paste the AI response first.', 'warning');
    return;
  }

  try {
    const count = loadSuggestionsFromRaw(responseText);
    if (els.llmResponse) {
      els.llmResponse.value = responseText;
    }
    updateResponseAutoLoadBadge(`Loaded ${count} suggestion${count === 1 ? '' : 's'}`, 'success');
    setCopilotParseStatus(`Loaded ${count} suggestion${count === 1 ? '' : 's'}.`, 'success');
    setLlmStatus('Parsed response and loaded suggestions.', 'success');
    els.groupList?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    setCopilotParseStatus(`Could not parse the response: ${error.message}`, 'error');
    setLlmStatus(`Could not parse the response: ${error.message}`, 'error');
  }
}

function setCopilotParseStatus(message, tone) {
  if (!els.copilotParseStatus) {
    return;
  }
  els.copilotParseStatus.textContent = message;
  els.copilotParseStatus.className = `status ${tone}`;
  els.copilotParseStatus.classList.remove('hidden');
}

async function onSendChatToLlm({ focusSuggestions }) {
  try {
    const provider = PRESET_META[els.llmProvider.value] ? els.llmProvider.value : DEFAULT_PROVIDER;

    if (provider === PRESET_MANUAL) {
      await onCopyManualPrompt();
      return;
    }

    const endpoint = els.llmEndpoint.value.trim();
    const model = els.llmModel.value.trim() || PRESET_META[provider].model;
    const apiKey = els.llmApiKey.value.trim();
    const systemPrompt = els.systemPromptInput.value.trim();
    const styleGuide = els.styleGuideInput?.value.trim() || '';
    const userPrompt = els.chatInput.value.trim();
    const temperature = parseOptionalNumber(els.llmTemperature.value.trim(), DEFAULT_TEMPERATURE, 'temperature');
    const maxTokens = parseOptionalInteger(els.llmMaxTokens.value.trim(), DEFAULT_MAX_TOKENS, 'max tokens');

    if (isApiKeyRequiredForPreset(provider) && !apiKey) {
      maybePromptForApiKey({ force: true });
      throw new Error('Enter an API key first.');
    }

    if (!endpoint) {
      ensureSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody);
      throw new Error('Enter a chat completions endpoint first.');
    }

    if (!systemPrompt && !userPrompt) {
      ensureSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody);
      throw new Error('Provide either a system instruction or a user message before sending.');
    }

    const hasManualSourceOverride = state.contextSourceMode === 'manual' && Boolean(els.documentContextInput.value.trim());
    const documentText = hasManualSourceOverride
      ? els.documentContextInput.value.trim()
      : await getDocumentTextForCurrentMode({ refresh: true });

    if (!documentText.trim() && !userPrompt) {
      throw new Error('Provide source text, a chat request, or both before sending.');
    }

    persistLlmSettings();
    setSendButtonsBusy(true, focusSuggestions);
    updateRequestFeedback('Sending', 'neutral', 'Request sent to the LLM. Waiting for a response…', true);
    setLlmStatus('Sending request to LLM...', 'neutral');
    updateResponseAutoLoadBadge('Waiting for response', 'neutral');

    const requestPayloads = buildCandidatePayloads({ model, systemPrompt, documentText, userPrompt, styleGuide, temperature, maxTokens });
    const headers = {
      'Content-Type': 'application/json',
    };

    if (apiKey) {
      headers.Authorization = `Bearer ${apiKey}`;
    }

    const { responseText } = await sendChatRequest(endpoint, headers, requestPayloads);

    updateRequestFeedback('Received', 'neutral', 'Response received. Validating and loading suggestions…', true);

    let json;
    try {
      json = JSON.parse(responseText);
    } catch (error) {
      ensureSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody);
      throw new Error('The endpoint did not return valid JSON.');
    }

    const assistantText = extractAssistantText(json).trim();
    if (!assistantText) {
      ensureSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody);
      throw new Error('The LLM response did not contain any assistant text.');
    }

    els.llmResponse.value = assistantText;

    const autoLoadResult = autoLoadLlmResponse(assistantText, { focusSuggestions });

    if (autoLoadResult.loaded) {
      updateResponseAutoLoadBadge(`Auto-loaded ${autoLoadResult.count} suggestion${autoLoadResult.count === 1 ? '' : 's'}`, 'success');
      updateRequestFeedback('Loaded', 'success', `Received response and loaded ${autoLoadResult.count} suggestion${autoLoadResult.count === 1 ? '' : 's'}.`);
      setLlmStatus('Response received and auto-loaded as suggestions. You can still review the raw response in advanced settings.', 'success');
    } else {
      ensureSectionExpanded(els.toggleAdvancedPanelBtn, els.advancedPanelBody);
      updateResponseAutoLoadBadge('Saved only', 'warning');
      updateRequestFeedback('Saved', 'warning', 'Response received and saved, but it could not be auto-loaded. Open advanced settings to review the raw output.');
      setLlmStatus(`Response received, but it was not auto-loaded: ${autoLoadResult.error}`, 'warning');
    }
  } catch (error) {
    updateResponseAutoLoadBadge('Request failed', 'error');
    updateRequestFeedback('Failed', 'error', `Request failed: ${error.message}`);
    setLlmStatus(`LLM request failed: ${error.message}`, 'error');
  } finally {
    setSendButtonsBusy(false, focusSuggestions);
  }
}

function isApiKeyRequiredForPreset(provider) {
  return Boolean(PRESET_META[provider]?.apiKeyRequired);
}

async function sendChatRequest(endpoint, headers, payloads) {
  let lastError = null;

  for (let i = 0; i < payloads.length; i += 1) {
    const payload = payloads[i];
    const response = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
    });

    const responseText = await response.text();

    if (response.ok) {
      return { responseText, attempt: i + 1 };
    }

    lastError = new Error(`HTTP ${response.status}: ${responseText || response.statusText}`);

    if (!shouldRetryWithAlternatePayload(responseText, i, payloads.length)) {
      throw lastError;
    }
  }

  throw lastError || new Error('The LLM request failed.');
}

function shouldRetryWithAlternatePayload(responseText, attemptIndex, payloadCount) {
  if (attemptIndex >= payloadCount - 1) {
    return false;
  }

  const text = String(responseText || '').toLowerCase();
  return text.includes('messages') && (text.includes('required') || text.includes('missing'));
}

function buildCandidatePayloads({ model, systemPrompt, documentText, userPrompt, styleGuide, temperature, maxTokens }) {
  const messages = buildLlmMessages(systemPrompt, documentText, userPrompt, styleGuide);
  if (!Array.isArray(messages) || !messages.length) {
    throw new Error('Could not build a valid messages array for the LLM request.');
  }

  const basePayload = {
    model,
    stream: false,
    temperature,
    max_tokens: maxTokens,
    messages,
  };

  const compactPayload = {
    model,
    messages,
    temperature,
    max_tokens: maxTokens,
    stream: false,
  };

  const alternateTokenPayload = {
    model,
    messages,
    temperature,
    max_completion_tokens: maxTokens,
    stream: false,
  };

  return [basePayload, compactPayload, alternateTokenPayload];
}

function autoLoadLlmResponse(responseText, { focusSuggestions }) {
  try {
    const count = loadSuggestionsFromRaw(responseText);
    els.jsonInput.value = responseText;
    setGlobalStatus(`Loaded ${count} suggestion${count === 1 ? '' : 's'} from the LLM response.`, 'success');

    if (focusSuggestions) {
      els.groupList?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    return { loaded: true, count };
  } catch (error) {
    return { loaded: false, error: error.message };
  }
}

function onCopyResponseToJson() {
  const responseText = els.llmResponse.value.trim();
  if (!responseText) {
    setLlmStatus('There is no LLM response to copy.', 'warning');
    return;
  }

  els.jsonInput.value = responseText;
  setGlobalStatus('Copied the LLM response into the JSON box.', 'success');
}

function onLoadResponseAsSuggestions() {
  const responseText = els.llmResponse.value.trim();
  if (!responseText) {
    setLlmStatus('There is no LLM response to load.', 'warning');
    return;
  }

  try {
    const count = loadSuggestionsFromRaw(responseText);
    els.jsonInput.value = responseText;
    updateResponseAutoLoadBadge(`Loaded ${count} suggestion${count === 1 ? '' : 's'}`, 'success');
    setGlobalStatus(`Loaded ${count} suggestion${count === 1 ? '' : 's'} from the LLM response.`, 'success');
    setLlmStatus('Loaded the edited LLM response as suggestions.', 'success');
  } catch (error) {
    updateResponseAutoLoadBadge('Manual load failed', 'error');
    setLlmStatus(`Could not load the LLM response as suggestions: ${error.message}`, 'error');
  }
}

function buildLlmMessages(systemPrompt, documentText, userPrompt, styleGuide) {
  const messages = [];

  if (systemPrompt) {
    messages.push({
      role: 'system',
      content: systemPrompt,
    });
  }

  const userContent = buildLlmUserMessage(documentText, userPrompt, styleGuide);
  if (!userContent.trim()) {
    throw new Error('The user message for the LLM request is empty.');
  }

  messages.push({
    role: 'user',
    content: userContent,
  });

  return messages;
}

function buildLlmUserMessage(documentText, userPrompt, styleGuide) {
  const parts = [];

  if (documentText.trim()) {
    parts.push(
      [
        'Use the following document text as the only source for any exact quotations, original values, and original_span values.',
        '<<<DOCUMENT',
        documentText.trim(),
        'DOCUMENT',
      ].join('\n')
    );
  }

  if (styleGuide && styleGuide.trim()) {
    parts.push(
      [
        'Follow this style guide unless a later instruction in this request explicitly overrides it.',
        '<<<STYLE_GUIDE',
        styleGuide.trim(),
        'STYLE_GUIDE',
      ].join('\n')
    );
  }

  if (userPrompt) {
    parts.push(userPrompt);
  } else if (documentText.trim()) {
    parts.push('Review the provided document text and return the requested edits.');
  }

  return parts.join('\n\n').trim();
}

function extractAssistantText(payload) {
  if (typeof payload === 'string') {
    return payload;
  }

  if (Array.isArray(payload?.choices) && payload.choices.length) {
    const firstChoice = payload.choices[0];
    const content = firstChoice?.message?.content;

    if (typeof content === 'string') {
      return content;
    }

    if (Array.isArray(content)) {
      return content
        .map((part) => {
          if (typeof part === 'string') {
            return part;
          }
          if (typeof part?.text === 'string') {
            return part.text;
          }
          if (typeof part?.content === 'string') {
            return part.content;
          }
          return '';
        })
        .join('');
    }
  }

  if (typeof payload?.output_text === 'string') {
    return payload.output_text;
  }

  if (typeof payload?.response === 'string') {
    return payload.response;
  }

  return JSON.stringify(payload, null, 2);
}

function parseOptionalNumber(value, fallback, label) {
  const effectiveValue = value || fallback;
  const parsed = Number(effectiveValue);

  if (!Number.isFinite(parsed)) {
    throw new Error(`Invalid ${label} value.`);
  }

  return parsed;
}

function parseOptionalInteger(value, fallback, label) {
  const effectiveValue = value || fallback;
  const parsed = Number(effectiveValue);

  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error(`Invalid ${label} value.`);
  }

  return parsed;
}

function ensureSectionExpanded(toggleButton, sectionBody) {
  if (!sectionBody.classList.contains('hidden')) {
    setToggleButtonState(toggleButton, true);
    return;
  }

  sectionBody.classList.remove('hidden');
  setToggleButtonState(toggleButton, true);
}

function updateResponseAutoLoadBadge(message, tone) {
  if (!els.responseAutoLoadBadge) {
    return;
  }

  els.responseAutoLoadBadge.textContent = message;
  els.responseAutoLoadBadge.className = `mini-badge ${tone || 'neutral'}`;
}

function updateRequestFeedback(stateLabel, tone, detail, isBusy = false) {
  if (!els.requestFeedback || !els.requestStateBadge || !els.requestProgressText) {
    return;
  }

  const isInitialPlaceholder = stateLabel === 'Ready' && detail === 'Nothing sent yet.';
  if (els.requestSection && !isInitialPlaceholder) {
    els.requestSection.classList.remove('hidden');
  }

  els.requestFeedback.className = `request-feedback ${tone || 'neutral'}${isBusy ? ' is-busy' : ''}`;
  els.requestStateBadge.textContent = stateLabel;
  els.requestStateBadge.className = `mini-badge ${tone || 'neutral'}`;
  els.requestProgressText.textContent = detail;
}

function setSendButtonsBusy(isBusy, focusSuggestions = false) {
  const activeButtons = [els.sendChatBtn, els.sendAndLoadBtn].filter(Boolean);
  if (!activeButtons.length) {
    return;
  }

  activeButtons.forEach((button) => {
    if (!button.dataset.defaultLabel) {
      button.dataset.defaultLabel = button.textContent;
    }
    button.disabled = isBusy;
    button.classList.remove('is-busy');
  });

  if (els.sendChatBtn) {
    els.sendChatBtn.classList.toggle('is-busy', isBusy && !focusSuggestions);
  }
  if (els.sendAndLoadBtn) {
    els.sendAndLoadBtn.classList.toggle('is-busy', isBusy);
  }

  if (isBusy) {
    if (els.sendChatBtn) {
      els.sendChatBtn.textContent = focusSuggestions ? 'Please wait…' : 'Sending…';
    }
    if (els.sendAndLoadBtn) {
      els.sendAndLoadBtn.textContent = 'Sending…';
    }
  } else {
    activeButtons.forEach((button) => {
      button.textContent = button.dataset.defaultLabel;
    });
  }
}



async function maybeRestoreSavedContextFromWord() {
  if (!els.documentContextInput?.value.trim()) {
    return;
  }

  try {
    if (state.contextSourceMode === 'selection') {
      const latestSelection = await getWordSelectionText();
      if (latestSelection.trim()) {
        els.documentContextInput.value = latestSelection;
        persistLlmSettings();
        updateSourceStatusSummary(`Ready to send the current highlighted text. Latest cached preview: ${formatCharacterCount(latestSelection)}.`);
        setLlmStatus('Restored your last selection and refreshed it from the current Word selection.', 'neutral');
        return;
      }

      updateSourceStatusSummary();
      setLlmStatus('Restored your last loaded selection text from local storage.', 'neutral');
      return;
    }

    if (state.contextSourceMode === 'body') {
      const latestBody = await getWordBodyText();
      if (latestBody.trim()) {
        els.documentContextInput.value = latestBody;
        persistLlmSettings();
        updateSourceStatusSummary(`Ready to send the whole document. Latest cached preview: ${formatCharacterCount(latestBody)}.`);
        setLlmStatus('Restored your last full-document source and refreshed it from the current Word document.', 'neutral');
        return;
      }

      updateSourceStatusSummary();
      setLlmStatus('Restored your last full-document source text from local storage.', 'neutral');
    }
  } catch (error) {
    setLlmStatus(`Restored saved source text from local storage, but could not refresh it from Word: ${error.message}`, 'warning');
  }
}

async function getWordSelectionText() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const ooxmlResult = selection.getOoxml();
    selection.load('text');
    await context.sync();

    const cleanFromOoxml = extractCleanVisibleTextFromOoxml(ooxmlResult.value);
    return cleanFromOoxml || selection.text || '';
  });
}

async function getWordBodyText() {
  return Word.run(async (context) => {
    const body = context.document.body;
    const ooxmlResult = body.getOoxml();
    body.load('text');
    await context.sync();

    const cleanFromOoxml = extractCleanVisibleTextFromOoxml(ooxmlResult.value);
    return cleanFromOoxml || body.text || '';
  });
}

function extractCleanVisibleTextFromOoxml(ooxml) {
  if (!ooxml || typeof ooxml !== 'string') {
    return '';
  }

  try {
    const parser = new DOMParser();
    const xml = parser.parseFromString(ooxml, 'application/xml');

    if (xml.getElementsByTagName('parsererror').length) {
      return '';
    }

    const bodyLikeNode =
      xml.getElementsByTagNameNS('*', 'body')[0] ||
      xml.documentElement;

    const parts = [];
    appendVisibleTextFromNode(bodyLikeNode, parts);

    return normalizeExtractedWordText(parts.join(''));
  } catch (error) {
    console.warn('Could not parse OOXML for clean text extraction:', error);
    return '';
  }
}

function appendVisibleTextFromNode(node, parts) {
  if (!node) {
    return;
  }

  if (node.nodeType === Node.TEXT_NODE) {
    parts.push(node.nodeValue || '');
    return;
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return;
  }

  const name = (node.localName || '').toLowerCase();

  if (name === 'del' || name === 'deltext' || name === 'movefrom') {
    return;
  }

  if (name === 't' || name === 'instrtext') {
    parts.push(node.textContent || '');
    return;
  }

  if (name === 'tab') {
    parts.push('\t');
    return;
  }

  if (name === 'br' || name === 'cr') {
    parts.push('\n');
    return;
  }

  const blockStartsOnNewLine = name === 'p' || name === 'tr';
  const endsWithTab = name === 'tc';

  if (blockStartsOnNewLine && parts.length && !endsWithLineBreak(parts)) {
    parts.push('\n');
  }

  for (const child of Array.from(node.childNodes || [])) {
    appendVisibleTextFromNode(child, parts);
  }

  if (endsWithTab && !endsWithTabChar(parts)) {
    parts.push('\t');
  }

  if (blockStartsOnNewLine && !endsWithLineBreak(parts)) {
    parts.push('\n');
  }
}

function endsWithLineBreak(parts) {
  const last = parts.length ? parts[parts.length - 1] : '';
  return /\n$/.test(last);
}

function endsWithTabChar(parts) {
  const last = parts.length ? parts[parts.length - 1] : '';
  return /\t$/.test(last);
}

function normalizeExtractedWordText(text) {
  return String(text || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

async function onFindSentence() {

  const group = getSelectedGroup();
  if (!group) return;

  group.confirmReady = false;

  try {
    const result = await getSentenceMatchInfo(group.original);
    group.matchCount = result.count;
    group.searchMode = result.mode;
    group.searchText = result.searchText || group.original;

    if (result.count === 0) {
      group.status = 'not found';
      updateGroupStatus(
        'Original sentence was not found in the document. Try a shorter original sentence or exclude footnote markers.',
        'error'
      );
    } else {
      group.status = result.count === 1 ? 'matched' : 'multiple matches';
      group.activeMatchIndex = Math.min(group.activeMatchIndex, result.count - 1);
      await selectSentenceMatch(group.searchText, group.activeMatchIndex, group.searchMode);
      updateGroupStatus(
        `Found ${result.count} sentence match${result.count === 1 ? '' : 'es'} using ${result.mode} search.`,
        result.count === 1 ? 'success' : 'warning'
      );
    }

    renderGroupList();
    renderSelectedGroup();
  } catch (error) {
    updateGroupStatus(`Find failed: ${error.message}`, 'error');
  }
}

async function onCycleMatch(step) {
  const group = getSelectedGroup();
  if (!group) return;

  if (!group.matchCount || group.matchCount < 2) {
    updateGroupStatus('There is no alternate sentence match to cycle through.', 'warning');
    return;
  }

  group.activeMatchIndex = (group.activeMatchIndex + step + group.matchCount) % group.matchCount;

  try {
    await selectSentenceMatch(group.searchText || group.original, group.activeMatchIndex, group.searchMode || 'exact');
    updateGroupStatus(
      `Selected sentence match ${group.activeMatchIndex + 1} of ${group.matchCount}.`,
      'success'
    );
    renderSelectedGroup();
  } catch (error) {
    updateGroupStatus(`Could not select the requested match: ${error.message}`, 'error');
  }
}

async function onCycleGroup(step) {
  if (!state.groups.length) {
    updateGroupStatus('Load suggestions before cycling groups.', 'warning');
    return;
  }

  const currentIndex = state.groups.findIndex((group) => group.id === state.selectedId);
  const safeCurrentIndex = currentIndex >= 0 ? currentIndex : 0;
  const nextIndex = Math.max(0, Math.min(safeCurrentIndex + step, state.groups.length - 1));

  if (nextIndex === safeCurrentIndex) {
    updateGroupStatus(
      step > 0 ? 'Already at the last group.' : 'Already at the first group.',
      'warning'
    );
    return;
  }

  state.selectedId = state.groups[nextIndex].id;
  resetAllConfirmReady();
  renderGroupList();
  renderSelectedGroup();
}


async function onApplySelected() {
  const group = getSelectedGroup();
  if (!group) return;

  const rationale = group.reasoning_summary || 'No edit rationale provided.';

  if (!group.confirmReady) {
    resetAllConfirmReady();
    group.confirmReady = true;
    updateGroupStatus(
      `Review the selected edit rationale and span preview, then click “Confirm apply replacement” to proceed. Rationale: ${rationale}`,
      'warning'
    );
    renderSelectedGroup();
    return;
  }

  try {
    const result = await applyMinimalReplacement(group);
    group.status = 'applied';
    group.confirmReady = false;

    updateGroupStatus(
      `Applied replacement. Track Changes: ${friendlyTrackingMode(result.trackingMode)}. Sentence matches: ${result.sentenceCount}. Span matches: ${result.spanCount}.`,
      result.trackingMode === 'Off' ? 'warning' : 'success'
    );

    renderGroupList();
    renderSelectedGroup();
  } catch (error) {
    group.status = 'error';
    group.confirmReady = false;
    updateGroupStatus(`Apply failed: ${error.message}`, 'error');
    renderGroupList();
    renderSelectedGroup();
  }
}

async function onApplyAllExactSingles() {
  if (!state.groups.length) {
    setGlobalStatus('Load suggestions before running batch apply.', 'warning');
    return;
  }

  let applied = 0;
  let skipped = 0;
  const failures = [];

  for (const group of state.groups) {
    try {
      const matchInfo = await getSentenceMatchInfo(group.original);
      group.matchCount = matchInfo.count;
      group.searchMode = matchInfo.mode;
      group.searchText = matchInfo.searchText || group.original;

      if (matchInfo.mode !== 'exact' || matchInfo.count !== 1) {
        group.status = matchInfo.count === 0 ? 'not found' : 'skipped';
        skipped += 1;
        continue;
      }

      const spanCheck = await getSpanMatchInfo(group.searchText, group.original_span, 0, 'exact');
      if (spanCheck.count !== 1) {
        group.status = 'skipped';
        skipped += 1;
        continue;
      }

      await applyMinimalReplacement(group);
      group.status = 'applied';
      applied += 1;
    } catch (error) {
      group.status = 'error';
      failures.push(`Group ${group.group}: ${error.message}`);
    }
  }

  renderGroupList();
  renderSelectedGroup();

  const summary = `Batch complete. Applied: ${applied}. Skipped: ${skipped}. Failed: ${failures.length}.`;
  if (failures.length) {
    setGlobalStatus(`${summary} First failure: ${failures[0]}`, 'warning');
  } else {
    setGlobalStatus(summary, applied ? 'success' : 'warning');
  }
}

function parseSuggestionJson(raw) {
  let text = raw.trim();
  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch) {
    text = fenceMatch[1].trim();
  }

  if (!text.startsWith('[')) {
    const start = text.indexOf('[');
    const end = text.lastIndexOf(']');
    if (start >= 0 && end > start) {
      text = text.slice(start, end + 1);
    }
  }

  const parsed = JSON.parse(text);
  if (!Array.isArray(parsed) || !parsed.length) {
    throw new Error('The JSON must be a non-empty array.');
  }

  return parsed;
}

function buildGroup(item, index) {
  if (!item || typeof item !== 'object') {
    throw new Error(`Suggestion ${index + 1} is not an object.`);
  }

  const original = cleanRequired(item.original, `Suggestion ${index + 1} is missing “original”.`);
  const replacement = cleanRequired(item.replacement, `Suggestion ${index + 1} is missing “replacement”.`);
  let originalSpan = cleanOptional(item.original_span);
  let replacementSpan = cleanOptional(item.replacement_span);

  const derived = deriveMinimalDiff(original, replacement);

  if (!originalSpan && !replacementSpan) {
    originalSpan = derived.originalSpan;
    replacementSpan = derived.replacementSpan;
  }

  if (!originalSpan && original !== replacement) {
    throw new Error(
      `Suggestion ${index + 1} needs “original_span”/“replacement_span” or a replaceable diff.`
    );
  }

  if (originalSpan && !original.includes(originalSpan)) {
    throw new Error(
      `Suggestion ${index + 1} has an original_span that does not appear inside original.`
    );
  }

  return {
    id: `group-${index + 1}`,
    group: item.group ?? index + 1,
    original,
    replacement,
    original_span: originalSpan,
    replacement_span: replacementSpan,
    reasoning_summary: cleanOptional(item.reasoning_summary),
    diff: derived,
    activeMatchIndex: 0,
    matchCount: null,
    searchMode: 'exact',
    searchText: original,
    status: 'loaded',
    confirmReady: false,
  };
}

function cleanRequired(value, errorMessage) {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(errorMessage);
  }
  return value;
}

function cleanOptional(value) {
  return typeof value === 'string' && value.length ? value : '';
}

function deriveMinimalDiff(original, replacement) {
  if (original === replacement) {
    return {
      originalSpan: '',
      replacementSpan: '',
      prefix: original,
      suffix: '',
    };
  }

  let start = 0;
  while (start < original.length && start < replacement.length && original[start] === replacement[start]) {
    start += 1;
  }

  let endOriginal = original.length - 1;
  let endReplacement = replacement.length - 1;
  while (
    endOriginal >= start &&
    endReplacement >= start &&
    original[endOriginal] === replacement[endReplacement]
  ) {
    endOriginal -= 1;
    endReplacement -= 1;
  }

  const originalSpan = original.slice(start, endOriginal + 1);
  const replacementSpan = replacement.slice(start, endReplacement + 1);

  return {
    originalSpan,
    replacementSpan,
    prefix: original.slice(0, start),
    suffix: original.slice(endOriginal + 1),
  };
}

function buildFallbackSearchTargets(text) {
  const cleaned = text.replace(/\s+/g, ' ').trim();
  const targets = [];

  const clauses = cleaned.split(/([.;:!?])/);
  let running = '';
  for (let i = 0; i < clauses.length; i += 1) {
    running += clauses[i];
    const candidate = running.trim();
    if (candidate.length >= 40 && candidate !== cleaned) {
      targets.push(candidate);
    }
  }

  const words = cleaned.split(' ');
  for (let n = Math.min(words.length, 18); n >= 6; n -= 1) {
    const candidate = words.slice(0, n).join(' ');
    if (candidate.length >= 30 && candidate !== cleaned) {
      targets.push(candidate);
    }
  }

  return [...new Set(targets)];
}


function getStatusTone(status) {
  const normalized = String(status || '').toLowerCase();

  if (normalized === 'applied') {
    return 'success';
  }

  if (normalized === 'error' || normalized === 'not found') {
    return 'error';
  }

  if (normalized === 'skipped') {
    return 'warning';
  }

  return 'neutral';
}

function renderGroupList() {
  els.groupList.innerHTML = '';

  if (!state.groups.length) {
    const emptyItem = document.createElement('li');
    emptyItem.className = 'group-list-empty';
    emptyItem.textContent = 'No suggestions loaded yet. Generate or paste suggestion JSON to begin.';
    els.groupList.appendChild(emptyItem);
    return;
  }

  for (const group of state.groups) {
    const li = document.createElement('li');
    li.className = group.id === state.selectedId ? 'selected' : '';
    li.dataset.groupId = group.id;
    li.tabIndex = 0;
    li.setAttribute('role', 'button');
    li.setAttribute('aria-pressed', group.id === state.selectedId ? 'true' : 'false');

    const selectGroup = () => {
      state.selectedId = group.id;
      resetAllConfirmReady();
      renderGroupList();
      renderSelectedGroup();
    };

    li.addEventListener('click', selectGroup);
    li.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        selectGroup();
      }
    });

    const topRow = document.createElement('div');
    topRow.className = 'group-row-top';

    const title = document.createElement('div');
    title.className = 'group-title';
    title.textContent = `Suggestion ${group.group}`;

    const statusPill = document.createElement('div');
    statusPill.className = `pill ${getStatusTone(group.status)}`;
    statusPill.textContent = `Status: ${group.status}`;

    topRow.appendChild(title);
    topRow.appendChild(statusPill);

    const snippet = document.createElement('div');
    snippet.className = 'group-snippet';
    snippet.textContent = truncate(group.original, 140);

    const markupPreview = document.createElement('div');
    markupPreview.className = 'group-markup-preview';
    markupPreview.innerHTML = renderGroupCardMarkupHtml(group);

    const reasoning = document.createElement('div');
    reasoning.className = 'group-reasoning';
    reasoning.textContent = group.reasoning_summary || 'No edit rationale provided.';

    const pillRow = document.createElement('div');
    pillRow.className = 'group-pill-row';

    const matchPill = document.createElement('div');
    matchPill.className = 'pill neutral';
    matchPill.textContent = group.matchCount == null
      ? 'Matches: not checked'
      : `Matches: ${group.matchCount}`;

    const modePill = document.createElement('div');
    modePill.className = 'pill neutral';
    modePill.textContent = `Search: ${group.searchMode || 'exact'}`;

    pillRow.appendChild(matchPill);
    pillRow.appendChild(modePill);

    li.appendChild(topRow);
    li.appendChild(snippet);
    li.appendChild(markupPreview);
    li.appendChild(reasoning);
    li.appendChild(pillRow);
    els.groupList.appendChild(li);
  }
}

function renderSelectedGroup() {
  const group = getSelectedGroup();
  if (!group) {
    els.detailEmpty.classList.remove('hidden');
    els.detailPane.classList.add('hidden');
    return;
  }

  els.detailEmpty.classList.add('hidden');
  els.detailPane.classList.remove('hidden');

  els.detailGroup.textContent = String(group.group);
  els.detailMatchCount.textContent =
    group.matchCount == null
      ? '–'
      : `${group.matchCount}${group.matchCount > 1 ? ` (selected ${group.activeMatchIndex + 1})` : ''}`;
  els.detailStatus.textContent = group.status;
  els.detailOriginal.textContent = group.original;
  els.detailReplacement.textContent = group.replacement;
  els.detailSpanPreview.innerHTML = renderSpanPreviewHtml(group);
  els.detailReasoning.textContent = group.reasoning_summary || 'No edit rationale provided.';

  els.applyBtn.textContent = group.confirmReady
    ? 'Confirm apply replacement'
    : 'Apply replacement';

  els.applyBtn.classList.toggle('confirm-action', group.confirmReady);

  const groupIndex = state.groups.findIndex((item) => item.id === group.id);
  if (els.prevGroupBtn) {
    els.prevGroupBtn.disabled = groupIndex <= 0;
  }
  if (els.nextGroupBtn) {
    els.nextGroupBtn.disabled = groupIndex < 0 || groupIndex >= state.groups.length - 1;
  }
}

function renderGroupCardMarkupHtml(group) {
  const spanText = group.original_span || '';
  const spanIndex = spanText ? group.original.indexOf(spanText) : -1;

  const prefixSource = spanIndex >= 0 ? group.original.slice(0, spanIndex) : group.diff.prefix || '';
  const suffixSource =
    spanIndex >= 0 ? group.original.slice(spanIndex + spanText.length) : group.diff.suffix || '';

  const compactPrefix = truncateStart(prefixSource, 45);
  const compactSuffix = truncateEnd(suffixSource, 45);

  const prefix = escapeHtml(compactPrefix);
  const suffix = escapeHtml(compactSuffix);
  const originalSpan = escapeHtml(spanText || '');
  const replacementSpan = escapeHtml(group.replacement_span || '');

  return `
    <div class="group-markup-line">
      <span class="diff-context">${prefix}</span><span class="diff-old">${originalSpan || '(empty)'}</span><span class="diff-new">${replacementSpan || '(empty)'}</span><span class="diff-context">${suffix}</span>
    </div>
  `;
}

function renderSpanPreviewHtml(group) {
  const spanText = group.original_span || '';
  const spanIndex = spanText ? group.original.indexOf(spanText) : -1;

  const prefix = escapeHtml(
    spanIndex >= 0 ? group.original.slice(0, spanIndex) : group.diff.prefix || ''
  );
  const suffix = escapeHtml(
    spanIndex >= 0 ? group.original.slice(spanIndex + spanText.length) : group.diff.suffix || ''
  );
  const originalSpan = escapeHtml(spanText || '');
  const replacementSpan = escapeHtml(group.replacement_span || '');

  return `
    <span class="diff-context">${prefix}</span><span class="diff-old">${originalSpan || '(empty)'}</span><span class="diff-context">${suffix}</span>
    <br />
    <span class="diff-context">${prefix}</span><span class="diff-new">${replacementSpan || '(empty)'}</span><span class="diff-context">${suffix}</span>
  `;
}

function resetAllConfirmReady() {
  for (const group of state.groups) {
    group.confirmReady = false;
  }
}

function getSelectedGroup() {
  return state.groups.find((group) => group.id === state.selectedId) || null;
}

function scrollSelectedGroupIntoView() {
  if (!els.groupList || !state.selectedId) {
    return;
  }

  const selectedItem = els.groupList.querySelector(`[data-group-id="${CSS.escape(state.selectedId)}"]`);
  if (selectedItem) {
    selectedItem.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const helper = document.createElement('textarea');
  helper.value = text;
  helper.setAttribute('readonly', 'readonly');
  helper.style.position = 'fixed';
  helper.style.top = '-9999px';
  helper.style.left = '-9999px';
  document.body.appendChild(helper);
  helper.focus();
  helper.select();

  let succeeded = false;
  try {
    succeeded = document.execCommand('copy');
  } finally {
    document.body.removeChild(helper);
  }

  if (!succeeded) {
    throw new Error('Clipboard access was unavailable.');
  }
}

function truncate(text, max) {
  if (!text) return '';
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function truncateStart(text, max) {
  if (!text) return '';
  return text.length > max ? `…${text.slice(text.length - Math.max(0, max - 1))}` : text;
}

function truncateEnd(text, max) {
  if (!text) return '';
  return text.length > max ? `${text.slice(0, Math.max(0, max - 1))}…` : text;
}

function setGlobalStatus(message, tone) {
  setStatus(els.globalStatus, message, tone);
}

function setLlmStatus(message, tone) {
  setStatus(els.llmStatus, message, tone);
}

function updateGroupStatus(message, tone) {
  setStatus(els.groupStatus, message, tone);
}

function setStatus(el, message, tone) {
  el.textContent = message;
  el.className = `status ${tone || 'neutral'}`;
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function friendlyTrackingMode(mode) {
  switch (mode) {
    case Word.ChangeTrackingMode.trackAll:
    case 'TrackAll':
      return 'Track all changes';
    case Word.ChangeTrackingMode.trackMineOnly:
    case 'TrackMineOnly':
      return 'Track my changes only';
    case Word.ChangeTrackingMode.off:
    case 'Off':
    default:
      return 'Off';
  }
}

async function getTrackingMode() {
  return Word.run(async (context) => {
    const doc = context.document;
    doc.load('changeTrackingMode');
    await context.sync();
    return doc.changeTrackingMode;
  });
}

async function getSentenceMatchInfo(original) {
  return Word.run(async (context) => {
    let results = context.document.body.search(original, exactSearchOptions());
    results.load('items/text');
    await context.sync();

    if (results.items.length) {
      return { count: results.items.length, mode: 'exact', searchText: original };
    }

    results = context.document.body.search(original, normalizedSearchOptions());
    results.load('items/text');
    await context.sync();

    if (results.items.length) {
      return { count: results.items.length, mode: 'normalized', searchText: original };
    }

    const fallbackTargets = buildFallbackSearchTargets(original);
    for (const candidate of fallbackTargets) {
      const fallback = context.document.body.search(candidate, normalizedSearchOptions());
      fallback.load('items/text');
      await context.sync();

      if (fallback.items.length) {
        return {
          count: fallback.items.length,
          mode: 'fallback',
          searchText: candidate,
        };
      }
    }

    return { count: 0, mode: 'none', searchText: original };
  });
}

async function getSpanMatchInfo(originalSentence, spanText, sentenceIndex, mode) {
  return Word.run(async (context) => {
    const sentenceRange = await resolveSentenceRange(context, originalSentence, sentenceIndex, mode);

    const spanMatches = sentenceRange.search(spanText, exactSearchOptions());
    spanMatches.load('items/text');
    await context.sync();

    if (spanMatches.items.length) {
      return { count: spanMatches.items.length, mode: 'exact' };
    }

    const normalized = sentenceRange.search(spanText, normalizedSearchOptions());
    normalized.load('items/text');
    await context.sync();

    return { count: normalized.items.length, mode: normalized.items.length ? 'normalized' : 'exact' };
  });
}

async function selectSentenceMatch(original, sentenceIndex, mode) {
  return Word.run(async (context) => {
    const sentenceRange = await resolveSentenceRange(context, original, sentenceIndex, mode);
    sentenceRange.select();
    await context.sync();
  });
}

async function applyMinimalReplacement(group) {
  return Word.run(async (context) => {
    const doc = context.document;
    doc.load('changeTrackingMode');
    await context.sync();

    const sentenceRange = await resolveSentenceRange(
      context,
      group.searchText || group.original,
      group.activeMatchIndex,
      group.searchMode || 'exact'
    );

    const spanText = group.original_span;
    const replacementText = group.replacement_span;

    if (!spanText && group.original !== group.replacement) {
      throw new Error(
        'This suggestion does not contain a replaceable original_span. Regenerate it with the provided system instructions.'
      );
    }

    const spanRange = await resolveSpanRange(context, sentenceRange, spanText);
    spanRange.select();
    spanRange.insertText(replacementText, Word.InsertLocation.replace);
    await context.sync();

    return {
      trackingMode: doc.changeTrackingMode,
      sentenceCount: group.matchCount ?? 1,
      spanCount: 1,
    };
  });
}

async function resolveSentenceRange(context, original, sentenceIndex, mode) {
  const effectiveMode = mode === 'normalized' || mode === 'fallback' ? 'normalized' : 'exact';
  const opts = effectiveMode === 'normalized' ? normalizedSearchOptions() : exactSearchOptions();

  const matches = context.document.body.search(original, opts);
  matches.load('items/text');
  await context.sync();

  if (!matches.items.length) {
    throw new Error('Original sentence was not found.');
  }

  const index = Math.max(0, Math.min(sentenceIndex || 0, matches.items.length - 1));
  return matches.items[index];
}

async function resolveSpanRange(context, sentenceRange, spanText) {
  const exactMatches = sentenceRange.search(spanText, exactSearchOptions());
  exactMatches.load('items/text');
  await context.sync();

  if (exactMatches.items.length === 1) {
    return exactMatches.items[0];
  }
  if (exactMatches.items.length > 1) {
    throw new Error(
      'The minimal span appears more than once inside the sentence. Widen original_span slightly so it is unique.'
    );
  }

  const normalizedMatches = sentenceRange.search(spanText, normalizedSearchOptions());
  normalizedMatches.load('items/text');
  await context.sync();

  if (normalizedMatches.items.length === 1) {
    return normalizedMatches.items[0];
  }
  if (normalizedMatches.items.length > 1) {
    throw new Error(
      'The minimal span matches multiple times inside the sentence. Widen original_span slightly so it is unique.'
    );
  }

  throw new Error('original_span was not found inside the matched sentence.');
}

function exactSearchOptions() {
  return {
    matchCase: true,
    matchWholeWord: false,
  };
}

function normalizedSearchOptions() {
  return {
    matchCase: false,
    matchWholeWord: false,
    ignorePunct: true,
    ignoreSpace: true,
  };
}