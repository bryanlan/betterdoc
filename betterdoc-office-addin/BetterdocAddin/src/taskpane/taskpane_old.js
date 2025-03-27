/* global Word, Office */

// Helper function to log messages to both the console and the debug window.
function logMessage(message) {
  const debugLog = document.getElementById("debugLog");
  if (debugLog) {
    const p = document.createElement("p");
    p.textContent = message;
    debugLog.appendChild(p);
  }
  console.log(message);
}

let paragraphs = [];            // Array of paragraph text from the document
let currentParagraphIndex = 0;  // Track which paragraph is currently selected
let reimaginedAllText = null;   // In "Reimagine All" mode, store combined text here
let isPromptEditorVisible = false;

// Fired when the Office.js library is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Hook up button events.
    document.getElementById("btnUp").onclick = onUp;
    document.getElementById("btnDown").onclick = onDown;
    document.getElementById("btnReimagineAll").onclick = onReimagineAll;
    document.getElementById("btnReimagine").onclick = onReimagine;
    document.getElementById("btnApply").onclick = onApply;
    document.getElementById("togglePromptBtn").onclick = togglePromptEditor;

    // Load paragraphs from the document.
    loadParagraphsFromDocument().catch((err) => {
      logMessage("Error loading paragraphs: " + err);
    });
  }
});

/**
 * Toggle the prompt editor window in the task pane UI.
 */
function togglePromptEditor() {
  isPromptEditorVisible = !isPromptEditorVisible;
  const promptSection = document.getElementById("prompt-section");
  promptSection.style.display = isPromptEditorVisible ? "block" : "none";
  logMessage("Prompt editor toggled: " + (isPromptEditorVisible ? "shown" : "hidden"));
}

/**
 * Loads all paragraphs in the Word document into our 'paragraphs' array.
 * Then highlights the first paragraph.
 */
async function loadParagraphsFromDocument() {
  await Word.run(async (context) => {
    const body = context.document.body;
    // Get all paragraphs.
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();

    paragraphs = [];
    for (let i = 0; i < paras.items.length; i++) {
      const p = paras.items[i];
      // Store each paragraph's text.
      paragraphs.push(p.text);
    }
    logMessage("Paragraphs loaded: " + JSON.stringify(paragraphs));

    // If we have at least one paragraph, highlight the first.
    if (paragraphs.length > 0) {
      currentParagraphIndex = 0;
      await highlightCurrentParagraph(context);
      updateUIWithParagraph();
    } else {
      document.getElementById("originalText").value = "";
      document.getElementById("modifiedText").value = "";
    }
  });
}

/**
 * Highlights the currently selected paragraph in the document (by selecting it).
 */
async function highlightCurrentParagraph(context) {
  if (!context) {
    return Word.run(async (ctx) => highlightCurrentParagraph(ctx));
  }

  if (currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    return;
  }

  // Get the paragraph text and trim whitespace
  let paragraphText = paragraphs[currentParagraphIndex].trim();
  
  // Define a maximum search length (e.g., 250 characters)
  const maxSearchLength = 250;
  if (paragraphText.length > maxSearchLength) {
    // Use only the first maxSearchLength characters for searching
    paragraphText = paragraphText.substring(0, maxSearchLength);
    logMessage("Truncated search text to: " + paragraphText);
  }

  logMessage("Searching for paragraph text: " + paragraphText);

  // Perform the search in the document body using the (possibly truncated) text.
  const results = context.document.body.search(paragraphText, { matchCase: false, matchWholeWord: false });
  results.load("items");
  await context.sync();

  if (results.items.length > 0) {
    results.items[0].select("Select");
    await context.sync();
    logMessage("Paragraph highlighted successfully.");
  } else {
    logMessage("No match found for the paragraph search.");
  }
}

/**
 * Updates the task pane text areas to show the selected paragraph text.
 */
function updateUIWithParagraph() {
  const originalBox = document.getElementById("originalText");
  const modifiedBox = document.getElementById("modifiedText");
  if (reimaginedAllText !== null) {
    // Reimagine All mode.
    originalBox.value = paragraphs.join("\n\n");
    modifiedBox.value = reimaginedAllText;
    logMessage("UI updated in Reimagine All mode.");
  } else {
    // Single paragraph mode.
    if (currentParagraphIndex >= 0 && currentParagraphIndex < paragraphs.length) {
      const text = paragraphs[currentParagraphIndex];
      originalBox.value = text;
      modifiedBox.value = text;
      logMessage("UI updated with paragraph (" + currentParagraphIndex + "): " + text);
    } else {
      originalBox.value = "";
      modifiedBox.value = "";
      logMessage("No valid paragraph index for UI update.");
    }
  }
}

/**
 * Move to the previous paragraph (Up).
 */
function onUp() {
  if (currentParagraphIndex > 0) {
    let newIndex = currentParagraphIndex - 1;
    
    // Skip empty or single-sentence paragraphs
    while (newIndex > 0 && isEmptyOrSingleSentence(paragraphs[newIndex])) {
      newIndex--;
    }
    
    currentParagraphIndex = newIndex;
    reimaginedAllText = null; // Exiting "Reimagine All" mode
    Word.run(async (context) => {
      await highlightCurrentParagraph(context);
      updateUIWithParagraph();
    }).catch((error) => console.error(error));
  }
}

/**
 * Move to the next paragraph (Down).
 */
function onDown() {
  if (currentParagraphIndex < paragraphs.length - 1) {
    let newIndex = currentParagraphIndex + 1;
    
    // Skip empty or single-sentence paragraphs
    while (newIndex < paragraphs.length - 1 && isEmptyOrSingleSentence(paragraphs[newIndex])) {
      newIndex++;
    }
    
    currentParagraphIndex = newIndex;
    reimaginedAllText = null; // Exiting "Reimagine All" mode
    Word.run(async (context) => {
      await highlightCurrentParagraph(context);
      updateUIWithParagraph();
    }).catch((error) => console.error(error));
  }
}

/**
 * Helper function to check if a paragraph is empty or contains only one sentence
 */
function isEmptyOrSingleSentence(text) {
  if (!text || text.trim().length === 0) {
    return true;
  }
  
  // Count sentences by looking for common sentence endings
  const sentences = text.split(/[.!?]+/).filter(sentence => sentence.trim().length > 0);
  return sentences.length <= 1;
}

/**
 * Call our local LLM-like service to reimagine a single paragraph with the given prompt.
 */
async function onReimagine() {
  // If Reimagine All is active, only act on highlighted text.
  if (reimaginedAllText !== null) {
    const selectedText = getSelectedTextInTextarea("modifiedText");
    if (selectedText && selectedText.trim().length > 0) {
      logMessage("Reimagining selected text in Reimagine All mode...");
      const newText = await callOllama(selectedText, getPrompt());
      replaceSelectedTextInTextarea("modifiedText", newText);
      logMessage("Reimagined text (selected): " + newText);
    }
    return;
  }
  // Otherwise, single paragraph mode.
  if (currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    logMessage("Invalid paragraph index: " + currentParagraphIndex);
    return;
  }
  const originalBox = document.getElementById("originalText");
  const prompt = getPrompt();
  logMessage("Calling LLM for paragraph: " + originalBox.value + " with prompt: " + prompt);
  const newText = await callOllama(originalBox.value, prompt);
  document.getElementById("modifiedText").value = newText;
  logMessage("Received LLM output: " + newText);
}

/**
 * Reimagine All: For each paragraph, call the LLM with the same prompt, then combine the results.
 */
async function onReimagineAll() {
  if (paragraphs.length === 0) {
    logMessage("No paragraphs to reimagine.");
    return;
  }
  const prompt = getPrompt();
  let allResults = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const newText = await callOllama(p, prompt);
    allResults.push(newText);
    logMessage("Paragraph " + i + " reimagined: " + newText);
  }
  reimaginedAllText = allResults.join("\n\n");
  updateUIWithParagraph();
}

/**
 * Apply:
 * - In Reimagine All mode, split the combined text and update each paragraph.
 * - Otherwise, update only the selected paragraph.
 */
async function onApply() {
  if (paragraphs.length === 0) {
    logMessage("No paragraphs available to apply changes.");
    return;
  }
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();
    if (reimaginedAllText !== null) {
      const bigText = document.getElementById("modifiedText").value;
      const updatedTexts = bigText.split("\n\n");
      for (let i = 0; i < paras.items.length && i < updatedTexts.length; i++) {
        paras.items[i].insertText(updatedTexts[i], "Replace");
        logMessage("Applied text to paragraph " + i + ": " + updatedTexts[i]);
      }
    } else {
      if (currentParagraphIndex >= 0 && currentParagraphIndex < paras.items.length) {
        const newText = document.getElementById("modifiedText").value;
        paras.items[currentParagraphIndex].insertText(newText, "Replace");
        logMessage("Applied text to paragraph " + currentParagraphIndex + ": " + newText);
      }
    }
    await context.sync();
    logMessage("Changes applied and document synced.");
    await loadParagraphsFromDocument();
  }).catch((error) => logMessage("Error in onApply: " + error));
}

/**
 * Helper to get the current prompt from the prompt textarea.
 */
function getPrompt() {
  // Add the hardcoded instruction to the end of whatever prompt the user has entered
  const userPrompt = document.getElementById("promptInput").value.trim();
  return `${userPrompt}\n\nONLY RESPOND WITH THE RESTRUCTURED PARAGRAPH`;
}

/**
 * Call a local server or remote service that wraps the Ollama functionality.
 */
async function callOllama(paragraphText, userPrompt) {
  try {
    const requestData = {
      paragraphText: paragraphText,
      userPrompt: userPrompt
    };
    logMessage("Sending request to LLM server: " + JSON.stringify(requestData));
    const response = await fetch("https://localhost:8000/ollama", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(requestData)
    });
    if (!response.ok) {
      logMessage("Error from LLM server: " + response.statusText);
      return paragraphText; // fallback - if error, return original text
    }
    const data = await response.json();
    logMessage("LLM server response: " + JSON.stringify(data));
    return data.result || paragraphText;
  } catch (err) {
    logMessage("Failed to call LLM service: " + err);
    return paragraphText;
  }
}

/**
 * Helper: get the currently selected text inside a <textarea> by ID.
 */
function getSelectedTextInTextarea(textareaId) {
  const textarea = document.getElementById(textareaId);
  return textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
}

/**
 * Helper: replace the selected text inside a <textarea> with a new string.
 */
function replaceSelectedTextInTextarea(textareaId, newText) {
  const textarea = document.getElementById(textareaId);
  const start = textarea.selectionStart;
  const end = textarea.selectionEnd;
  const oldValue = textarea.value;
  textarea.value = oldValue.substring(0, start) + newText + oldValue.substring(end);
  const newCursorPos = start + newText.length;
  textarea.selectionStart = newCursorPos;
  textarea.selectionEnd = newCursorPos;
}
