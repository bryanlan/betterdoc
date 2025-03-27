/* global Word, Office */

let paragraphs = [];            // Array of paragraph text from the doc
let currentParagraphIndex = 0;  // Track which paragraph is currently selected
let reimaginedAllText = null;   // If we do "Reimagine All," store combined text here
let isPromptEditorVisible = false;

// Fired when the Office.js library is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Once the add-in is loaded in Word, we can hook up our button events
    document.getElementById("btnUp").onclick = onUp;
    document.getElementById("btnDown").onclick = onDown;
    document.getElementById("btnReimagineAll").onclick = onReimagineAll;
    document.getElementById("btnReimagine").onclick = onReimagine;
    document.getElementById("btnApply").onclick = onApply;
    document.getElementById("togglePromptBtn").onclick = togglePromptEditor;

    // Load paragraphs from the doc
    loadParagraphsFromDocument().catch((err) => {
      console.log("Error loading paragraphs: ", err);
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
}

/**
 * Loads all paragraphs in the Word document into our 'paragraphs' array.
 * Then highlights the first paragraph.
 */
async function loadParagraphsFromDocument() {
  await Word.run(async (context) => {
    const body = context.document.body;
    // Get all paragraphs
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();

    paragraphs = [];
    for (let i = 0; i < paras.items.length; i++) {
      const p = paras.items[i];
      // Store each paragraph's text
      paragraphs.push(p.text);
    }

    // If we have at least one paragraph, highlight the first
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
 * Highlights the currently selected paragraph in the doc (by selecting it).
 */
async function highlightCurrentParagraph(context) {
  if (!context) {
    return Word.run(async (ctx) => highlightCurrentParagraph(ctx));
  }

  if (currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    return;
  }

  const paragraphText = paragraphs[currentParagraphIndex];

  const results = context.document.body.search(paragraphText, { matchCase: false, matchWholeWord: false });
  results.load("items");
  await context.sync();

  if (results.items.length > 0) {
    results.items[0].select("Select");
    await context.sync();
  }
}

/**
 * Updates the task pane text areas to show the selected paragraph text.
 */
function updateUIWithParagraph() {
  const originalBox = document.getElementById("originalText");
  const modifiedBox = document.getElementById("modifiedText");

  if (reimaginedAllText !== null) {
    // We are in "Reimagine All" mode
    originalBox.value = paragraphs.join("\n\n");
    modifiedBox.value = reimaginedAllText;
  } else {
    // Single paragraph mode
    if (currentParagraphIndex >= 0 && currentParagraphIndex < paragraphs.length) {
      const text = paragraphs[currentParagraphIndex];
      originalBox.value = text;
      modifiedBox.value = text;
    } else {
      originalBox.value = "";
      modifiedBox.value = "";
    }
  }
}

/**
 * Move to the previous paragraph (Up).
 */
function onUp() {
  if (currentParagraphIndex > 0) {
    currentParagraphIndex--;
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
    currentParagraphIndex++;
    reimaginedAllText = null; // Exiting "Reimagine All" mode
    Word.run(async (context) => {
      await highlightCurrentParagraph(context);
      updateUIWithParagraph();
    }).catch((error) => console.error(error));
  }
}

/**
 * Call our local LLM-like service to reimagine a single paragraph with the given prompt.
 */
async function onReimagine() {
  // If "Reimagine All" is in effect, do nothing unless user highlights text in the "modifiedText"
  if (reimaginedAllText !== null) {
    const selectedText = getSelectedTextInTextarea("modifiedText");
    if (selectedText && selectedText.trim().length > 0) {
      console.log("Reimagining selected text in reimaginedAll mode...");
      const newText = await callOllama(selectedText, getPrompt());
      replaceSelectedTextInTextarea("modifiedText", newText);
    }
    return;
  }

  // Otherwise, single paragraph mode:
  if (currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    return;
  }

  const originalBox = document.getElementById("originalText");
  const modifiedBox = document.getElementById("modifiedText");

  const paragraphText = originalBox.value;
  const prompt = getPrompt();

  // Call the LLM
  const newText = await callOllama(paragraphText, prompt);
  modifiedBox.value = newText;
}

/**
 * Reimagine All: For each paragraph, call the LLM with the same prompt, then combine the results.
 */
async function onReimagineAll() {
  if (paragraphs.length === 0) {
    return;
  }

  const prompt = getPrompt();
  let allResults = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const newText = await callOllama(p, prompt);
    allResults.push(newText);
  }

  reimaginedAllText = allResults.join("\n\n");
  updateUIWithParagraph();
}

/**
 * Apply: 
 * - If Reimagine All mode is active, split the big text by double newlines and place them 
 *   back into the doc paragraphs. 
 * - Otherwise, just apply the single paragraph's text.
 */
async function onApply() {
  if (paragraphs.length === 0) {
    return;
  }

  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();

    if (reimaginedAllText !== null) {
      // "Reimagine All" scenario
      const bigText = document.getElementById("modifiedText").value;
      const updatedTexts = bigText.split("\n\n");

      // We'll just apply them in order to paragraphs
      for (let i = 0; i < paras.items.length && i < updatedTexts.length; i++) {
        paras.items[i].insertText(updatedTexts[i], "Replace");
      }
    } else {
      // Single paragraph scenario
      if (currentParagraphIndex >= 0 && currentParagraphIndex < paras.items.length) {
        const newText = document.getElementById("modifiedText").value;
        paras.items[currentParagraphIndex].insertText(newText, "Replace");
      }
    }

    await context.sync();

    // Re-load paragraphs to reflect changes
    await loadParagraphsFromDocument();
  });
}

/**
 * Helper to get the current prompt from the prompt textarea.
 */
function getPrompt() {
  return document.getElementById("promptInput").value.trim();
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

    const response = await fetch("http://localhost:8000/ollama", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(requestData)
    });

    if (!response.ok) {
      console.error("Error from LLM server:", response.statusText);
      return paragraphText; // fallback - if error, just return original
    }

    const data = await response.json();
    return data.result || paragraphText; 
  } catch (err) {
    console.error("Failed to call LLM service:", err);
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

  // Replace
  textarea.value = oldValue.substring(0, start) + newText + oldValue.substring(end);

  // Move the cursor to just after the inserted text
  const newCursorPos = start + newText.length;
  textarea.selectionStart = newCursorPos;
  textarea.selectionEnd = newCursorPos;
} 