/* global Word, Office */

// Global variables.
let paragraphs = [];            // Array of paragraph text from the document.
let currentParagraphIndex = 0;  // Track which paragraph is currently selected.
let reimaginedAllText = null;   // In "Reimagine All" mode, store combined text here.
let isPromptEditorVisible = false;

// New globals for multi‐paragraph selection and LLM outputs.
let selectedParagraphs = new Set();  // Set of paragraph indexes selected for reimagination.
let paragraphLLMOutputs = {};          // Mapping: paragraph index -> modified text.

// Add new global to store formatting information
let paragraphFormatting = {};  // Map: paragraph index -> formatting info

// Add to the global variables section
let availableModels = [];
let selectedModel = ""; // Will store the currently selected model

// Fired when the Office.js library is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    logMessage("Office Add-in initialized successfully");
    // Hook up existing button events.
    document.getElementById("btnUp").onclick = onUp;
    document.getElementById("btnDown").onclick = onDown;
    document.getElementById("btnReimagine").onclick = onReimagine;
    document.getElementById("btnApply").onclick = onApply;
    // New buttons.
    document.getElementById("btnSelectAll").onclick = onSelectAll;
    document.getElementById("btnUnselectAll").onclick = onUnselectAll;
    document.getElementById("btnApplyAll").onclick = onApplyAll;
    // Checkbox for current paragraph.
    document.getElementById("checkboxReimagine").onchange = onCheckboxChange;
    // Save any user edits when they leave the modified text box.
    document.getElementById("modifiedText").onblur = onModifiedTextBlur;
    // Prompt selection dropdown
    document.getElementById("promptSelect").onchange = onPromptSelectChange;
    // Model selection dropdown
    document.getElementById("modelSelect").onchange = onModelSelectChange;

    // Load paragraphs from the document.
    loadParagraphsFromDocument().then(() => {
      if (paragraphs.length > 0) {
        currentParagraphIndex = 0;
        updateUIForSelectedParagraph();
      }
    }).catch(error => {
      logMessage("Error loading initial document: " + error);
    });
    
    // Load available models
    loadAvailableModels().catch((err) => {
      logMessage("Error loading models: " + err);
    });

    // Add refresh button handler
    document.getElementById("refreshButton").onclick = onRefresh;
  }
});

const PROMPTS = {
  professional: "Rewrite this and make it sound more professional while not leaving out any info. Do NOT increase the word count, reduce if possible.",
  concise: "Reduce the word count as much as possible while retaining the key points",
  custom: ""
};

/**
 * Handles changes to the prompt selection dropdown
 */
function onPromptSelectChange() {
  const select = document.getElementById("promptSelect");
  const customSection = document.getElementById("custom-prompt-section");
  const customInput = document.getElementById("customPromptInput");
  
  if (select.value === "custom") {
    customSection.style.display = "block";
    if (customInput.value === "") {
      customInput.value = "Replace this with your own prompt";
    }
  } else {
    customSection.style.display = "none";
  }
  logMessage("Prompt selection changed to: " + select.value);
}

/**
 * Gets the current prompt based on the dropdown selection
 */
function getPrompt() {
  const select = document.getElementById("promptSelect");
  const customInput = document.getElementById("customPromptInput");
  
  let prompt;
  if (select.value === "custom") {
    prompt = customInput.value.trim();
    // If custom prompt is empty or default, fall back to professional prompt
    if (!prompt || prompt === "Replace this with your own prompt") {
      prompt = PROMPTS.professional;
    }
  } else {
    prompt = PROMPTS[select.value];
  }
  
  logMessage("Using prompt: " + prompt);
  return `${prompt}\n\nONLY RESPOND WITH THE RESTRUCTURED PARAGRAPH. DO NOT INCREASE THE WORD COUNT. DO NOT ANSWER QUESTIONS IN THE TEXT!!!.`;
}

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

/**
 * Helper function to check if a paragraph should be skipped
 * Returns true if the paragraph:
 * - is empty, or
 * - has fewer than 7 words, or
 * - ends with a question mark
 */
function shouldSkipParagraph(text) {
  if (!text || text.trim().length === 0) {
    return true;
  }
  
  const trimmedText = text.trim();
  
  // Skip if ends with question mark
  if (trimmedText.endsWith('?')) {
    return true;
  }
  
  // Count words (split by whitespace and filter out empty strings)
  const words = trimmedText.split(/\s+/).filter(word => word.length > 0);
  const MIN_WORDS = 7;
  
  return words.length < MIN_WORDS;
}

/**
 * Detects and stores formatting information for a paragraph
 */
async function detectFormatting(context, paragraph, index) {
  try {
    // First, get character-by-character formatting
    const ranges = paragraph.getTextRanges([" "]);  // Split by spaces instead of characters
    ranges.load("items");
    await context.sync();
    
    // First pass: collect all formatting to find mode
    const formatDistribution = new Map();
    const formattingByRange = [];
    
    for (let i = 0; i < ranges.items.length; i++) {
      const range = ranges.items[i];
      range.load(["text", "font/bold", "font/italic", "font/size", "font/name"]);
      await context.sync();
      
      const formatting = {
        text: range.text,
        bold: range.font.bold,
        italic: range.font.italic,
        size: range.font.size,
        name: range.font.name
      };
      formattingByRange.push(formatting);
      
      const formatKey = JSON.stringify({
        bold: formatting.bold,
        italic: formatting.italic,
        size: formatting.size
      });
      
      formatDistribution.set(formatKey, (formatDistribution.get(formatKey) || 0) + 1);
    }
    
    // Find modal formatting
    let modalFormat = null;
    let maxCount = 0;
    for (const [key, count] of formatDistribution) {
      if (count > maxCount) {
        maxCount = count;
        modalFormat = JSON.parse(key);
      }
    }
    
    // Initialize formatting info for this paragraph
    paragraphFormatting[index] = {
      defaultFont: modalFormat,
      formattedClauses: []
    };
    
    logMessage(`\nFormatting analysis for paragraph ${index + 1}:`);
    logMessage(`Modal formatting: ${JSON.stringify(modalFormat)}`);
    
    // Second pass: detect non-modal phrases
    let currentPhrase = null;
    
    for (let i = 0; i < formattingByRange.length; i++) {
      const current = formattingByRange[i];
      
      // Check if current range deviates from modal
      const isNonModal = 
        current.bold !== modalFormat.bold ||
        current.italic !== modalFormat.italic ||
        current.size !== modalFormat.size ||
        // If previous word was non-modal and this has null where previous had non-modal value
        (currentPhrase && 
         ((current.bold === null && currentPhrase.formatting.bold !== modalFormat.bold) ||
          (current.italic === null && currentPhrase.formatting.italic !== modalFormat.italic) ||
          (current.size === null && currentPhrase.formatting.size !== modalFormat.size)));
      
      if (isNonModal) {
        if (!currentPhrase) {
          // Start new phrase
          currentPhrase = {
            text: current.text,
            formatting: {
              bold: current.bold,
              italic: current.italic,
              size: current.size
            }
          };
        } else {
          // Continue current phrase
          currentPhrase.text += current.text;
        }
      } else if (currentPhrase) {
        // End current phrase
        paragraphFormatting[index].formattedClauses.push({
          text: currentPhrase.text.trim(),
          formatting: currentPhrase.formatting
        });
        currentPhrase = null;
      }
    }
    
    // Don't forget last phrase
    if (currentPhrase) {
      paragraphFormatting[index].formattedClauses.push({
        text: currentPhrase.text.trim(),
        formatting: currentPhrase.formatting
      });
    }
    
    // Log non-modal phrases
    if (paragraphFormatting[index].formattedClauses.length > 0) {
      logMessage("\nNon-modal formatted phrases:");
      paragraphFormatting[index].formattedClauses.forEach(phrase => {
        logMessage(`"${phrase.text}": ${JSON.stringify(phrase.formatting)}`);
      });
    } else {
      logMessage("\nNo non-modal formatting detected");
    }
    
    return paragraphFormatting[index].formattedClauses;
    
  } catch (error) {
    logMessage(`Error in detectFormatting: ${error.toString()}`);
    throw error;
  }
}

/**
 * Loads all paragraphs in the Word document into our 'paragraphs' array
 */
async function loadParagraphsFromDocument() {
  logMessage("Loading paragraphs from document...");
  paragraphs = [];
  
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("text");
    await context.sync();

    for (let i = 0; i < paras.items.length; i++) {
      paragraphs.push(paras.items[i].text);
    }
    
    logMessage(`Loaded ${paragraphs.length} paragraphs from document`);
  }).catch((error) => {
    logMessage("Error loading paragraphs: " + error);
  });
  
  return paragraphs;
}

/**
 * Highlights the current paragraph in the document.
 */
async function highlightCurrentParagraph() {
  if (paragraphs.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    return;
  }
  
  await Word.run(async (context) => {
    try {
      // Clear any existing highlights first
      const body = context.document.body;
      body.style.backgroundColor = "white";
      await context.sync();
      
      // Get the text to search for - limit to first 255 characters to avoid "too long" error
      const paragraphText = paragraphs[currentParagraphIndex];
      const searchText = paragraphText.substring(0, Math.min(255, paragraphText.length));
      
      if (searchText.trim().length < 3) {
        logMessage("Paragraph text too short to highlight");
        return;
      }
      
      logMessage(`Searching for paragraph text: ${searchText.length} chars`);
      
      // Search for the paragraph text
      const searchResults = body.search(searchText);
      searchResults.load("items");
  await context.sync();
      
      if (searchResults.items.length > 0) {
        // Highlight the first match
        searchResults.items[0].style.backgroundColor = "#FFFF00";
    await context.sync();
        logMessage("Paragraph highlighted successfully");
  } else {
        logMessage("Could not find paragraph to highlight");
      }
    } catch (error) {
      logMessage(`Error highlighting paragraph: ${error}`);
  }
  }).catch((error) => logMessage(`Error in highlightCurrentParagraph: ${error}`));
}

/**
 * Updates the UI with the current paragraph's text.
 */
function updateUIWithParagraph() {
  document.getElementById("paragraphHeader").textContent = `Paragraph ${currentParagraphIndex + 1}`;
  document.getElementById("originalText").value = paragraphs[currentParagraphIndex];
  document.getElementById("modifiedText").value = paragraphLLMOutputs[currentParagraphIndex]?.text || "No changes from original text";
  document.getElementById("checkboxReimagine").checked = selectedParagraphs.has(currentParagraphIndex);
  logMessage(`UI updated with paragraph ${currentParagraphIndex + 1}`);
}

/**
 * Updates the list (displayed below the "Paragraphs to be reimagined" label)
 * of selected paragraphs. Each list item is clickable and jumps to that paragraph.
 */
function updateSelectedParagraphList() {
  const listDiv = document.getElementById("selectedParagraphsList");
  listDiv.innerHTML = "";
  if (selectedParagraphs.size === 0) {
    listDiv.textContent = "No paragraphs selected for reimagination.";
    return;
  }
  const ul = document.createElement("ul");
  selectedParagraphs.forEach(idx => {
    const li = document.createElement("li");
    li.textContent = "Paragraph " + (idx + 1);
    li.style.cursor = "pointer";
    // When clicked, jump to the corresponding paragraph.
    li.onclick = () => goToParagraph(idx);
    ul.appendChild(li);
  });
  listDiv.appendChild(ul);
}

/**
 * Jumps to the specified paragraph index.
 */
function goToParagraph(index) {
  currentParagraphIndex = index;
  Word.run(async (context) => {
    await highlightCurrentParagraph(context);
    updateUIForSelectedParagraph();
  }).catch((error) => console.error(error));
  logMessage("Jumped to paragraph " + (index + 1));
}

/**
 * Handler for changes to the "Reimagine" checkbox.
 */
function onCheckboxChange() {
  const checkbox = document.getElementById("checkboxReimagine");
  if (checkbox.checked) {
    selectedParagraphs.add(currentParagraphIndex);
  } else {
    selectedParagraphs.delete(currentParagraphIndex);
  }
  updateSelectedParagraphList();
}

/**
 * Saves any user edits to the modified text box for the current paragraph.
 */
function onModifiedTextBlur() {
  const modifiedBox = document.getElementById("modifiedText");
  const originalBox = document.getElementById("originalText");
  
  // Don't save if it's our placeholder text
  if (modifiedBox.value === "No changes from original text") {
    return;
  }
  
  // If the modified text is the same as original, show placeholder
  if (modifiedBox.value.trim() === originalBox.value.trim()) {
    modifiedBox.value = "No changes from original text";
    delete paragraphLLMOutputs[currentParagraphIndex];
  } else {
  paragraphLLMOutputs[currentParagraphIndex] = modifiedBox.value;
  }
  
  logMessage(`Saved modified text for paragraph ${currentParagraphIndex + 1}`);
}

/**
 * Moves to the previous paragraph (Up).
 */
function onUp() {
  if (currentParagraphIndex > 0) {
    let newIndex = currentParagraphIndex - 1;
    while (newIndex > 0 && shouldSkipParagraph(paragraphs[newIndex])) {
      newIndex--;
    }
    currentParagraphIndex = newIndex;
    reimaginedAllText = null;
    
    Word.run(async (context) => {
      await highlightCurrentParagraph(context);
      updateUIForSelectedParagraph();
    }).catch((error) => logMessage("Error in onUp: " + error));
  }
}

/**
 * Moves to the next paragraph (Down).
 */
function onDown() {
  if (currentParagraphIndex < paragraphs.length - 1) {
    let newIndex = currentParagraphIndex + 1;
    while (newIndex < paragraphs.length - 1 && shouldSkipParagraph(paragraphs[newIndex])) {
      newIndex++;
    }
    currentParagraphIndex = newIndex;
    reimaginedAllText = null;
    
    Word.run(async (context) => {
      await highlightCurrentParagraph(context);
      updateUIForSelectedParagraph();
    }).catch((error) => logMessage("Error in onDown: " + error));
  }
}

/**
 * Processes all selected paragraphs by sending them sequentially to the LLM.
 */
async function onReimagine() {
  if (selectedParagraphs.size === 0) {
    logMessage("No paragraphs selected for reimagination.");
    return;
  }
  
  const prompt = getPrompt();
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paras = body.paragraphs;
      paras.load("items");
      await context.sync();
      
  for (let idx of selectedParagraphs) {
        try {
    const originalText = paragraphs[idx];
          logMessage(`Processing paragraph ${idx + 1}`);
          
          // First detect formatting
          await detectFormatting(context, paras.items[idx], idx);
          
          // Get rewritten text from LLM
          let newText = await callOllama(originalText, prompt);
          
          // Check word count
          const originalWordCount = countWords(originalText);
          const newWordCount = countWords(newText);
          
          // If new text is more than 10% longer, ask for a shorter version
          if (newWordCount > originalWordCount * 1.1) {
            logMessage(`Word count exceeded: Original=${originalWordCount}, New=${newWordCount}`);
            
            // Create a follow-up prompt asking for shorter text
            const followUpPrompt = `${prompt}\n\nYou failed to do as instructed and exceeded the word count. Try again and reduce your word count. The original had ${originalWordCount} words, yours had ${newWordCount}.`;
            
            // Call LLM again with the follow-up prompt
            newText = await callOllama(originalText, followUpPrompt);
            
            // Log the new word count
            const revisedWordCount = countWords(newText);
            logMessage(`Revised word count: ${revisedWordCount}`);
          }
          
          // Process each non-modal phrase to find its match in new text
          const formattingMatches = [];
          for (const phrase of paragraphFormatting[idx].formattedClauses) {
            // Prepare clean versions for comparison (no punctuation, lowercase)
            const cleanPhrase = phrase.text.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
            const cleanNewText = newText.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
            
            // Check if the exact phrase exists as a whole word in the new text
            const phraseRegex = new RegExp(`\\b${escapeRegExp(cleanPhrase)}\\b`, 'i');
            
            if (phraseRegex.test(cleanNewText)) {
              // Direct match found - no need to call LLM
              logMessage(`Direct match found for phrase: "${phrase.text}"`);
              
              // Find the actual case in the new text
              const match = newText.match(new RegExp(`\\b[^.,;:!?()]*${escapeRegExp(cleanPhrase)}[^.,;:!?()]*\\b`, 'i'));
              const exactMatch = match ? match[0].trim() : phrase.text;
              
              formattingMatches.push({
                newText: exactMatch,
                formatting: phrase.formatting
              });
            } else {
              // No direct match, use LLM to find equivalent
              const matchPrompt = createClauseMatchPrompt(originalText, newText, phrase.text);
              logMessage("Finding match for formatted phrase via LLM");
              
              let matchingClause = await callOllama(phrase.text, matchPrompt);
              
              // Check if the matching clause is too long compared to original
              const originalPhraseWordCount = countWords(phrase.text);
              const matchingClauseWordCount = countWords(matchingClause);
              
              // If the matching clause is more than 30% longer than the original phrase
              if (matchingClauseWordCount > originalPhraseWordCount * 1.3 && matchingClauseWordCount > originalPhraseWordCount + 1) {
                logMessage(`Matching clause too verbose: Original=${originalPhraseWordCount} words, Match=${matchingClauseWordCount} words`);
                
                // Create a follow-up prompt asking for a more concise match
                const followUpPrompt = `${matchPrompt}\n\nYour matching text was too long. The original formatted text had ${originalPhraseWordCount} words, but your match had ${matchingClauseWordCount} words.\n\nPlease find a more concise match with approximately ${originalPhraseWordCount} words.`;
                
                // Call LLM again with the follow-up prompt
                const shorterMatchingClause = await callOllama(phrase.text, followUpPrompt);
                
                // Log the revised word count
                const revisedMatchWordCount = countWords(shorterMatchingClause);
                logMessage(`Revised match word count: ${revisedMatchWordCount} (original: ${originalPhraseWordCount})`);
                
                // Verify the shortened response actually exists in the new text
                // Clean both texts for comparison (remove punctuation, lowercase)
                const cleanShorterMatch = shorterMatchingClause.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
                const cleanNewText = newText.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
                
                // Only use the shorter match if it exists in the new text
                if (cleanNewText.includes(cleanShorterMatch)) {
                  matchingClause = shorterMatchingClause;
                  logMessage(`Using shorter match: "${shorterMatchingClause}"`);
                } else {
                  logMessage(`Shorter match not found in new text, using original match: "${matchingClause}"`);
                }
              }
              
              formattingMatches.push({
                newText: matchingClause.trim(),
                formatting: phrase.formatting
              });
            }
          }
          
          // Store both the new text and formatting matches
          paragraphLLMOutputs[idx] = {
            text: newText,
            formattingMatches: formattingMatches
          };
          
          // Update UI if we're on this paragraph
    if (idx === currentParagraphIndex) {
      document.getElementById("modifiedText").value = newText;
    }
          
        } catch (error) {
          logMessage(`Error processing paragraph ${idx + 1}: ${error.toString()}`);
        }
      }
    });
  } catch (error) {
    logMessage(`Error in onReimagine: ${error.toString()}`);
  }
}

/**
 * Helper function to count words in a string
 */
function countWords(text) {
  return text.trim().split(/\s+/).filter(word => word.length > 0).length;
}

/**
 * Creates a prompt to find matching clause in new text with focused context
 */
function createClauseMatchPrompt(originalText, newText, originalClause) {
  // Split texts into sentences
  const originalSentences = splitIntoSentences(originalText);
  const newSentences = splitIntoSentences(newText);
  
  // Find which sentence contains the formatted phrase
  let containingSentence = "";
  let sentenceIndex = -1;
  for (let i = 0; i < originalSentences.length; i++) {
    if (originalSentences[i].includes(originalClause)) {
      containingSentence = originalSentences[i];
      sentenceIndex = i;
      break;
    }
  }
  
  // If not found, use the whole text
  if (sentenceIndex === -1) {
    containingSentence = originalText;
  }
  
  // Calculate relative position in original text
  const relativePosition = sentenceIndex / originalSentences.length;
  
  // Select relevant sentences from new text
  let relevantNewSentences;
  if (newSentences.length <= 4) {
    // If 4 or fewer sentences, use all
    relevantNewSentences = newSentences;
  } else {
    // Calculate window size and position
    const windowSize = Math.ceil(newSentences.length / 2);
    const estimatedPosition = Math.floor(relativePosition * newSentences.length);
    
    // Calculate window boundaries
    const halfWindow = Math.floor(windowSize / 2);
    let windowStart = Math.max(0, estimatedPosition - halfWindow);
    let windowEnd = Math.min(newSentences.length, windowStart + windowSize);
    
    // Adjust if window is at the end
    if (windowEnd === newSentences.length) {
      windowStart = Math.max(0, newSentences.length - windowSize);
    }
    
    // Extract the window of sentences
    relevantNewSentences = newSentences.slice(windowStart, windowEnd);
  }
  
  // Join the relevant sentences
  const relevantNewText = relevantNewSentences.join(' ');
  
  // Create the focused prompt
  return `FIND THE EXACT MATCHING PHRASE ONLY.

The original text contained this formatted phrase: "${originalClause}"
It appeared in this sentence: "${containingSentence}"

In the new text, find the equivalent phrase that matches in meaning and importance.
Relevant portion of new text: "${relevantNewText}"

RESPOND ONLY WITH THE MATCHING PHRASE FROM THE NEW TEXT.
DO NOT include any explanation or additional text.
RETURN ONLY THE SHORTEST MATCHING PHRASE.`;
}

/**
 * Helper function to split text into sentences
 */
function splitIntoSentences(text) {
  // Split on periods, question marks, and exclamation points
  // but handle common abbreviations and decimal numbers
  const sentenceEnders = /[.!?](?=\s+|$)/g;
  const sentences = [];
  let start = 0;
  let match;
  
  // Use regex to find sentence boundaries
  const regex = /[.!?](?=\s|$)/g;
  while ((match = regex.exec(text)) !== null) {
    // Check if this period is part of an abbreviation
    const isPeriodInAbbreviation = 
      text[match.index] === '.' && 
      match.index > 0 && 
      /[A-Za-z]/.test(text[match.index-1]) && 
      match.index < text.length-2 && 
      /[A-Za-z]/.test(text[match.index+2]);
    
    // If not an abbreviation, it's a sentence boundary
    if (!isPeriodInAbbreviation) {
      sentences.push(text.substring(start, match.index + 1).trim());
      start = match.index + 1;
    }
  }
  
  // Add the last sentence if there's text remaining
  if (start < text.length) {
    sentences.push(text.substring(start).trim());
  }
  
  return sentences;
}

/**
 * Selects all paragraphs for reimagination.
 */
function onSelectAll() {
  for (let i = 0; i < paragraphs.length; i++) {
    // Only select paragraphs that meet our requirements
    if (!shouldSkipParagraph(paragraphs[i])) {
    selectedParagraphs.add(i);
    }
  }
  updateSelectedParagraphList();
  // Update checkbox based on current paragraph
  const checkbox = document.getElementById("checkboxReimagine");
  checkbox.checked = !shouldSkipParagraph(paragraphs[currentParagraphIndex]);
  logMessage("All valid paragraphs selected for reimagination.");
}

/**
 * Unselects all paragraphs.
 */
function onUnselectAll() {
  selectedParagraphs.clear();
  updateSelectedParagraphList();
  document.getElementById("checkboxReimagine").checked = false;
  logMessage("All paragraphs unselected.");
}

/**
 * Applies changes only to the current paragraph: replaces its text with the saved modified text.
 */
async function onApply() {
  if (paragraphs.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    logMessage("No valid paragraph selected to apply changes.");
    return;
  }
  
  if (!paragraphLLMOutputs.hasOwnProperty(currentParagraphIndex)) {
    logMessage(`No modified text for paragraph ${currentParagraphIndex + 1}`);
    return;
  }
  
  // Get the current text from the UI (might have been manually edited)
  const modifiedTextElement = document.getElementById("modifiedText");
  const currentModifiedText = modifiedTextElement ? modifiedTextElement.value : null;
  
  logMessage(`Applying changes to paragraph ${currentParagraphIndex + 1}`);
  
  await Word.run(async (context) => {
    try {
      // Get all paragraphs
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();
      
      if (currentParagraphIndex >= paras.items.length) {
        logMessage(`Error: Paragraph index ${currentParagraphIndex} out of bounds`);
        return;
      }
      
      // Get the paragraph and output
      const paragraph = paras.items[currentParagraphIndex];
      const output = paragraphLLMOutputs[currentParagraphIndex];
      
      // Use modified text from UI if available, otherwise use stored LLM output
      const textToInsert = currentModifiedText || output.text;
      
      // Replace the text
      paragraph.insertText(textToInsert, "Replace");
    await context.sync();
    
      // Apply formatting if available
      if (output.formattingMatches && output.formattingMatches.length > 0) {
        await applyFormatting(context, paragraph, output.formattingMatches);
      }
      
      // Update only this paragraph in our source cache
      paragraphs[currentParagraphIndex] = textToInsert;
      
      // Clear this entry from LLM outputs since it's now applied
    delete paragraphLLMOutputs[currentParagraphIndex];
    selectedParagraphs.delete(currentParagraphIndex);
      
      logMessage(`Successfully applied changes to paragraph ${currentParagraphIndex + 1}`);
      
      // Update UI
      updateUIForSelectedParagraph();
    updateSelectedParagraphList();
      
    } catch (error) {
      logMessage(`Error in onApply: ${error}`);
    }
  }).catch((error) => logMessage(`Error in Word.run: ${error}`));
}

/**
 * Applies changes to all selected paragraphs.
 */
async function onApplyAll() {
  if (paragraphs.length === 0) {
    logMessage("No paragraphs available to apply changes.");
    return;
  }
  await Word.run(async (context) => {
    // Only get paragraphs from the main document body, not headers/footers
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();
    
    // Track how many paragraphs we've processed
    let appliedCount = 0;
    let skippedCount = 0;
    
    for (let idx of selectedParagraphs) {
      if (paragraphLLMOutputs.hasOwnProperty(idx) && idx < paras.items.length) {
        try {
          const output = paragraphLLMOutputs[idx];
          const paragraph = paras.items[idx];
          
          // Verify this is a paragraph in the main document body
          // by checking if it has text that matches what we expect
          paragraph.load("text");
          await context.sync();
          
          // If paragraph text doesn't match what we expect, it might be in a header/footer
          // or the document structure might have changed
          if (!paragraph.text.includes(paragraphs[idx].substring(0, 30))) {
            logMessage(`Skipping paragraph ${idx + 1} - text doesn't match expected content`);
            skippedCount++;
            continue;
          }
          
          // Insert the new text
          paragraph.insertText(output.text, "Replace");
          await context.sync();
          
          // Apply formatting
          if (output.formattingMatches) {
            try {
              await applyFormatting(context, paragraph, output.formattingMatches);
            } catch (error) {
              logMessage(`Error applying formatting to paragraph ${idx + 1}: ${error}`);
            }
          } else {
            // If no special formatting, apply the original paragraph's default font settings
            const defaultFont = paragraphFormatting[idx]?.defaultFont;
            if (defaultFont) {
              paragraph.font.size = defaultFont.size;
              paragraph.font.name = defaultFont.name;
              paragraph.font.bold = defaultFont.bold;
              paragraph.font.italic = defaultFont.italic;
              await context.sync();
              logMessage(`Applied default font settings to paragraph ${idx + 1}`);
            }
          }
          
          logMessage(`Applied text and formatting to paragraph ${idx + 1}`);
        delete paragraphLLMOutputs[idx];
          appliedCount++;
        } catch (error) {
          logMessage(`Error processing paragraph ${idx + 1}: ${error}`);
          skippedCount++;
        }
      } else {
        logMessage(`No modified text for paragraph ${idx + 1}`);
        skippedCount++;
      }
    }
    
    await context.sync();
    logMessage(`Apply All completed: ${appliedCount} paragraphs updated, ${skippedCount} skipped`);
    
    // Reload paragraphs and update UI
    await loadParagraphsFromDocument();
    // Clear all selections since we've applied everything
    selectedParagraphs.clear();
    updateSelectedParagraphList();
  }).catch((error) => logMessage("Error in onApplyAll: " + error));
}

/**
 * Loads available models from the global variable
 */
async function loadAvailableModels() {
  try {
    logMessage("Starting to load models...");
    
    // Use the global variable defined in HTML
    if (window.availableModels && Array.isArray(window.availableModels)) {
      availableModels = window.availableModels;
      logMessage(`Available models loaded: ${availableModels.length}`);
      logMessage(`Models: ${JSON.stringify(availableModels)}`);
      
      // Populate the dropdown
      const modelSelect = document.getElementById("modelSelect");
      modelSelect.innerHTML = ""; // Clear existing options
      
      availableModels.forEach(model => {
        const option = document.createElement("option");
        option.value = model;
        option.textContent = model;
        modelSelect.appendChild(option);
        logMessage(`Added option: ${model}`);
      });
      
      // Set the first model as selected
      if (availableModels.length > 0) {
        selectedModel = availableModels[0];
        modelSelect.value = selectedModel;
        logMessage(`Selected model: ${selectedModel}`);
      }
    } else {
      throw new Error("Global models variable not found or not an array");
    }
  } catch (error) {
    logMessage(`Error loading models: ${error}`);
    
    // Add a default model as fallback
    selectedModel = "qwq:latest";
    logMessage("Using fallback model: " + selectedModel);
    
    const modelSelect = document.getElementById("modelSelect");
    modelSelect.innerHTML = ""; // Clear existing options
    const option = document.createElement("option");
    option.value = selectedModel;
    option.textContent = selectedModel;
    modelSelect.appendChild(option);
    modelSelect.value = selectedModel;
  }
}

/**
 * Handles changes to the model selection dropdown
 */
function onModelSelectChange() {
  const select = document.getElementById("modelSelect");
  selectedModel = select.value;
  logMessage("Model selection changed to: " + selectedModel);
}

/**
 * Calls our local LLM-like service with a paragraph and prompt.
 */
async function callOllama(paragraphText, userPrompt) {
  try {
    const requestData = {
      paragraphText: paragraphText,
      userPrompt: userPrompt,
      model: selectedModel // Add the selected model to the request
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
      return paragraphText; // fallback: return original text on error.
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
 * Helper: gets the currently selected text in a <textarea> by ID.
 */
function getSelectedTextInTextarea(textareaId) {
  const textarea = document.getElementById(textareaId);
  return textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
}

/**
 * Helper: replaces the selected text in a <textarea> with new text.
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

/**
 * Applies formatting to a paragraph based on the formatting matches.
 */
async function applyFormatting(context, paragraph, formattingMatches) {
  try {
    for (const match of formattingMatches) {
      try {
        if (!match.newText || match.newText.trim().length < 1) {
          logMessage("Skipping empty formatting match");
          continue;
        }
        
        // Limit search text to 255 characters to avoid "too long" error
        const searchText = match.newText.trim().substring(0, Math.min(255, match.newText.trim().length));
        
        if (searchText.length < 3) {
          logMessage(`Search text too short: "${searchText}"`);
          continue;
        }
        
        logMessage(`Searching for text to format: "${searchText}"`);
        
        // Search for the text
        const ranges = paragraph.search(searchText, {matchCase: false, matchWholeWord: false});
        ranges.load("items");
        await context.sync();
        
        if (ranges.items.length > 0) {
          const range = ranges.items[0];
          
          // Apply formatting
          if (match.formatting.bold !== null) range.font.bold = match.formatting.bold;
          if (match.formatting.italic !== null) range.font.italic = match.formatting.italic;
          if (match.formatting.size !== null) range.font.size = match.formatting.size;
          
          await context.sync();
          logMessage(`Applied formatting to: "${searchText}"`);
        } else {
          logMessage(`Could not find text to format: "${searchText}"`);
        }
      } catch (matchError) {
        logMessage(`Error processing format match: ${matchError}`);
        // Continue with next match
      }
    }
  } catch (error) {
    logMessage(`Error in applyFormatting: ${error}`);
  }
}

/**
 * Helper function to escape special characters in a string for use in a RegExp
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Debug function to log paragraph info
 */
function debugParagraphInfo() {
  logMessage("------ DEBUG INFO ------");
  logMessage(`paragraphs array length: ${paragraphs.length}`);
  logMessage(`currentParagraphIndex: ${currentParagraphIndex}`);
  
  if (paragraphs.length > 0 && currentParagraphIndex >= 0 && currentParagraphIndex < paragraphs.length) {
    logMessage(`Current paragraph content: "${paragraphs[currentParagraphIndex].substring(0, 30)}..."`);
  } else {
    logMessage("Current paragraph content: [INVALID INDEX]");
  }
  logMessage("------------------------");
}

/**
 * Navigates to the next paragraph.
 */
function onNextParagraph() {
  if (paragraphs.length === 0) return;
  currentParagraphIndex = (currentParagraphIndex + 1) % paragraphs.length;
  debugParagraphInfo();
  
  try {
    updateUIForSelectedParagraph();
    logMessage(`Navigated to paragraph ${currentParagraphIndex + 1}`);
  } catch (error) {
    logMessage(`Error in onNextParagraph: ${error}`);
  }
}

/**
 * Navigates to the previous paragraph.
 */
function onPrevParagraph() {
  if (paragraphs.length === 0) return;
  currentParagraphIndex = (currentParagraphIndex - 1 + paragraphs.length) % paragraphs.length;
  debugParagraphInfo();
  
  try {
    updateUIForSelectedParagraph();
    logMessage(`Navigated to paragraph ${currentParagraphIndex + 1}`);
  } catch (error) {
    logMessage(`Error in onPrevParagraph: ${error}`);
  }
}

/**
 * Updates the UI to show the text for the currently selected paragraph.
 */
function updateUIForSelectedParagraph() {
  if (paragraphs.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphs.length) {
    document.getElementById("originalText").value = "";
    document.getElementById("modifiedText").value = "";
    document.getElementById("paragraphLabel").textContent = "No paragraph selected";
    document.getElementById("reimagineCheckbox").checked = false;
    return;
  }
  
  // Update the paragraph number display
  document.getElementById("paragraphLabel").textContent = 
    `Paragraph ${currentParagraphIndex + 1} of ${paragraphs.length}`;
  
  // Display the original text
  document.getElementById("originalText").value = paragraphs[currentParagraphIndex];
  
  // Display the modified text if available, otherwise show empty box
  if (paragraphLLMOutputs.hasOwnProperty(currentParagraphIndex)) {
    document.getElementById("modifiedText").value = paragraphLLMOutputs[currentParagraphIndex].text;
  } else {
    document.getElementById("modifiedText").value = "";
  }
  
  // Update the reimagine checkbox
  document.getElementById("reimagineCheckbox").checked = selectedParagraphs.has(currentParagraphIndex);
  
  // Update highlights in the document
  highlightCurrentParagraph();
}

/**
 * Refreshes the source cache and clears the reimagined cache.
 */
async function onRefresh() {
  logMessage("Refreshing document...");
  
  // Clear the reimagined cache
  paragraphLLMOutputs = {};
  
  // Reload paragraphs from document
  await loadParagraphsFromDocument();
  
  // Update UI to reflect changes
  updateUIForSelectedParagraph();
  updateSelectedParagraphList();
  
  logMessage("Document refreshed - all paragraphs reloaded and LLM outputs cleared");
}

/**
 * For long paragraphs, breaks the text into searchable chunks.
 */
function getSearchableChunks(text, maxLength = 255) {
  const chunks = [];
  let start = 0;
  
  while (start < text.length) {
    // Find a good breaking point (space, period, etc.)
    let end = Math.min(start + maxLength, text.length);
    
    if (end < text.length) {
      // Try to find a space to break at
      const spaceIndex = text.lastIndexOf(' ', end);
      if (spaceIndex > start && spaceIndex > end - 50) {
        end = spaceIndex;
      }
    }
    
    chunks.push(text.substring(start, end));
    start = end;
  }
  
  return chunks;
}
