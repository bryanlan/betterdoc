/* global Word, Office */

// Global variables.
let paragraphCache = [];         // Array of objects containing paragraph data
let currentParagraphIndex = 0;   // Track which paragraph is currently selected.
let selectedParagraphs = new Set(); // Set of paragraph indexes selected for reimagination.
let paragraphFormatting = {};    // Map: paragraph index -> formatting info
let availableModels = [];
let selectedModel = "";          // Will store the currently selected model

// Each paragraphCache entry has the structure:
// {
//   source: "",        // Original text from document
//   reimagined: {      // LLM output if it exists (null if not reimagined)
//     text: "",
//     formattingMatches: []
//   },
//   reimagineState: false  // Boolean for checkbox state
// }

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
      if (paragraphCache.length > 0) {
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
 * Loads all paragraphs in the Word document into our 'paragraphCache' array
 */
async function loadParagraphsFromDocument() {
  logMessage("Loading paragraphs from document...");
  paragraphCache = [];
  
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("text");
    await context.sync();

    for (let i = 0; i < paras.items.length; i++) {
      paragraphCache.push({
        source: paras.items[i].text,
        reimagined: null,
        reimagineState: false
      });
    }
    
    logMessage(`Loaded ${paragraphCache.length} paragraphs from document`);
  }).catch((error) => {
    logMessage("Error loading paragraphs: " + error);
  });
  
  // Clear selected paragraphs since we're refreshing everything
  selectedParagraphs.clear();
  
  return paragraphCache;
}

/**
 * Selects the current paragraph in the document.
 */
async function selectCurrentParagraph(context) {
  if (paragraphCache.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphCache.length) {
    logMessage("Selection skipped: Invalid index or empty cache.");
    return;
  }

  const body = context.document.body;

  try {
    // --- Step 1: Get the text to search for ---
    const paragraphText = paragraphCache[currentParagraphIndex].source;
    // Use the first 100 characters for searching, as before
    const searchText = paragraphText.substring(0, Math.min(100, paragraphText.length));

    if (searchText.trim().length < 3) {
      logMessage(`Selection skipped: Paragraph text too short ("${searchText.substring(0, 10)}...")`);
      return;
    }

    logMessage(`Searching for text to select: "${searchText.substring(0, 30)}..." (${searchText.length} chars)`);

    // --- Step 2: Search for the current paragraph's text ---
    const searchOptions = { matchCase: false, matchWholeWord: false };
    const searchResults = body.search(searchText, searchOptions);
    // We only need the items themselves to select one
    searchResults.load("items");
    await context.sync();

    if (searchResults.items.length > 0) {
      const firstMatch = searchResults.items[0];

      // --- Step 3: Select the found text ---
      firstMatch.select(); // The core change: select the range
      await context.sync();
      logMessage(`Text selected successfully: "${firstMatch.text.substring(0, 30)}..."`);

    } else {
      logMessage(`Could not find paragraph text in document to select: "${searchText.substring(0, 30)}..."`);
    }
  } catch (error) {
    // Log the full error for better debugging
    logMessage(`--- ERROR during text selection ---`);
    logMessage(`Error message: ${error.message}`);
    if (error instanceof OfficeExtension.Error) {
      logMessage("OfficeExtension Error Debug Info: " + JSON.stringify(error.debugInfo));
    }
    console.error("Selection Error Details:", error);
    logMessage(`--- End ERROR ---`);
  }
}

/**
 * Updates the UI with the current paragraph's text.
 */
function updateUIForSelectedParagraph() {
  if (paragraphCache.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphCache.length) {
    document.getElementById("originalText").value = "";
    document.getElementById("modifiedText").value = "";
    document.getElementById("paragraphLabel").textContent = "No paragraph selected";
    document.getElementById("checkboxReimagine").checked = false;
    return;
  }
  
  // Update the paragraph number display
  document.getElementById("paragraphLabel").textContent = 
    `Paragraph ${currentParagraphIndex + 1} of ${paragraphCache.length}`;
  
  // Display the original text
  document.getElementById("originalText").value = paragraphCache[currentParagraphIndex].source;
  
  // Display reimagined text if available, otherwise show empty box
  const reimaginedText = paragraphCache[currentParagraphIndex].reimagined?.text;
  document.getElementById("modifiedText").value = reimaginedText || "";
  
  // Update the reimagine checkbox
  document.getElementById("checkboxReimagine").checked = paragraphCache[currentParagraphIndex].reimagineState;
  
  // Update highlights in the document - this is now handled by the calling functions
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
  if (index >= 0 && index < paragraphCache.length && index !== currentParagraphIndex) {
      currentParagraphIndex = index;
      Word.run(async (context) => {
        await selectCurrentParagraph(context); // Use the new function
        await context.sync(); // Ensure selection syncs
        updateUIForSelectedParagraph();
        logMessage("Jumped to paragraph " + (index + 1));
      }).catch((error) => logMessage("Error in goToParagraph: " + error));
  } else if (index === currentParagraphIndex) {
      // If jumping to the same index, still ensure selection and UI update
       Word.run(async (context) => {
        await selectCurrentParagraph(context); 
        await context.sync(); 
        updateUIForSelectedParagraph();
      }).catch((error) => logMessage("Error re-selecting current paragraph: " + error));
  } else {
       logMessage(`Invalid index for goToParagraph: ${index}`);
       updateUIForSelectedParagraph(); // Still update UI if jump is invalid
  }
}

/**
 * Handler for changes to the "Reimagine" checkbox.
 */
function onCheckboxChange() {
  const checkbox = document.getElementById("checkboxReimagine");
  paragraphCache[currentParagraphIndex].reimagineState = checkbox.checked;
  
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
  
  // Don't save if it's empty
  if (!modifiedBox.value.trim()) {
    paragraphCache[currentParagraphIndex].reimagined = null;
    return;
  }
  
  // If the modified text is the same as original, clear reimagined
  if (modifiedBox.value.trim() === originalBox.value.trim()) {
    paragraphCache[currentParagraphIndex].reimagined = null;
    modifiedBox.value = "";
  } else {
    // Save the modified text to our cache
    paragraphCache[currentParagraphIndex].reimagined = {
      text: modifiedBox.value,
      formattingMatches: paragraphCache[currentParagraphIndex].reimagined?.formattingMatches || []
    };
  }
  
  logMessage(`Saved modified text for paragraph ${currentParagraphIndex + 1}`);
}

/**
 * Moves to the previous paragraph (Up).
 */
function onUp() {
  if (currentParagraphIndex > 0) {
    let newIndex = currentParagraphIndex - 1;
    while (newIndex > 0 && shouldSkipParagraph(paragraphCache[newIndex].source)) {
      newIndex--;
    }
    if (newIndex === 0 && shouldSkipParagraph(paragraphCache[newIndex].source)) {
       newIndex = currentParagraphIndex; // Avoid getting stuck on skippable first item
    }
     // Only update if the index actually changed to avoid unnecessary selection calls
    if (newIndex !== currentParagraphIndex) {
        currentParagraphIndex = newIndex;
        Word.run(async (context) => {
          await selectCurrentParagraph(context); // Use the new function
          await context.sync(); // Ensure selection syncs
          updateUIForSelectedParagraph(); 
        }).catch((error) => logMessage("Error in onUp: " + error));
    } else {
         // If index didn't change (e.g., blocked by skippable first item), just update UI
         updateUIForSelectedParagraph();
    }
  } else {
      updateUIForSelectedParagraph(); // Update UI even if at the boundary
  }
}

/**
 * Moves to the next paragraph (Down).
 */
function onDown() {
  if (currentParagraphIndex < paragraphCache.length - 1) {
    let newIndex = currentParagraphIndex + 1;
     // Skipping logic remains the same
    while (newIndex < paragraphCache.length - 1 && shouldSkipParagraph(paragraphCache[newIndex].source)) {
      newIndex++;
    }
     if (newIndex === paragraphCache.length - 1 && shouldSkipParagraph(paragraphCache[newIndex].source)) {
       newIndex = currentParagraphIndex; // Avoid getting stuck on skippable last item
    }
     // Only update if the index actually changed
    if (newIndex !== currentParagraphIndex) {
        currentParagraphIndex = newIndex;
        Word.run(async (context) => {
          await selectCurrentParagraph(context); // Use the new function
          await context.sync(); // Ensure selection syncs
          updateUIForSelectedParagraph();
        }).catch((error) => logMessage("Error in onDown: " + error));
    } else {
         // If index didn't change, just update UI
         updateUIForSelectedParagraph();
    }
  } else {
     updateUIForSelectedParagraph(); // Update UI even if at the boundary
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
          const originalText = paragraphCache[idx].source;
          logMessage(`Processing paragraph ${idx + 1}`);
          
          // First detect formatting
          await detectFormatting(context, paras.items[idx], idx);
          
          // Initialize reimagined structure if it doesn't exist
          if (!paragraphCache[idx].reimagined) {
            paragraphCache[idx].reimagined = {
              text: "",
              formattingMatches: []
            };
          }
          
          try {
            // Get rewritten text from LLM with retry logic
            // **** Combine the base prompt with the actual paragraph text ****
            const combinedPrompt = `${prompt}\n\nParagraph to rewrite:\n${originalText}`;
            const ollamaResult = await callOllama(combinedPrompt);
            
            // Check for errors from callOllama first
            if (ollamaResult.error) {
              throw new Error(`Ollama call failed: ${ollamaResult.error}`);
            }
            
            // Extract the actual response text
            let newText = ollamaResult.response; 

            if (!newText || newText.trim().length === 0) {
              // Check if rawData exists and has a response, maybe parsing was slightly off
              if (ollamaResult.rawData && ollamaResult.rawData.response) {
                 newText = ollamaResult.rawData.response;
                 logMessage("Used response from rawData as fallback.");
              } else {
                 throw new Error("Empty response from LLM");
              }
            }
            
            // Check word count
            const originalWordCount = countWords(originalText);
            const newWordCount = countWords(newText);
            
            // If new text is more than 10% longer, ask for a shorter version
            if (newWordCount > originalWordCount * 1.1) {
              logMessage(`Word count exceeded: Original=${originalWordCount}, New=${newWordCount}`);
              
              // Create a follow-up prompt asking for shorter text
              const followUpBasePrompt = `${prompt}\n\nYou failed to do as instructed and exceeded the word count. Try again and reduce your word count. The original had ${originalWordCount} words, yours had ${newWordCount}.`;
              // **** Combine follow-up prompt with the problematic *new* text ****
              const combinedFollowUpPrompt = `${followUpBasePrompt}\n\nRewrite this text specifically:\n${newText}`;
              
              // Call LLM again with the follow-up prompt
              const shorterOllamaResult = await callOllama(combinedFollowUpPrompt);
              
              // Handle potential error from the second call
              if (shorterOllamaResult.error) {
                 logMessage(`Warning: Follow-up Ollama call failed: ${shorterOllamaResult.error}. Using previous text.`);
              } else {
                  const shorterText = shorterOllamaResult.response;
                  if (shorterText && shorterText.trim().length > 0) {
                    newText = shorterText;
                  }
              }
              
              // Log the new word count
              const revisedWordCount = countWords(newText);
              logMessage(`Revised word count: ${revisedWordCount}`);
            }
            
            // Store the new text
            paragraphCache[idx].reimagined.text = newText;
            
            // Process each non-modal phrase to find its match in new text
            const formattingMatches = [];
            for (const phrase of paragraphFormatting[idx].formattedClauses) {
              try {
                // Prepare clean versions for comparison
                const cleanPhrase = phrase.text.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
                const cleanNewText = newText.toLowerCase().replace(/[.,;:!?()]/g, '').trim();
                
                // Check if the exact phrase exists as a whole word in the new text
                const phraseRegex = new RegExp(`\\b${escapeRegExp(cleanPhrase)}\\b`, 'i');
                
                if (phraseRegex.test(cleanNewText)) {
                  // Direct match found
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
                  // **** The matchPrompt already includes necessary context ****
                  const matchPrompt = createClauseMatchPrompt(originalText, newText, phrase.text);
                  const matchingClauseResult = await callOllama(matchPrompt);
                  let matchingClause;
                  
                  if (matchingClauseResult.error) {
                     logMessage(`Warning: Clause match Ollama call failed: ${matchingClauseResult.error}. Using original phrase.`);
                     matchingClause = phrase.text; // Fallback to original on error
                  } else {
                      matchingClause = matchingClauseResult.response;
                      if (!matchingClause || matchingClause.trim().length === 0) {
                        // If LLM fails to provide a clause, use original phrase
                        matchingClause = phrase.text;
                      }
                  }
                  
                  formattingMatches.push({
                    newText: matchingClause.trim(),
                    formatting: phrase.formatting
                  });
                }
              } catch (phraseError) {
                logMessage(`Error processing phrase "${phrase.text}": ${phraseError}`);
                // Add original phrase if there's an error
                formattingMatches.push({
                  newText: phrase.text,
                  formatting: phrase.formatting
                });
              }
            }
            
            // Store the formatting matches
            paragraphCache[idx].reimagined.formattingMatches = formattingMatches;
            
          } catch (llmError) {
            logMessage(`LLM processing error for paragraph ${idx + 1}: ${llmError}`);
            // Clear reimagined content on error
            paragraphCache[idx].reimagined = null;
            continue;
          }
          
          // Update UI if we're on this paragraph
          if (idx === currentParagraphIndex) {
            updateUIForSelectedParagraph();
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
  for (let i = 0; i < paragraphCache.length; i++) {
    // Only select paragraphs that meet our requirements
    if (!shouldSkipParagraph(paragraphCache[i].source)) {
      selectedParagraphs.add(i);
      paragraphCache[i].reimagineState = true;
    }
  }
  updateSelectedParagraphList();
  
  // Update checkbox based on current paragraph
  const checkbox = document.getElementById("checkboxReimagine");
  checkbox.checked = !shouldSkipParagraph(paragraphCache[currentParagraphIndex].source) && 
                    paragraphCache[currentParagraphIndex].reimagineState;
  logMessage("All valid paragraphs selected for reimagination.");
}

/**
 * Unselects all paragraphs.
 */
function onUnselectAll() {
  selectedParagraphs.clear();
  paragraphCache.forEach(para => para.reimagineState = false);
  updateSelectedParagraphList();
  document.getElementById("checkboxReimagine").checked = false;
  logMessage("All paragraphs unselected.");
}

/**
 * Applies changes only to the current paragraph
 */
async function onApply() {
  if (paragraphCache.length === 0 || currentParagraphIndex < 0 || currentParagraphIndex >= paragraphCache.length) {
    logMessage("No valid paragraph selected to apply changes.");
    return;
  }
  
  if (!paragraphCache[currentParagraphIndex].reimagined?.text) {
    logMessage(`No modified text for paragraph ${currentParagraphIndex + 1}`);
    return;
  }
  
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
      
      // Get the paragraph and reimagined text
      const paragraph = paras.items[currentParagraphIndex];
      const reimaginedText = paragraphCache[currentParagraphIndex].reimagined.text;
      
      // Replace the text
      paragraph.insertText(reimaginedText, "Replace");
      await context.sync();
      
      // Apply formatting if available
      if (paragraphCache[currentParagraphIndex].reimagined.formattingMatches?.length > 0) {
        await applyFormatting(context, paragraph, paragraphCache[currentParagraphIndex].reimagined.formattingMatches);
      }
      
      // Update source cache with the new text
      paragraphCache[currentParagraphIndex].source = reimaginedText;
      
      // Clear reimagined content and state
      paragraphCache[currentParagraphIndex].reimagined = null;
      paragraphCache[currentParagraphIndex].reimagineState = false;
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
 * Applies changes to all selected paragraphs
 */
async function onApplyAll() {
  if (paragraphCache.length === 0) {
    logMessage("No paragraphs available to apply changes.");
    return;
  }
  
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();
    
    let appliedCount = 0;
    let skippedCount = 0;
    
    for (let idx of selectedParagraphs) {
      if (paragraphCache[idx].reimagined?.text && idx < paras.items.length) {
        try {
          const paragraph = paras.items[idx];
          const reimaginedText = paragraphCache[idx].reimagined.text;
          
          // Replace the text
          paragraph.insertText(reimaginedText, "Replace");
          await context.sync();
          
          // Apply formatting
          if (paragraphCache[idx].reimagined.formattingMatches?.length > 0) {
            await applyFormatting(context, paragraph, paragraphCache[idx].reimagined.formattingMatches);
          }
          
          // Update source cache with the new text
          paragraphCache[idx].source = reimaginedText;
          
          // Clear reimagined content and state
          paragraphCache[idx].reimagined = null;
          paragraphCache[idx].reimagineState = false;
          
          logMessage(`Applied changes to paragraph ${idx + 1}`);
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
    
    // Clear all selections since we've applied everything
    selectedParagraphs.clear();
    updateSelectedParagraphList();
    updateUIForSelectedParagraph();
    
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
async function callOllama(prompt, maxRetries = 3) {
  const selectedModel = document.getElementById("modelSelect").value;
  if (!selectedModel) {
    logMessage("Error: No model selected.");
    return { error: "No model selected" };
  }
  const url = `http://localhost:11434/api/generate`; // Ensure this matches your Ollama endpoint

  const body = JSON.stringify({
    model: selectedModel,
    prompt: prompt,
    stream: false, // Ensure response is not streamed for easier handling
  });

  // Log the full prompt being sent
  logMessage(`Sending request to Ollama (${selectedModel}). Full Prompt: ${prompt}`);

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: body,
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // Get raw text first
      const rawResponseText = await response.text();
      logMessage(`Raw server response received: ${rawResponseText}`); // Log the raw response

      // Now parse the raw text
      const data = JSON.parse(rawResponseText); 
      
      if (!data || !data.response) {
        logMessage(`Warning: Ollama response missing 'response' field. Full data: ${JSON.stringify(data)}`);
        // Depending on expected behavior, you might return null or throw an error
        // For now, let's return the structure indicating an issue but not crash
        return { response: null, error: "Invalid response structure from Ollama", rawData: data }; 
      }
      
      logMessage("Server response parsed successfully.");
      return data; // Return the parsed data

    } catch (error) {
      logMessage(`Attempt ${attempt} failed: ${error.message}`);
      if (attempt === maxRetries) {
        logMessage("Max retries reached. Failing.");
        return { error: error.message };
      }
      const delay = Math.pow(2, attempt - 1) * 1000; // Exponential backoff
      logMessage(`Retrying in ${delay / 1000} seconds...`);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  // Should not be reached if maxRetries > 0, but added for safety
  return { error: "Failed after multiple retries" }; 
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
        
        // Limit search text to 100 characters to avoid "too long" error
        const searchText = match.newText.trim().substring(0, Math.min(100, match.newText.trim().length));
        
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
          // Try searching for chunks if the text is too long
          const chunks = getSearchableChunks(match.newText);
          let found = false;
          
          for (const chunk of chunks) {
            if (chunk.trim().length < 3) continue;
            
            const chunkRanges = paragraph.search(chunk, {matchCase: false, matchWholeWord: false});
            chunkRanges.load("items");
            await context.sync();
            
            if (chunkRanges.items.length > 0) {
              const range = chunkRanges.items[0];
              
              // Apply formatting
              if (match.formatting.bold !== null) range.font.bold = match.formatting.bold;
              if (match.formatting.italic !== null) range.font.italic = match.formatting.italic;
              if (match.formatting.size !== null) range.font.size = match.formatting.size;
              
              await context.sync();
              logMessage(`Applied formatting to chunk: "${chunk}"`);
              found = true;
              break;
            }
          }
          
          if (!found) {
            logMessage(`Could not find text to format: "${searchText}"`);
          }
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
  logMessage(`paragraphCache array length: ${paragraphCache.length}`);
  logMessage(`currentParagraphIndex: ${currentParagraphIndex}`);
  logMessage(`Selected paragraphs: ${Array.from(selectedParagraphs).join(', ')}`);
  
  if (paragraphCache.length > 0 && currentParagraphIndex >= 0 && currentParagraphIndex < paragraphCache.length) {
    const current = paragraphCache[currentParagraphIndex];
    logMessage(`Current paragraph:`);
    logMessage(`- Source: "${current.source.substring(0, 30)}..."`);
    logMessage(`- Reimagined: ${current.reimagined ? "Yes" : "No"}`);
    logMessage(`- ReimagineState: ${current.reimagineState}`);
  } else {
    logMessage("Current paragraph: [INVALID INDEX]");
  }
  logMessage("------------------------");
}

/**
 * Navigates to the next paragraph.
 */
function onNextParagraph() {
  if (paragraphCache.length === 0) return;
  currentParagraphIndex = (currentParagraphIndex + 1) % paragraphCache.length;
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
  if (paragraphCache.length === 0) return;
  currentParagraphIndex = (currentParagraphIndex - 1 + paragraphCache.length) % paragraphCache.length;
  debugParagraphInfo();
  
  try {
    updateUIForSelectedParagraph();
    logMessage(`Navigated to paragraph ${currentParagraphIndex + 1}`);
  } catch (error) {
    logMessage(`Error in onPrevParagraph: ${error}`);
  }
}

/**
 * Refreshes the source cache and clears reimagined content
 */
async function onRefresh() {
  logMessage("Refreshing document...");
  
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();
    
    // Create new cache array
    const newCache = [];
    
    for (let i = 0; i < paras.items.length; i++) {
      // Check if paragraph should be skipped before adding to cache
      const text = paras.items[i].text;
      newCache.push({
        source: text,
        reimagined: null,
        reimagineState: false
      });
    }
    
    // Replace the old cache with the new one
    paragraphCache = newCache;
    
    // Clear selected paragraphs
    selectedParagraphs.clear();
    
    // Update UI
    updateUIForSelectedParagraph();
    updateSelectedParagraphList();
    
    logMessage("Document refreshed - all paragraphs reloaded and reimagined content cleared");
  }).catch((error) => logMessage("Error in onRefresh: " + error));
}

/**
 * For long paragraphs, breaks the text into searchable chunks.
 */
function getSearchableChunks(text, maxLength = 100) {
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
