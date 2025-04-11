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
    // Still split by spaces for analysis
    const ranges = paragraph.getTextRanges([" "], false); 
    ranges.load("items/text, items/font/bold, items/font/italic, items/font/size, items/font/underline"); 
    await context.sync();

    // Calculate Modal Format (remains the same logic)
    const formatDistribution = new Map();
    for (let i = 0; i < ranges.items.length; i++) {
      const range = ranges.items[i];
      const formatKey = JSON.stringify({
        bold: range.font.bold,
        italic: range.font.italic,
        size: range.font.size,
        underline: range.font.underline 
      });
      formatDistribution.set(formatKey, (formatDistribution.get(formatKey) || 0) + range.text.length); 
    }
    let modalFormat = null;
    let maxCount = 0;
    for (const [key, count] of formatDistribution) {
      if (count > maxCount) {
        maxCount = count;
        modalFormat = JSON.parse(key);
      }
    }
    // --- End Modal Format Calculation ---

    paragraphFormatting[index] = { defaultFont: modalFormat, formattedClauses: [] };
    logMessage(`\nFormatting analysis for paragraph ${index + 1}:`);
    logMessage(`Modal formatting: ${JSON.stringify(modalFormat)}`);
    
    // --- Revised Second pass for better phrase merging --- 
    let currentPhraseText = "";
    let currentPhraseFormatting = null;

    for (let i = 0; i < ranges.items.length; i++) {
      const range = ranges.items[i];
      const currentText = range.text;
      const currentFormat = { 
         bold: range.font.bold, 
         italic: range.font.italic, 
         size: range.font.size, 
         underline: range.font.underline 
      };
      const currentFormatKey = JSON.stringify(currentFormat);
      const modalFormatKey = JSON.stringify(modalFormat);
      const isNonModal = (currentFormatKey !== modalFormatKey);

      if (isNonModal) {
        // If starting a new phrase OR if the format changes from the previous non-modal
        if (!currentPhraseFormatting || JSON.stringify(currentPhraseFormatting) !== currentFormatKey) {
           // Finalize the previous phrase if it existed
           if (currentPhraseFormatting) { 
              // Trim only trailing spaces, preserve internal ones for matching
              paragraphFormatting[index].formattedClauses.push({
                 text: currentPhraseText.replace(/\s+$/, ''), 
                 formatting: currentPhraseFormatting 
              });
           }
           // Start the new phrase
           currentPhraseText = currentText;
           currentPhraseFormatting = currentFormat;
        } else {
           // Continue the current phrase (same non-modal format)
           currentPhraseText += currentText;
        }
      } else { // Current range IS modal
        // If we were building a non-modal phrase, finalize it now
        if (currentPhraseFormatting) {
           paragraphFormatting[index].formattedClauses.push({
              text: currentPhraseText.replace(/\s+$/, ''),
              formatting: currentPhraseFormatting
           });
           currentPhraseText = "";
           currentPhraseFormatting = null;
        }
        // Append the modal text (like spaces) to the *next* potential non-modal phrase start
        // Or just ignore it if we want clauses to be purely non-modal text? Let's ignore for now.
      }
    }
    
    // Don't forget the last phrase if it was non-modal
    if (currentPhraseFormatting) {
      paragraphFormatting[index].formattedClauses.push({
        text: currentPhraseText.replace(/\s+$/, ''),
        formatting: currentPhraseFormatting
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
      if (!paragraphFormatting[index]) {
          paragraphFormatting[index] = { defaultFont: null, formattedClauses: [] };
      }
      paragraphFormatting[index].formattedClauses = []; 
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
  logMessage("Starting reimagination process...");
  if (selectedParagraphs.size === 0) {
    logMessage("No paragraphs selected for reimagination.");
    return;
  }

  // Disable buttons during processing
  // TODO: Add logic to disable/enable buttons

  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    // **** Load hyperlinks along with items ****
    paras.load("items, items/hyperlinks"); 
    await context.sync();

    let processedCount = 0;
    let errorCount = 0;

    // Create a copy for iteration as selectedParagraphs might change
    const paragraphsToProcess = Array.from(selectedParagraphs); 

    for (let idx of paragraphsToProcess) {
      // Bounds check for cache and actual paragraphs
      if (idx < 0 || idx >= paragraphCache.length || idx >= paras.items.length) {
        logMessage(`Skipping invalid index ${idx}. Cache size: ${paragraphCache.length}, Doc paras: ${paras.items.length}`);
        selectedParagraphs.delete(idx); // Remove invalid index
        errorCount++;
        continue;
      }

      // Check if paragraph is still intended for reimagining
      if (!paragraphCache[idx].reimagineState) {
        // This check might be redundant if paragraphsToProcess uses selectedParagraphs, 
        // but good for safety if selection changes during the async process.
        logMessage(`Skipping paragraph ${idx + 1} as it's no longer marked for reimagining.`);
        continue; 
      }

      const paragraph = paras.items[idx];
      const paragraphRange = paragraph.getRange();

      // --- Call NEW detectFormatting ---
      logMessage(`Detecting formatting for paragraph ${idx + 1}...`);
      try {
        await detectFormatting(context, paragraph, idx);
      } catch (fmtError) {
        if (!paragraphFormatting[idx]) { paragraphFormatting[idx] = { defaultFont: null, formattedClauses: [] }; }
      }
      // --- End NEW detectFormatting ---

      // --- Get Hyperlinks (using getHyperlinkRanges) --- 
      const hyperlinkRanges = paragraphRange.getHyperlinkRanges();
      hyperlinkRanges.load("items/text, items/hyperlink");
      // --- End Get Hyperlinks --- 

      // Load paragraph text needed for multiple steps
      paragraph.load("text");
      await context.sync(); // Sync after format detection and link setup

      const originalText = paragraph.text;
      let prompt = getPrompt(); // Get base prompt instructions

      // ---- Revised Leading Formatting Check ----
      let leadingFormatText = ""; 
      let textToSendToLLM = originalText;
      let leadingFormat = null; 
      let leadingLength = 0;
      let currentPos = 0; // Track position in originalText

      const initialClauses = paragraphFormatting[idx]?.formattedClauses || [];
      const modalFormatKeyCheck = JSON.stringify(paragraphFormatting[idx]?.defaultFont);

      for (let i = 0; i < initialClauses.length; i++) {
          const clause = initialClauses[i];
          // Check if clause text starts exactly at the current position
          if (clause.text && originalText.substring(currentPos).startsWith(clause.text)) {
              // Is this clause non-modal?
              const clauseFormatKey = JSON.stringify(clause.formatting);
              if (clauseFormatKey !== modalFormatKeyCheck) {
                  // It's a leading non-modal clause, append it
                  leadingFormatText += clause.text;
                  currentPos += clause.text.length;
                  if (leadingFormat === null) { // Store the format of the very first one
                    leadingFormat = clause.formatting;
                  }
                  // Check for spaces following the clause which might also be part of the lead
                  const remainingText = originalText.substring(currentPos);
                  const spaceMatch = remainingText.match(/^\s+/);
                  if (spaceMatch) {
                     leadingFormatText += spaceMatch[0];
                     currentPos += spaceMatch[0].length;
                  }

              } else {
                  // We hit a modal clause at the start, stop checking
                  break;
              }
          } else {
              // Clause doesn't start immediately after previous leading text, stop
              break;
          }
      }
      leadingLength = currentPos; // Use final position as length

      if (leadingLength > 0) {
          textToSendToLLM = originalText.substring(leadingLength);
          logMessage(`Detected leading formatting to preserve: "${leadingFormatText}"`);
          logMessage(`Text sent to LLM: "${textToSendToLLM.substring(0, 50)}..."`);
      } else {
          logMessage(`No distinct leading formatting detected.`);
          leadingFormat = null; // Ensure it's null if none found
          leadingFormatText = null; // Explicitly nullify if none found
      }
      // ---- End Revised Leading Formatting Check ----

      // ---- Link Detection and Handling (using getHyperlinkRanges) ----
      let linkToPreserve = null;
      if (hyperlinkRanges.items && hyperlinkRanges.items.length > 0) {
        const firstLinkRange = hyperlinkRanges.items[0];
        const linkText = firstLinkRange.text;
        const linkAddress = firstLinkRange.hyperlink;

        if (linkText && linkText.trim() && linkAddress) {
            linkToPreserve = { text: linkText, address: linkAddress };
            logMessage(`Detected hyperlink via getHyperlinkRanges: Text='${linkToPreserve.text}', Address='${linkToPreserve.address}'`);
            prompt += `\\n\\nImportant: This paragraph contains a hyperlink with the exact text "${linkToPreserve.text}". Do NOT change the words within this specific phrase.`;
        }
      }
      // ---- End Link Detection ----

      try {
        // **** Combine prompt with the potentially shortened textToSendToLLM ****
        const combinedPrompt = `${prompt}\\n\\nParagraph to rewrite:\\n${textToSendToLLM}`;

        let newText = null;
        let ollamaResult = null;
        let retries = 0;
        const maxRetries = 3; // Max retries for link text preservation

        while (retries <= maxRetries) {
          logMessage(`Calling Ollama for paragraph ${idx + 1}, attempt ${retries + 1}...`);
          ollamaResult = await callOllama(combinedPrompt); 

          if (ollamaResult.error) {
            throw new Error(`Ollama call failed on attempt ${retries}: ${ollamaResult.error}`);
          }

          newText = ollamaResult.response;

          if (linkToPreserve && newText && !newText.includes(linkToPreserve.text)) {
            retries++;
            if (retries > maxRetries) {
              logMessage(`Warning: LLM failed to preserve link text "${linkToPreserve.text}" after ${maxRetries} retries for paragraph ${idx + 1}. Reverting to original text PART for LLM.`);
              newText = textToSendToLLM; // Revert only the LLM part to its original
              // Need to signal that original text was used for this part
              paragraphCache[idx].revertedLLMPart = true; 
              break; 
            } else {
              logMessage(`Warning: Link text "${linkToPreserve.text}" not found in LLM response for paragraph ${idx + 1}. Retrying (${retries}/${maxRetries})...`);
              await new Promise(resolve => setTimeout(resolve, 500 * retries)); 
            }
          } else {
             if (linkToPreserve && newText && newText.includes(linkToPreserve.text)){
                logMessage(`Link text "${linkToPreserve.text}" preserved successfully on attempt ${retries + 1}.`);
             }
             paragraphCache[idx].revertedLLMPart = false; // Mark as not reverted
            break; 
          }
        } // End retry while loop

        let finalLLMText = newText; // Holds the result for the part sent to LLM

        // Fallback check if newText is still invalid
        if (!finalLLMText || finalLLMText.trim().length === 0) {
           if (ollamaResult && ollamaResult.rawData && ollamaResult.rawData.response) {
              finalLLMText = ollamaResult.rawData.response;
              logMessage("Used response from rawData as fallback.");
              if (linkToPreserve && !finalLLMText.includes(linkToPreserve.text)) {
                 logMessage(`Warning: Link text "${linkToPreserve.text}" still not found after rawData fallback. Reverting LLM part to original.`);
                 finalLLMText = textToSendToLLM; 
                 paragraphCache[idx].revertedLLMPart = true; 
              }
           } else {
               logMessage(`Warning: Empty response from LLM for paragraph ${idx + 1} after checks/retries. Reverting LLM part to original.`);
               finalLLMText = textToSendToLLM; 
               paragraphCache[idx].revertedLLMPart = true; 
           }
        }

        // --- Word count check adjustments ---
        const originalFullWordCount = countWords(originalText); // Use full original for comparison baseline
        const currentLLMWordCount = countWords(finalLLMText); 
        const combinedWordCount = countWords(leadingFormatText || "") + currentLLMWordCount;
        const wordCountThreshold = 1.1; // Allow 10% increase

        // Only perform word count check if the LLM part wasn't reverted
        if (!paragraphCache[idx].revertedLLMPart && 
            combinedWordCount > originalFullWordCount * wordCountThreshold) {
           logMessage(`Combined word count exceeded for paragraph ${idx + 1}. Original: ${originalFullWordCount}, Combined (Leading+LLM): ${combinedWordCount}. Requesting shorter version...`);
           
           const followUpBasePrompt = `${prompt}\\n\\nYou failed to do as instructed and exceeded the word count. Try again and make the text shorter. The original paragraph had ${originalFullWordCount} words, your combined response implies ${combinedWordCount} words. Only provide the rewritten text for the section you were asked to process.`;
           let combinedFollowUpPrompt = `${followUpBasePrompt}\\n\\nRewrite this text specifically to be shorter:\\n${finalLLMText}`; // Send only the LLM's output for shortening
            if (linkToPreserve) {
                combinedFollowUpPrompt += `\\n\\nRemember: Keep the exact phrase "${linkToPreserve.text}" unchanged.`;
            }

           const shorterOllamaResult = await callOllama(combinedFollowUpPrompt);

           if (shorterOllamaResult.error) {
              logMessage(`Warning: Follow-up Ollama call failed: ${shorterOllamaResult.error}. Using previous (long) LLM text part.`);
           } else {
               const shorterText = shorterOllamaResult.response;
               if (shorterText && shorterText.trim().length > 0) {
                   // Check link preservation AGAIN for the shorter text
                   if (linkToPreserve && !shorterText.includes(linkToPreserve.text)) {
                       logMessage(`Warning: Link text "${linkToPreserve.text}" lost during shortening. Using previous (long) LLM text part.`);
                   } else {
                       logMessage(`Successfully shortened LLM text part for paragraph ${idx + 1}. New word count (LLM part): ${countWords(shorterText)}`);
                       finalLLMText = shorterText; // Use the shorter version for the LLM part
                   }
               } else {
                  logMessage("Warning: Follow-up call provided empty response. Using previous (long) LLM text part.");
               }
           }
        }

        // --- Revised Calculate Formatting Matches --- 
        let formattingMatches = [];
        if (finalLLMText !== textToSendToLLM) { // Only calculate if LLM changed its part
           logMessage(`Calculating formatting matches within LLM response part...`);
           let searchStartPos = 0; // Track position within originalText to correlate clauses
           for (let i = 0; i < initialClauses.length; i++) {
               const clause = initialClauses[i];
               const clausePos = originalText.indexOf(clause.text, searchStartPos);

               // Skip the clause(s) IF they constituted the leading format text
               if (leadingLength > 0 && clausePos !== -1 && clausePos < leadingLength) {
                   logMessage(`Skipping clause ("${clause.text}") because it's within the leading format block.`);
                   searchStartPos = clausePos + clause.text.length; // Update search position
                   continue; 
               }
               searchStartPos = clausePos + clause.text.length; // Update search pos for next iteration
               
               // Check direct match in LLM output
               if (clause.text && finalLLMText.includes(clause.text)) { 
                   formattingMatches.push({ newText: clause.text, format: clause.formatting });
                   logMessage(`Found direct match for "${clause.text}" in LLM output.`);
               } else { // Attempt LLM clause match only if direct match fails
                   logMessage(`Direct match for "${clause.text}" not found. Attempting LLM clause match...`);
                   try {
                       const matchPrompt = createClauseMatchPrompt(originalText, finalLLMText, clause.text);
                       const matchingClauseResult = await callOllama(matchPrompt); 
                       let matchingClauseText;
                       
                       if (matchingClauseResult.error) { /* log warning */ logMessage(`Warning: Clause match Ollama call failed: ${matchingClauseResult.error}. Skipping format for "${clause.text}".`); } 
                       else {
                           matchingClauseText = matchingClauseResult.response?.trim();
                           if (!matchingClauseText || matchingClauseText.length === 0 || matchingClauseText.toLowerCase().includes("not found") || matchingClauseText.toLowerCase().includes("no match")) { 
                              /* log warning */ logMessage(`Warning: LLM returned empty or no clause match for "${clause.text}". Skipping.`); 
                           } else {
                               // **** VALIDATION STEP ****
                               if (finalLLMText.includes(matchingClauseText)) {
                                   logMessage(`LLM identified matching clause: "${matchingClauseText}" AND validated it exists in response.`);
                                   formattingMatches.push({ 
                                       newText: matchingClauseText, 
                                       format: clause.formatting 
                                   });
                               } else {
                                   logMessage(`LLM identified clause "${matchingClauseText}", but it was NOT FOUND in the final LLM text ("${finalLLMText.substring(0,50)}..."). Discarding match.`);
                               }
                           }
                       }
                   } catch (matchError) { /* log error */ logMessage(`Error during LLM clause matching for "${clause.text}": ${matchError}. Skipping.`); }
               }
           }
           logMessage(`Found ${formattingMatches.length} validated formatting matches within LLM output.`);
        } else { logMessage(`Skipping formatting match calculation for paragraph ${idx+1} as original LLM part was used.`); }
        // --- End Revised Calculate Formatting Matches --- 

        // --- Store final results in cache --- 
        const combinedFinalText = (leadingFormatText || "") + finalLLMText;
        paragraphCache[idx].reimagined = {
          text: combinedFinalText,
          formattingMatches: formattingMatches, 
          link: (linkToPreserve && combinedFinalText.includes(linkToPreserve.text)) ? linkToPreserve : null,
          leadingFormatText: leadingFormatText,
          leadingFormat: leadingFormat,
          revertedLLMPart: paragraphCache[idx].revertedLLMPart // Carry over revert status
        };
        
        processedCount++;
        logMessage(`Successfully processed paragraph ${idx + 1}.`);

      } catch (error) {
        logMessage(`LLM processing error loop for paragraph ${idx + 1}: ${error.message}`);
        console.error(error);
        paragraphCache[idx].reimagined = { 
            text: originalText, 
            error: error.message, 
            link: linkToPreserve, 
            formattingMatches: [],
            leadingFormatText: null, 
            leadingFormat: null,
            revertedLLMPart: true // Mark as reverted on error
        }; 
        errorCount++;
      }
    } // End loop
    logMessage(`Reimagination complete: ${processedCount} processed, ${errorCount} errors.`);
    updateSelectedParagraphList(); 
    updateUIForSelectedParagraph(); 
  }).catch((error) => {
    logMessage(`Error in Word.run for onReimagine: ${error}`);
    console.error("Word.run Reimagination Error:", error);
    if (error instanceof OfficeExtension.Error) {
        logMessage("OfficeExtension Error Debug Info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/**
 * Helper function to count words in a string
 */
function countWords(str) {
  if (!str) return 0;
  // Basic word count, splits on spaces and filters empty strings
  return str.trim().split(/\s+/).filter(Boolean).length;
}

/**
 * Creates a prompt to find matching clause in new text with focused context
 */
function createClauseMatchPrompt(originalText, newText, originalClause) {
  // Split texts into sentences
  const originalSentences = splitIntoSentences(originalText);
  const newSentences = splitIntoSentences(newText);
  
  // Find the sentence containing the clause
  let containingSentence = originalSentences.find(s => s.includes(originalClause)) || originalText;

  // Calculate relative position in original text
  const relativePosition = originalSentences.indexOf(containingSentence) / originalSentences.length;
  
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
It appeared in this source text: "${containingSentence}" 

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
  
  const cacheEntry = paragraphCache[currentParagraphIndex];
  const reimaginedText = cacheEntry.reimagined?.text;
  const leadingFormatText = cacheEntry.reimagined?.leadingFormatText;
  const leadingFormat = cacheEntry.reimagined?.leadingFormat;
  
  if (!reimaginedText) {
    logMessage(`No modified/reimagined text available for paragraph ${currentParagraphIndex + 1}`);
    return;
  }
  
  logMessage(`Applying changes to paragraph ${currentParagraphIndex + 1}`);
  
  await Word.run(async (context) => {
    try {
      // Get the specific paragraph using its document index if available, otherwise use currentParagraphIndex
      const docIndex = cacheEntry.documentIndex ?? currentParagraphIndex; // Prefer docIndex if cached
       const paras = context.document.body.paragraphs;
       paras.load("items");
       await context.sync();

      if (docIndex >= paras.items.length) {
         // This might happen if the document changed significantly since refresh
         logMessage(`Error: Paragraph document index ${docIndex} is out of bounds (total: ${paras.items.length}). Cannot apply.`);
         // Optionally try falling back to currentParagraphIndex if different? For now, just error out.
         return; 
      }

      const paragraph = paras.items[docIndex];
      
      // Replace the text
      paragraph.insertText(reimaginedText, "Replace");
      await context.sync();
      logMessage("Paragraph text replaced.");

       // --- Reapply Leading Formatting if necessary --- 
      if (leadingFormatText && leadingFormat) {
          logMessage(`Attempting to reapply leading format for "${leadingFormatText}"`);
          const searchResults = paragraph.search(leadingFormatText, { matchCase: true }); 
          searchResults.load("items");
          const paragraphRange = paragraph.getRange("Start"); 
          paragraphRange.load("start");
          await context.sync();

          if (searchResults.items.length > 0) {
              const leadingRange = searchResults.items[0];
              leadingRange.load("start"); 
              await context.sync();

              if (leadingRange.start === paragraphRange.start) {
                  logMessage("Found leading text at paragraph start. Applying format...");
                  // **** Apply underline ****
                  if (leadingFormat.bold !== null) leadingRange.font.bold = leadingFormat.bold;
                  if (leadingFormat.italic !== null) leadingRange.font.italic = leadingFormat.italic;
                  if (leadingFormat.size !== null) leadingRange.font.size = leadingFormat.size;
                  if (leadingFormat.underline !== null) leadingRange.font.underline = leadingFormat.underline; // Add underline
                  await context.sync();
                  logMessage("Leading formatting reapplied.");

                  // --- Explicitly apply modal formatting to the rest --- 
                  try {
                    const defaultFont = paragraphFormatting[currentParagraphIndex]?.defaultFont;
                    if (defaultFont) {
                        logMessage(`Applying modal format ${JSON.stringify(defaultFont)} to rest of paragraph.`);
                        const restRange = leadingRange.getRange("After").expandTo(paragraph.getRange("End"));
                        restRange.load("text"); 
                        await context.sync();
                        logMessage(`Applying modal to range: "${restRange.text.substring(0,50)}..."`);

                        // **** Reset underline ****
                        if (defaultFont.bold !== null) restRange.font.bold = defaultFont.bold;
                        if (defaultFont.italic !== null) restRange.font.italic = defaultFont.italic;
                        if (defaultFont.size !== null) restRange.font.size = defaultFont.size;
                        if (defaultFont.underline !== null) restRange.font.underline = defaultFont.underline; // Add underline reset
                        await context.sync();
                        logMessage("Modal formatting applied to rest of paragraph.");
                    } else {
                        logMessage("Warning: Could not find default/modal font info to apply to rest of paragraph.");
                    }
                  } catch (resetError) {
                     logMessage(`Error applying modal format to rest of paragraph: ${resetError}`);
                  }
                  // --- End applying modal format --- 

              } else {
                   logMessage(`Warning: Found "${leadingFormatText}" but not at the exact start of the paragraph. Cannot reapply leading format.`);
              }
          } else {
              logMessage(`Warning: Could not find leading text "${leadingFormatText}" after insertion to reapply formatting.`);
          }
      }
      // --- End Reapply Leading Formatting ---

      // --- Reapply Link if necessary --- 
      const linkInfo = cacheEntry.reimagined?.link;
      if (linkInfo && linkInfo.text && linkInfo.address) {
          logMessage(`Attempting to reapply link for "${linkInfo.text}"`);
          const linkSearchResults = paragraph.search(linkInfo.text, { matchCase: true }); 
          linkSearchResults.load("items");
          await context.sync();

          if (linkSearchResults.items.length > 0) {
              const linkRange = linkSearchResults.items[0];
              linkRange.hyperlink = linkInfo.address;
              await context.sync();
              logMessage(`Successfully reapplied hyperlink to "${linkInfo.text}".`);
          } else {
              logMessage(`Warning: Could not find the exact text "${linkInfo.text}" in paragraph ${currentParagraphIndex + 1} after insertion to reapply the link.`);
          }
      }
      // ---- End Link Reapplication ----

      // --- Apply other formatting --- 
      // TODO: Review applyFormatting logic - does it handle the potentially modified text correctly?
      // Only apply if the LLM part wasn't reverted?
      if (cacheEntry.reimagined?.formattingMatches?.length > 0 && !cacheEntry.reimagined?.revertedLLMPart) {
        logMessage("Applying other formatting matches...");
        await applyFormatting(context, paragraph, cacheEntry.reimagined.formattingMatches);
      } else if (cacheEntry.reimagined?.revertedLLMPart) {
         logMessage("Skipping other formatting application as original LLM part was used.");
      }
      
      // --- Update Cache and UI --- 
      paragraphCache[currentParagraphIndex].source = reimaginedText;
      paragraphCache[currentParagraphIndex].reimagined = null; 
      paragraphCache[currentParagraphIndex].reimagineState = false;
      selectedParagraphs.delete(currentParagraphIndex);
      document.getElementById("modifiedText").value = "";
      document.getElementById("checkboxReimagine").checked = false;
      logMessage(`Successfully applied changes to paragraph ${currentParagraphIndex + 1}`);
      updateUIForSelectedParagraph(); 
      updateSelectedParagraphList(); 
      
    } catch (error) {
      logMessage(`Error in onApply for paragraph ${currentParagraphIndex + 1}: ${error.message}`);
       if (error instanceof OfficeExtension.Error) {
          logMessage("OfficeExtension Error Debug Info: " + JSON.stringify(error.debugInfo));
       }
       console.error("onApply Error:", error);
    }
  }).catch((error) => {
      logMessage(`Error in Word.run for onApply: ${error}`);
      console.error("Word.run onApply Error:", error);
  });
}

/**
 * Applies changes to all selected paragraphs
 */
async function onApplyAll() {
  if (paragraphCache.length === 0) {
    logMessage("No paragraphs available to apply changes.");
    return;
  }
  if (selectedParagraphs.size === 0) {
     logMessage("No paragraphs selected to apply changes.");
     return;
  }

  logMessage(`Starting Apply All for ${selectedParagraphs.size} paragraphs...`);
  
  // Store selected paragraphs in array since we'll be clearing the set during the process
  const selectedParaArray = Array.from(selectedParagraphs);
  let appliedCount = 0;
  let skippedCount = 0;

  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    // Load items once for efficiency if possible, though indices might shift if paragraphs are added/deleted
    // Consider loading within loop if document structure changes are likely during apply all
    paras.load("items"); 
    await context.sync();
    const totalDocParas = paras.items.length;

    for (let idx of selectedParaArray) {
        // Check cache validity
        if (idx < 0 || idx >= paragraphCache.length) {
             logMessage(`Skipping invalid cache index ${idx} during Apply All.`);
             skippedCount++;
             selectedParagraphs.delete(idx); // Clean up selection
             continue;
        }

        const cacheEntry = paragraphCache[idx];
        const reimaginedText = cacheEntry.reimagined?.text;
        const leadingFormatText = cacheEntry.reimagined?.leadingFormatText;
        const leadingFormat = cacheEntry.reimagined?.leadingFormat;

        if (!reimaginedText) {
            logMessage(`Skipping paragraph ${idx + 1}: No reimagined text found.`);
            skippedCount++;
            // Clear state even if skipped
             cacheEntry.reimagined = null;
             cacheEntry.reimagineState = false;
             selectedParagraphs.delete(idx);
            continue;
        }

        try {
          // Get the specific paragraph using its document index if available
           const docIndex = cacheEntry.documentIndex ?? idx; 
           if (docIndex >= totalDocParas) {
              logMessage(`Error: Paragraph document index ${docIndex} (for cache index ${idx}) is out of bounds (total: ${totalDocParas}). Cannot apply.`);
              skippedCount++;
               // Clear state
               cacheEntry.reimagined = null;
               cacheEntry.reimagineState = false;
               selectedParagraphs.delete(idx);
              continue; 
           }
           const paragraph = paras.items[docIndex];

          // Replace the text
          paragraph.insertText(reimaginedText, "Replace");
          // Syncing inside the loop can be slow but ensures each step completes
          await context.sync(); 

          // --- Reapply Leading Formatting --- 
          if (leadingFormatText && leadingFormat) {
              logMessage(`ApplyAll: Reapplying leading format for para ${idx + 1}`);
              const searchResults = paragraph.search(leadingFormatText, { matchCase: true }); 
              searchResults.load("items");
              const paragraphRange = paragraph.getRange("Start"); 
              paragraphRange.load("start");
              await context.sync();
              if (searchResults.items.length > 0) {
                  const leadingRange = searchResults.items[0];
                  leadingRange.load("start"); 
                  await context.sync();
                  if (leadingRange.start === paragraphRange.start) {
                      if (leadingFormat.bold !== null) leadingRange.font.bold = leadingFormat.bold;
                      if (leadingFormat.italic !== null) leadingRange.font.italic = leadingFormat.italic;
                      if (leadingFormat.size !== null) leadingRange.font.size = leadingFormat.size;
                      if (leadingFormat.underline !== null) leadingRange.font.underline = leadingFormat.underline; // Add underline
                      await context.sync();
                      logMessage(`ApplyAll: Leading format reapplied for para ${idx + 1}.`);

                      // --- Explicitly apply modal formatting to the rest --- 
                      try {
                        const defaultFont = paragraphFormatting[idx]?.defaultFont;
                        if (defaultFont) {
                            logMessage(`ApplyAll: Applying modal format ${JSON.stringify(defaultFont)} to rest of para ${idx + 1}.`);
                            const restRange = leadingRange.getRange("After").expandTo(paragraph.getRange("End"));
                            restRange.load("text");
                            await context.sync();
                            logMessage(`ApplyAll: Applying modal to range: "${restRange.text.substring(0,50)}..."`);
                            
                            // **** Reset underline ****
                            if (defaultFont.bold !== null) restRange.font.bold = defaultFont.bold;
                            if (defaultFont.italic !== null) restRange.font.italic = defaultFont.italic;
                            if (defaultFont.size !== null) restRange.font.size = defaultFont.size;
                            if (defaultFont.underline !== null) restRange.font.underline = defaultFont.underline; // Add underline reset
                            await context.sync();
                            logMessage(`ApplyAll: Modal formatting applied to rest of para ${idx + 1}.`);
                        } else {
                           logMessage(`ApplyAll Warning: Could not find default font info for para ${idx + 1}.`);
                        }
                      } catch (resetError) {
                         logMessage(`ApplyAll Error applying modal format to rest of para ${idx + 1}: ${resetError}`);
                      }
                      // --- End applying modal format --- 

                  } else { logMessage(`ApplyAll Warning: Found leading text for para ${idx + 1} but not at start.`); }
              } else { logMessage(`ApplyAll Warning: Could not find leading text for para ${idx + 1}.`); }
          }
          // --- End Reapply Leading Formatting ---

          // --- Reapply Link --- 
          const linkInfo = cacheEntry.reimagined?.link;
          if (linkInfo && linkInfo.text && linkInfo.address) {
             logMessage(`ApplyAll: Reapplying link for para ${idx + 1}`);
              const linkSearchResults = paragraph.search(linkInfo.text, { matchCase: true }); 
              linkSearchResults.load("items");
              await context.sync();
              if (linkSearchResults.items.length > 0) {
                  const linkRange = linkSearchResults.items[0];
                  linkRange.hyperlink = linkInfo.address;
                  await context.sync();
                  logMessage(`ApplyAll: Link reapplied for para ${idx + 1}.`);
              } else { logMessage(`ApplyAll Warning: Could not find link text for para ${idx + 1}.`); }
          }
          // --- End Link Reapplication ---

          // --- Apply other formatting --- 
           if (cacheEntry.reimagined?.formattingMatches?.length > 0 && !cacheEntry.reimagined?.revertedLLMPart) {
              logMessage(`ApplyAll: Applying other formatting matches for paragraph ${idx + 1}...`);
              await applyFormatting(context, paragraph, cacheEntry.reimagined.formattingMatches);
              await context.sync(); // Sync after formatting
           } else if (cacheEntry.reimagined?.revertedLLMPart) {
              logMessage(`ApplyAll: Skipping other formatting for paragraph ${idx + 1} as original LLM part was used.`);
           }
           
          // --- Update Cache and State --- 
          cacheEntry.source = reimaginedText;
          cacheEntry.reimagined = null;
          cacheEntry.reimagineState = false;
          selectedParagraphs.delete(idx); 
          logMessage(`Applied changes to paragraph ${idx + 1}`);
          appliedCount++;

        } catch (error) {
          logMessage(`Error processing paragraph ${idx + 1} during Apply All: ${error.message}`);
          if (error instanceof OfficeExtension.Error) {
             logMessage("OfficeExtension Error Debug Info: " + JSON.stringify(error.debugInfo));
          }
           console.error(`Apply All Error - Para ${idx + 1}:`, error);
          skippedCount++;
           // Attempt to clear state even on error
           cacheEntry.reimagined = null; 
           cacheEntry.reimagineState = false;
           selectedParagraphs.delete(idx);
        }
    } // End loop
    
    // Final sync after loop (though syncs happened inside)
    await context.sync();
    logMessage(`Apply All completed: ${appliedCount} paragraphs updated, ${skippedCount} skipped/errors.`);
    
    // Clear all selections (should be empty now, but belt-and-suspenders)
    selectedParagraphs.clear();
    
    // Explicitly clear UI elements for the *current* paragraph if it was processed
    document.getElementById("modifiedText").value = "";
    document.getElementById("checkboxReimagine").checked = false;
    
    // Update UI
    updateSelectedParagraphList(); // Show empty list
    updateUIForSelectedParagraph(); // Update display for current index
    
  }).catch((error) => {
      logMessage("Error in Word.run for onApplyAll: " + error);
       if (error instanceof OfficeExtension.Error) {
          logMessage("OfficeExtension Error Debug Info: " + JSON.stringify(error.debugInfo));
       }
       console.error("Word.run onApplyAll Error:", error);
      // Attempt to clear UI state even on top-level error
      selectedParagraphs.clear();
      updateSelectedParagraphList();
      updateUIForSelectedParagraph();
  });
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
          if (match.format.bold !== null) range.font.bold = match.format.bold;
          if (match.format.italic !== null) range.font.italic = match.format.italic;
          if (match.format.size !== null) range.font.size = match.format.size;
          if (match.format.underline !== null) range.font.underline = match.format.underline; // Add underline
          
          await context.sync();
          logMessage(`Applied formatting to: "${range.text.substring(0, 50)}..."`); // Log applied range text
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
              if (match.format.bold !== null) range.font.bold = match.format.bold;
              if (match.format.italic !== null) range.font.italic = match.format.italic;
              if (match.format.size !== null) range.font.size = match.format.size;
              if (match.format.underline !== null) range.font.underline = match.format.underline; // Add underline
              
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
        logMessage(`Error processing format match for "${match.newText}": ${matchError}`);
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
