const express = require("express");
const { spawn } = require("child_process");
const https = require("https");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const { PythonShell } = require("python-shell");

const app = express();
app.use(express.json());
app.use(cors());

// Import Azure configuration from confidential.py
let CLIENT_ID, AZURE_SCOPE, AZURE_API_BASE, AZURE_API_VERSION, AZURE_DEFAULT_MODEL;

try {
  // Use Python to read the confidential.py file
  const configScript = `
import sys
import json
try:
    from confidential import CLIENT_ID, API_BASE, API_VERSION, DEFAULT_MODEL, OAUTH_SCOPE
    config = {
        "CLIENT_ID": CLIENT_ID,
        "API_BASE": API_BASE,
        "API_VERSION": API_VERSION,
        "DEFAULT_MODEL": DEFAULT_MODEL,
        "OAUTH_SCOPE": OAUTH_SCOPE
    }
    print(json.dumps(config))
except ImportError:
    print(json.dumps({"error": "confidential.py not found"}))
    sys.exit(1)
`;

  const configPath = path.join(__dirname, 'get_config.py');
  fs.writeFileSync(configPath, configScript);
  
  const result = require('child_process').execSync('python get_config.py', { 
    cwd: __dirname,
    encoding: 'utf8' 
  });
  
  const config = JSON.parse(result.trim());
  if (config.error) {
    throw new Error(config.error);
  }
  
  CLIENT_ID = config.CLIENT_ID;
  AZURE_SCOPE = config.OAUTH_SCOPE;
  AZURE_API_BASE = config.API_BASE;
  AZURE_API_VERSION = config.API_VERSION;
  AZURE_DEFAULT_MODEL = config.DEFAULT_MODEL;
  
  // Clean up temporary file
  fs.unlinkSync(configPath);
  
} catch (error) {
  console.error("❌ Failed to load configuration from confidential.py");
  console.error("Please copy confidential_sample.py to confidential.py and update with your Azure configuration.");
  process.exit(1);
}

// Token cache file path
const TOKEN_CACHE_FILE = path.join(__dirname, 'azure_token_cache.json');

let cachedAzureToken = null;
let tokenExpiry = null;
let authPromise = null;

// Load cached token from file on startup
function loadCachedToken() {
  try {
    if (fs.existsSync(TOKEN_CACHE_FILE)) {
      const cacheData = JSON.parse(fs.readFileSync(TOKEN_CACHE_FILE, 'utf8'));
      if (cacheData.token && cacheData.expiry) {
        const expiryDate = new Date(cacheData.expiry);
        if (expiryDate > new Date()) {
          cachedAzureToken = cacheData.token;
          tokenExpiry = expiryDate;
          console.log("✅ Loaded valid cached token from file");
          return true;
        } else {
          console.log("⚠️ Cached token has expired");
        }
      }
    }
  } catch (error) {
    console.log("⚠️ Could not load cached token:", error.message);
  }
  return false;
}

// Save token to file
function saveCachedToken(token, expiry) {
  try {
    const cacheData = {
      token: token,
      expiry: expiry.toISOString()
    };
    fs.writeFileSync(TOKEN_CACHE_FILE, JSON.stringify(cacheData, null, 2));
    console.log("✅ Token cached to file for future use");
  } catch (error) {
    console.log("⚠️ Could not save token to cache:", error.message);
  }
}

// Function to call Python script for token acquisition
async function callPythonTokenFetcher() {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(__dirname, 'token_fetch.py');
    const options = {
      pythonPath: 'python',
      args: ['--scope', AZURE_SCOPE]
    };

    console.log("🐍 Calling Python script for Azure authentication...");
    console.log("   (This will open a browser window for authentication)");

    PythonShell.run(scriptPath, options)
      .then(results => {
        if (results && results.length > 0) {
          try {
            const tokenInfo = JSON.parse(results[0]);
            if (tokenInfo.success) {
              resolve(tokenInfo);
            } else {
              reject(new Error(`Python auth failed: ${tokenInfo.error_description}`));
            }
          } catch (parseError) {
            reject(new Error(`Failed to parse Python response: ${parseError.message}`));
          }
        } else {
          reject(new Error('No response from Python script'));
        }
      })
      .catch(error => {
        reject(new Error(`Python script error: ${error.message}`));
      });
  });
}

// Function to get Azure AD token - Python bridge approach
async function getAzureAccessToken() {
  const now = new Date();
  
  // Return cached token if still valid (with 5 minute buffer)
  if (cachedAzureToken && tokenExpiry && now < new Date(tokenExpiry.getTime() - 5 * 60 * 1000)) {
    console.log("✅ Using cached Azure token");
    return cachedAzureToken;
  }

  // If we're already in the middle of auth, wait for it
  if (authPromise) {
    console.log("⏳ Authentication already in progress, waiting...");
    return await authPromise;
  }

  // Call Python script for authentication
  authPromise = (async () => {
    try {
      console.log("\n" + "=".repeat(80));
      console.log("🔐 AZURE AUTHENTICATION REQUIRED");
      console.log("=".repeat(80));
      console.log("Calling Python script for authentication...");
      console.log("   (A browser window will open automatically)");
      console.log("   (This is a one-time setup that will be cached for future use)");
      console.log("-".repeat(80));

      const tokenInfo = await callPythonTokenFetcher();
      
      cachedAzureToken = tokenInfo.access_token;
      if (tokenInfo.expires_on) {
        tokenExpiry = new Date(tokenInfo.expires_on * 1000); // Convert Unix timestamp to Date
      } else {
        // Default to 1 hour if no expiry provided
        tokenExpiry = new Date(Date.now() + 60 * 60 * 1000);
      }
      
      saveCachedToken(cachedAzureToken, tokenExpiry);
      
      console.log("\n" + "=".repeat(80));
      console.log("🎉 AUTHENTICATION SUCCESSFUL!");
      console.log("=".repeat(80));
      console.log("✅ Azure token acquired and cached for future use");
      console.log("✅ Server is now ready to process Azure AI requests");
      console.log("=".repeat(80) + "\n");
      
      return cachedAzureToken;
      
    } catch (error) {
      console.error('\n❌ Error acquiring Azure token:', error.message);
      console.log("\nTroubleshooting:");
      console.log("1. Make sure you have access to the CoreOS AI Pioneers subscription");
      console.log("2. Check that your Microsoft work account has the right permissions");
      console.log("3. Verify you're connected to the internet");
      console.log("4. Make sure Python and required packages (msal) are installed");
      console.log("5. Try restarting the server if authentication seems stuck\n");
      throw new Error('Failed to acquire Azure AD token. Please check your Azure access permissions.');
    } finally {
      authPromise = null;
    }
  })();

  return await authPromise;
}

// Load cached token on startup
loadCachedToken();

// Function to check if model is Azure-based
function isAzureModel(model) {
  return model && (
    model.toLowerCase().includes('azure') || 
    model === 'gpt-4.1' || 
    model === 'gpt-4.1 (Azure)'
  );
}

// Function to call Azure AI Foundry - Updated with working API structure
async function callAzureAI(userPrompt, paragraphText) {
  const token = await getAzureAccessToken();
  
  // If paragraphText is empty or undefined, treat userPrompt as the complete message
  let messages;
  if (!paragraphText || paragraphText.trim() === "") {
    messages = [
      {
        role: "user",
        content: userPrompt
      }
    ];
  } else {
    // Original format with separate system and user messages
    messages = [
      {
        role: "system",
        content: paragraphText
      },
      {
        role: "user", 
        content: userPrompt
      }
    ];
  }

  // Updated to use working API endpoint and structure
  const apiUrl = `${AZURE_API_BASE}/chat/completions?api-version=${AZURE_API_VERSION}`;
  
  const requestBody = {
    model: AZURE_DEFAULT_MODEL,
    messages: messages,
    max_tokens: 1000,
    temperature: 0.7
  };

  console.log(`Making Azure API request to: ${apiUrl}`);
  console.log(`Request body:`, JSON.stringify(requestBody, null, 2));

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error("Error from Azure AI server:", response.status, response.statusText, errorText);
    throw new Error(`Azure AI API error: ${response.status} - ${response.statusText}`);
  }

  const data = await response.json();
  console.log("Azure AI server response:", JSON.stringify(data, null, 2));
  
  // Extract response from Azure AI format
  if (data && data.choices && data.choices.length > 0 && data.choices[0].message) {
    return data.choices[0].message.content.trim();
  } else {
    throw new Error("Invalid response structure from Azure AI");
  }
}

// Function to check if Ollama is available
async function checkOllamaAvailable() {
  try {
    const response = await fetch('http://localhost:11434/api/version', {
      method: 'GET',
      timeout: 5000 // 5 second timeout
    });
    return response.ok;
  } catch (error) {
    return false;
  }
}

// Function to call local Ollama
async function callOllama(userPrompt, paragraphText, model) {
  // Check if Ollama is available first
  const ollamaAvailable = await checkOllamaAvailable();
  
  if (!ollamaAvailable) {
    console.log(`⚠️ Ollama not available for model: ${model}`);
    return "Sorry, you need to install and run Ollama to use local models. Please visit https://ollama.ai to download and install Ollama, then try again.";
  }

  try {
    const response = await fetch('http://localhost:11434/api/generate', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: model,
        prompt: userPrompt,
        system: paragraphText,
        stream: false
      }),
    });

    if (!response.ok) {
      if (response.status === 404) {
        return `Sorry, the model "${model}" is not installed in Ollama. Please run "ollama pull ${model}" to install it, or choose a different model.`;
      }
      throw new Error(`Ollama API error: ${response.status} - ${response.statusText}`);
    }

    const data = await response.json();
    console.log("Ollama server response:", data);
    
    if (data && data.response) {
      return data.response.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
    } else if (data && data.result) {
      return data.result;
    } else {
      throw new Error("Invalid response from Ollama");
    }
  } catch (error) {
    console.error(`Error calling Ollama for model ${model}:`, error.message);
    
    // Return user-friendly error messages
    if (error.message.includes('ECONNREFUSED') || error.message.includes('fetch failed')) {
      return "Sorry, Ollama is not running. Please start Ollama and try again.";
    } else if (error.message.includes('timeout')) {
      return "Sorry, Ollama is taking too long to respond. Please check if Ollama is running properly.";
    } else {
      return `Sorry, there was an error with the local model "${model}". Please try again or choose a different model.`;
    }
  }
}

app.post("/ollama", async (req, res) => {
  try {
    const { paragraphText, userPrompt, model } = req.body;
    
    // Use the provided model or fall back to Azure default
    const modelToUse = model || AZURE_DEFAULT_MODEL;
    
    console.log(`Using model: ${modelToUse}`);
    console.log(`Processing paragraph: ${paragraphText ? paragraphText.substring(0, 50) + '...' : 'N/A'}`);
    console.log(`User prompt: ${userPrompt ? userPrompt.substring(0, 100) + '...' : 'N/A'}`);
    
    let cleanedOutput = "";
    
    if (isAzureModel(modelToUse)) {
      console.log("Routing to Azure AI Foundry...");
      cleanedOutput = await callAzureAI(userPrompt, paragraphText);
    } else {
      console.log("Routing to local Ollama...");
      cleanedOutput = await callOllama(userPrompt, paragraphText, modelToUse);
    }
    
    res.json({ result: cleanedOutput });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: `Failed to process request: ${error.message}` });
  }
});

const PORT = 8000;

// Set up paths to your development certificate and key.
const certPath = path.join("C:", "Users", "blangl", ".office-addin-dev-certs", "localhost.crt");
const keyPath = path.join("C:", "Users", "blangl", ".office-addin-dev-certs", "localhost.key");

const options = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath)
};

https.createServer(options, app).listen(PORT, () => {
  console.log(`Multi-model server listening on https://localhost:${PORT}`);
  console.log("Supported models:");
  console.log(`- Azure model: "gpt-4.1 (Azure)" -> Azure AI Foundry (${AZURE_API_BASE})`);
  console.log("- Local models: gemma3, qwq, deepseek-r1 -> Ollama (optional - install from https://ollama.ai)");
  console.log(`- Default model: ${AZURE_DEFAULT_MODEL}`);
  console.log(`- Token cache: ${TOKEN_CACHE_FILE}`);
  console.log("\n✅ Azure models work without Ollama");
  console.log("⚠️  Local models require Ollama to be installed and running");
});
