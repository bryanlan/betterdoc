const express = require("express");
const { spawn } = require("child_process");
const https = require("https");
const fs = require("fs");
const path = require("path");
const cors = require("cors"); // Import the CORS middleware

const app = express();
app.use(express.json());
app.use(cors()); // Enable CORS for all routes

app.post("/ollama", async (req, res) => {
  try {
    const { paragraphText, userPrompt, model } = req.body;
    
    // Use the provided model or fall back to a default
    const modelToUse = model || "qwq:latest";
    
    console.log(`Using model: ${modelToUse}`);
    console.log(`Processing paragraph: ${paragraphText.substring(0, 50)}...`);
    
    const response = await fetch('http://localhost:11434/api/generate', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: modelToUse,
        prompt: userPrompt,
        system: paragraphText,
        stream: false
      }),
    });

    if (!response.ok) {
      console.error("Error from LLM server:", response.statusText);
      return res.status(500).json({ error: 'Failed to process request' });
    }

    const data = await response.json();
    console.log("LLM server response:", data);
    
    // Check if data.response exists before trying to use replace
    let cleanedOutput = "";
    if (data && data.response) {
      // Remove <think> sections and trim the result
      cleanedOutput = data.response.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
    } else if (data && data.result) {
      // Some models might return result instead of response
      cleanedOutput = data.result;
    } else {
      // Fallback if neither exists
      cleanedOutput = "No valid response from model";
    }
    
    res.json({ result: cleanedOutput });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to process request' });
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
  console.log(`Ollama server listening on https://localhost:${PORT}`);
});
