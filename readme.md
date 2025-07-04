# BetterDoc - AI-Powered Word Document Enhancement Tool

BetterDoc is a Microsoft Office Word add-in that leverages AI technology to intelligently reimagine and improve your document content while preserving formatting, hyperlinks, and document structure.

## Features

### Core Functionality
- **Paragraph-by-Paragraph Processing**: Navigate through your document paragraph by paragraph with intelligent skipping of short paragraphs, questions, and empty content
- **AI Content Enhancement**: Use Azure AI Foundry or local Ollama models to rewrite content with customizable prompts
- **Formatting Preservation**: Automatically detects and preserves text formatting (bold, italic, underline, font sizes)
- **Hyperlink Protection**: Detects hyperlinks and ensures they remain intact during content processing
- **Flexible Model Support**: Choose between cloud-based Azure models or local Ollama models

### Smart Processing
- **Selective Processing**: Only processes paragraphs with 7+ words, skips questions and empty paragraphs
- **Word Count Control**: Prevents AI from increasing word count beyond original content
- **Retry Logic**: Automatically retries if hyperlink text is lost during processing
- **Batch Operations**: Select and process multiple paragraphs simultaneously

### User Interface
- **Visual Navigation**: Up/Down arrow buttons to move through paragraphs
- **Real-time Preview**: View original and modified text side-by-side
- **Document Highlighting**: Current paragraph is selected in the document for easy reference
- **Model Selection**: Choose from Azure AI Foundry or local Ollama models

## Architecture

The project consists of two main components:

### 1. Office Add-in (`betterdoc-office-addin/`)
- **Frontend**: HTML/CSS/JavaScript task pane interface
- **Backend Integration**: Communicates with AI services via proxy server
- **Office.js Integration**: Uses Microsoft Office JavaScript API for document manipulation

### 2. Multi-Model Proxy Server (`ollama-service/`)
- **Azure AI Foundry Integration**: Secure connection to Azure AI Foundry using Python MSAL authentication
- **Ollama Integration**: Optional local model support via Ollama API
- **Smart Routing**: Automatically routes requests to Azure or Ollama based on model selection
- **CORS Handling**: Enables cross-origin requests from the Office add-in

## Quick Start

### Prerequisites
- Microsoft Word (Office 365 or Office 2019+)
- Node.js (18+)
- Python 3.7+ with `msal` and `requests` packages
- Azure AI Foundry access (for cloud models)
- Ollama (optional, for local models)

### 1. Clone and Setup
```bash
git clone <repository-url>
cd betterdoc2
cp confidential_sample.py confidential.py
```

### 2. Configure Azure AI Foundry
Edit `confidential.py` with your Azure details:
```python
CLIENT_ID = "your-azure-ad-client-id"
TENANT_ID = "your-azure-ad-tenant-id"
API_BASE = "https://your-endpoint.azure-api.net/api/projects/your-project/models"
API_VERSION = "2024-05-01-preview"
DEFAULT_MODEL = "gpt-4.1"
```

### 3. Install Dependencies
```bash
# Python dependencies
pip install msal requests

# Node.js dependencies
cd ollama-service
npm install
```

### 4. Start the Server
```bash
cd ollama-service
node server.js
```

### 5. Install Office Add-in
```bash
cd betterdoc-office-addin/BetterdocAddin
npm install
npm run build
```

**Install Development Certificates** (Required for HTTPS):
```bash
# Navigate to the betterdoc directory
cd betterdoc2

# Install the Office Add-in development certificates
npx office-addin-dev-certs install
```
This will:
- Generate SSL certificates in `%USERPROFILE%\.office-addin-dev-certs\`
- Install the CA certificate for trusted access to `https://localhost`
- Create the required `localhost.crt` and `localhost.key` files

Then sideload in Word:
- Insert > My Add-ins > Upload My Add-in
- Select `manifest.xml`

## Configuration Details

### Azure AI Foundry Setup

1. **Azure AD Application Registration**:
   - Create an app registration in Azure AD
   - Note the `CLIENT_ID` and `TENANT_ID`
   - Configure API permissions for Azure AI Foundry

2. **Azure AI Foundry Project**:
   - Create an Azure AI Foundry project
   - Get the API endpoint URL
   - Ensure your account has access to the project

3. **Authentication**:
   - Uses Python MSAL for interactive authentication
   - Tokens are cached automatically for future use
   - One-time browser authentication setup

### Local Models (Optional)

1. **Install Ollama**:
   ```bash
   # Visit https://ollama.ai for installation instructions
   ```

2. **Pull Models**:
   ```bash
   ollama pull gemma3:12b
   ollama pull qwq:latest
   ollama pull deepseek-r1:14b
   ```

3. **Start Ollama**:
   ```bash
   ollama serve
   ```

## Available Models

### Azure Models
- **gpt-4.1 (Azure)** - Powered by Azure AI Foundry (no local setup required)

### Local Models (Requires Ollama)
- **gemma3:12b** - Google's Gemma model (12B parameters)
- **gemma3:27b** - Google's Gemma model (27B parameters)  
- **qwq:latest** - QwQ reasoning model
- **deepseek-r1:14b** - DeepSeek R1 model (14B parameters)
- **deepseek-r1:32b** - DeepSeek R1 model (32B parameters)

## Usage Guide

### Basic Workflow

1. **Open the Add-in**: In Word, find "Betterdoc" in your add-ins and open the task pane

2. **First-Time Setup**: 
   - Select an Azure model for immediate use
   - For local models, ensure Ollama is installed and running

3. **Navigate Document**: Use the ↑/↓ buttons to move through paragraphs
   - Automatically skips short paragraphs (<7 words), questions, and empty content
   - Current paragraph is highlighted in the document

4. **Select Content**: Check the "Reimagine" checkbox for paragraphs you want to process
   - Use "Select All" to mark all valid paragraphs
   - Use "Unselect All" to clear selections

5. **Configure Settings**:
   - **Model Selection**: Choose between Azure AI Foundry or local Ollama models
   - **Prompt Selection**: Choose from predefined prompts or create custom ones
     - "Make More Professional": Enhances formal tone
     - "Reduce Word Count": Makes content more concise
     - "Custom Prompt": Write your own instructions

6. **Process Content**: Click "Reimagine Now" to send selected paragraphs to the AI
   - Azure models work immediately (after one-time authentication)
   - Local models require Ollama to be running

7. **Review and Apply**:
   - Review AI-generated content in the "Modified/LLM Text" area
   - Edit the content if needed
   - Click "Apply" for current paragraph or "Apply All" for batch processing

### Error Handling

The system provides helpful error messages:
- **Ollama not installed**: "Sorry, you need to install and run Ollama to use local models..."
- **Model not available**: "Sorry, the model 'gemma3:12b' is not installed in Ollama..."
- **Azure authentication needed**: Automatic browser-based authentication flow

## Testing

### Test Azure AI Foundry Connection
```bash
python test_azure_ai.py
```

### Test Token Authentication
```bash
cd ollama-service
python token_fetch.py --scope "your-client-id/access_api"
```

### Test Server
```bash
cd ollama-service
node server.js
```

## Security

- **Confidential Configuration**: All sensitive data is in `confidential.py` (gitignored)
- **Token Caching**: Azure tokens are cached securely for reuse
- **No Hardcoded Secrets**: Use `confidential_sample.py` as a template
- **Automatic Authentication**: Browser-based OAuth flow for Azure

## Troubleshooting

### Azure AI Foundry Issues
- Ensure `confidential.py` is configured correctly
- Check Azure AD app permissions
- Verify Azure AI Foundry project access
- Complete browser authentication when prompted

### Local Model Issues
- Install Ollama from https://ollama.ai
- Run `ollama serve` to start the service
- Pull required models with `ollama pull <model-name>`
- Check that Ollama is running on port 11434

### Office Add-in Issues
- Ensure manifest.xml is properly sideloaded
- Check browser console for JavaScript errors
- Verify server is running on https://localhost:8000
- Use "Refresh Source" button to reload document content

## Development

### Project Structure
```
betterdoc2/
├── betterdoc-office-addin/          # Office add-in
│   └── BetterdocAddin/
│       ├── manifest.xml             # Office add-in manifest
│       └── src/taskpane/
│           ├── taskpane.html        # Main UI
│           ├── taskpane.js          # Core logic
│           └── styles.css           # Styling
├── ollama-service/                  # Multi-model proxy server
│   ├── server.js                    # Main server (Node.js)
│   └── token_fetch.py               # Azure authentication (Python)
├── confidential.py                  # Your Azure configuration (gitignored)
├── confidential_sample.py           # Configuration template
├── test_azure_ai.py                 # Test Azure connection
└── README.md                        # This file
```

### Adding New Models

1. **Azure Models**: Update `isAzureModel()` function in `server.js`
2. **Local Models**: Add to `window.availableModels` array in `taskpane.html`
3. **Custom Prompts**: Modify the prompts dropdown in `taskpane.html`

## License

[Add your license information here]
