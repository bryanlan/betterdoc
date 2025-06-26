# Azure AI Foundry Setup Guide

## Initial Setup

1. **Copy the configuration template:**
   ```bash
   cp confidential_sample.py confidential.py
   ```

2. **Update confidential.py with your Azure details:**
   - `CLIENT_ID`: Your Azure AD application client ID
   - `TENANT_ID`: Your Azure AD tenant ID
   - `API_BASE`: Your Azure AI Foundry API base URL
   - `API_VERSION`: API version (usually "2024-05-01-preview")
   - `DEFAULT_MODEL`: Your preferred model (e.g., "gpt-4.1")

3. **Install Python dependencies:**
   ```bash
   pip install msal requests
   ```

4. **Install Node.js dependencies:**
   ```bash
   cd ollama-service
   npm install python-shell
   ```

## Configuration Details

### Azure AD Application
- Get `CLIENT_ID` and `TENANT_ID` from your Azure AD app registration
- Ensure the app has the correct API permissions for Azure AI Foundry

### Azure AI Foundry API
- Get `API_BASE` from your Azure AI Foundry project
- Format: `https://your-endpoint.azure-api.net/api/projects/your-project/models`
- Use the correct `API_VERSION` for your deployment

### OAuth Scope
- The scope is automatically generated as `{CLIENT_ID}/access_api`
- This should match your Azure AD app registration scopes

## Testing

1. **Test Python authentication:**
   ```bash
   python test_azure_ai.py
   ```

2. **Test the token fetcher:**
   ```bash
   cd ollama-service
   python token_fetch.py --scope "your-client-id/access_api"
   ```

3. **Start the server:**
   ```bash
   cd ollama-service
   node server.js
   ```

## Security Notes

- **Never commit `confidential.py` to version control**
- The file is already in `.gitignore`
- Token cache files are also excluded from git
- Keep your Azure credentials secure and rotate them regularly

## Troubleshooting

- If you get "confidential.py not found" errors, make sure you copied and configured the file
- Ensure Python can import the confidential module (it should be in the same directory)
- Check that your Azure AD app has the correct permissions
- Verify your API endpoint URL is correct for your Azure AI Foundry project 