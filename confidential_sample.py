# Azure AI Foundry Configuration - SAMPLE FILE
# 
# INSTRUCTIONS:
# 1. Copy this file to 'confidential.py'
# 2. Replace all the dummy values below with your actual Azure configuration
# 3. Add 'confidential.py' to your .gitignore file
# 4. Never commit the real confidential.py file to version control

# Azure AD Application Configuration
# Get these values from your Azure AD app registration
CLIENT_ID = "your-client-id-here"
TENANT_ID = "your-tenant-id-here"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Azure AI Foundry API Configuration
# Get these values from your Azure AI Foundry project
API_BASE = "https://your-api-endpoint.azure-api.net/api/projects/your-project/models"
API_VERSION = "2024-05-01-preview"  # or your preferred API version
DEFAULT_MODEL = "gpt-4.1"  # or your preferred model

# OAuth Scope
# This should match the scope defined in your Azure AD app registration
OAUTH_SCOPE = f"{CLIENT_ID}/access_api" 