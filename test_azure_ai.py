#!/usr/bin/env python3
"""
Simple test script for Azure AI Foundry API
Based on the CoreOS API documentation - Updated with working authentication
"""

import msal
import requests
import json
import os

# Import configuration from confidential.py
try:
    from confidential import CLIENT_ID, TENANT_ID, AUTHORITY, API_BASE, API_VERSION, DEFAULT_MODEL, OAUTH_SCOPE
except ImportError:
    print("ERROR: confidential.py not found!")
    print("Please copy confidential_sample.py to confidential.py and update with your Azure configuration.")
    exit(1)

def acquire_token_interactive():
    """Interactive Azure login using MSAL"""
    print("\n--- Interactive Azure Login (MSAL) ---")
    
    # Use specific scopes to avoid admin consent requirements
    scope = [OAUTH_SCOPE]
    
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    # Try silent authentication first
    accounts = app.get_accounts()
    result = None
    
    if accounts:
        print("Found cached account, trying silent authentication...")
        result = app.acquire_token_silent(scope, account=accounts[0])
    
    if not result:
        print("Silent authentication failed, starting interactive authentication...")
        print("A browser window will open for authentication...")
        result = app.acquire_token_interactive(scopes=scope)
    
    if "access_token" in result:
        print("✅ Login successful. Token acquired.")
        return result["access_token"]
    else:
        print(f"❌ Failed to acquire token: {result.get('error_description', result)}")
        return None

def main():
    print("Testing Azure AI Foundry API access...")
    print("=" * 50)
    
    # Step 1: Authentication
    print("Step 1: Authenticating with Azure AD...")
    
    # Check for environment token first
    bearer_token = os.environ.get("APIM_BEARER_TOKEN")
    if bearer_token:
        print("✅ Using bearer token from environment variable APIM_BEARER_TOKEN")
    else:
        bearer_token = acquire_token_interactive()
        if not bearer_token:
            return False
    
    # Step 2: Test API call
    print("\nStep 2: Testing API call...")
    
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json"
    }
    
    # Simple test prompt - using the configured default model
    test_data = {
        "model": DEFAULT_MODEL,
        "messages": [
            {
                "role": "user", 
                "content": "Rewrite this text to be more professional: 'Hey, this is a test to see if the API works.'"
            }
        ],
        "max_tokens": 100,
        "temperature": 0.7
    }
    
    # Construct the full API URL
    api_url = f"{API_BASE}/chat/completions?api-version={API_VERSION}"
    
    print(f"Making request to: {api_url}")
    print(f"Request data: {json.dumps(test_data, indent=2)}")
    
    try:
        response = requests.post(
            api_url,
            headers=headers,
            json=test_data,
            timeout=30
        )
        
        print(f"\nResponse status code: {response.status_code}")
        
        if response.status_code == 200:
            print("✅ API call successful!")
            result_data = response.json()
            
            # Print response details
            print("\n--- Chat Completion Response ---")
            if "choices" in result_data and len(result_data["choices"]) > 0:
                for choice in result_data.get("choices", []):
                    msg = choice.get("message", {})
                    if "content" in msg:
                        print(f"\n🤖 AI Response: {msg['content']}")
                    if "reasoning_content" in msg:
                        print(f"🧠 Reasoning: {msg['reasoning_content']}")
            
            # Print usage info if available
            if "usage" in result_data:
                usage = result_data["usage"]
                print(f"\n📊 Usage: {usage}")
            
            if "model" in result_data:
                print(f"🔧 Model: {result_data['model']}")
                
        elif response.status_code == 401:
            print("❌ Authentication error (401)")
            print("This means your token is invalid or you don't have permission to access this API")
            print("Response:", response.text)
            
        elif response.status_code == 403:
            print("❌ Forbidden error (403)")
            print("This means you don't have permission to access this specific resource")
            print("Response:", response.text)
            
        elif response.status_code == 404:
            print("❌ Not Found error (404)")
            print("This means the API endpoint doesn't exist or the resource is not available")
            print("Response:", response.text)
            
        else:
            print(f"❌ API call failed with status {response.status_code}")
            print("Response:", response.text)
            
        return response.status_code == 200
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Network error: {e}")
        return False
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

if __name__ == "__main__":
    print("Azure AI Foundry API Test Script")
    print("This will test your access to the CoreOS AI Pioneers API")
    print()
    
    success = main()
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 Test completed successfully!")
        print("The Azure AI Foundry API is working for your account.")
        print("\nNext steps:")
        print("1. Your authentication is working correctly")
        print("2. The API endpoint is accessible")
        print("3. You can now use this in your Office add-in server")
    else:
        print("❌ Test failed!")
        print("\nTroubleshooting:")
        print("1. Make sure you have access to the CoreOS AI Pioneers subscription")
        print("2. Check if your Microsoft work account has the right permissions")
        print("3. Verify you're connected to the internet")
        print("4. Try setting APIM_BEARER_TOKEN environment variable if you have a token") 