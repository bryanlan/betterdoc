#!/usr/bin/env python3
"""
Updated test script for Azure AI Foundry API
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
    
    bearer_token = acquire_token_interactive()
    if not bearer_token:
        return False
    
    # Step 2: Test API call
    print("\nStep 2: Testing API call...")
    
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json"
    }
    
    # Simple test prompt
    test_data = {
        "model": DEFAULT_MODEL,
        "messages": [
            {
                "role": "user", 
                "content": "Hello, can you help me test this API connection?"
            }
        ],
        "max_tokens": 100,
        "temperature": 0.7
    }
    
    # Construct the full API URL
    api_url = f"{API_BASE}/chat/completions?api-version={API_VERSION}"
    
    print(f"Making request to: {api_url}")
    
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
            
            if "choices" in result_data and len(result_data["choices"]) > 0:
                content = result_data["choices"][0]["message"]["content"]
                print(f"\n🤖 AI Response: {content}")
                
        else:
            print(f"❌ API call failed with status {response.status_code}")
            print("Response:", response.text)
            
        return response.status_code == 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    
    if success:
        print("\n🎉 Test completed successfully!")
    else:
        print("\n❌ Test failed!") 