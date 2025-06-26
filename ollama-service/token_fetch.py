#!/usr/bin/env python3
"""
Token fetcher for Azure AI Foundry API
Returns access token as JSON for Node.js consumption
"""

import msal
import json
import sys
import os
import argparse

# Import configuration from confidential.py
try:
    from confidential import CLIENT_ID, TENANT_ID, AUTHORITY
except ImportError:
    print("ERROR: confidential.py not found!")
    print("Please copy confidential_sample.py to confidential.py and update with your Azure configuration.")
    sys.exit(1)

def acquire_token_interactive(scope):
    """Interactive Azure login using MSAL"""
    
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    # Try silent authentication first
    accounts = app.get_accounts()
    result = None
    
    if accounts:
        # Try silent first
        result = app.acquire_token_silent(scope, account=accounts[0])
    
    if not result:
        # Need interactive authentication
        result = app.acquire_token_interactive(scopes=scope)
    
    return result

def main():
    parser = argparse.ArgumentParser(description='Fetch Azure access token')
    parser.add_argument('--scope', required=True, help='OAuth scope to request')
    args = parser.parse_args()
    
    scope = [args.scope]
    
    try:
        result = acquire_token_interactive(scope)
        
        if "access_token" in result:
            # Return token info as JSON
            token_info = {
                "access_token": result["access_token"],
                "expires_on": result.get("expires_on"),
                "success": True
            }
            print(json.dumps(token_info))
        else:
            # Return error info as JSON
            error_info = {
                "success": False,
                "error": result.get("error"),
                "error_description": result.get("error_description", "Unknown error")
            }
            print(json.dumps(error_info))
            sys.exit(1)
            
    except Exception as e:
        # Return exception info as JSON
        error_info = {
            "success": False,
            "error": "exception",
            "error_description": str(e)
        }
        print(json.dumps(error_info))
        sys.exit(1)

if __name__ == "__main__":
    main() 