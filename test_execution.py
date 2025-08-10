#!/usr/bin/env python3

# Simple test to verify the environment
print("Testing Python execution...")

try:
    import sys
    print(f"Python version: {sys.version}")
    print(f"Python executable: {sys.executable}")
    
    # Test importing the packages
    print("\nTesting package imports...")
    
    import os
    print("✓ os imported")
    
    import json
    print("✓ json imported")
    
    from dotenv import load_dotenv
    print("✓ python-dotenv imported")
    
    from docx import Document
    print("✓ python-docx imported")
    
    import tiktoken
    print("✓ tiktoken imported")
    
    from openai import AzureOpenAI
    print("✓ openai imported")
    
    print("\nAll packages imported successfully!")
    
    # Test .env loading
    load_dotenv()
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    if endpoint:
        print(f"✓ Environment variables loaded (endpoint starts with: {endpoint[:30]}...)")
    else:
        print("✗ Environment variables not found")
        
except ImportError as e:
    print(f"✗ Import error: {e}")
except Exception as e:
    print(f"✗ Error: {e}")