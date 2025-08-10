#!/usr/bin/env python3

import os
import sys

# Add current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the notebook code
from test_execution import *  # Just to verify imports work

# Now run the actual proofreading code
import os
import json
import re
import math
from docx import Document
import tiktoken
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

print("=== Proofreading Tool Starting ===")

# Test environment setup
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_API_VERSION = os.getenv("AZURE_API_VERSION")
MODEL_NAME = os.getenv("MODEL_NAME")

print(f"Endpoint: {AZURE_OPENAI_ENDPOINT[:50] + '...' if AZURE_OPENAI_ENDPOINT else 'None'}")
print(f"API Key: {'*' * 20 if AZURE_OPENAI_API_KEY else 'None'}")
print(f"API Version: {AZURE_API_VERSION}")
print(f"Model Name: {MODEL_NAME}")

# Check if we can initialize the client
try:
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_API_VERSION,
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    print("✓ Azure OpenAI client initialized successfully")
except Exception as e:
    print(f"✗ Failed to initialize Azure OpenAI client: {e}")

# Check for input document
INPUT_DOCUMENT_PATH = "to_proof.docx"
if os.path.exists(INPUT_DOCUMENT_PATH):
    print(f"✓ Found input document: {INPUT_DOCUMENT_PATH}")
else:
    print(f"✗ Input document not found: {INPUT_DOCUMENT_PATH}")
    print("Available files:")
    for f in os.listdir('.'):
        if f.endswith('.docx'):
            print(f"  - {f}")

print("\n=== Environment Check Complete ===")