#!/usr/bin/env python3
"""
Configuration management for the Document Proofreader
"""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class Config:
    """Configuration class to manage all settings"""
    
    # Azure OpenAI API settings (loaded from .env file)
    AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
    AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
    AZURE_API_VERSION = os.getenv("AZURE_API_VERSION")
    MODEL_NAME = os.getenv("MODEL_NAME")
    
    # Document parameters
    INPUT_DOCUMENT_PATH = "to_proof.docx"
    OUTPUT_DIRECTORY = "proofreading_results"
    
    # Segmentation parameters
    MAX_TOKENS_PER_SEGMENT = 1500
    
    # Pricing parameters per million tokens
    COST_TEXT_INPUT = 7.9158  # $ per million tokens
    COST_CACHED_INPUT = 3.9579  # $ per million tokens
    COST_OUTPUT = 31.663105  # $ per million tokens
    
    # Proofreading instruction templates
    SYSTEM_PROMPT = """You are an expert proofreader and editor with excellent attention to detail working on documents written using Australian spelling. 
Your task is to identify and correct errors in the document, including:
1. Spelling mistakes and typos
2. Grammar and punctuation errors
3. Awkward or unclear phrasing
4. Inconsistencies in terminology, style, or formatting
5. Potential factual errors or logical inconsistencies

Provide specific corrections for each issue you identify, and explain your reasoning when necessary."""

    USER_PROMPT_TEMPLATE = """
# Document Content to Proofread
{content}

## Instructions
Please carefully proofread the above document content. For each issue you find:
1. Quote the problematic text
2. Explain the issue
3. Provide a corrected version

Format your response as follows:

----
## Identified Issues

### Issue 1: [Issue type - spelling, grammar, clarity, etc.]
**Original**: "[problematic text]"
**Issue**: [explanation of the problem]
**Suggestion**: "[corrected text]"

### Issue 2: [Issue type]
...
If no issues are found in a section, please state: "No issues found in this section."
"""

    @classmethod
    def validate(cls):
        """Validate that all required configuration is present"""
        missing = []
        
        if not cls.AZURE_OPENAI_ENDPOINT:
            missing.append("AZURE_OPENAI_ENDPOINT")
        if not cls.AZURE_OPENAI_API_KEY:
            missing.append("AZURE_OPENAI_API_KEY")
        if not cls.AZURE_API_VERSION:
            missing.append("AZURE_API_VERSION")
        if not cls.MODEL_NAME:
            missing.append("MODEL_NAME")
            
        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")
        
        return True
    
    @classmethod
    def display(cls):
        """Display current configuration (safely)"""
        print("=== Configuration ===")
        print(f"Endpoint: {cls.AZURE_OPENAI_ENDPOINT[:50] + '...' if cls.AZURE_OPENAI_ENDPOINT else 'None'}")
        print(f"API Key: {'*' * 20 if cls.AZURE_OPENAI_API_KEY else 'None'}")
        print(f"API Version: {cls.AZURE_API_VERSION}")
        print(f"Model Name: {cls.MODEL_NAME}")
        print(f"Input Document: {cls.INPUT_DOCUMENT_PATH}")
        print(f"Output Directory: {cls.OUTPUT_DIRECTORY}")
        print(f"Max Tokens per Segment: {cls.MAX_TOKENS_PER_SEGMENT}")
        print("=" * 20)