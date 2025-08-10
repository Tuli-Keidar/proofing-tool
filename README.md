# Document Proofreader

A Python tool that uses Azure OpenAI to proofread Word documents with intelligent segmentation and detailed error reporting.

## Features

- **Document Structure Analysis**: Automatically detects headings and sections
- **Intelligent Segmentation**: Splits large documents based on token limits while preserving structure
- **Comprehensive Proofreading**: Identifies spelling, grammar, formatting, and consistency issues
- **Multiple Output Formats**: Generates Markdown, DOCX, and JSON reports
- **Cost Tracking**: Estimates and tracks OpenAI API usage costs
- **Modular Architecture**: Clean, maintainable code structure

## Project Structure

```
ProofingTool/
├── .env                    # Environment variables (API keys, etc.)
├── requirements.txt        # Python dependencies
├── config.py              # Configuration management
├── utils.py               # Utility functions (cost calculation, etc.)
├── proofreader.py         # Main DocumentProofreader class
├── main.py                # Entry point and CLI interface
├── 25_03_22_proofer.ipynb # Original Jupyter notebook (legacy)
└── README.md              # This file
```

## Setup

### 1. Install Dependencies

```bash
pip install --break-system-packages python-docx tiktoken openai python-dotenv
```

Or using requirements.txt:
```bash
pip install --break-system-packages -r requirements.txt
```

### 2. Configure Environment Variables

Create a `.env` file in the project directory with your Azure OpenAI credentials:

```env
AZURE_OPENAI_ENDPOINT=https://your-resource.cognitiveservices.azure.com/openai/deployments/your-deployment/chat/completions?api-version=2024-10-21
AZURE_OPENAI_API_KEY=your-api-key-here
AZURE_API_VERSION=2024-10-21
MODEL_NAME=gpt-4o-2
```

## Usage

### Command Line Interface

```bash
# Show current configuration
python3 main.py --config

# Create and proofread a test document
python3 main.py --test

# Proofread a specific document
python3 main.py --input path/to/document.docx

# Specify output directory
python3 main.py --input document.docx --output results/

# Proofread the default document (to_proof.docx)
python3 main.py
```

### Programmatic Usage

```python
from config import Config
from proofreader import DocumentProofreader

# Initialize
proofreader = DocumentProofreader()

# Proofread a document
results = proofread_document("path/to/document.docx", "output/directory")

# Access results
print(f"Total cost: ${results['costs']['total_cost']:.4f}")
```

## Configuration

The `config.py` file contains all configuration settings:

- **API Settings**: Loaded from `.env` file
- **Document Paths**: Input and output locations
- **Segmentation**: Token limits per segment
- **Pricing**: Cost calculation parameters
- **Prompts**: System and user prompts for OpenAI

You can modify these settings by either:
1. Editing the `.env` file (recommended for API credentials)
2. Modifying `config.py` directly (for other settings)

## Output Files

The tool generates several output files:

- `complete_proofreading_report.md` - Combined Markdown report
- `complete_proofreading_report.docx` - Combined Word document
- `complete_proofreading_report.json` - Machine-readable results
- `proofreading_segment_X.md` - Individual segment reports
- `proofreading_segment_X.json` - Individual segment data

## Cost Tracking

The tool automatically calculates and reports:
- Input token count
- Output token count
- Estimated API costs based on current Azure OpenAI pricing
- Cost breakdown by segment

## Error Handling

- **Missing Documents**: Creates test documents when input files aren't found
- **API Errors**: Graceful error handling with detailed error messages
- **Configuration Issues**: Validates environment variables on startup
- **Large Documents**: Automatic segmentation for documents exceeding token limits

## Troubleshooting

### Common Issues

1. **Import Errors**: Make sure all dependencies are installed with `--break-system-packages` flag
2. **API Errors**: Check your `.env` file contains correct Azure OpenAI credentials
3. **Permission Errors**: Ensure write permissions for output directory

### Jupyter Notebook Issues

If the Jupyter notebook isn't working:
- Use the Python scripts instead (`python3 main.py`)
- The modular structure provides better error handling and debugging

## Development

### Adding New Features

1. **Configuration**: Add new settings to `config.py`
2. **Utilities**: Add helper functions to `utils.py`
3. **Core Logic**: Extend `DocumentProofreader` class in `proofreader.py`
4. **CLI**: Add new command-line options to `main.py`

### Testing

```bash
# Test with a generated document
python3 main.py --test

# Test configuration
python3 main.py --config

# Test with your own document
python3 main.py --input your_document.docx
```

## License

This tool is for document proofreading purposes. Please ensure you comply with Azure OpenAI terms of service and usage policies.
