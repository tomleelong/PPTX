# PowerPoint Creator for Claude Code

A Python tool for creating PowerPoint presentations programmatically using the `python-pptx` library. This project is designed to work seamlessly with Claude Code for automated presentation generation.

## Features

- **Programmatic PowerPoint Creation**: Generate presentations entirely through code
- **Multiple Slide Types**: Support for title slides, bullet point slides, text slides, and image slides
- **JSON-Based Outlines**: Create presentations from structured JSON files
- **Command Line Interface**: Easy-to-use CLI for quick presentation generation
- **Customizable Styling**: Modify colors, fonts, and layouts
- **Claude Code Integration**: Designed specifically for use with Claude Code workflows

## Installation

### Prerequisites

- Python 3.10 or higher
- Virtual environment (recommended)

### Setup

1. Clone or navigate to the project directory:
   ```bash
   cd /path/to/PPTX
   ```

2. Create and activate a virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Command Line Interface

#### Create a sample presentation:
```bash
python -m src.cli --sample my_presentation.pptx
```

#### Create from JSON outline:
```bash
python -m src.cli --json examples/sample_outline.json --output my_presentation.pptx
```

#### Generate sample JSON file:
```bash
python -m src.cli --sample-json my_outline.json
```

### Python API

```python
from src.pptx_creator import PPTXCreator

# Create a new presentation
creator = PPTXCreator()

# Add title slide
creator.add_title_slide(
    title="My Presentation",
    subtitle="Created with Python"
)

# Add content slide with bullet points
creator.add_content_slide(
    title="Key Points",
    content=[
        "First important point",
        "Second important point",
        "Third important point"
    ]
)

# Add text slide
creator.add_text_slide(
    title="Details",
    text_content="This slide contains paragraph text instead of bullet points."
)

# Save the presentation
creator.save_presentation("my_presentation.pptx")
```

### JSON Outline Format

Create presentations from structured JSON files:

```json
{
  "title": "Presentation Title",
  "subtitle": "Optional Subtitle",
  "slides": [
    {
      "type": "content",
      "title": "Slide Title",
      "content": ["Bullet point 1", "Bullet point 2"]
    },
    {
      "type": "text",
      "title": "Text Slide",
      "content": "Paragraph content goes here."
    },
    {
      "type": "image",
      "title": "Image Slide",
      "image_path": "path/to/image.png"
    }
  ]
}
```

## Examples

Run the basic example:
```bash
cd examples
python basic_example.py
```

This creates a sample presentation demonstrating various slide types.

## Project Structure

```
PPTX/
├── src/
│   ├── __init__.py
│   ├── pptx_creator.py    # Main PowerPoint creation class
│   └── cli.py             # Command-line interface
├── tests/
│   └── test_pptx_creator.py
├── examples/
│   ├── basic_example.py
│   └── sample_outline.json
├── requirements.txt
├── pyproject.toml
└── README.md
```

## Testing

Run the test suite:
```bash
# Activate virtual environment first
source venv/bin/activate

# Run tests
python -m pytest tests/ -v
```

## Development

### Code Formatting

Format code with Black:
```bash
python -m black src/ tests/ examples/
```

### Type Checking

Run MyPy for type checking:
```bash
python -m mypy src/
```

### Import Sorting

Sort imports with isort:
```bash
python -m isort src/ tests/ examples/
```

## Dependencies

- **python-pptx**: Core PowerPoint creation library
- **pytest**: Testing framework
- **black**: Code formatting
- **isort**: Import sorting
- **mypy**: Static type checking

## Contributing

1. Ensure all tests pass
2. Follow PEP 8 style guidelines
3. Add type hints to new functions
4. Update documentation for new features

## License

This project is created for use with Claude Code. Please refer to your Claude Code license for usage terms.