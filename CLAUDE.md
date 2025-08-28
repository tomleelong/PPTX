# CLAUDE.md - PowerPoint Creator Project

## Project Overview
This is a PowerPoint creation toolkit built with Python and the python-pptx library. It allows programmatic generation of presentations through code, CLI, or JSON outlines.

## Environment Setup
- **Virtual Environment**: Always use `source venv/bin/activate` before running any Python commands
- **Dependencies**: Installed via `requirements.txt` with latest stable versions
- **Python Version**: 3.10+ required

## Project Structure
```
src/
├── pptx_creator.py    # Main PowerPoint creation class (includes template support)
└── cli.py             # Command-line interface

tests/
└── test_pptx_creator.py

examples/
├── basic_example.py
└── sample_outline.json

templates/
├── README.md          # Template usage instructions
└── *.pptx             # User template files (gitignored)

content/
└── *.md, *.txt, etc.  # Private notes and source materials (gitignored)
```

### Content Folder
The `content/` directory is for storing private notes, meeting minutes, and source documents that you're converting into presentations. This folder is gitignored to keep sensitive materials out of version control. Store your committee notes, agendas, and other confidential documents here while working on presentations.

## Development Commands

### Testing
```bash
source venv/bin/activate
python -m pytest tests/ -v
```

### Code Quality
```bash
# Format code
python -m black src/ tests/ examples/

# Sort imports
python -m isort src/ tests/ examples/

# Type checking
python -m mypy src/
```

### Running Examples
```bash
# CLI sample presentation
python -m src.cli --sample output.pptx

# From JSON outline
python -m src.cli --json examples/sample_outline.json --output report.pptx

# Basic Python example
cd examples && python basic_example.py
```

## Code Patterns

### Creating Presentations
Always use the `PPTXCreator` class:

**Basic presentations:**
```python
from src.pptx_creator import PPTXCreator

creator = PPTXCreator()
creator.add_title_slide("Title", "Subtitle")
creator.add_content_slide("Content", ["Point 1", "Point 2"])
creator.save_presentation("output.pptx")
```

**Template-based presentations:**
```python
from src.pptx_creator import PPTXCreator

# Use existing template for professional styling
creator = PPTXCreator("templates/my_template.pptx")
creator.add_title_slide("Professional Title", "Uses template styling")
creator.add_content_slide("Content", ["Point 1", "Point 2"])
creator.save_presentation("styled_output.pptx")
```

### JSON Structure
Use this format for outline-based presentations:
```json
{
  "title": "Presentation Title",
  "subtitle": "Optional Subtitle",
  "slides": [
    {"type": "content", "title": "Title", "content": ["Point 1"]},
    {"type": "text", "title": "Title", "content": "Paragraph text"},
    {"type": "image", "title": "Title", "image_path": "path/to/image.png"}
  ]
}
```

## Common Tasks

### Using Templates
1. **Prepare your template:**
   - Open template in PowerPoint
   - Delete all content slides (keep only slide master and layouts)
   - Save as cleaned template in `templates/` directory

2. **Use in code:**
   ```python
   creator = PPTXCreator("templates/your_template.pptx")
   ```

3. **Template files are gitignored** - your company/personal templates stay private

### Adding New Slide Types
1. Add method to `PPTXCreator` class in `src/pptx_creator.py`
2. Update `create_from_outline` method to handle new type
3. Add corresponding test in `tests/test_pptx_creator.py`
4. Update CLI help text and examples

### Extending CLI
- Modify `src/cli.py` for new command-line options
- Follow existing argument parsing patterns
- Add help text and examples

### Testing Requirements
- All new features must have tests
- Tests must pass before committing
- Use `pytest` with descriptive test names
- Test both success and error cases

## Dependencies to Remember
- **python-pptx**: Core PowerPoint creation (v1.0.2+)
- **pytest**: Testing framework
- **black**: Code formatting
- **isort**: Import sorting  
- **mypy**: Type checking

## File Naming Conventions
- PowerPoint files: `*.pptx`
- JSON outlines: `*_outline.json` or `*.json`
- Test files: `test_*.py`
- Example files: `*_example.py`

## Error Handling
- Always validate file paths before using images
- Handle missing JSON keys gracefully
- Provide meaningful error messages
- Log important operations with appropriate levels

## Performance Notes
- Virtual environment activation is required for all operations
- Large presentations may take time to generate
- Image files should be optimized before inclusion

## Integration with Claude Code
This project is specifically designed for Claude Code workflows:
- CLI interface for quick generation
- JSON outlines for structured content
- Python API for programmatic control
- Comprehensive error handling and logging