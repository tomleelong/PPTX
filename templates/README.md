# Templates Directory

Place your PowerPoint template files (`.pptx`) in this directory to use them with the PPTXCreator.

## How to Use Templates

1. **Add your template file** to this directory (e.g., `my_template.pptx`)

2. **Clean your template** (important!):
   - Open your template in PowerPoint
   - Delete all content slides, keeping only the slide master and layouts
   - Save the file

3. **Use in your code**:
   ```python
   from src.pptx_creator import PPTXCreator
   
   creator = PPTXCreator("templates/my_template.pptx")
   creator.add_title_slide("My Presentation", "Using template styling")
   creator.save_presentation("output.pptx")
   ```

## Template Preparation Tips

- **Keep slide masters and layouts** - These provide the styling, fonts, colors, and branding
- **Remove content slides** - Delete all slides with actual content to avoid mixing old content with new
- **Test your template** - Create a simple presentation to ensure it works correctly

## File Structure

```
templates/
├── README.md          # This file
├── my_template.pptx   # Your template files (gitignored)
└── other_template.pptx
```

Template `.pptx` files are automatically ignored by git to keep personal/company templates private while allowing the folder structure to be shared.