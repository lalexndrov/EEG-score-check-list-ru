# Dynamic EEG Form Generator

A flexible GUI application for creating customizable forms based on JSON templates, with support for generating Word and PDF documents from form data.

## Features

- 🎨 **Dynamic Form Generation**: Create forms from JSON templates without modifying code
- 📝 **Multiple Field Types**: Support for text, number, select, checkbox, textarea, and date fields
- 📄 **Document Export**: Generate Word (.docx) and PDF documents from form data
- 🔄 **Template System**: Use Word templates with placeholders for consistent document formatting
- 🖱️ **User-Friendly Interface**: Scrollable GUI with organized sections
- 💾 **Data Management**: JSON output with clipboard integration
- 🌐 **Multilingual Support**: Full UTF-8 support for international characters

## Quick Start

### Prerequisites

- Python 3.7 or higher
- pip package manager

### Installation

1. **Clone or download the project files**
2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   python dynamic_form.py
   ```

## Project Structure

```
eeg-form-generator/
├── dynamic_form.py          # Main GUI application
├── document_generator.py    # Document generation module
├── eeg_template.json       # Form template configuration
├── requirements.txt        # Python dependencies
├── README.md              # This file
└── word_template.docx     # Word template (auto-generated)
```

## Usage

### Creating Forms

1. **Launch the application:** `python dynamic_form.py`
2. **Fill out the form fields** based on your needs
3. **Use the action buttons:**
   - **Generate JSON**: Creates JSON output and copies to clipboard
   - **Clear Form**: Resets all fields to default values
   - **Save to Word**: Exports data to a Word document
   - **Save to PDF**: Exports data to a PDF document

### Customizing Forms

Edit the `eeg_template.json` file to modify the form structure:

```json
{
  "form_title": "Your Custom Form Title",
  "sections": [
    {
      "id": "section_id",
      "title": "Section Title",
      "fields": [
        {
          "id": "field_id",
          "label": "Field Label:",
          "type": "text|number|select|checkbox|textarea|date",
          "default": "default_value",
          "required": true|false
        }
      ]
    }
  ]
}
```

### Field Types

| Type | Description | Additional Properties |
|------|-------------|----------------------|
| `text` | Single-line text input | - |
| `number` | Numeric input (auto-converted) | - |
| `select` | Dropdown list | `options: ["opt1", "opt2"]` |
| `checkbox` | Boolean checkbox | - |
| `textarea` | Multi-line text input | `height: 4` (rows) |
| `date` | Date input | Use `"today"` for current date |

### Document Templates

#### Word Templates
- Place placeholders in format: `{{field_id}}`
- Example: `Patient: {{patient_name}}`
- Template file: `word_template.docx` (auto-generated if missing)

#### PDF Generation
- Automatically formatted with sections and tables
- No template needed - generates structured layout

## Advanced Usage

### Loading Custom Templates

1. **Via Menu**: File → Load Template...
2. **Via Code**: Modify `template_file` parameter in `DynamicFormGUI()`

### Programmatic Usage

```python
from document_generator import DocumentGenerator

# Create generator instance
generator = DocumentGenerator()

# Generate documents
data = {"patient_name": "John Doe", "age": 45}
generator.generate_word_document(data, "output.docx")
generator.generate_pdf_document(data, "output.pdf")
```

### Creating Custom Word Templates

```python
from document_generator import DocumentGenerator

generator = DocumentGenerator()
generator.create_word_template("my_template.docx")
```

## Configuration

### Window Settings
Modify in template JSON:
```json
"window_config": {
  "width": 800,
  "height": 900
}
```

### Default Values
- `"today"` - Current date for date fields
- `""` - Empty string for text fields
- `false` - Boolean for checkboxes
- First option for select fields

## Troubleshooting

### Common Issues

**"Template loading error"**
- Ensure `eeg_template.json` exists and is valid JSON
- Check file permissions

**"Error generating Word document"**
- Install python-docx: `pip install python-docx`
- Verify template file exists and is accessible

**"Error generating PDF document"**
- Install reportlab: `pip install reportlab`
- Check output directory permissions

**GUI not responsive**
- Large forms may take time to load
- Try reducing the number of fields or sections

### Dependencies Issues

If you encounter import errors:
```bash
# Install all dependencies
pip install --upgrade -r requirements.txt

# Or install individually
pip install python-docx reportlab
```

## Development

### Adding New Field Types

1. **Update JSON template** with new type
2. **Modify `create_field()` method** in `dynamic_form.py`
3. **Update `collect_data()` method** for data handling

### Extending Document Generation

1. **Modify `DocumentGenerator` class** in `document_generator.py`
2. **Add new methods** for different output formats
3. **Update GUI buttons** to call new methods

## Examples

### Medical Form Template
```json
{
  "form_title": "Medical Assessment Form",
  "sections": [
    {
      "title": "Patient Demographics",
      "fields": [
        {"id": "name", "label": "Full Name:", "type": "text"},
        {"id": "dob", "label": "Date of Birth:", "type": "date"},
        {"id": "gender", "label": "Gender:", "type": "select", 
         "options": ["Male", "Female", "Other"]}
      ]
    }
  ]
}
```

### Survey Template
```json
{
  "form_title": "Customer Satisfaction Survey",
  "sections": [
    {
      "title": "Feedback",
      "fields": [
        {"id": "rating", "label": "Overall Rating:", "type": "select",
         "options": ["1", "2", "3", "4", "5"]},
        {"id": "comments", "label": "Comments:", "type": "textarea", "height": 6}
      ]
    }
  ]
}
```

## License

This project is provided as-is for educational and professional use.

## Contributing

1. Fork the repository
2. Create feature branch
3. Make changes
4. Test thoroughly
5. Submit pull request

## Support

For issues and questions:
- Check troubleshooting section
- Review template JSON syntax
- Verify all dependencies are installed

---

**Version**: 1.0  
**Last Updated**: 2024