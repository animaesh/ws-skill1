# PowerPoint Skill Generator

A Python skill that automatically populates PowerPoint decks with data-driven graphs. The skill analyzes data structure and selects the most appropriate visualization type.

## Features

- Automatically analyzes data structure to determine optimal graph types
- Supports multiple data formats (CSV, Excel, JSON)
- Generates various chart types (bar, line, pie, scatter, etc.)
- Uses PowerPoint templates for consistent styling
- Handles multiple data series and categories
- **RBC Branding**: Official RBC colors and fonts for professional presentations

## RBC Branding

The skill now includes official RBC branding guidelines:

### 🎨 **RBC Color Palette**
- **Primary Blue**: #005DAA (Medium Persian Blue)
- **Accent Yellow**: #FFD200 (Cyber Yellow)  
- **White**: #FFFFFF
- Additional supporting colors for charts and backgrounds

### 🔤 **Typography**
- **Primary Font**: Arial (most commonly used in RBC presentations)
- **Data Labels**: Calibri
- **Sizes**: Title (24pt), Heading (16pt), Body (12pt)

### 📊 **Chart Styling**
- All charts use RBC's official color scheme
- Professional formatting matching RBC investor relations materials
- Consistent fonts and styling throughout presentations

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
```python
from ppt_skill_generator import PowerPointSkillGenerator

# Initialize the skill (with RBC branding by default)
generator = PowerPointSkillGenerator()

# Generate presentation from data
generator.create_presentation(
    data_source="data.xlsx",
    template_path="template.pptx",
    output_path="output.pptx"
)
```

### RBC-Styled Presentations
```python
# Create RBC-styled financial presentation
generator.create_custom_chart(
    chart_type='line',
    data_source=financial_data,
    config={'x_column': 'Month', 'y_column': 'Revenue'},
    output_path="rbc_financial_analysis.pptx"
)
```

### Test RBC Branding
```bash
python test_rbc_branding.py
```

## Supported Chart Types

- Bar charts (vertical/horizontal)
- Line charts  
- Pie charts
- Scatter plots
- Area charts
- Histograms

All charts are automatically styled with RBC colors and fonts.
