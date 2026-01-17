# PowerPoint Skill Generator

A Python skill that automatically populates PowerPoint decks with data-driven graphs. The skill analyzes data structure and selects the most appropriate visualization type.

## Features

- Automatically analyzes data structure to determine optimal graph types
- Supports multiple data formats (CSV, Excel, JSON)
- Generates various chart types (bar, line, pie, scatter, etc.)
- Uses PowerPoint templates for consistent styling
- Handles multiple data series and categories

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```python
from ppt_skill_generator import PowerPointSkillGenerator

# Initialize the skill
generator = PowerPointSkillGenerator()

# Generate presentation from data
generator.create_presentation(
    data_source="data.xlsx",
    template_path="template.pptx",
    output_path="output.pptx"
)
```

## Supported Chart Types

- Bar charts (vertical/horizontal)
- Line charts
- Pie charts
- Scatter plots
- Area charts
- Histograms
