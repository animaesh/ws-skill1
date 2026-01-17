"""
PowerPoint Skill Generator

A Python skill that automatically populates PowerPoint decks with data-driven graphs.
The skill analyzes data structure and selects the most appropriate visualization type.
"""

from .ppt_skill_generator import PowerPointSkillGenerator
from .data_analyzer import DataAnalyzer, ChartType
from .ppt_handler import PowerPointHandler

__version__ = "1.0.0"
__author__ = "PowerPoint Skill Generator"

__all__ = [
    'PowerPointSkillGenerator',
    'DataAnalyzer', 
    'ChartType',
    'PowerPointHandler'
]
