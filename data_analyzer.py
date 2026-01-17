import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Any
from enum import Enum

class ChartType(Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"
    SCATTER = "scatter"
    AREA = "area"
    HISTOGRAM = "histogram"

class DataAnalyzer:
    def __init__(self):
        self.chart_rules = {
            ChartType.BAR: self._should_use_bar,
            ChartType.LINE: self._should_use_line,
            ChartType.PIE: self._should_use_pie,
            ChartType.SCATTER: self._should_use_scatter,
            ChartType.AREA: self._should_use_area,
            ChartType.HISTOGRAM: self._should_use_histogram
        }
    
    def analyze_data(self, data: pd.DataFrame) -> Dict[str, Any]:
        """Analyze data structure and return insights for chart selection."""
        analysis = {
            'shape': data.shape,
            'dtypes': data.dtypes.to_dict(),
            'numeric_columns': data.select_dtypes(include=[np.number]).columns.tolist(),
            'categorical_columns': data.select_dtypes(include=['object', 'category']).columns.tolist(),
            'datetime_columns': data.select_dtypes(include=['datetime64']).columns.tolist(),
            'missing_values': data.isnull().sum().to_dict(),
            'unique_counts': {col: data[col].nunique() for col in data.columns}
        }
        
        # Additional analysis
        analysis['numeric_stats'] = data[analysis['numeric_columns']].describe().to_dict() if analysis['numeric_columns'] else {}
        analysis['correlation_matrix'] = data[analysis['numeric_columns']].corr().to_dict() if len(analysis['numeric_columns']) > 1 else {}
        
        return analysis
    
    def recommend_chart_type(self, data: pd.DataFrame, target_column: str = None) -> Tuple[ChartType, float]:
        """Recommend the best chart type based on data structure."""
        analysis = self.analyze_data(data)
        
        scores = {}
        for chart_type, rule_func in self.chart_rules.items():
            score = rule_func(data, analysis, target_column)
            scores[chart_type] = score
        
        best_chart = max(scores, key=scores.get)
        confidence = scores[best_chart]
        
        return best_chart, confidence
    
    def _should_use_bar(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for bar chart suitability."""
        score = 0.0
        
        # Good for categorical vs numerical
        if (len(analysis['categorical_columns']) >= 1 and 
            len(analysis['numeric_columns']) >= 1):
            score += 0.8
        
        # Good for comparing values across categories
        if target_column and target_column in analysis['numeric_columns']:
            cat_cols = [col for col in analysis['categorical_columns'] if col != target_column]
            if cat_cols:
                score += 0.6
        
        # Not too many categories
        if analysis['categorical_columns']:
            max_unique = max(analysis['unique_counts'][col] for col in analysis['categorical_columns'])
            if max_unique <= 10:
                score += 0.4
            elif max_unique <= 20:
                score += 0.2
        
        return min(score, 1.0)
    
    def _should_use_line(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for line chart suitability."""
        score = 0.0
        
        # Good for time series data
        if analysis['datetime_columns']:
            score += 0.9
        
        # Good for showing trends over ordered categories
        if len(analysis['numeric_columns']) >= 1:
            score += 0.5
        
        # Good for continuous data
        if target_column and target_column in analysis['numeric_columns']:
            score += 0.3
        
        return min(score, 1.0)
    
    def _should_use_pie(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for pie chart suitability."""
        score = 0.0
        
        # Only for single numerical column with categories
        if (len(analysis['numeric_columns']) == 1 and 
            len(analysis['categorical_columns']) == 1):
            score += 0.7
        
        # Limited number of categories (2-7 ideal)
        if analysis['categorical_columns']:
            unique_counts = [analysis['unique_counts'][col] for col in analysis['categorical_columns']]
            if any(2 <= count <= 7 for count in unique_counts):
                score += 0.6
            elif any(count <= 10 for count in unique_counts):
                score += 0.3
        
        # Data should represent parts of a whole
        if target_column and target_column in analysis['numeric_columns']:
            col_data = data[target_column].dropna()
            if (col_data >= 0).all():  # All non-negative values
                score += 0.2
        
        return min(score, 1.0)
    
    def _should_use_scatter(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for scatter plot suitability."""
        score = 0.0
        
        # Need at least 2 numeric columns
        if len(analysis['numeric_columns']) >= 2:
            score += 0.8
        
        # Good for showing relationships between variables
        if len(analysis['numeric_columns']) >= 2:
            score += 0.4
        
        # Good for large datasets
        if data.shape[0] > 50:
            score += 0.2
        
        return min(score, 1.0)
    
    def _should_use_area(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for area chart suitability."""
        score = 0.0
        
        # Similar to line but good for showing magnitude
        if analysis['datetime_columns']:
            score += 0.7
        
        # Good for cumulative data
        if len(analysis['numeric_columns']) >= 1:
            score += 0.5
        
        return min(score, 1.0)
    
    def _should_use_histogram(self, data: pd.DataFrame, analysis: Dict, target_column: str = None) -> float:
        """Score for histogram suitability."""
        score = 0.0
        
        # Good for single numeric column distribution
        if len(analysis['numeric_columns']) == 1:
            score += 0.8
        
        # Good for showing frequency distribution
        if target_column and target_column in analysis['numeric_columns']:
            score += 0.6
        
        # Need sufficient data points
        if data.shape[0] >= 20:
            score += 0.3
        
        return min(score, 1.0)
    
    def get_chart_config(self, data: pd.DataFrame, chart_type: ChartType, target_column: str = None) -> Dict[str, Any]:
        """Generate configuration for the selected chart type."""
        analysis = self.analyze_data(data)
        config = {'chart_type': chart_type.value}
        
        if chart_type == ChartType.BAR:
            config.update({
                'x_column': analysis['categorical_columns'][0] if analysis['categorical_columns'] else data.columns[0],
                'y_column': target_column or analysis['numeric_columns'][0] if analysis['numeric_columns'] else data.columns[1],
                'orientation': 'vertical'
            })
        
        elif chart_type == ChartType.LINE:
            config.update({
                'x_column': analysis['datetime_columns'][0] if analysis['datetime_columns'] else data.columns[0],
                'y_column': target_column or analysis['numeric_columns'][0] if analysis['numeric_columns'] else data.columns[1],
            })
        
        elif chart_type == ChartType.PIE:
            config.update({
                'labels_column': analysis['categorical_columns'][0] if analysis['categorical_columns'] else data.columns[0],
                'values_column': target_column or analysis['numeric_columns'][0] if analysis['numeric_columns'] else data.columns[1],
            })
        
        elif chart_type == ChartType.SCATTER:
            numeric_cols = analysis['numeric_columns']
            config.update({
                'x_column': numeric_cols[0] if len(numeric_cols) > 0 else data.columns[0],
                'y_column': numeric_cols[1] if len(numeric_cols) > 1 else data.columns[1],
            })
        
        elif chart_type == ChartType.AREA:
            config.update({
                'x_column': analysis['datetime_columns'][0] if analysis['datetime_columns'] else data.columns[0],
                'y_column': target_column or analysis['numeric_columns'][0] if analysis['numeric_columns'] else data.columns[1],
            })
        
        elif chart_type == ChartType.HISTOGRAM:
            config.update({
                'column': target_column or analysis['numeric_columns'][0] if analysis['numeric_columns'] else data.columns[0],
                'bins': min(30, max(5, data.shape[0] // 10))
            })
        
        return config
