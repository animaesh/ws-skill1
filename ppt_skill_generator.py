import pandas as pd
import os
from typing import Dict, List, Any, Optional, Union
from data_analyzer import DataAnalyzer, ChartType
from ppt_handler import PowerPointHandler

class PowerPointSkillGenerator:
    def __init__(self, template_path: str = None):
        """
        Initialize the PowerPoint Skill Generator.
        
        Args:
            template_path: Path to PowerPoint template file (optional)
        """
        self.data_analyzer = DataAnalyzer()
        self.ppt_handler = PowerPointHandler(template_path)
        self.data = None
        self.analysis = None
    
    def load_data(self, data_source: Union[str, pd.DataFrame], **kwargs) -> None:
        """
        Load data from various sources.
        
        Args:
            data_source: Path to file (CSV, Excel, JSON) or pandas DataFrame
            **kwargs: Additional arguments for pandas read functions
        """
        if isinstance(data_source, pd.DataFrame):
            self.data = data_source
        elif isinstance(data_source, str):
            file_ext = os.path.splitext(data_source)[1].lower()
            
            if file_ext == '.csv':
                self.data = pd.read_csv(data_source, **kwargs)
            elif file_ext in ['.xlsx', '.xls']:
                self.data = pd.read_excel(data_source, **kwargs)
            elif file_ext == '.json':
                self.data = pd.read_json(data_source, **kwargs)
            else:
                raise ValueError(f"Unsupported file format: {file_ext}")
        else:
            raise ValueError("data_source must be a file path or pandas DataFrame")
        
        # Analyze the loaded data
        self.analysis = self.data_analyzer.analyze_data(self.data)
    
    def recommend_visualizations(self, target_column: str = None, max_charts: int = 5) -> List[Dict[str, Any]]:
        """
        Recommend visualizations based on data structure.
        
        Args:
            target_column: Specific column to focus on (optional)
            max_charts: Maximum number of chart recommendations
            
        Returns:
            List of chart configurations
        """
        if self.data is None:
            raise ValueError("No data loaded. Call load_data() first.")
        
        recommendations = []
        
        # Get primary recommendation
        primary_chart, confidence = self.data_analyzer.recommend_chart_type(self.data, target_column)
        primary_config = self.data_analyzer.get_chart_config(self.data, primary_chart, target_column)
        primary_config['confidence'] = confidence
        recommendations.append(primary_config)
        
        # Get additional recommendations for different perspectives
        if len(self.analysis['numeric_columns']) >= 2:
            # Add scatter plot for relationships
            scatter_config = self.data_analyzer.get_chart_config(self.data, ChartType.SCATTER)
            scatter_config['confidence'] = 0.7
            if scatter_config not in recommendations:
                recommendations.append(scatter_config)
        
        if len(self.analysis['categorical_columns']) >= 1 and len(self.analysis['numeric_columns']) >= 1:
            # Add bar chart for categorical comparison
            bar_config = self.data_analyzer.get_chart_config(self.data, ChartType.BAR)
            bar_config['confidence'] = 0.8
            if bar_config not in recommendations:
                recommendations.append(bar_config)
        
        if len(self.analysis['numeric_columns']) == 1:
            # Add histogram for distribution
            hist_config = self.data_analyzer.get_chart_config(self.data, ChartType.HISTOGRAM)
            hist_config['confidence'] = 0.6
            if hist_config not in recommendations:
                recommendations.append(hist_config)
        
        return recommendations[:max_charts]
    
    def create_presentation(self, data_source: Union[str, pd.DataFrame], 
                          output_path: str = None, title: str = "Data Visualization Report",
                          template_path: str = None, target_column: str = None,
                          auto_recommend: bool = True, chart_configs: List[Dict] = None) -> str:
        """
        Create a PowerPoint presentation with data visualizations.
        
        Args:
            data_source: Path to data file or pandas DataFrame
            output_path: Output file path (optional)
            title: Presentation title
            template_path: PowerPoint template path (optional)
            target_column: Specific column to focus on (optional)
            auto_recommend: Whether to automatically recommend charts
            chart_configs: Custom chart configurations (overrides auto_recommend)
            
        Returns:
            Path to the generated presentation
        """
        # Load data
        self.load_data(data_source)
        
        # Initialize PPT handler with template
        if template_path:
            self.ppt_handler = PowerPointHandler(template_path)
        
        # Get chart configurations
        if chart_configs is None:
            if auto_recommend:
                chart_configs = self.recommend_visualizations(target_column)
            else:
                # Use single best recommendation
                best_chart, _ = self.data_analyzer.recommend_chart_type(self.data, target_column)
                chart_configs = [self.data_analyzer.get_chart_config(self.data, best_chart, target_column)]
        
        # Create presentation
        if output_path is None:
            output_path = f"{title.replace(' ', '_')}.pptx"
        
        self.ppt_handler.create_presentation_from_data(self.data, chart_configs, title)
        
        # Move to final output path if different
        if output_path != f"{title.replace(' ', '_')}.pptx":
            os.rename(f"{title.replace(' ', '_')}.pptx", output_path)
        
        return output_path
    
    def create_custom_chart(self, chart_type: str, data_source: Union[str, pd.DataFrame],
                           config: Dict[str, Any], output_path: str = None) -> str:
        """
        Create a presentation with a single custom chart.
        
        Args:
            chart_type: Type of chart ('bar', 'line', 'pie', 'scatter', 'area', 'histogram')
            data_source: Path to data file or pandas DataFrame
            config: Chart configuration
            output_path: Output file path (optional)
            
        Returns:
            Path to the generated presentation
        """
        # Load data
        self.load_data(data_source)
        
        # Validate chart type
        try:
            chart_enum = ChartType(chart_type)
        except ValueError:
            raise ValueError(f"Invalid chart type: {chart_type}. Valid types: {[t.value for t in ChartType]}")
        
        # Get configuration
        full_config = self.data_analyzer.get_chart_config(self.data, chart_enum)
        full_config.update(config)
        
        # Create presentation
        title = f"{chart_type.title()} Chart Analysis"
        if output_path is None:
            output_path = f"{chart_type}_chart.pptx"
        
        self.ppt_handler.create_presentation_from_data(self.data, [full_config], title)
        
        # Move to final output path if different
        if output_path != f"{title.replace(' ', '_')}.pptx":
            os.rename(f"{title.replace(' ', '_')}.pptx", output_path)
        
        return output_path
    
    def get_data_insights(self) -> Dict[str, Any]:
        """
        Get detailed insights about the loaded data.
        
        Returns:
            Dictionary containing data analysis insights
        """
        if self.data is None:
            raise ValueError("No data loaded. Call load_data() first.")
        
        insights = self.analysis.copy()
        
        # Add additional insights
        insights['recommendations'] = {
            'primary_chart': self.data_analyzer.recommend_chart_type(self.data)[0].value,
            'confidence': self.data_analyzer.recommend_chart_type(self.data)[1],
            'data_quality': {
                'completeness': (1 - sum(insights['missing_values'].values()) / (insights['shape'][0] * insights['shape'][1])) * 100,
                'numeric_ratio': len(insights['numeric_columns']) / insights['shape'][1] * 100,
                'categorical_ratio': len(insights['categorical_columns']) / insights['shape'][1] * 100
            }
        }
        
        return insights
    
    def batch_process(self, data_files: List[str], output_dir: str = "output", 
                     template_path: str = None) -> List[str]:
        """
        Process multiple data files and create presentations.
        
        Args:
            data_files: List of data file paths
            output_dir: Directory to save presentations
            template_path: PowerPoint template path (optional)
            
        Returns:
            List of generated presentation paths
        """
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        generated_files = []
        
        for file_path in data_files:
            try:
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                output_path = os.path.join(output_dir, f"{file_name}_analysis.pptx")
                
                self.create_presentation(
                    data_source=file_path,
                    output_path=output_path,
                    title=f"{file_name} Analysis",
                    template_path=template_path
                )
                
                generated_files.append(output_path)
                print(f"Generated: {output_path}")
                
            except Exception as e:
                print(f"Error processing {file_path}: {str(e)}")
        
        return generated_files
