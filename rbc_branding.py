"""
RBC Brand Configuration for PowerPoint Skill Generator
Contains official RBC colors, fonts, and styling guidelines
"""

from typing import Dict, List, Tuple
import matplotlib.pyplot as plt
import seaborn as sns

class RBCBranding:
    """RBC Brand Guidelines and Configuration"""
    
    # Official RBC Color Palette
    COLORS = {
        'primary_blue': '#005DAA',      # Medium Persian Blue
        'accent_yellow': '#FFD200',     # Cyber Yellow
        'white': '#FFFFFF',             # White
        'light_blue': '#E6F2FF',       # Light blue variant for backgrounds
        'dark_blue': '#003D73',        # Dark blue variant for text
        'gray_light': '#F5F5F5',       # Light gray
        'gray_medium': '#CCCCCC',       # Medium gray
        'gray_dark': '#666666',        # Dark gray
        'black': '#000000',            # Black for high contrast
    }
    
    # RGB values for matplotlib
    RGB_COLORS = {
        'primary_blue': (0, 93, 170),
        'accent_yellow': (255, 210, 0),
        'white': (255, 255, 255),
        'light_blue': (230, 242, 255),
        'dark_blue': (0, 61, 115),
        'gray_light': (245, 245, 245),
        'gray_medium': (204, 204, 204),
        'gray_dark': (102, 102, 102),
        'black': (0, 0, 0),
    }
    
    # Typography
    FONTS = {
        'primary': 'Arial',           # Most commonly used in RBC presentations
        'secondary': 'Calibri',       # Alternative for data labels
        'heading': 'Arial Bold',      # For titles and headers
        'body': 'Arial',             # For body text
        'data_labels': 'Calibri',     # For chart labels
    }
    
    # Font sizes
    FONT_SIZES = {
        'title': 24,
        'subtitle': 18,
        'heading': 16,
        'body': 12,
        'caption': 10,
        'data_label': 11,
        'axis_label': 10,
        'legend': 10,
    }
    
    # Chart color palettes for different data types
    CHART_PALETTES = {
        'primary': ['#005DAA', '#FFD200', '#666666', '#CCCCCC', '#003D73'],
        'sequential': ['#E6F2FF', '#B3D9FF', '#80C0FF', '#4DA7FF', '#1A8EFF', '#0075E6', '#005DAA', '#00458E', '#002D72'],
        'diverging': ['#003D73', '#005DAA', '#4DA7FF', '#FFFFFF', '#FFD200', '#FFB300', '#FF9400'],
        'categorical': ['#005DAA', '#FFD200', '#666666', '#CCCCCC', '#003D73', '#FF9400', '#4DA7FF', '#FFB300'],
    }
    
    @classmethod
    def get_color_palette(cls, palette_type: str = 'primary', n_colors: int = 5) -> List[str]:
        """Get a color palette for charts"""
        if palette_type in cls.CHART_PALETTES:
            palette = cls.CHART_PALETTES[palette_type]
            if n_colors <= len(palette):
                return palette[:n_colors]
            else:
                # Repeat colors if more are needed
                return (palette * ((n_colors // len(palette)) + 1))[:n_colors]
        else:
            # Default to primary palette
            return cls.CHART_PALETTES['primary'][:n_colors]
    
    @classmethod
    def setup_matplotlib_style(cls):
        """Configure matplotlib with RBC branding"""
        # Set color palette
        sns.set_palette(cls.CHART_PALETTES['primary'])
        
        # Set default parameters
        plt.rcParams.update({
            'font.family': 'sans-serif',
            'font.sans-serif': cls.FONTS['primary'],
            'font.size': cls.FONT_SIZES['body'],
            'axes.titlesize': cls.FONT_SIZES['heading'],
            'axes.labelsize': cls.FONT_SIZES['axis_label'],
            'xtick.labelsize': cls.FONT_SIZES['data_label'],
            'ytick.labelsize': cls.FONT_SIZES['data_label'],
            'legend.fontsize': cls.FONT_SIZES['legend'],
            'figure.titlesize': cls.FONT_SIZES['title'],
            'axes.titleweight': 'bold',
            'axes.labelweight': 'normal',
            'text.color': cls.COLORS['dark_blue'],
            'axes.labelcolor': cls.COLORS['dark_blue'],
            'xtick.color': cls.COLORS['dark_blue'],
            'ytick.color': cls.COLORS['dark_blue'],
            'axes.edgecolor': cls.COLORS['gray_medium'],
            'axes.facecolor': cls.COLORS['white'],
            'figure.facecolor': cls.COLORS['white'],
            'savefig.facecolor': cls.COLORS['white'],
            'savefig.edgecolor': 'none',
        })
    
    @classmethod
    def get_chart_style_config(cls, chart_type: str) -> Dict:
        """Get chart-specific styling configuration"""
        base_config = {
            'title_font': cls.FONTS['heading'],
            'title_size': cls.FONT_SIZES['heading'],
            'title_color': cls.COLORS['dark_blue'],
            'label_font': cls.FONTS['data_labels'],
            'label_size': cls.FONT_SIZES['data_label'],
            'label_color': cls.COLORS['dark_blue'],
            'grid_color': cls.COLORS['gray_light'],
            'grid_alpha': 0.3,
            'spine_color': cls.COLORS['gray_medium'],
            'background_color': cls.COLORS['white'],
        }
        
        if chart_type == 'bar':
            base_config.update({
                'bar_colors': cls.get_color_palette('categorical', 8),
                'edge_color': cls.COLORS['white'],
                'edge_width': 0.5,
            })
        elif chart_type == 'line':
            base_config.update({
                'line_colors': cls.get_color_palette('primary', 5),
                'line_width': 2.5,
                'marker_size': 6,
                'marker_color': cls.COLORS['primary_blue'],
            })
        elif chart_type == 'pie':
            base_config.update({
                'pie_colors': cls.get_color_palette('categorical', 8),
                'wedge_props': {'edgecolor': cls.COLORS['white'], 'linewidth': 1},
            })
        elif chart_type == 'scatter':
            base_config.update({
                'scatter_colors': cls.get_color_palette('primary', 3),
                'scatter_alpha': 0.7,
                'scatter_edge_color': cls.COLORS['white'],
                'scatter_edge_width': 0.5,
            })
        elif chart_type == 'area':
            base_config.update({
                'area_colors': cls.get_color_palette('sequential', 5),
                'area_alpha': 0.7,
            })
        elif chart_type == 'histogram':
            base_config.update({
                'hist_color': cls.COLORS['primary_blue'],
                'hist_edge_color': cls.COLORS['white'],
                'hist_alpha': 0.8,
            })
        elif chart_type == 'stacked_bar':
            base_config.update({
                'bar_colors': cls.get_color_palette('categorical', 8),
                'edge_color': cls.COLORS['white'],
                'edge_width': 0.5,
            })
        elif chart_type == 'stacked_area':
            base_config.update({
                'area_colors': cls.get_color_palette('sequential', 5),
                'area_alpha': 0.7,
            })
        
        return base_config
    
    @classmethod
    def apply_rbc_theme_to_axis(cls, ax, chart_type: str):
        """Apply RBC styling to matplotlib axis"""
        config = cls.get_chart_style_config(chart_type)
        
        # Set spine colors
        for spine in ax.spines.values():
            spine.set_color(config['spine_color'])
            spine.set_linewidth(0.8)
        
        # Set grid
        ax.grid(True, color=config['grid_color'], alpha=config['grid_alpha'], linestyle='-', linewidth=0.5)
        ax.set_axisbelow(True)
        
        # Set background
        ax.set_facecolor(config['background_color'])
        
        # Set tick parameters
        ax.tick_params(colors=config['label_color'], labelsize=config['label_size'])
        
        # Set label colors
        ax.xaxis.label.set_color(config['label_color'])
        ax.yaxis.label.set_color(config['label_color'])
        
        return config
