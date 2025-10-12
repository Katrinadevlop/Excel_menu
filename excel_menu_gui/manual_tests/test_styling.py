#!/usr/bin/env python3
"""
Test script to validate UI styling functionality
"""

import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel
from PySide6.QtCore import Qt

from app.gui.ui_styles import (
    AppStyles, StyleManager, ButtonStyles, ComponentStyles,
    LayoutStyles, StyleSheets
)


class TestWindow(QMainWindow):
    """Simple test window to validate styling"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("UI Styling Test")
        
        # Apply main window styling
        StyleManager.setup_main_window(self)
        
        # Create central widget and layout
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        LayoutStyles.apply_margins(layout, LayoutStyles.DEFAULT_MARGINS)
        layout.setSpacing(AppStyles.DEFAULT_SPACING)
        
        # Test caption label
        caption = QLabel("Test Caption Label")
        ComponentStyles.style_caption_label(caption)
        layout.addWidget(caption)
        
        # Test different button styles
        toolbar_btn = QPushButton("Toolbar Button")
        StyleManager.style_toolbar_button(toolbar_btn)
        layout.addWidget(toolbar_btn)
        
        action_btn = QPushButton("Action Button")  
        StyleManager.style_action_button(action_btn)
        layout.addWidget(action_btn)
        
        browse_btn = QPushButton("Browse Button")
        StyleManager.style_browse_button(browse_btn)
        layout.addWidget(browse_btn)
        
        # Test direct button styling
        custom_btn = QPushButton("Custom Styled Button")
        ButtonStyles.apply_button_style(custom_btn, {
            "background-color": "#4CAF50",
            "color": "white",
            "padding": "10px 20px",
            "border-radius": "5px",
            "font-weight": "bold"
        })
        layout.addWidget(custom_btn)
        
        # Info label
        info_label = QLabel("All styling components loaded successfully!")
        info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(info_label)
        
        # Resize to reasonable size
        self.resize(400, 300)


def main():
    """Main test function"""
    app = QApplication(sys.argv)
    
    # Test stylesheet generation
    main_css = StyleSheets.get_main_stylesheet()
    print(f"✓ Main stylesheet generated: {len(main_css)} characters")
    
    # Test constants
    print(f"✓ Window size: {AppStyles.WINDOW_DEFAULT_SIZE}")
    print(f"✓ Button padding: {ButtonStyles.DEFAULT_PADDING}")
    print(f"✓ Border radius: {AppStyles.DEFAULT_BORDER_RADIUS}px")
    
    # Create and show test window
    window = TestWindow()
    window.show()
    
    print("✓ Test window created and styled successfully!")
    print("Close the window to complete the test.")
    
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
