import sys
import os
import openai
import random
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QTextEdit, QPushButton, QLabel, 
                           QComboBox, QMessageBox, QTextBrowser, QProgressBar,
                           QInputDialog, QLineEdit)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QPalette, QColor, QIcon
from dotenv import load_dotenv
from markdown2 import markdown

class OpenAIThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, client, messages):
        super().__init__()
        self.client = client
        self.messages = messages

    def run(self):
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=self.messages
            )
            self.finished.emit(response.choices[0].message.content)
        except Exception as e:
            self.error.emit(str(e))

class ExcelFormulaAssistant(QMainWindow):
    PRIMARY_COLOR = "#6654f5"
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Formula Assistant")
        self.setMinimumSize(800, 600)
        self.setStyleSheet(f"""
            QMainWindow {{background: white;}}
            QLabel {{font-weight: bold; color: #333;}}
            QScrollBar:vertical {{                
                border: none;
                background: #f0f0f0;
                width: 10px;
                margin: 0;
            }}
            QScrollBar::handle:vertical {{                
                background: #c1c1c1;
                min-height: 30px;
                border-radius: 5px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{                
                height: 0;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{                
                background: none;
            }}
            QComboBox {{                
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 4px;
                min-width: 150px;
            }}
            QComboBox::drop-down {{                
                border: none;
                padding-right: 10px;
            }}
            QComboBox::down-arrow {{                
                border: none;
                background: #6654f5;
                width: 12px;
                height: 12px;
                border-radius: 6px;
            }}
            QComboBox QAbstractItemView {{                
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
                selection-background-color: #6654f5;
                selection-color: white;
            }}
            QPushButton {{
                background-color: {self.PRIMARY_COLOR};
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #5143d4;
            }}
            QComboBox {{
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 4px;
                min-width: 150px;
            }}
            QTextEdit {{
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 8px;
                font-family: Arial;
                background-color: white;
            }}
            QTextBrowser {{
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 8px;
                font-family: Arial;
                background-color: #f8f9fa;
            }}
        """)
       
        
        # Get application path based on whether we're running as executable or script
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(os.path.abspath(sys.executable))
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        env_path = os.path.join(application_path, ".env")
        
        # Try to get API key from .env if it exists
        self.api_key = None
        if os.path.exists(env_path):
            load_dotenv(env_path)
            env_key = os.getenv('OPENAI_API_KEY')
            if env_key and env_key.strip():  # Only use key from .env if it's not empty
                self.api_key = env_key.strip()
        
        # If no valid API key found, prompt user
        if not self.api_key:
            self.get_api_key_from_user()
        else:
            # Test the API key by initializing OpenAI
            try:
                self.initialize_openai()
            except Exception:
                # If API key is invalid, prompt user for a new one
                self.get_api_key_from_user()

        self.setup_ui()

    def setup_ui(self):
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Top bar layout
        top_bar = QHBoxLayout()
        
        # Left side - Mode selection
        mode_group = QHBoxLayout()
        mode_label = QLabel("Mode:")
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Generate Formula", "Explain Formula"])
        mode_group.addWidget(mode_label)
        mode_group.addWidget(self.mode_combo)
        top_bar.addLayout(mode_group)
        
        top_bar.addStretch()
        
        # Right side - Utility buttons
        util_buttons = QHBoxLayout()
        
        # Start Fresh button
        self.reset_button = QPushButton("Start Fresh")
        self.reset_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.reset_button.clicked.connect(self.reset_form)
        self.reset_button.setStyleSheet(self.get_secondary_button_style())
        
        # Random Task button
        self.random_task_btn = QPushButton("Random Task")
        self.random_task_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.random_task_btn.clicked.connect(self.generate_random_task)
        
        util_buttons.addWidget(self.random_task_btn)
        util_buttons.addWidget(self.reset_button)
        top_bar.addLayout(util_buttons)
        
        layout.addLayout(top_bar)

        # Input section
        input_section = QVBoxLayout()
        input_section.setSpacing(5)
        
        # Input header with label and copy button
        input_header = QHBoxLayout()
        input_label = QLabel("Input:")
        # Input buttons
        input_buttons = QHBoxLayout()
        input_buttons.setSpacing(5)
        
        input_copy_btn = QPushButton()
        input_copy_btn.setIcon(QIcon.fromTheme("edit-copy"))
        input_copy_btn.setFixedSize(24, 24)
        input_copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        input_copy_btn.setStyleSheet("background: transparent; border: none;")
        input_copy_btn.clicked.connect(lambda: self.copy_text(self.input_text))
        
        input_clear_btn = QPushButton("×")
        input_clear_btn.setFixedSize(24, 24)
        input_clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        input_clear_btn.setStyleSheet("""
            QPushButton {
                background: #eee;
                border: 1px solid #ddd;
                border-radius: 12px;
                color: #666;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #ff4444;
                color: white;
                border-color: #ff4444;
            }
        """)
        
        input_buttons.addWidget(input_copy_btn)
        input_buttons.addWidget(input_clear_btn)
        
        input_header.addWidget(input_label)
        input_header.addStretch()
        input_header.addLayout(input_buttons)
        
        input_section.addLayout(input_header)
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("For Generate Mode: Describe what you want the Excel formula to do\n"
                                         "For Explain Mode: Paste the Excel formula you want to understand")
        self.input_text.setMaximumHeight(100)
        input_section.addWidget(self.input_text)
        layout.addLayout(input_section)
        
        # Connect input clear button after input_text is created
        input_clear_btn.clicked.connect(self.input_text.clear)

        # Process button and loading indicator in center
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.process_button = QPushButton("Process")
        self.process_button.setFixedWidth(200)
        self.process_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.process_button.clicked.connect(self.process_request)
        button_layout.addWidget(self.process_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # Loading indicator centered
        loading_layout = QHBoxLayout()
        loading_layout.addStretch()
        self.loading_indicator = QProgressBar()
        self.loading_indicator.setFixedWidth(200)
        self.loading_indicator.setTextVisible(False)
        self.loading_indicator.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ccc;
                border-radius: 4px;
                height: 8px;
            }
            QProgressBar::chunk {
                background-color: #6654f5;
                border-radius: 4px;
            }
        """)
        self.loading_indicator.hide()
        loading_layout.addWidget(self.loading_indicator)
        loading_layout.addStretch()
        layout.addLayout(loading_layout)
        # Output section
        output_section = QVBoxLayout()
        output_section.setSpacing(5)
        
        # Output header with label and copy button
        output_header = QHBoxLayout()
        output_label = QLabel("Output:")
        # Output buttons
        output_buttons = QHBoxLayout()
        output_buttons.setSpacing(5)
        
        output_copy_btn = QPushButton()
        output_copy_btn.setIcon(QIcon.fromTheme("edit-copy"))
        output_copy_btn.setFixedSize(24, 24)
        output_copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        output_copy_btn.setStyleSheet("background: transparent; border: none;")
        output_copy_btn.clicked.connect(lambda: self.copy_text(self.output_text))
        
        output_clear_btn = QPushButton("×")
        output_clear_btn.setFixedSize(24, 24)
        output_clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        output_clear_btn.setStyleSheet("""
            QPushButton {
                background: #eee;
                border: 1px solid #ddd;
                border-radius: 12px;
                color: #666;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #ff4444;
                color: white;
                border-color: #ff4444;
            }
        """)
        
        output_buttons.addWidget(output_copy_btn)
        output_buttons.addWidget(output_clear_btn)
        
        output_header.addWidget(output_label)
        output_header.addStretch()
        output_header.addLayout(output_buttons)
        
        output_section.addLayout(output_header)
        self.output_text = QTextBrowser()
        
        # Connect output clear button after output_text is created
        output_clear_btn.clicked.connect(self.output_text.clear)
        self.output_text.setOpenExternalLinks(True)
        self.output_text.setMinimumHeight(300)
        output_section.addWidget(self.output_text)
        layout.addLayout(output_section)
        
        # Set up loading animation
        self.loading_timer = QTimer()
        self.loading_timer.timeout.connect(self.update_loading_indicator)
        self.loading_value = 0

    def get_secondary_button_style(self):
        return f"""
            QPushButton {{                
                background-color: #6c757d;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{                
                background-color: #5a6268;
            }}
        """

    def get_api_key_from_user(self):
        while True:
            api_key, ok = QInputDialog.getText(
                self,
                "OpenAI API Key Required",
                "Please enter your OpenAI API Key:\n\nYou can get it from: https://platform.openai.com/api-keys",
                QLineEdit.EchoMode.Password
            )
            
            if not ok:
                # User clicked Cancel
                if QMessageBox.question(
                    self,
                    "Exit Confirmation",
                    "The application requires an API key to function. Do you want to exit?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                ) == QMessageBox.StandardButton.Yes:
                    sys.exit(0)
                continue
            
            if not api_key.strip():
                QMessageBox.warning(self, "Invalid Input", "API key cannot be empty.")
                continue
            
            # Save the API key to .env file
            env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
            with open(env_path, "w") as f:
                f.write(f"OPENAI_API_KEY={api_key}")
            
            self.api_key = api_key
            self.initialize_openai()
            break
    
    def initialize_openai(self):
        self.client = openai.OpenAI(api_key=self.api_key)
    
    def copy_text(self, text_widget):
        text = text_widget.toPlainText()
        QApplication.clipboard().setText(text)

    def generate_random_task(self):
        tasks = [
            "Calculate the sum of sales for the last 3 months",
            "Find the average of values in column A if they are greater than 100",
            "Count how many times 'Yes' appears in column B",
            "Look up a value in table A and return corresponding value from table B",
            "Calculate the percentage of completed tasks (marked as 'Done')",
            "Find the latest date in a range of cells",
            "Calculate the running total of values in column A",
            "Count unique values in a range",
            "Calculate the difference between two dates in days",
            "Find the third highest value in a range"
        ]
        self.input_text.setText(random.choice(tasks))
        if self.mode_combo.currentText() != "Generate Formula":
            self.mode_combo.setCurrentText("Generate Formula")

    def reset_form(self):
        """Reset all form fields and hide loading indicator"""
        self.input_text.clear()
        self.output_text.clear()
        self.loading_indicator.hide()
        self.process_button.setEnabled(True)
        self.loading_timer.stop()
        self.loading_value = 0
        self.loading_indicator.setValue(0)
        self.mode_combo.setCurrentText("Generate Formula")

    def update_loading_indicator(self):
        """Update the loading animation"""
        self.loading_value = (self.loading_value + 5) % 101
        self.loading_indicator.setValue(self.loading_value)

    def process_request(self):
        if not self.api_key:
            QMessageBox.warning(self, "Error", "Please set your OpenAI API key first!")
            return

        mode = self.mode_combo.currentText()
        user_input = self.input_text.toPlainText().strip()

        if not user_input:
            QMessageBox.warning(self, "Error", "Please provide input text!")
            return

        # Show loading indicator and disable process button
        self.loading_indicator.show()
        self.process_button.setEnabled(False)
        self.loading_timer.start(50)

        if mode == "Generate Formula":
            messages = [
                {"role": "system", "content": "You are an Excel formula expert. When asked to generate a formula, provide ONLY the Excel formula without any explanation. The formula should be on a single line and start with '='."},
                {"role": "user", "content": f"Generate an Excel formula for the following requirement: {user_input}"}
            ]
        else:  # Explain Formula
            messages = [
                {"role": "system", "content": "You are an Excel formula expert. Provide a detailed explanation of how the given Excel formula works, breaking down each component and function used."},
                {"role": "user", "content": f"Explain how this Excel formula works in detail: {user_input}"}
            ]

        # Create and start the OpenAI thread
        self.openai_thread = OpenAIThread(self.client, messages)
        self.openai_thread.finished.connect(self.handle_response)
        self.openai_thread.error.connect(self.handle_error)
        self.openai_thread.start()

    def handle_response(self, content):
        # Convert the response to markdown and set as HTML
        markdown_content = markdown(content)
        self.output_text.setHtml(markdown_content)
        
        # Hide loading indicator and enable process button
        self.loading_indicator.hide()
        self.process_button.setEnabled(True)
        self.loading_timer.stop()
        self.loading_value = 0
        self.loading_indicator.setValue(0)

    def handle_error(self, error_message):
        QMessageBox.critical(self, "Error", f"An error occurred: {error_message}")
        # Hide loading indicator and enable process button on error
        self.loading_indicator.hide()
        self.process_button.setEnabled(True)
        self.loading_timer.stop()
        self.loading_value = 0
        self.loading_indicator.setValue(0)

def main():
    app = QApplication(sys.argv)
    window = ExcelFormulaAssistant()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
