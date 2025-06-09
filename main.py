import sys
import os
import glob
import shutil
import csv
import json
import re
import fitz  # PyMuPDF
import atexit
import tempfile
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QLabel, QVBoxLayout, QWidget,
    QMessageBox, QSizePolicy, QStatusBar, QMenuBar, QMenu, QSlider, QHBoxLayout,
    QTabWidget, QGroupBox, QFormLayout, QScrollArea, QSplashScreen, QTextEdit, QFrame,
    QTextBrowser,QScrollArea
)
from PySide6.QtGui import (
    QAction, QPixmap, QImage, QKeySequence, QWheelEvent,
    QFont, QIcon, QColor, QShortcut
)
from PySide6.QtCore import Qt, QTimer, Signal, QSize

# Application constants
APP_NAME = 'Better HR'
CONFIG_FILE = 'config.json'
STATE_FILE = 'session_state.json'

# Default configuration
DEFAULT_CONFIG = {
    "reason_map": {
        "1": "UnsatisfactoryEducation",
        "2": "UnsatisfactoryExperience",
        "3": "Other"
    },
    "hold_folder": "HoldForReview",
    "email_map": {},
    "max_undo": 100,
    "default_zoom": 1.5
}

class PdfViewer(QLabel):
    zoomChanged = Signal(float)
    
    def __init__(self):
        super().__init__()
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumSize(600, 800)
        self.setStyleSheet("""
            background-color: #f0f0f0; 
            border: 1px solid #ccc;
            border-radius: 4px;
        """)
        self.doc = None
        self.page_idx = 0
        self.zoom_level = DEFAULT_CONFIG['default_zoom']
        self.setFocusPolicy(Qt.StrongFocus)
        self.setText("No PDF loaded")
        self.setFont(QFont("Arial", 14))

    def load_pdf(self, path):
        # Close previous document
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass
            finally:
                self.doc = None
                
        if not path or not os.path.exists(path):
            self.clear()
            self.setText("File not found or invalid path")
            return
            
        try:
            self.doc = fitz.open(path)
            self.page_idx = 0
            self.show_page()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unable to load PDF:\n{e}")
            self.setText("Error loading PDF")

    def show_page(self):
        if not self.doc:
            self.setText("No document loaded")
            return
            
        try:
            if 0 <= self.page_idx < len(self.doc):
                page = self.doc.load_page(self.page_idx)
                mat = fitz.Matrix(self.zoom_level, self.zoom_level)
                pix = page.get_pixmap(matrix=mat)
                fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
                pm = QPixmap.fromImage(img)
                self.setPixmap(pm)
                self.setFixedSize(pm.size())  
                # or alternatively: self.adjustSize()
                self.zoomChanged.emit(self.zoom_level)

            else:
                self.setText("Invalid page number")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error showing page:\n{e}")
            self.setText("Error rendering page")

    def set_zoom(self, level):
        """Set zoom level with constraints"""
        self.zoom_level = max(0.5, min(3.0, round(level, 1)))
        if self.doc:
            self.show_page()
            
    def zoom_in(self):
        self.set_zoom(self.zoom_level + 0.1)
        
    def zoom_out(self):
        self.set_zoom(self.zoom_level - 0.1)
        
    def reset_zoom(self):
        self.set_zoom(DEFAULT_CONFIG['default_zoom'])

    def keyPressEvent(self, event):
        if not self.doc:
            return
            
        # Navigation
        if event.key() == Qt.Key_Right:
            if self.page_idx < len(self.doc) - 1:
                self.page_idx += 1
                self.show_page()
        elif event.key() == Qt.Key_Left:
            if self.page_idx > 0:
                self.page_idx -= 1
                self.show_page()
                
        # Zoom controls
        elif event.modifiers() & Qt.ControlModifier:
            if event.key() == Qt.Key_Equal:  # Ctrl+Plus
                self.zoom_in()
            elif event.key() == Qt.Key_Minus:  # Ctrl+Minus
                self.zoom_out()
            elif event.key() == Qt.Key_0:  # Ctrl+0
                self.reset_zoom()
            else:
                super().keyPressEvent(event)
        else:
            super().keyPressEvent(event)
            
    def wheelEvent(self, event):
        """Zoom with Ctrl+Mouse wheel"""
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y() / 120
            self.set_zoom(self.zoom_level + delta * 0.1)
            event.accept()
        else:
            super().wheelEvent(event)

class HelpTab(QWidget):
    def __init__(self, config):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Create a scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        scroll_layout = QVBoxLayout(content)
        scroll_layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel(f"{APP_NAME} - Help & Documentation")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(18)
        title.setFont(title_font)
        title.setStyleSheet("color: white; margin-bottom: 20px;")
        scroll_layout.addWidget(title)
        
        # Introduction
        intro = QTextEdit()
        intro.setReadOnly(True)
        intro.setHtml(f"""
        <p>Welcome to <b>{APP_NAME}</b> - a tool designed to streamline your CV review process with transparent 
        rejection reasons and candidate feedback capabilities.</p>
        
        <p>This tool helps you efficiently categorize CVs, track decisions, and generate reports for 
        candidate feedback while maintaining a complete audit trail of your decisions.</p>
        """)
        intro.setStyleSheet("background: transparent; border: none; font-size: 12pt;")
        scroll_layout.addWidget(intro)
        
        # Key Bindings Section
        keys_group = QGroupBox("Key Bindings")
        keys_group.setStyleSheet("""
            QGroupBox { 
                font-weight: bold; 
                font-size: 12pt;
                margin-top: 15px;
            }
        """)
        keys_layout = QFormLayout()
        keys_layout.setVerticalSpacing(10)
        keys_layout.setHorizontalSpacing(20)
        
        # Navigation
        keys_layout.addRow(QLabel("<span style='font-size: 11pt; font-weight: bold; color: white;'>Navigation</span>"), QLabel(""))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>→ Right Arrow</span>"), QLabel("<span style='font-size: 11pt;'>Next page</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>← Left Arrow</span>"), QLabel("<span style='font-size: 11pt;'>Previous page</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>1, 2, 3</span>"), QLabel("<span style='font-size: 11pt;'>Reject with reason</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+Z</span>"), QLabel("<span style='font-size: 11pt;'>Undo last action</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+B</span>"), QLabel("<span style='font-size: 11pt;'>Hold for review</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+H</span>"), QLabel("<span style='font-size: 11pt;'>Show undo history</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Esc</span>"), QLabel("<span style='font-size: 11pt;'>Return to main view</span>"))
        
        # Zoom Controls
        keys_layout.addRow(QLabel("<span style='font-size: 11pt; font-weight: bold; color: white;'>Zoom Controls</span>"), QLabel(""))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl++</span>"), QLabel("<span style='font-size: 11pt;'>Zoom in</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+-</span>"), QLabel("<span style='font-size: 11pt;'>Zoom out</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+0</span>"), QLabel("<span style='font-size: 11pt;'>Reset zoom</span>"))
        keys_layout.addRow(QLabel("<span style='font-size: 11pt;'>Ctrl+Mouse Wheel</span>"), QLabel("<span style='font-size: 11pt;'>Adjust zoom</span>"))
        
        keys_group.setLayout(keys_layout)
        scroll_layout.addWidget(keys_group)
        
        # Rejection Reasons
        reasons_group = QGroupBox("Rejection Reasons")
        reasons_group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; margin-top: 15px; }")
        reasons_layout = QFormLayout()
        reasons_layout.setVerticalSpacing(8)
        reasons_layout.setHorizontalSpacing(20)
        
        for key, reason in config['reason_map'].items():
            reasons_layout.addRow(
                QLabel(f"<span style='font-size: 11pt;'>Key {key}</span>"), 
                QLabel(f"<span style='font-size: 11pt;'>{reason}</span>")
            )
        
        reasons_group.setLayout(reasons_layout)
        scroll_layout.addWidget(reasons_group)
        
        # Workflow Guide
        workflow_group = QGroupBox("Workflow Guide")
        workflow_group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; margin-top: 15px; }")
        workflow_layout = QVBoxLayout()
        workflow_layout.setSpacing(10)
        
        steps = [
            "1. <b>Open Folder</b>: Select a folder containing CVs (PDF files)",
            "2. <b>Navigate</b>: Use arrow keys to browse CV pages",
            "3. <b>Categorize</b>:",
            "   - Press 1, 2, or 3 to reject with specific reason",
            "   - Press Ctrl+B to hold for review",
            "4. <b>Correct Mistakes</b>: Use Ctrl+Z to undo actions",
            "5. <b>Export Data</b>: Generate CSV for rejected candidates",
            "6. <b>Import Emails</b>: Match candidate names to email addresses"
        ]
        
        for step in steps:
            label = QLabel(f"<span style='font-size: 11pt;'>{step}</span>")
            label.setStyleSheet("margin-left: 15px;" if step.startswith("   -") else "")
            label.setTextFormat(Qt.RichText)
            workflow_layout.addWidget(label)
        
        workflow_group.setLayout(workflow_layout)
        scroll_layout.addWidget(workflow_group)
        
        # Tips
        tips_group = QGroupBox("Tips & Best Practices")
        tips_group.setStyleSheet("QGroupBox { font-weight: bold; font-size: 12pt; margin-top: 15px; }")
        tips_layout = QVBoxLayout()
        tips_layout.setSpacing(8)
        
        tips = [
            "• Use consistent naming for CV files to improve email matching accuracy",
            "• Regularly export your rejection list to provide timely candidate feedback",
            "• The undo history (Ctrl+H) shows your last 10 actions",
            "• Zoom controls help examine details in complex CV layouts",
            "• You can import multiple email lists throughout the session"
        ]
        
        for tip in tips:
            label = QLabel(f"<span style='font-size: 11pt;'>{tip}</span>")
            label.setTextFormat(Qt.RichText)
            tips_layout.addWidget(label)
        
        tips_group.setLayout(tips_layout)
        scroll_layout.addWidget(tips_group)
        
        # Add spacer
        scroll_layout.addStretch()
        
        # Set scroll content
        scroll.setWidget(content)
        layout.addWidget(scroll)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.statusBar().showMessage('Ready')
        self.setStyleSheet("""
            QMainWindow {
                background-color: grey;
            }
            QTabWidget::pane {
                border: none;
            }
            QGroupBox {
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 1em;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
                font-weight: bold;
            }
            QSlider::groove:horizontal {
                border: 1px solid #bbb;
                background: white;
                height: 8px;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #3498db;
                border: 1px solid #2980b9;
                width: 16px;
                margin: -4px 0;
                border-radius: 8px;
            }
            QLabel {
                font-size: 11pt;
            }
        """)
        
        # Initialize state
        self.cv_list = []
        self.current_index = -1
        self.undo_stack = []
        self.base_folder = ''
        self.cand_email_map = {}
        self.state_file = STATE_FILE
        self.temp = None
        
        # Load configuration
        self.config = self.load_config()
        
        # Setup UI
        self.setup_ui()
        
        # Load session state if available
        self.load_session_state()
        self.update_ui()
        
        # Setup auto-save timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.save_session_state)
        self.timer.start(30000)  # 30 seconds
        
        # Register cleanup
        atexit.register(self.cleanup_temp)
        
    def cleanup_temp(self):
        """Clean up temporary files on exit"""
        if self.temp and os.path.exists(self.temp):
            try:
                os.remove(self.temp)
            except Exception:
                pass
                
    def load_config(self):
        """Load configuration with fallback to defaults"""
        cfg = DEFAULT_CONFIG.copy()
        
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # Merge with defaults
                    cfg.update(loaded)
            except Exception as e:
                QMessageBox.warning(self, "Config Error", f"Error loading config:\n{e}")
        
        # Validate and save config
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Config Error", f"Couldn't save config:\n{e}")
            
        return cfg
        
    def setup_ui(self):
        """Initialize the user interface"""
        # Create main tab widget
        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        self.setCentralWidget(tabs)
        
        # Create viewer tab
        vt = QWidget()
        vl = QVBoxLayout(vt)
        vl.setContentsMargins(10, 10, 10, 10)
        
        # Count label
        self.count_lbl = QLabel('0')
        self.count_lbl.setAlignment(Qt.AlignCenter)
        self.count_lbl.setStyleSheet("""
            font-size: 14px; 
            font-weight: bold; 
            padding: 8px;
            background-color: #e3f2fd;
            border-radius: 4px;
            margin-bottom: 10px;
        """)
        vl.addWidget(self.count_lbl)
        
        # PDF viewer
        self.viewer = PdfViewer()
        # Make a scroll area around the PDF viewer
        pdf_scroll = QScrollArea()
        pdf_scroll.setWidgetResizable(False)    # ← disable automatic stretching
        pdf_scroll.setAlignment(Qt.AlignCenter)
        pdf_scroll.setWidget(self.viewer)


        vl.addWidget(pdf_scroll, 1)

        
        # Zoom controls
        zoom_frame = QFrame()
        zoom_frame.setFrameShape(QFrame.StyledPanel)
        zoom_frame.setStyleSheet("background: white; border-radius: 4px;")
        hl = QHBoxLayout(zoom_frame)
        hl.setContentsMargins(10, 5, 10, 5)
        
        hl.addWidget(QLabel('Zoom:'))
        
        self.zs = QSlider(Qt.Horizontal)
        self.zs.setRange(50, 300)
        self.zs.setValue(int(self.config['default_zoom'] * 100))
        self.zs.valueChanged.connect(lambda v: self.viewer.set_zoom(v/100))
        hl.addWidget(self.zs, 1)
        
        self.zl = QLabel(f"{int(self.config['default_zoom'] * 100)}%")
        self.zl.setMinimumWidth(40)
        self.zl.setAlignment(Qt.AlignCenter)
        self.zl.setStyleSheet("font-weight: bold;")
        hl.addWidget(self.zl)
        
        # Connect zoom change signal
        self.viewer.zoomChanged.connect(lambda z: self.zl.setText(f"{int(z * 100)}%"))
        
        vl.addWidget(zoom_frame)
        tabs.addTab(vt, "CV Viewer")
        
        # Help tab
        tabs.addTab(HelpTab(self.config), "Help")
        
        # Create toolbar with icons
        tb = self.addToolBar('Main Tools')
        tb.setIconSize(QSize(24, 24))
        tb.setMovable(False)
        tb.setStyleSheet("""
            QToolBar {
                background: black;
                border-bottom: 1px solid #dee2e6;
                padding: 5px;
            }
            QToolButton {
                padding: 5px;
            }
        """)
        
        # Create actions
        actions = [
            ('Open Folder', 'folder', self.open_folder, "Ctrl+O"),
            ('Import Emails', 'mail', self.import_emails, "Ctrl+I"),
            ('Undo', 'undo', self.undo, "Ctrl+Z"),
            ('Hold', 'pause', self.hold, "Ctrl+B"),
            ('Export CSV', 'file', self.export_csv, "Ctrl+E"),
            ('History', 'history', self.show_history, "Ctrl+H"),
            ('Help', 'help', lambda: tabs.setCurrentIndex(1), "F1")
        ]
        
        for name, icon_name, fn, shortcut in actions:
            a = QAction(QIcon(), name, self)  # Placeholder for icons
            a.triggered.connect(fn)
            a.setShortcut(QKeySequence(shortcut))
            tb.addAction(a)
        
        # Reason shortcuts
        for k in self.config['reason_map']:
            QShortcut(QKeySequence(k), self, activated=lambda k=k: self.move_current(k))
            
        # Escape key to return to main view
        QShortcut(QKeySequence(Qt.Key_Escape), self, activated=lambda: tabs.setCurrentIndex(0))
        
    def open_folder(self):
        """Open a folder containing CVs"""
        d = QFileDialog.getExistingDirectory(self, "Select CV Folder")
        if not d:
            return
            
        # Create splash screen
        splash = QSplashScreen(QPixmap(400, 200))
        splash.showMessage("Loading CVs...\nPlease wait", 
                          Qt.AlignCenter | Qt.AlignBottom, Qt.black)
        splash.show()
        QApplication.processEvents()
        
        try:
            self.base_folder = d
            self.cv_list = sorted(glob.glob(os.path.join(d, '*.pdf')))
            self.current_index = 0 if self.cv_list else -1
            self.undo_stack.clear()
            
            if not self.cv_list:
                QMessageBox.information(self, "No PDFs", "No PDF files found in selected folder")
                
            self.save_session_state()
            self.update_ui()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load folder:\n{e}")
        finally:
            splash.finish(self)
            
    def update_ui(self):
        """Update the UI to reflect current state"""
        if not self.cv_list or self.current_index < 0 or self.current_index >= len(self.cv_list):
            self.viewer.setText("No CVs loaded")
            self.count_lbl.setText('0')
            self.statusBar().showMessage('No CVs loaded')
            return
            
        try:
            p = self.cv_list[self.current_index]
            self.viewer.load_pdf(p)
            self.count_lbl.setText(f"{len(self.cv_list)}")
            
            # Get candidate email if available
            name = os.path.splitext(os.path.basename(p))[0]
            email = self.cand_email_map.get(name) or self.config['email_map'].get(name, '')
            
            # Update status
            msg = f"Evaluating: {os.path.basename(p)}"
            if email:
                msg += f" | Email: {email}"
            self.statusBar().showMessage(msg)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to update UI:\n{e}")
            
    def import_emails(self):
        """Import candidate emails from CSV"""
        if not self.base_folder:
            QMessageBox.information(self, "Info", "Please open a folder first")
            return
            
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Candidate Emails", "", 
            "CSV Files (*.csv);;All Files (*)"
        )
        
        if not path:
            return
            
        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                if 'name' not in reader.fieldnames or 'email' not in reader.fieldnames:
                    QMessageBox.warning(
                        self, "Invalid Format", 
                        "CSV must contain 'name' and 'email' columns"
                    )
                    return
                    
                new_emails = 0
                for row in reader:
                    # Normalize name for matching
                    name_key = self.normalize_name(row['name'])
                    email = row['email'].strip()
                    
                    if name_key and email:
                        # Update session cache
                        self.cand_email_map[name_key] = email
                        # Update persistent config
                        self.config['email_map'][name_key] = email
                        new_emails += 1
                
                # Save updated config
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=2)
                    
                QMessageBox.information(
                    self, "Import Complete", 
                    f"Imported {new_emails} email addresses\n"
                    f"Total emails in system: {len(self.config['email_map'])}"
                )
                
                # Update UI to show current candidate's email
                self.update_ui()
                self.save_session_state()
                
        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import emails:\n{e}")
            
    def normalize_name(self, name):
        """Create consistent name key for matching"""
        # Remove non-alphanumeric characters (keep spaces)
        clean = re.sub(r'[^\w\s]', '', name, flags=re.UNICODE)
        # Convert to lowercase
        clean = clean.lower()
        # Remove extra spaces
        clean = re.sub(r'\s+', ' ', clean).strip()
        return clean
        
    def unique_dest(self, directory, name):
        """Generate a unique filename in destination directory"""
        base, ext = os.path.splitext(name)
        candidate = name
        counter = 1
        
        while os.path.exists(os.path.join(directory, candidate)):
            candidate = f"{base}_{counter}{ext}"
            counter += 1
            
        return candidate
        
    def move_current(self, reason_key):
        
        """Move current CV to rejection folder"""
        if not self.cv_list or self.current_index < 0 or self.current_index >= len(self.cv_list):
            return
            
        if reason_key not in self.config['reason_map']:
            QMessageBox.warning(self, "Invalid Reason", f"Unknown reason key: {reason_key}")
            return
            
        src = self.cv_list[self.current_index]
        if not os.path.exists(src):
            QMessageBox.warning(self, "File Missing", f"File not found:\n{src}")
            return
            
        folder = self.config['reason_map'][reason_key]
                # Release the file lock by closing the open PDF
        try:
            if self.viewer.doc:
                self.viewer.doc.close()
        except Exception:
            pass
        finally:
            self.viewer.doc = None
            self.viewer.clear()

        dest_dir = os.path.join(self.base_folder, 'Rejected', folder)
        
        try:
            os.makedirs(dest_dir, exist_ok=True)
            name = os.path.basename(src)
            unique_name = self.unique_dest(dest_dir, name)
            dest = os.path.join(dest_dir, unique_name)
            
            # Move file and record operation
            shutil.move(src, dest)
            operation = {
                'src': src,
                'dest': dest,
                'type': 'reject',
                'reason': reason_key,
                'filename': name,
                'position': self.current_index
            }
            self.undo_stack.append(operation)
            
            # Enforce max undo limit
            if len(self.undo_stack) > self.config['max_undo']:
                self.undo_stack.pop(0)
            
            self.cv_list.pop(self.current_index)
            
            # Adjust current index
            if not self.cv_list:
                self.current_index = -1
            elif self.current_index >= len(self.cv_list):
                self.current_index = len(self.cv_list) - 1
                
            self.save_session_state()
            self.update_ui()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to move file:\n{e}")
            
    def undo(self):
        """Undo the last operation"""
        if not self.undo_stack:
            QMessageBox.information(self, "Undo", "Nothing to undo")
            return
            
        # Get last operation
        operation = self.undo_stack.pop()
        src = operation['src']
        dest = operation['dest']
        
        try:
            # Ensure original directory exists
            os.makedirs(os.path.dirname(src), exist_ok=True)
            
            # Check if destination file still exists
            if not os.path.exists(dest):
                QMessageBox.warning(self, "File Missing", 
                    f"Original file not found:\n{dest}\n\n"
                    "It may have been moved or deleted outside the application.")
                return
                
            # Check if source location is available
            if os.path.exists(src):
                response = QMessageBox.question(
                    self, "File Conflict",
                    f"A file already exists at:\n{src}\n\n"
                    "Do you want to overwrite it?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if response != QMessageBox.Yes:
                    return
                # Remove existing file
                try:
                    os.remove(src)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to remove file:\n{e}")
                    return
                    
            # Move file back
            shutil.move(dest, src)
            
            # Add back to file list at original position
            insert_index = operation['position']
            if insert_index > len(self.cv_list):
                insert_index = len(self.cv_list)
            self.cv_list.insert(insert_index, src)
            
            # Set current index to the restored file
            self.current_index = insert_index
            
            self.save_session_state()
            self.update_ui()
            
            # Show confirmation
            self.statusBar().showMessage(f"Undo successful: Restored {operation['filename']}", 3000)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to undo:\n{e}")
            
    def show_history(self):
        """Show undo history"""
        if not self.undo_stack:
            QMessageBox.information(self, "Undo History", "No operations in history")
            return
            
        history = "Recent operations (most recent last):\n\n"
        for i, op in enumerate(self.undo_stack[-10:], 1):  # Show last 10
            op_type = "Rejected" if op['type'] == 'reject' else "Held"
            reason = self.config['reason_map'].get(op['reason'], op['reason']) if op['type'] == 'reject' else "Review"
            history += f"{i}. {op_type}: {op['filename']}\n   Reason: {reason}\n\n"
            
        QMessageBox.information(self, "Undo History", history)
        
    def hold(self):
        """Move current CV to hold folder"""
        if not self.cv_list or self.current_index < 0 or self.current_index >= len(self.cv_list):
            return
            
        src = self.cv_list[self.current_index]
        if not os.path.exists(src):
            QMessageBox.warning(self, "File Missing", f"File not found:\n{src}")
            return

        # 1) Close the PDF in the viewer to release the file handle
        try:
            if self.viewer.doc:
                self.viewer.doc.close()
        except Exception:
            pass
        finally:
            self.viewer.doc = None
            self.viewer.clear()

        hold_dir = os.path.join(self.base_folder, self.config['hold_folder'])
        
        try:
            os.makedirs(hold_dir, exist_ok=True)
            name = os.path.basename(src)
            unique_name = self.unique_dest(hold_dir, name)
            dest = os.path.join(hold_dir, unique_name)
            
            # Move file and record operation
            shutil.move(src, dest)
            operation = {
                'src': src,
                'dest': dest,
                'type': 'hold',
                'filename': name,
                'position': self.current_index
            }
            self.undo_stack.append(operation)
            
            # Enforce max undo limit
            if len(self.undo_stack) > self.config['max_undo']:
                self.undo_stack.pop(0)
            
            self.cv_list.pop(self.current_index)
            
            # Adjust current index
            if not self.cv_list:
                self.current_index = -1
            elif self.current_index >= len(self.cv_list):
                self.current_index = len(self.cv_list) - 1
                
            self.save_session_state()
            self.update_ui()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to hold file:\n{e}")
            
    def export_csv(self):
        """Export rejection data to CSV"""
        if not self.base_folder:
            QMessageBox.information(self, "Info", "Open a folder first.")
            return
            
        rows = []
        rejected_dir = os.path.join(self.base_folder, 'Rejected')
        
        if not os.path.exists(rejected_dir):
            QMessageBox.information(self, "Export CSV", "No rejected CVs to export.")
            return
            
        for key, folder in self.config['reason_map'].items():
            path = os.path.join(rejected_dir, folder)
            if os.path.isdir(path):
                for fname in os.listdir(path):
                    if fname.lower().endswith('.pdf'):
                        # Get base name without extension
                        base_name = os.path.splitext(fname)[0]
                        
                        # Find email from multiple sources
                        email = (
                            self.cand_email_map.get(base_name) or
                            self.config['email_map'].get(base_name) or
                            ""
                        )
                        
                        rows.append({
                            'Filename': fname,
                            'ReasonKey': key,
                            'ReasonText': folder,
                            'Email': email
                        })
        
        if not rows:
            QMessageBox.information(self, "Export CSV", "No rejected CVs to export.")
            return
            
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save CSV", 'rejection_list.csv', "CSV Files (*.csv)"
        )
        
        if not out_path:
            return
            
        # Ensure .csv extension
        if not out_path.lower().endswith('.csv'):
            out_path += '.csv'
            
        try:
            with open(out_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Filename', 'ReasonKey', 'ReasonText', 'Email'])
                writer.writeheader()
                writer.writerows(rows)
                
            QMessageBox.information(self, "Export Complete", 
                f"CSV exported with {len(rows)} entries\n"
                f"Emails found: {sum(1 for row in rows if row['Email'])}/{len(rows)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save CSV:\n{e}")
            
    def save_session_state(self):
        """Save current session state to file"""
        if not self.base_folder:
            return
            
        # Create temporary file first to prevent corruption
        if not self.temp:
            self.temp = tempfile.mktemp(suffix='.tmp', prefix='cv_sorter_')
            
        state = {
            'base_folder': self.base_folder,
            'cv_list': self.cv_list,
            'current_index': self.current_index,
            'undo_stack': self.undo_stack,
            'cand_email_map': self.cand_email_map,
            'viewer_zoom': self.viewer.zoom_level
        }
        
        try:
            # Save to temp file first
            with open(self.temp, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2)
                
            # Atomically replace existing state file
            if os.path.exists(self.state_file):
                os.remove(self.state_file)
            shutil.move(self.temp, self.state_file)
            
        except Exception as e:
            QMessageBox.warning(self, "State Error", f"Failed to save session state: {e}")
            
    def load_session_state(self):
        """Load session state from file if available"""
        if not os.path.exists(self.state_file):
            return
            
        try:
            with open(self.state_file, 'r', encoding='utf-8') as f:
                state = json.load(f)
                
            # Basic validation
            if not isinstance(state, dict):
                return
                
            # Restore state
            self.base_folder = state.get('base_folder', '')
            self.cv_list = state.get('cv_list', [])
            self.current_index = state.get('current_index', -1)
            self.undo_stack = state.get('undo_stack', [])
            self.cand_email_map = state.get('cand_email_map', {})
            
            # Validate current index
            if self.current_index >= len(self.cv_list):
                self.current_index = max(0, len(self.cv_list) - 1) if self.cv_list else -1
                
            # Restore zoom level if available
            zoom_level = state.get('viewer_zoom', DEFAULT_CONFIG["default_zoom"])
            self.viewer.zoom_level = max(0.5, min(3.0, zoom_level))
            self.zs.setValue(int(self.viewer.zoom_level * 100))
            
            if self.base_folder:
                QMessageBox.information(self, "Session Restored", 
                                       "Previous session state has been restored")
            
        except Exception as e:
            QMessageBox.warning(self, "State Error", f"Failed to load session state: {e}")
            
    def closeEvent(self, event):
        """Handle application close"""
        self.save_session_state()
        self.timer.stop()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.showMaximized()
    sys.exit(app.exec_())