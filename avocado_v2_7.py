import sys
import os
from PyQt5.QtCore import Qt
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()  # automatically looks for .env in the current dir
import csv
import time
import requests
import unicodedata
import re
from pathlib import Path

# ✅ Now import everything else
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QTextCursor, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QTextEdit, QFileDialog, QMessageBox, QProgressBar, QGroupBox, 
    QGridLayout, QFrame, QCheckBox, QVBoxLayout, QHBoxLayout, QTabWidget,
    QSizePolicy, QSpacerItem
)

# ✅ Nothing PyQt5-related (e.g., QPixmap, QFont) must appear before the above block


# Full logo path (safe across platforms)
logo_path = os.path.join(os.path.dirname(__file__), "avocado_logo.png")

# Declare logo_pixmap as global variable (will be initialized in main)
logo_pixmap = None

class WorkerThread(QThread):
    """Worker thread for OCLC operations without blocking UI"""
    progress_update = pyqtSignal(str)
    progress_value = pyqtSignal(int)
    workflow_complete = pyqtSignal(str, int, int, int)
    workflow_error = pyqtSignal(str)

    
    def __init__(self, operation_type, app_instance):
        super().__init__()
        self.operation_type = operation_type
        self.app = app_instance
        self.should_stop = False

    def run(self):
        """Execute operation in separate thread"""
        try:
            if self.operation_type == "complete_workflow":
                self.run_complete_workflow()
        except Exception as e:
            self.workflow_error.emit(f"Unexpected error: {str(e)}")
    
    def stop(self):
        """Stop operation"""
        self.should_stop = True
    
    def run_complete_workflow(self):
        """Execute complete workflow"""
        try:
            self.progress_update.emit("AVOCADO Professional - Complete Workflow Started")
            self.progress_update.emit("=" * 60)
            
            # Phase 1: Authentication
            self.progress_update.emit("Phase 1: Authenticating with OCLC...")
            self.progress_value.emit(5)
            
            if not self.app.fetch_oclc_token():
                self.workflow_error.emit("Failed to authenticate with OCLC API")
                return
                
            self.progress_update.emit("OCLC authentication successful")
            self.progress_value.emit(10)
            
            if self.should_stop:
                return
            
            # Phase 2: Read CSV file
            self.progress_update.emit("Phase 2: Processing CSV file...")
            
            try:
                with open(self.app.input_file, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    
                    # Validate headers
                    expected_headers = {'OCLC #', 'Author', 'Title'}
                    if not expected_headers.issubset(set(reader.fieldnames)):
                        self.workflow_error.emit(f"CSV must contain columns: {', '.join(expected_headers)}")
                        return
                    
                    # Read books
                    raw_books = list(reader)
                    
                # Filter valid books
                books = []
                for book in raw_books:
                    if any(v.strip() for v in book.values() if v):
                        books.append(book)
                
                if not books:
                    self.workflow_error.emit("No valid books found in CSV")
                    return
                    
            except Exception as e:
                self.workflow_error.emit(f"Error reading CSV: {str(e)}")
                return
            
            total_books = len(books)
            self.progress_update.emit(f"Found {total_books} books to process")
            self.progress_value.emit(15)
            
            if self.should_stop:
                return
            
            # Phase 3: Search for OCLC numbers
            self.progress_update.emit("Phase 3: Searching for OCLC numbers...")
            
            oclc_results = []
            found_oclc = 0
            
            for i, book in enumerate(books):
                if self.should_stop:
                    return
                    
                existing_oclc = book.get("OCLC #", "").strip()
                title = book.get("Title", "").strip()
                author = book.get("Author", "").strip()
                
                display_title = title[:40] + "..." if len(title) > 40 else title
                self.progress_update.emit(f"Processing {i+1}/{total_books}: {display_title}")
                
                if existing_oclc:
                    self.progress_update.emit(f"OCLC already present: {existing_oclc}")
                    book["OCLC #"] = existing_oclc
                    found_oclc += 1
                elif title and author:
                    oclc_number = self.app.search_oclc(title, author)
                    book["OCLC #"] = oclc_number or ""
                    if oclc_number:
                        self.progress_update.emit(f"OCLC found: {oclc_number}")
                        found_oclc += 1
                    else:
                        self.progress_update.emit("No OCLC found")
                else:
                    self.progress_update.emit("Insufficient data for search")
                    book["OCLC #"] = ""
                
                oclc_results.append(book)
                
                # Progress 15-50% for OCLC search
                progress = 15 + int((i + 1) / total_books * 35)
                self.progress_value.emit(progress)
                
                time.sleep(0.3)  # Rate limiting
            
            self.progress_update.emit(f"Phase 3 complete: {found_oclc}/{total_books} OCLC numbers found")
            
            if self.should_stop:
                return
            
            # Phase 4: Download complete metadata
            self.progress_update.emit("Phase 4: Downloading complete metadata...")
            self.progress_value.emit(50)
            
            # Extract valid OCLC numbers
            oclc_numbers = []
            for result in oclc_results:
                oclc_num = result.get("OCLC #", "").strip()
                if oclc_num:
                    oclc_numbers.append((oclc_num, result))
            
            if not oclc_numbers:
                self.progress_update.emit("No OCLC numbers to download metadata")
                # Save basic results
                self.save_basic_results(oclc_results)
                return
            
            # Process metadata
            complete_records = []
            metadata_complete = 0
            
            for i, (oclc_num, original_book) in enumerate(oclc_numbers):
                if self.should_stop:
                    return
                    
                self.progress_update.emit(f"Downloading metadata {i+1}/{len(oclc_numbers)}: OCLC {oclc_num}")
                
                try:
                    metadata = self.app.fetch_metadata_json(oclc_num)
                    record = self.app.parse_complete_record(metadata or {}, oclc_num, original_book)
                    complete_records.append(record)
                    
                    if record.get("Title") and record.get("Publisher"):
                        self.progress_update.emit(f"Complete: {record['Title'][:30]}...")
                        metadata_complete += 1
                    else:
                        self.progress_update.emit("Partial metadata")
                        
                except Exception as e:
                    self.progress_update.emit(f"Error in metadata: {str(e)}")
                    # Create basic record on error
                    record = self.app.create_basic_record(original_book, oclc_num)
                    complete_records.append(record)
                
                # Progress 50-90% for metadata
                progress = 50 + int((i + 1) / len(oclc_numbers) * 40)
                self.progress_value.emit(progress)
                
                time.sleep(0.3)  # Rate limiting
            
            # Phase 5: Save results
            self.progress_update.emit("Phase 5: Saving final file...")
            self.progress_value.emit(90)
            
            output_file = self.save_complete_results(complete_records)
            
            self.progress_value.emit(100)
            self.progress_update.emit("=" * 60)
            self.progress_update.emit("COMPLETE WORKFLOW FINISHED!")
            self.progress_update.emit(f"File: {Path(output_file).name}")
            self.progress_update.emit(f"Total: {len(complete_records)} | OCLC: {found_oclc} | Metadata: {metadata_complete}")
            self.progress_update.emit("=" * 60)
            
            # Emit completion signal
            self.workflow_complete.emit(output_file, len(complete_records), found_oclc, metadata_complete)
            
        except Exception as e:
            self.workflow_error.emit(f"Workflow error: {str(e)}")
    
    def save_basic_results(self, results):
        """Save basic results without metadata"""
        input_name = Path(self.app.input_file).stem
        timestamp = int(time.time())
        output_file = Path(self.app.output_dir) / f"{input_name}_avocado_basic_{timestamp}.csv"
        
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            if results:
                fieldnames = results[0].keys()
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(results)
        
        self.workflow_complete.emit(str(output_file), len(results), 0, 0)
    
    def save_complete_results(self, records):
        """Save complete results with metadata"""
        input_name = Path(self.app.input_file).stem
        timestamp = int(time.time())
        output_file = Path(self.app.output_dir) / f"{input_name}_avocado_professional_{timestamp}.csv"
        
        # Field order for consistent output
        fieldnames = [
            "OCLC #", "Title", "Creator", "Contributor", "Publisher", 
            "Date", "Language", "Subjects", "Type", "Format", 
            "ISBN", "ISSN", "Edition", "URL"
        ]
        
        with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(records)
        
        return str(output_file)

class AvocadoProfessional(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # State variables
        self.wskey = ""
        self.wssecret = ""
        self.input_file = ""
        self.output_dir = str(Path.home() / "Downloads")
        self.access_token = None
        self.worker_thread = None
        
        # Load credentials
        self.load_credentials()
        
        # Initialize UI
        self.init_ui()
        self.apply_professional_styles()
        
    def init_ui(self):
        """Initialize professional interface"""
        self.setWindowTitle("AVOCADO v2.7 - Archivo Venezuela OCLC & Data Organizer")
        self.setGeometry(100, 100, 1200, 800)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Professional header
        self.create_professional_header(main_layout)
        
        # Content area
        self.create_content_area(main_layout)
        
    def create_professional_header(self, main_layout):
        """Create clean professional header"""
        header = QFrame()
        header.setFixedHeight(80)
        header.setObjectName("professionalHeader")
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(25, 0, 25, 0)
        
        # Title section
        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(2)
        
        title_label = QLabel("AVOCADO v2.7")
        title_label.setObjectName("headerTitle")
        
        subtitle_label = QLabel("Archivo Venezuela OCLC & Data Organizer")
        subtitle_label.setObjectName("headerSubtitle")
        
        features_label = QLabel("Complete Workflow | Advanced Metadata | Venezuelan Archives")
        features_label.setObjectName("headerFeatures")
        
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        title_layout.addWidget(features_label)
        title_layout.addStretch()
        
        # Logo
        logo_label = QLabel()
        if logo_pixmap and not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap)
        else:
            logo_label.setText("LOGO")
        
        # Status indicator
        self.status_indicator = QLabel("●")
        self.status_indicator.setObjectName("statusIndicator")
        self.update_connection_status(False)
        
        # Add to header
        header_layout.addWidget(title_widget)
        header_layout.addStretch()
        header_layout.addWidget(self.status_indicator)
        header_layout.addWidget(logo_label)
        
        main_layout.addWidget(header)
    
    def create_content_area(self, main_layout):
        """Create content area with professional tabs"""
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(25, 25, 25, 25)
        
        # Professional tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("professionalTabs")
        content_layout.addWidget(self.tab_widget)
        
        # Create tabs
        self.create_setup_tab()
        self.create_complete_workflow_tab()
        self.create_advanced_tab()
        self.create_about_tab()
        
        main_layout.addWidget(content_widget)
    
    def create_setup_tab(self):
        """Professional setup tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(25)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # Welcome section
        welcome_group = QGroupBox("Welcome to AVOCADO Professional")
        welcome_group.setObjectName("welcomeGroup")
        welcome_layout = QVBoxLayout(welcome_group)
        
        welcome_text = QLabel(
            "AVOCADO Professional provides a complete and integrated workflow for managing "
            "bibliographic metadata for Venezuelan archives.\n\n"
            "Connect with OCLC WorldCat to find, download, and organize bibliographic data "
            "with professional precision and efficiency."
        )
        welcome_text.setWordWrap(True)
        welcome_text.setObjectName("welcomeText")
        welcome_layout.addWidget(welcome_text)
        
        # Quick guide
        guide_text = QLabel(
            "Quick Start Guide:\n"
            "1. Configure your OCLC credentials below\n"
            "2. Download the template CSV and fill in your book list\n"
            "3. Go to 'Complete Workflow' tab for automatic processing\n"
            "4. Get complete metadata with a single click"
        )
        guide_text.setObjectName("guideText")
        welcome_layout.addWidget(guide_text)
        
        # Credentials section
        cred_group = QGroupBox("OCLC WorldCat API Credentials")
        cred_group.setObjectName("credentialsGroup")
        cred_layout = QGridLayout(cred_group)
        cred_layout.setSpacing(15)
        
        # WSKey
        cred_layout.addWidget(QLabel("WSKey:"), 0, 0)
        self.wskey_input = QLineEdit()
        self.wskey_input.setEchoMode(QLineEdit.Password)
        self.wskey_input.setPlaceholderText("Enter your OCLC WSKey")
        self.wskey_input.setText(self.wskey)
        self.wskey_input.textChanged.connect(self.on_credentials_changed)
        cred_layout.addWidget(self.wskey_input, 0, 1)
        
        # Secret
        cred_layout.addWidget(QLabel("Secret:"), 1, 0)
        self.wssecret_input = QLineEdit()
        self.wssecret_input.setEchoMode(QLineEdit.Password)
        self.wssecret_input.setPlaceholderText("Enter your OCLC Secret")
        self.wssecret_input.setText(self.wssecret)
        self.wssecret_input.textChanged.connect(self.on_credentials_changed)
        cred_layout.addWidget(self.wssecret_input, 1, 1)
        
        # Save credentials checkbox
        self.save_creds_checkbox = QCheckBox("Save credentials securely (.env file)")
        self.save_creds_checkbox.setChecked(bool(self.wskey and self.wssecret))
        self.save_creds_checkbox.setObjectName("saveCredsCheckbox")
        cred_layout.addWidget(self.save_creds_checkbox, 2, 0, 1, 2)
        
        # Test button
        self.test_btn = QPushButton("Test Connection")
        self.test_btn.setObjectName("testButton")
        self.test_btn.clicked.connect(self.test_connection)
        cred_layout.addWidget(self.test_btn, 3, 1)
        
        # File settings
        file_group = QGroupBox("File Settings")
        file_group.setObjectName("fileGroup")
        file_layout = QGridLayout(file_group)
        file_layout.setSpacing(15)
        
        file_layout.addWidget(QLabel("Output Directory:"), 0, 0)
        self.output_dir_input = QLineEdit(self.output_dir)
        self.output_dir_input.textChanged.connect(lambda: setattr(self, 'output_dir', self.output_dir_input.text()))
        file_layout.addWidget(self.output_dir_input, 0, 1)
        
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_output_dir)
        file_layout.addWidget(browse_btn, 0, 2)
        
        # Template download
        template_btn = QPushButton("Download AVOCADO Template")
        template_btn.setObjectName("templateButton")
        template_btn.clicked.connect(self.download_template)
        file_layout.addWidget(template_btn, 1, 0, 1, 3)
        
        # Add all groups
        layout.addWidget(welcome_group)
        layout.addWidget(cred_group)
        layout.addWidget(file_group)
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "Setup")
    
    def create_complete_workflow_tab(self):
        """Complete workflow tab - heart of AVOCADO Professional"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(25)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # Title section
        title_frame = QFrame()
        title_frame.setObjectName("workflowTitleFrame")
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(20, 15, 20, 15)
        
        title_label = QLabel("Complete Professional Workflow")
        title_label.setObjectName("workflowTitle")
        
        desc_label = QLabel(
            "Integrated processing that combines automatic OCLC number search "
            "with complete download of professional bibliographic metadata."
        )
        desc_label.setObjectName("workflowDesc")
        desc_label.setWordWrap(True)
        
        title_layout.addWidget(title_label)
        title_layout.addWidget(desc_label)
        
        # Input file section
        input_group = QGroupBox("Select Book List")
        input_group.setObjectName("inputGroup")
        input_layout = QGridLayout(input_group)
        input_layout.setSpacing(15)
        
        input_layout.addWidget(QLabel("CSV File:"), 0, 0)
        self.input_file_input = QLineEdit()
        self.input_file_input.setPlaceholderText("Select CSV file with your book list")
        self.input_file_input.textChanged.connect(lambda: setattr(self, 'input_file', self.input_file_input.text()))
        input_layout.addWidget(self.input_file_input, 0, 1)
        
        browse_input_btn = QPushButton("Browse")
        browse_input_btn.clicked.connect(self.browse_input_file)
        input_layout.addWidget(browse_input_btn, 0, 2)
        
        # Main process button
        self.process_btn = QPushButton("Process Complete Professional Workflow")
        self.process_btn.setObjectName("primaryProcessButton")
        self.process_btn.clicked.connect(self.start_complete_workflow)
        
        # Progress section
        progress_group = QGroupBox("Processing Progress")
        progress_group.setObjectName("progressGroup")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("professionalProgressBar")
        self.progress_label = QLabel("Ready to process...")
        self.progress_label.setObjectName("progressLabel")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.progress_label)
        
        # Results section
        results_group = QGroupBox("Processing Results")
        results_group.setObjectName("resultsGroup")
        results_layout = QVBoxLayout(results_group)
        
        self.results_text = QTextEdit()
        self.results_text.setObjectName("professionalResults")
        self.results_text.setReadOnly(True)
        self.results_text.setMinimumHeight(250)
        results_layout.addWidget(self.results_text)
        
        # Control buttons
        controls_layout = QHBoxLayout()
        
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.results_text.clear)
        
        controls_layout.addWidget(self.stop_btn)
        controls_layout.addStretch()
        controls_layout.addWidget(clear_btn)
        
        # Add to layout
        layout.addWidget(title_frame)
        layout.addWidget(input_group)
        layout.addWidget(self.process_btn)
        layout.addWidget(progress_group)
        layout.addWidget(results_group)
        layout.addLayout(controls_layout)
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "Complete Workflow")
    
    def create_advanced_tab(self):
        """Advanced options tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(25)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # Title
        title_label = QLabel("Advanced Options")
        title_label.setObjectName("advancedTitle")
        layout.addWidget(title_label)
        
        # Individual steps
        steps_group = QGroupBox("Individual Step Processing")
        steps_group.setObjectName("stepsGroup")
        steps_layout = QVBoxLayout(steps_group)
        
        # Step 1
        step1_layout = QHBoxLayout()
        step1_btn = QPushButton("Step 1: Find OCLC Numbers Only")
        step1_btn.setObjectName("stepButton")
        step1_btn.clicked.connect(lambda: QMessageBox.information(self, "AVOCADO Professional", 
                                                                 "Advanced functionality in development.\nUse Complete Workflow for best experience."))
        step1_layout.addWidget(step1_btn)
        
        step1_desc = QLabel("Find OCLC numbers without downloading complete metadata")
        step1_desc.setObjectName("stepDesc")
        step1_layout.addWidget(step1_desc)
        step1_layout.addStretch()
        steps_layout.addLayout(step1_layout)
        
        # Step 2
        step2_layout = QHBoxLayout()
        step2_btn = QPushButton("Step 2: Download Metadata Only")
        step2_btn.setObjectName("stepButton")
        step2_btn.clicked.connect(lambda: QMessageBox.information(self, "AVOCADO Professional", 
                                                                 "Advanced functionality in development.\nUse Complete Workflow for best experience."))
        step2_layout.addWidget(step2_btn)
        
        step2_desc = QLabel("Download metadata from existing OCLC numbers")
        step2_desc.setObjectName("stepDesc")
        step2_layout.addWidget(step2_desc)
        step2_layout.addStretch()
        steps_layout.addLayout(step2_layout)
        
        # OCLC numbers input
        oclc_group = QGroupBox("OCLC Numbers for Individual Processing")
        oclc_group.setObjectName("oclcGroup")
        oclc_layout = QVBoxLayout(oclc_group)
        
        oclc_label = QLabel("Enter OCLC numbers (one per line):")
        oclc_layout.addWidget(oclc_label)
        
        self.oclc_numbers_text = QTextEdit()
        self.oclc_numbers_text.setObjectName("oclcNumbersInput")
        self.oclc_numbers_text.setMaximumHeight(150)
        self.oclc_numbers_text.setPlaceholderText("12345678\n87654321\n...")
        oclc_layout.addWidget(self.oclc_numbers_text)
        
        # OCLC buttons
        oclc_buttons = QHBoxLayout()
        load_btn = QPushButton("Load from CSV")
        load_btn.clicked.connect(self.load_oclc_from_csv)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.oclc_numbers_text.clear)
        
        oclc_buttons.addWidget(load_btn)
        oclc_buttons.addStretch()
        oclc_buttons.addWidget(clear_btn)
        oclc_layout.addLayout(oclc_buttons)
        
        # Advanced progress
        self.advanced_progress = QProgressBar()
        self.advanced_progress.setObjectName("advancedProgressBar")
        
        # Add to layout
        layout.addWidget(steps_group)
        layout.addWidget(oclc_group)
        layout.addWidget(self.advanced_progress)
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "Advanced")
    
    def create_about_tab(self):
        """About tab with professional design"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(25)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # Title section
        title_label = QLabel("AVOCADO v2.7")
        title_label.setObjectName("aboutTitle")
        layout.addWidget(title_label)
        
        subtitle_label = QLabel("Archivo Venezuela OCLC & Data Organizer")
        subtitle_label.setObjectName("aboutSubtitle")
        layout.addWidget(subtitle_label)
        
        version_label = QLabel("Professional Version - Clean Interface + Complete Functionality")
        version_label.setObjectName("aboutVersion")
        layout.addWidget(version_label)
        
        # About content
        about_group = QGroupBox("About AVOCADO Professional")
        about_group.setObjectName("aboutGroup")
        about_layout = QVBoxLayout(about_group)
        
        about_text = QLabel(
            "AVOCADO Professional is the most advanced version of the specialized tool "
            "for Venezuelan cultural institutions, libraries, and archives.\n\n"
            "It combines a clean and professional interface with complete functionality "
            "for processing bibliographic metadata through OCLC WorldCat integration.\n\n"
            "Purpose: Complete bibliographic metadata management\n"
            "Focus: Venezuelan cultural heritage\n"
            "Features: Complete automated workflow\n"
            "Security: Secure credential storage\n"
            "Metadata: 14+ bibliographic information fields\n"
            "Interface: Clean professional and intuitive design"
        )
        about_text.setWordWrap(True)
        about_text.setObjectName("aboutText")
        about_layout.addWidget(about_text)
        
        # Technical specs
        tech_group = QGroupBox("Technical Specifications")
        tech_group.setObjectName("techGroup")
        tech_layout = QVBoxLayout(tech_group)
        
        tech_text = QLabel(
            "Complete metadata extraction: OCLC #, Title, Creator, Contributor, Publisher, Date, URL, ISBN, ISSN, Language, Subjects, Type, Format, Edition\n\n"
            "UTF-8 CSV output for international character support\n"
            "Automated workflow with progress tracking\n"
            "Error handling and recovery for large datasets\n"
            "Cross-platform compatibility (Windows, Mac, Linux)\n"
            "OCLC WorldCat Discovery API v2 integration\n"
        )
        tech_text.setWordWrap(True)
        tech_text.setObjectName("techText")
        tech_layout.addWidget(tech_text)
        
        layout.addWidget(about_group)
        layout.addWidget(tech_group)
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "About")
    
    def apply_professional_styles(self):
        """Apply clean professional styling without shadows and transforms"""
        self.setStyleSheet("""
            /* Main window - Clean professional theme */
            QMainWindow {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #fafbfa, stop: 1 #f4f6f4);
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
                font-size: 13px;
            }
            
            /* Professional Header */
            QFrame#professionalHeader {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1a3d1a, stop: 0.3 #2a5a2a, stop: 0.7 #4a7a4a, stop: 1 #5a8a5a);
                border: none;
                border-bottom: 3px solid #ff8c42;
            }
            
            QLabel#headerTitle {
                font-size: 22px;
                font-weight: 700;
                color: white;
                margin-bottom: 4px;
            }
            
            QLabel#headerSubtitle {
                font-size: 13px;
                color: #e8f5e8;
                font-weight: 500;
                margin-bottom: 2px;
            }
            
            QLabel#headerFeatures {
                font-size: 11px;
                color: #d0e8d0;
                font-style: italic;
            }
            
            QLabel#statusIndicator {
                font-size: 16px;
                margin-right: 12px;
                font-weight: bold;
            }
            
            /* Clean Tabs */
            QTabWidget#professionalTabs::pane {
                border: 2px solid #e8f2e8;
                background-color: white;
                border-radius: 10px;
                margin-top: 8px;
            }
            
            QTabWidget#professionalTabs QTabBar::tab {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #f9fcf9, stop: 1 #f0f5f0);
                border: 2px solid #dae8da;
                border-bottom: none;
                padding: 14px 24px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 600;
                font-size: 13px;
                color: #2a5a2a;
                min-width: 120px;
            }
            
            QTabWidget#professionalTabs QTabBar::tab:selected {
                background: white;
                color: #1a3d1a;
                border-color: #e8f2e8;
                border-bottom: 2px solid white;
                font-weight: bold;
                margin-bottom: -2px;
            }
            
            QTabWidget#professionalTabs QTabBar::tab:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #f7faf7);
                color: #1a3d1a;
            }
            
            /* Clean Group Boxes */
            QGroupBox {
                font-weight: 600;
                font-size: 15px;
                color: #1a3d1a;
                border: 2px solid #e8f2e8;
                border-radius: 10px;
                margin-top: 18px;
                padding-top: 18px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #fbfefb);
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 18px;
                padding: 5px 14px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #4a7a4a, stop: 1 #3a6a3a);
                color: white;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
            }
            
            QGroupBox#welcomeGroup {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #f7fdf7, stop: 1 #ecf7ec);
                border-color: #d0e8d0;
            }
            
            QGroupBox#credentialsGroup {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #fffaf7, stop: 1 #fdf4ef);
                border-color: #ffcc99;
            }
            
            QGroupBox#credentialsGroup::title {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ff8c42, stop: 1 #e67a38);
            }
            
            /* Clean Input Fields */
            QLineEdit {
                border: 2px solid #e5ebe5;
                border-radius: 6px;
                padding: 10px 14px;
                font-size: 13px;
                background-color: white;
                color: #1a3d1a;
                font-weight: 500;
                selection-background-color: #d0e8d0;
                min-height: 16px;
            }
            
            QLineEdit:focus {
                border-color: #4a7a4a;
                background-color: #fbfefb;
                color: #0d2b0d;
            }
            
            QLineEdit:hover {
                border-color: #5a8a5a;
                background-color: #fdfefd;
            }
            
            QLineEdit::placeholder {
                color: #8fa58f;
                font-style: italic;
                font-weight: normal;
            }
            
            /* Clean Text Areas */
            QTextEdit {
                border: 2px solid #e5ebe5;
                border-radius: 6px;
                padding: 10px;
                font-size: 12px;
                background-color: white;
                selection-background-color: #d0e8d0;
                color: #1a3d1a;
                line-height: 1.3;
            }
            
            QTextEdit#professionalResults {
                font-family: 'SF Mono', 'Monaco', 'Inconsolata', 'Roboto Mono', 'Courier New', monospace;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #fafcfa);
                border-color: #d5e5d5;
                font-size: 11px;
            }
            
            QTextEdit#oclcNumbersInput {
                font-family: 'SF Mono', 'Monaco', 'Inconsolata', 'Roboto Mono', 'Courier New', monospace;
                background: #fbfefb;
                border-color: #d0e8d0;
                font-size: 12px;
            }
            
            QTextEdit:focus {
                border-color: #4a7a4a;
                background-color: #fbfefb;
            }
            
            /* Clean Buttons */
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ff9552, stop: 1 #ff7a28);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-size: 13px;
                font-weight: 600;
                min-width: 100px;
                min-height: 18px;
            }
            
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ff8542, stop: 1 #ff6814);
            }
            
            QPushButton:pressed {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #e65f0a, stop: 1 #cc5500);
            }
            
            QPushButton#primaryProcessButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #4a7a4a, stop: 1 #3a6a3a);
                font-size: 16px;
                font-weight: bold;
                padding: 16px 32px;
                min-width: 300px;
                border-radius: 8px;
                margin: 8px 0;
                min-height: 22px;
            }
            
            QPushButton#primaryProcessButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #5a8a5a, stop: 1 #4a7a4a);
            }
            
            QPushButton#primaryProcessButton:pressed {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #2a5a2a, stop: 1 #1a3d1a);
            }
            
            QPushButton#testButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #66b366, stop: 1 #4d9c4d);
                min-width: 120px;
            }
            
            QPushButton#testButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #7abf7a, stop: 1 #66b366);
            }
            
        QPushButton#templateButton {
            background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                stop: 0 #4CAF50, stop: 1 #388E3C);
            min-width: 240px;
        }
            
            QPushButton#stopButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ff5252, stop: 1 #e64545);
            }
            
            QPushButton#stepButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #42a5f5, stop: 1 #1e88e5);
                min-width: 180px;
            }
            
            QPushButton:disabled {
                background: #cccccc;
                color: #666666;
            }
            
            /* Clean Labels */
            QLabel {
                color: #1a3d1a;
                font-weight: 500;
                font-size: 13px;
                line-height: 1.4;
            }
            
            QLabel#welcomeText {
                font-size: 14px;
                color: #2a5a2a;
                line-height: 1.5;
            }
            
            QLabel#guideText {
                font-size: 13px;
                color: #1a3d1a;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #f2f8f2, stop: 1 #eaf5ea);
                padding: 12px;
                border-radius: 6px;
                border-left: 3px solid #4a7a4a;
                margin: 12px 0;
                line-height: 1.4;
            }
            
            QLabel#workflowTitle {
                font-size: 22px;
                font-weight: bold;
                color: #1a3d1a;
                margin-bottom: 8px;
            }
            
            QLabel#workflowDesc {
                font-size: 15px;
                color: #2a5a2a;
                font-weight: 500;
                line-height: 1.5;
            }
            
            QLabel#advancedTitle {
                font-size: 18px;
                font-weight: bold;
                color: #1a3d1a;
                margin-bottom: 12px;
            }
            
            QLabel#stepDesc {
                font-size: 11px;
                color: #777;
                font-style: italic;
                margin-left: 8px;
            }
            
            QLabel#progressLabel {
                font-size: 12px;
                color: #2a5a2a;
                font-weight: 500;
                margin-top: 4px;
            }
            
            /* About page styles */
            QLabel#aboutTitle {
                font-size: 34px;
                font-weight: bold;
                color: #1a3d1a;
            }
            
            QLabel#aboutSubtitle {
                font-size: 18px;
                color: #4a7a4a;
                font-weight: 500;
            }
            
            QLabel#aboutVersion {
                font-size: 13px;
                color: #8fa58f;
                font-style: italic;
            }
            
            QLabel#aboutText, QLabel#techText {
                font-size: 13px;
                color: #2a5a2a;
                line-height: 1.6;
            }
            
            /* Clean Progress Bars */
            QProgressBar {
                border: 2px solid #e5ebe5;
                border-radius: 6px;
                text-align: center;
                font-size: 12px;
                font-weight: bold;
                background-color: #fafcfa;
                color: #1a3d1a;
                height: 22px;
            }
            
            QProgressBar#professionalProgressBar {
                height: 26px;
                font-size: 13px;
            }
            
            QProgressBar::chunk {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #4a7a4a, stop: 0.5 #5a8a5a, stop: 1 #ff8c42);
                border-radius: 4px;
                margin: 2px;
            }
            
            /* Clean Checkbox */
            QCheckBox {
                color: #1a3d1a;
                font-size: 13px;
                font-weight: 500;
                spacing: 8px;
            }
            
            QCheckBox#saveCredsCheckbox {
                color: #4a7a4a;
                font-weight: 600;
                margin: 8px 0px;
                font-size: 14px;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #e5ebe5;
                background-color: white;
            }
            
            QCheckBox::indicator:hover {
                border-color: #4a7a4a;
                background-color: #fbfefb;
            }
            
            QCheckBox::indicator:checked {
                background-color: #4a7a4a;
                border-color: #4a7a4a;
            }
            
            /* Special frames */
            QFrame#workflowTitleFrame {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #f2f8f2, stop: 1 #eaf5ea);
                border: 2px solid #d0e8d0;
                border-radius: 8px;
                margin-bottom: 8px;
            }
            
            /* Connected status indicator */
            QLabel#statusIndicator[connected="true"] {
                color: #4CAF50;
            }
            
            QLabel#statusIndicator[connected="false"] {
                color: #F44336;
            }
        """)
    
    # Helper methods
    def load_credentials(self):
        """Load credentials from .env - CLEAN VERSION"""
        env_files = ['.env', os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')]
        
        for env_file in env_files:
            if os.path.exists(env_file):
                try:
                    with open(env_file, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if line and not line.startswith('#') and '=' in line:
                                key, value = line.split('=', 1)
                                key = key.strip()
                                value = value.strip().strip('"\'')
                                
                                if key == 'OCLC_WSKEY':
                                    self.wskey = value
                                elif key == 'OCLC_WSSECRET':
                                    self.wssecret = value
                                elif key == 'OUTPUT_DIR':
                                    self.output_dir = value
                    return
                except Exception:
                    pass
    
    def save_credentials(self):
        """Save credentials to .env"""
        try:
            env_content = f"""# AVOCADO v2.7 - OCLC Credentials
# Keep this file secure and do not commit to version control

OCLC_WSKEY={self.wskey}
OCLC_WSSECRET={self.wssecret}

# Output directory
OUTPUT_DIR={self.output_dir}
"""
            with open('.env', 'w', encoding='utf-8') as f:
                f.write(env_content)
            return True
        except Exception:
            return False
    
    # Event handlers
    def on_credentials_changed(self):
        """Handle credential changes"""
        self.wskey = self.wskey_input.text().strip()
        self.wssecret = self.wssecret_input.text().strip()
        self.update_connection_status(False)
    
    def update_connection_status(self, connected):
        """Update connection status indicator"""
        if connected:
            self.status_indicator.setText("● Connected")
            self.status_indicator.setProperty("connected", "true")
        else:
            self.status_indicator.setText("● Disconnected")
            self.status_indicator.setProperty("connected", "false")
        
        # Force style refresh
        self.status_indicator.style().unpolish(self.status_indicator)
        self.status_indicator.style().polish(self.status_indicator)
    
    def test_connection(self):
        """Test OCLC connection"""
        if not self.wskey or not self.wssecret:
            QMessageBox.warning(self, "AVOCADO Professional", 
                              "Please enter both credentials to test connection.")
            return
        
        # Save credentials if checked
        if self.save_creds_checkbox.isChecked():
            self.save_credentials()
        
        self.test_btn.setText("Testing...")
        self.test_btn.setEnabled(False)
        
        QTimer.singleShot(100, self.do_connection_test)
    
    def do_connection_test(self):
        """Perform actual connection test"""
        try:
            if self.fetch_oclc_token():
                self.update_connection_status(True)
                QMessageBox.information(self, "AVOCADO Professional - Success", 
                                      "Connection successful!\n\n"
                                      "Your OCLC credentials are valid.\n"
                                      "You can proceed to complete workflow.")
            else:
                self.update_connection_status(False)
                QMessageBox.critical(self, "AVOCADO Professional - Error", 
                                   "Authentication failed\n\n"
                                   "Please verify your OCLC credentials.")
        except Exception as e:
            self.update_connection_status(False)
            QMessageBox.critical(self, "AVOCADO Professional - Error", 
                               f"Connection error:\n{str(e)}")
        finally:
            self.test_btn.setText("Test Connection")
            self.test_btn.setEnabled(True)
    
    def browse_input_file(self):
        """Browse for input file"""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select CSV file", "", 
            "CSV files (*.csv);;All files (*.*)"
        )
        if filename:
            self.input_file_input.setText(filename)
            self.input_file = filename
    
    def browse_output_dir(self):
        """Browse for output directory"""
        dirname = QFileDialog.getExistingDirectory(self, "Select output directory")
        if dirname:
            self.output_dir_input.setText(dirname)
            self.output_dir = dirname
    
    def download_template(self):
        """Download CSV template"""
        template_content = """OCLC #,Author,Title
,García Márquez, Gabriel,Cien años de soledad
,Allende, Isabel,La casa de los espíritus
,Vargas Llosa, Mario,Conversación en la catedral
,Borges, Jorge Luis,Ficciones
,Cortázar, Julio,Rayuela
,Carpentier, Alejo,El reino de este mundo
,Fuentes, Carlos,La muerte de Artemio Cruz
,Uslar Pietri, Arturo,Las lanzas coloradas
,Gallegos, Rómulo,Doña Bárbara
,Díaz Rodríguez, Manuel,Ídolos rotos"""
        
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save AVOCADO Professional Template", 
            "avocado_professional_template.csv", 
            "CSV files (*.csv)"
        )
        
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                    f.write(template_content)
                QMessageBox.information(self, "AVOCADO Professional", 
                                      f"Template saved successfully!\n\n"
                                      f"File: {Path(filename).name}\n\n"
                                      f"Fill in your data and use Complete Workflow.")
            except Exception as e:
                QMessageBox.critical(self, "AVOCADO Professional", 
                                   f"Error saving template:\n{str(e)}")
    
    def start_complete_workflow(self):
        """Start complete professional workflow"""
        # Validations
        if not self.input_file:
            QMessageBox.warning(self, "AVOCADO Professional", 
                              "Please select a CSV file first.")
            return
            
        if not os.path.exists(self.input_file):
            QMessageBox.warning(self, "AVOCADO Professional", 
                              "Selected file does not exist.")
            return
            
        if not self.wskey or not self.wssecret:
            QMessageBox.warning(self, "AVOCADO Professional", 
                              "Please configure your OCLC credentials first.")
            return
        
        # Prepare UI
        self.process_btn.setText("Processing...")
        self.process_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.results_text.clear()
        
        # Update credentials and directory
        self.wskey = self.wskey_input.text().strip()
        self.wssecret = self.wssecret_input.text().strip()
        self.output_dir = self.output_dir_input.text().strip()
        
        # Create output directory
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "AVOCADO Professional", 
                               f"Could not create output directory:\n{str(e)}")
            self.reset_ui()
            return
        
        # Start worker thread
        self.worker_thread = WorkerThread("complete_workflow", self)
        self.worker_thread.progress_update.connect(self.update_progress_text)
        self.worker_thread.progress_value.connect(self.progress_bar.setValue)
        self.worker_thread.workflow_complete.connect(self.on_workflow_complete)
        self.worker_thread.workflow_error.connect(self.on_workflow_error)
        self.worker_thread.start()
    
    def stop_processing(self):
        """Stop processing"""
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.stop()
            self.worker_thread.wait(3000)
            self.update_progress_text("Processing stopped by user")
        self.reset_ui()
    
    def update_progress_text(self, message):
        """Update progress text"""
        self.progress_label.setText(message)
        self.results_text.append(message)
        
        # Auto-scroll
        cursor = self.results_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.results_text.setTextCursor(cursor)
        
        QApplication.processEvents()
    
    def on_workflow_complete(self, output_file, total, oclc_found, metadata_complete):
        """Handle successful completion"""
        self.reset_ui()
        self.update_connection_status(True)
        
        QMessageBox.information(self, "AVOCADO Professional - Complete!", 
                              f"Professional workflow completed successfully!\n\n"
                              f"File: {Path(output_file).name}\n"
                              f"Total processed: {total}\n"
                              f"With OCLC numbers: {oclc_found}\n"
                              f"With complete metadata: {metadata_complete}\n\n"
                              f"Your professional metadata file is ready!")
    
    def on_workflow_error(self, error_message):
        """Handle workflow error"""
        self.reset_ui()
        QMessageBox.critical(self, "AVOCADO Professional - Error", 
                           f"An error occurred:\n\n{error_message}")
    
    def reset_ui(self):
        """Reset UI after processing"""
        self.process_btn.setText("Process Complete Professional Workflow")
        self.process_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        if self.worker_thread:
            self.worker_thread = None
    
    def load_oclc_from_csv(self):
        """Load OCLC numbers from CSV"""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select CSV with OCLC numbers", "", 
            "CSV files (*.csv)"
        )
        if filename:
            try:
                oclc_numbers = []
                with open(filename, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        oclc = row.get('OCLC #', '').strip()
                        if oclc and oclc.isdigit():
                            oclc_numbers.append(oclc)
                
                if oclc_numbers:
                    self.oclc_numbers_text.clear()
                    self.oclc_numbers_text.insertPlainText('\n'.join(oclc_numbers))
                    QMessageBox.information(self, "AVOCADO Professional", 
                                          f"Loaded {len(oclc_numbers)} OCLC numbers")
                else:
                    QMessageBox.warning(self, "AVOCADO Professional", 
                                      "No valid OCLC numbers found")
                    
            except Exception as e:
                QMessageBox.critical(self, "AVOCADO Professional", 
                                   f"Error loading file:\n{str(e)}")
    
    # OCLC API methods - FIXED VERSION
    def fetch_oclc_token(self):
        """Get OCLC token - CLEAN"""
        try:
            token_url = "https://oauth.oclc.org/token"
            headers = {"Content-Type": "application/x-www-form-urlencoded"}
            payload = {
                "grant_type": "client_credentials",
                "scope": "wcapi:view_bib"
            }
            
            response = requests.post(token_url, 
                                   auth=(self.wskey, self.wssecret), 
                                   data=payload, headers=headers, timeout=30)
            
            if response.status_code == 200:
                self.access_token = response.json().get("access_token")
                return True
            return False
        except Exception:
            return False
    
    def search_oclc(self, title, author):
        """Search for OCLC number - CLEAN VERSION"""
        try:
            # Clean search terms
            title_clean = self.clean_search_term(title)
            author_clean = self.clean_search_term(author)
            
            # Multiple search strategies
            queries = [
                f'ti:"{title_clean}" AND au:"{author_clean}"',
                f'ti:{title_clean} AND au:{author_clean}',
                f'"{title_clean}" AND "{author_clean}"',
                f'{title_clean} {author_clean}',
            ]
            
            for query in queries:
                result = self._search_with_query(query)
                if result:
                    return result
                time.sleep(0.2)
                    
            return None
        except Exception:
            return None
    
    def _search_with_query(self, query):
        """Perform search with specific query - FIXED API PARAMETERS"""
        try:
            url = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs"
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json"
            }
            params = {
                "q": query,
                "limit": 10,
                "offset": 1,  # FIXED: Must start from 1, not 0
                "orderBy": "bestMatch"  # FIXED: Use bestMatch instead of relevance
            }
            
            response = requests.get(url, headers=headers, params=params, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                bibs = data.get("bibRecords", [])
                
                if bibs:
                    identifier = bibs[0].get("identifier", {})
                    oclc_number = identifier.get("oclcNumber") if identifier else None
                    return str(oclc_number) if oclc_number else None
                    
            return None
        except Exception:
            return None
    
    def fetch_metadata_json(self, oclc_number):
        """Get metadata JSON for OCLC number"""
        try:
            url = f"https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs/{oclc_number}"
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json"
            }
            
            response = requests.get(url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                return response.json()
            return None
        except Exception:
            return None
    
    def parse_complete_record(self, json_data, oclc_number, original_book):
        """Parse complete record - CLEAN VERSION"""
        record = {
            "OCLC #": str(oclc_number),
            "Title": "",
            "Creator": "",
            "Contributor": "",
            "Publisher": "",
            "Date": "",
            "Language": "",
            "Subjects": "",
            "Type": "",
            "Format": "",
            "ISBN": "",
            "ISSN": "",
            "Edition": "",
            "URL": f"https://www.worldcat.org/oclc/{oclc_number}"
        }
        
        try:
            if not json_data:
                return self.create_basic_record(original_book, oclc_number)
            
            # Title
            self.extract_title(json_data, record)
            
            # Creator
            self.extract_creator(json_data, record)
            
            # Contributors
            self.extract_contributors(json_data, record)
            
            # Publisher
            self.extract_publisher(json_data, record)
            
            # Other metadata
            self.extract_other_metadata(json_data, record)
            
            return record
            
        except Exception:
            return self.create_basic_record(original_book, oclc_number)
    
    def extract_title(self, json_data, record):
        """Extract title"""
        try:
            title_data = json_data.get("title", {})
            if isinstance(title_data, dict):
                main_titles = title_data.get("mainTitles", [])
                if main_titles and isinstance(main_titles, list) and len(main_titles) > 0:
                    first_title = main_titles[0]
                    if isinstance(first_title, dict):
                        title_text = first_title.get("text", "")
                        if title_text:
                            clean_title = title_text.split(" / ")[0].strip()
                            record["Title"] = self.clean_text(clean_title)
        except Exception:
            pass
    
    def extract_creator(self, json_data, record):
        """Extract creator"""
        try:
            contributor_data = json_data.get("contributor", {})
            if isinstance(contributor_data, dict):
                creators = contributor_data.get("creators", [])
                if creators and isinstance(creators, list) and len(creators) > 0:
                    creator = creators[0]
                    if isinstance(creator, dict):
                        first_name = ""
                        last_name = ""
                        
                        if "firstName" in creator and isinstance(creator["firstName"], dict):
                            first_name = creator["firstName"].get("text", "")
                        if "secondName" in creator and isinstance(creator["secondName"], dict):
                            last_name = creator["secondName"].get("text", "")
                        
                        full_name = f"{first_name} {last_name}".strip()
                        if full_name:
                            record["Creator"] = self.clean_text(full_name)
        except Exception:
            pass
    
    def extract_contributors(self, json_data, record):
        """Extract contributors"""
        try:
            contributor_data = json_data.get("contributor", {})
            if isinstance(contributor_data, dict):
                contributors = contributor_data.get("contributors", [])
                names = []
                
                if contributors and isinstance(contributors, list):
                    for c in contributors[:5]:
                        if isinstance(c, dict):
                            name_obj = c.get("name", {})
                            if isinstance(name_obj, dict):
                                name = name_obj.get("text", "")
                                if name:
                                    names.append(self.clean_text(name))
                
                if names:
                    record["Contributor"] = " ; ".join(names)
        except Exception:
            pass
    
    def extract_publisher(self, json_data, record):
        """Extract publisher - MULTIPLE METHODS"""
        publisher = ""
        
        try:
            # Method 1: publishers array
            publishers = json_data.get("publishers", [])
            if publishers and isinstance(publishers, list) and len(publishers) > 0:
                pub = publishers[0]
                if isinstance(pub, dict):
                    pub_name = pub.get("publisherName", {})
                    if isinstance(pub_name, dict):
                        publisher = pub_name.get("text", "")
            
            # Method 2: publication array
            if not publisher:
                publication = json_data.get("publication", [])
                if publication and isinstance(publication, list) and len(publication) > 0:
                    pub = publication[0]
                    if isinstance(pub, dict):
                        pub_text = pub.get("publisher", "")
                        if pub_text:
                            publisher = str(pub_text)
            
            # Method 3: direct publisher field
            if not publisher:
                pub_direct = json_data.get("publisher")
                if pub_direct:
                    if isinstance(pub_direct, list) and pub_direct:
                        publisher = str(pub_direct[0])
                    elif isinstance(pub_direct, str):
                        publisher = pub_direct
            
            # Method 4: placeOfPublication
            if not publisher:
                place_pub = json_data.get("placeOfPublication", [])
                if place_pub and isinstance(place_pub, list) and len(place_pub) > 0:
                    if isinstance(place_pub[0], dict):
                        pub_text = place_pub[0].get("publisher", "")
                        if pub_text:
                            publisher = str(pub_text)
            
            # Method 5: search in title field
            if not publisher:
                title_data = json_data.get("title", {})
                if isinstance(title_data, dict):
                    titles = title_data.get("mainTitles", [])
                    if titles and isinstance(titles, list) and len(titles) > 0:
                        if isinstance(titles[0], dict):
                            full_title = titles[0].get("text", "")
                            if " : " in full_title:
                                parts = full_title.split(" : ")
                                if len(parts) > 1:
                                    potential_pub = parts[-1].strip()
                                    if len(potential_pub) < 100:
                                        publisher = potential_pub
            
            record["Publisher"] = self.clean_text(publisher)
            
        except Exception:
            pass
    
    def extract_other_metadata(self, json_data, record):
        """Extract remaining metadata"""
        try:
            # Date
            date_data = json_data.get("date", {})
            if isinstance(date_data, dict):
                pub_date = date_data.get("publicationDate", "")
                if pub_date:
                    record["Date"] = str(pub_date)
            
            # Language
            languages = json_data.get("language", [])
            if languages and isinstance(languages, list) and len(languages) > 0:
                lang = languages[0]
                if isinstance(lang, dict):
                    lang_code = lang.get("languageCode", "")
                    if lang_code:
                        record["Language"] = str(lang_code)
                elif isinstance(lang, str):
                    record["Language"] = lang
            
            # Subjects
            subjects = json_data.get("subject", [])
            subject_list = []
            if subjects and isinstance(subjects, list):
                for subj in subjects[:5]:
                    if isinstance(subj, dict):
                        subj_name = subj.get("subjectName", {})
                        if isinstance(subj_name, dict):
                            subj_text = subj_name.get("text", "")
                            if subj_text:
                                subject_list.append(self.clean_text(subj_text))
                    elif isinstance(subj, str):
                        subject_list.append(self.clean_text(subj))
            record["Subjects"] = " ; ".join(subject_list)
            
            # Type
            item_type = json_data.get("itemType", {})
            if isinstance(item_type, dict):
                type_text = item_type.get("text", "")
                if type_text:
                    record["Type"] = str(type_text)
            
            # Format
            formats = json_data.get("format", [])
            if formats and isinstance(formats, list) and len(formats) > 0:
                fmt = formats[0]
                if isinstance(fmt, dict):
                    fmt_text = fmt.get("text", "")
                    if fmt_text:
                        record["Format"] = str(fmt_text)
                elif isinstance(fmt, str):
                    record["Format"] = fmt
            
            # ISBN/ISSN
            self.extract_identifiers(json_data, record)
            
            # Edition
            edition_info = json_data.get("edition", "")
            if isinstance(edition_info, list) and edition_info:
                edition_info = edition_info[0]
            if isinstance(edition_info, dict):
                edition_info = edition_info.get("text", "")
            if edition_info:
                record["Edition"] = self.clean_text(str(edition_info))
                
        except Exception:
            pass
    
    def extract_identifiers(self, json_data, record):
        """Extract ISBN and ISSN"""
        try:
            isbn_list = []
            issn_list = []
            
            identifier = json_data.get("identifier", {})
            if isinstance(identifier, dict):
                # Direct lists
                isbns = identifier.get("isbns", [])
                if isinstance(isbns, list):
                    isbn_list.extend([str(isbn) for isbn in isbns if isbn])
                
                issns = identifier.get("issns", [])
                if isinstance(issns, list):
                    issn_list.extend([str(issn) for issn in issns if issn])
                
                # Items array
                items = identifier.get("items", [])
                if isinstance(items, list):
                    for item in items:
                        if isinstance(item, dict):
                            item_type = item.get("type", "").lower()
                            value = item.get("value", "")
                            if item_type == "isbn" and value:
                                isbn_list.append(str(value))
                            elif item_type == "issn" and value:
                                issn_list.append(str(value))
            
            # Clean and deduplicate
            if isbn_list:
                record["ISBN"] = "; ".join(sorted(set(isbn_list)))
            if issn_list:
                record["ISSN"] = "; ".join(sorted(set(issn_list)))
                
        except Exception:
            pass
    
    def create_basic_record(self, original_book, oclc_number=""):
        """Create basic record"""
        return {
            "OCLC #": oclc_number,
            "Title": original_book.get("Title", ""),
            "Creator": original_book.get("Author", ""),
            "Contributor": "",
            "Publisher": "",
            "Date": "",
            "Language": "",
            "Subjects": "",
            "Type": "",
            "Format": "",
            "ISBN": "",
            "ISSN": "",
            "Edition": "",
            "URL": f"https://www.worldcat.org/oclc/{oclc_number}" if oclc_number else ""
        }
    
    def clean_text(self, text):
        """Clean text"""
        if not text:
            return ""
        
        if not isinstance(text, str):
            text = str(text)
        
        text = unicodedata.normalize('NFC', text)
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text
    
    def clean_search_term(self, term):
        """Clean search terms"""
        cleaned = re.sub(r'[^\w\sáéíóúüñÁÉÍÓÚÜÑ]', ' ', term)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned

def main():
    """Launch AVOCADO Professional"""
    global logo_pixmap
    
    try:
        app = QApplication(sys.argv)
        app.setApplicationName("AVOCADO v2.7")
        app.setApplicationVersion("2.7")
        app.setOrganizationName("Archivo Venezuela")
        app.setFont(QFont("Segoe UI", 10))
        
        # Now that QApplication exists, load logo
        from PyQt5.QtCore import Qt
        logo_pixmap = QPixmap(logo_path).scaledToHeight(48, Qt.SmoothTransformation)
        
        # Create main window
        window = AvocadoProfessional()
        window.show()
        window.raise_()
        window.activateWindow()
        
        sys.exit(app.exec_())
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()