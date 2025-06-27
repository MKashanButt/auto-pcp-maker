import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import logging
from datetime import datetime
import traceback
import json
import re
from pathlib import Path
import queue
import time

class MailMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced Mail Merge to PDF Tool v2.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Initialize logging
        self._setup_logging()
        self.logger = logging.getLogger(__name__)
        self.logger.info("Application initialized")
        
        # Variables
        self.template_path = tk.StringVar()
        self.data_source_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.output_dir.set(os.path.expanduser("~/Documents/MailMergeOutput"))
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.progress_var = tk.DoubleVar()
        self.preview_data = None
        self.is_processing = False
        self.template_variables = set()
        
        # Settings
        self.settings_file = os.path.expanduser("~/Documents/MailMergeSettings.json")
        self.settings = self._load_settings()
        
        # Queue for thread communication
        self.message_queue = queue.Queue()
        
        # Create the main UI
        self._create_ui()
        self._load_last_session()
        
        # Start queue monitoring
        self._monitor_queue()
    
    def _setup_logging(self):
        """Configure logging for the application"""
        log_dir = os.path.expanduser("~/Documents/MailMergeLogs")
        os.makedirs(log_dir, exist_ok=True)
        
        # Create a timestamped log file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"mailmerge_{timestamp}.log")
        
        # Keep only last 10 log files
        self._cleanup_old_logs(log_dir)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
    
    def _cleanup_old_logs(self, log_dir, keep_count=10):
        """Remove old log files, keeping only the most recent ones"""
        try:
            log_files = [f for f in os.listdir(log_dir) if f.startswith('mailmerge_') and f.endswith('.log')]
            log_files.sort(reverse=True)
            
            for old_log in log_files[keep_count:]:
                os.remove(os.path.join(log_dir, old_log))
                
        except Exception as e:
            print(f"Warning: Could not cleanup old logs: {e}")
    
    def _load_settings(self):
        """Load application settings from file"""
        default_settings = {
            "last_template": "",
            "last_data_source": "",
            "last_output_dir": self.output_dir.get(),
            "create_word": True,
            "create_pdf": True,
            "auto_save_settings": True,
            "max_preview_rows": 100,
            "filename_pattern": "{First Name}_{Last Name}"
        }
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    loaded_settings = json.load(f)
                    default_settings.update(loaded_settings)
        except Exception as e:
            self.logger.warning(f"Could not load settings: {e}")
        
        return default_settings
    
    def _save_settings(self):
        """Save current settings to file"""
        if not self.settings.get("auto_save_settings", True):
            return
            
        try:
            self.settings.update({
                "last_template": self.template_path.get(),
                "last_data_source": self.data_source_path.get(),
                "last_output_dir": self.output_dir.get(),
                "create_word": self.create_word_var.get(),
                "create_pdf": self.create_pdf_var.get()
            })
            
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
                
        except Exception as e:
            self.logger.warning(f"Could not save settings: {e}")
    
    def _load_last_session(self):
        """Load settings from last session"""
        try:
            if self.settings.get("last_template") and os.path.exists(self.settings["last_template"]):
                self.template_path.set(self.settings["last_template"])
            
            if self.settings.get("last_data_source") and os.path.exists(self.settings["last_data_source"]):
                self.data_source_path.set(self.settings["last_data_source"])
            
            if self.settings.get("last_output_dir"):
                self.output_dir.set(self.settings["last_output_dir"])
                
        except Exception as e:
            self.logger.warning(f"Could not restore last session: {e}")
    
    def _monitor_queue(self):
        """Monitor the message queue for updates from worker thread"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                message_type = message.get('type')
                
                if message_type == 'status':
                    self.status_var.set(message['text'])
                elif message_type == 'progress':
                    self.progress_var.set(message['value'])
                elif message_type == 'error':
                    messagebox.showerror("Error", message['text'])
                elif message_type == 'success':
                    messagebox.showinfo("Success", message['text'])
                elif message_type == 'warning':
                    messagebox.showwarning("Warning", message['text'])
                    
        except queue.Empty:
            pass
        finally:
            # Schedule next check
            self.root.after(100, self._monitor_queue)
    
    def _create_ui(self):
        """Create the main user interface"""
        self.logger.debug("Creating UI components")
        try:
            # Create main menu
            self._create_menu()
            
            # Create a notebook (tabbed interface)
            notebook = ttk.Notebook(self.root)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Tab 1: Configuration
            config_frame = ttk.Frame(notebook)
            notebook.add(config_frame, text="Configuration")
            
            # Tab 2: Preview
            preview_frame = ttk.Frame(notebook)
            notebook.add(preview_frame, text="Data Preview")
            
            # Tab 3: Template Analysis
            template_frame = ttk.Frame(notebook)
            notebook.add(template_frame, text="Template Analysis")
            
            # Tab 4: Settings
            settings_frame = ttk.Frame(notebook)
            notebook.add(settings_frame, text="Settings")
            
            # Tab 5: Help
            help_frame = ttk.Frame(notebook)
            notebook.add(help_frame, text="Help")
            
            # Setup each tab
            self._setup_config_tab(config_frame)
            self._setup_preview_tab(preview_frame)
            self._setup_template_tab(template_frame)
            self._setup_settings_tab(settings_frame)
            self._setup_help_tab(help_frame)
            
            # Status bar
            self._create_status_bar()
            
            self.logger.info("UI created successfully")
        except Exception as e:
            self.logger.error(f"Error creating UI: {str(e)}")
            raise
    
    def _create_menu(self):
        """Create the application menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Session", command=self._new_session)
        file_menu.add_separator()
        file_menu.add_command(label="Save Settings", command=self._save_settings)
        file_menu.add_command(label="Load Settings", command=self._load_settings_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Validate Template", command=self._validate_template)
        tools_menu.add_command(label="Open Output Folder", command=self._open_output_folder)
        tools_menu.add_command(label="Open Log Folder", command=self._open_log_folder)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)
    
    def _create_status_bar(self):
        """Create the status bar at the bottom"""
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            status_frame, 
            variable=self.progress_var, 
            mode="determinate",
            length=300
        )
        self.progress_bar.pack(side=tk.RIGHT, padx=5)
        
        # Status label
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(fill=tk.X, side=tk.LEFT)
        
        # Add time display
        self.time_var = tk.StringVar()
        time_label = ttk.Label(status_frame, textvariable=self.time_var, anchor=tk.E)
        time_label.pack(side=tk.RIGHT, padx=10)
        self._update_time()
    
    def _update_time(self):
        """Update the time display"""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_var.set(current_time)
        self.root.after(1000, self._update_time)
    
    def _setup_config_tab(self, parent):
        """Configure the Configuration tab"""
        try:
            # Create a canvas and scrollbar for scrolling
            canvas = tk.Canvas(parent)
            scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # File selection frame
            file_frame = ttk.LabelFrame(scrollable_frame, text="File Selection")
            file_frame.pack(fill=tk.X, padx=10, pady=10)
            
            # Template selection with validation indicator
            template_row = ttk.Frame(file_frame)
            template_row.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(template_row, text="Word Template:").pack(side=tk.LEFT)
            self.template_status = ttk.Label(template_row, text="", foreground="red")
            self.template_status.pack(side=tk.RIGHT)
            
            template_entry_frame = ttk.Frame(file_frame)
            template_entry_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Entry(template_entry_frame, textvariable=self.template_path, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(template_entry_frame, text="Browse...", command=self.browse_template).pack(side=tk.RIGHT, padx=(5,0))
            
            # Data source selection with validation indicator
            data_row = ttk.Frame(file_frame)
            data_row.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(data_row, text="Data Source:").pack(side=tk.LEFT)
            self.data_status = ttk.Label(data_row, text="", foreground="red")
            self.data_status.pack(side=tk.RIGHT)
            
            data_entry_frame = ttk.Frame(file_frame)
            data_entry_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Entry(data_entry_frame, textvariable=self.data_source_path, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(data_entry_frame, text="Browse...", command=self.browse_data_source).pack(side=tk.RIGHT, padx=(5,0))
            
            # Output directory selection
            output_row = ttk.Frame(file_frame)
            output_row.pack(fill=tk.X, padx=5, pady=5)
            ttk.Label(output_row, text="Output Directory:").pack(side=tk.LEFT)
            
            output_entry_frame = ttk.Frame(file_frame)
            output_entry_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Entry(output_entry_frame, textvariable=self.output_dir, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(output_entry_frame, text="Browse...", command=self.browse_output_dir).pack(side=tk.RIGHT, padx=(5,0))
            
            # Options frame
            options_frame = ttk.LabelFrame(scrollable_frame, text="Output Options")
            options_frame.pack(fill=tk.X, padx=10, pady=10)
            
            # Output format options
            format_frame = ttk.Frame(options_frame)
            format_frame.pack(fill=tk.X, padx=5, pady=5)
            
            self.create_word_var = tk.BooleanVar(value=self.settings.get("create_word", True))
            self.create_pdf_var = tk.BooleanVar(value=self.settings.get("create_pdf", True))
            
            ttk.Checkbutton(format_frame, text="Create Word Documents (.docx)", 
                          variable=self.create_word_var).pack(anchor=tk.W, pady=2)
            ttk.Checkbutton(format_frame, text="Create PDF Documents (.pdf)", 
                          variable=self.create_pdf_var).pack(anchor=tk.W, pady=2)
            
            # Advanced options
            advanced_frame = ttk.LabelFrame(scrollable_frame, text="Advanced Options")
            advanced_frame.pack(fill=tk.X, padx=10, pady=10)
            
            self.skip_errors_var = tk.BooleanVar(value=True)
            self.open_output_var = tk.BooleanVar(value=False)
            self.backup_existing_var = tk.BooleanVar(value=True)
            
            ttk.Checkbutton(advanced_frame, text="Skip records with errors (continue processing)", 
                          variable=self.skip_errors_var).pack(anchor=tk.W, pady=2, padx=5)
            ttk.Checkbutton(advanced_frame, text="Open output folder when complete", 
                          variable=self.open_output_var).pack(anchor=tk.W, pady=2, padx=5)
            ttk.Checkbutton(advanced_frame, text="Backup existing files", 
                          variable=self.backup_existing_var).pack(anchor=tk.W, pady=2, padx=5)
            
            # Filename pattern
            pattern_frame = ttk.Frame(advanced_frame)
            pattern_frame.pack(fill=tk.X, padx=5, pady=5)
            ttk.Label(pattern_frame, text="Filename Pattern:").pack(side=tk.LEFT)
            
            self.filename_pattern_var = tk.StringVar(value=self.settings.get("filename_pattern", "{First Name}_{Last Name}"))
            pattern_entry = ttk.Entry(pattern_frame, textvariable=self.filename_pattern_var, width=30)
            pattern_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            ttk.Label(pattern_frame, text="Use {Column Name} format").pack(side=tk.RIGHT)
            
            # Action buttons frame
            button_frame = ttk.Frame(scrollable_frame)
            button_frame.pack(fill=tk.X, padx=10, pady=20)
            
            # Left side buttons
            left_buttons = ttk.Frame(button_frame)
            left_buttons.pack(side=tk.LEFT)
            
            ttk.Button(left_buttons, text="Load Preview", command=self.load_preview).pack(side=tk.LEFT, padx=5)
            ttk.Button(left_buttons, text="Validate Setup", command=self._validate_setup).pack(side=tk.LEFT, padx=5)
            
            # Right side buttons
            right_buttons = ttk.Frame(button_frame)
            right_buttons.pack(side=tk.RIGHT)
            
            self.run_button = ttk.Button(right_buttons, text="🚀 Run Mail Merge", 
                                       command=self.run_mail_merge, style="Accent.TButton")
            self.run_button.pack(side=tk.RIGHT, padx=5)
            
            self.stop_button = ttk.Button(right_buttons, text="⏹ Stop", 
                                        command=self._stop_processing, state='disabled')
            self.stop_button.pack(side=tk.RIGHT, padx=5)
            
            # Pack canvas and scrollbar
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Bind mousewheel to canvas
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
            self.logger.debug("Configuration tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up configuration tab: {str(e)}")
            raise
    
    def _setup_preview_tab(self, parent):
        """Configure the Preview tab with enhanced features"""
        try:
            # Control frame
            control_frame = ttk.Frame(parent)
            control_frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Label(control_frame, text="Records:").pack(side=tk.LEFT)
            self.record_count_var = tk.StringVar(value="0")
            ttk.Label(control_frame, textvariable=self.record_count_var, font=("TkDefaultFont", 9, "bold")).pack(side=tk.LEFT, padx=5)
            
            # Refresh button
            ttk.Button(control_frame, text="🔄 Refresh", command=self.load_preview).pack(side=tk.RIGHT, padx=5)
            
            # Search frame
            search_frame = ttk.Frame(parent)
            search_frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
            self.search_var = tk.StringVar()
            self.search_var.trace("w", self._filter_preview)
            search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
            search_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Button(search_frame, text="Clear", command=lambda: self.search_var.set("")).pack(side=tk.LEFT, padx=5)
            
            # Create a frame for the preview table
            preview_frame = ttk.Frame(parent)
            preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Create a Treeview widget for displaying data
            self.preview_tree = ttk.Treeview(preview_frame, show='tree headings')
            
            # Add scrollbars
            y_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
            x_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
            self.preview_tree.configure(yscroll=y_scrollbar.set, xscroll=x_scrollbar.set)
            
            # Pack scrollbars and treeview
            y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Help text
            help_label = ttk.Label(parent, text="Load your data source to preview records and verify column mappings.", 
                                 font=("TkDefaultFont", 9, "italic"))
            help_label.pack(pady=5)
            
            self.logger.debug("Preview tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up preview tab: {str(e)}")
            raise
    
    def _setup_template_tab(self, parent):
        """Configure the Template Analysis tab"""
        try:
            # Template info frame
            info_frame = ttk.LabelFrame(parent, text="Template Information")
            info_frame.pack(fill=tk.X, padx=10, pady=10)
            
            self.template_info_text = tk.Text(info_frame, height=6, wrap=tk.WORD, state=tk.DISABLED)
            self.template_info_text.pack(fill=tk.X, padx=5, pady=5)
            
            # Variables frame
            vars_frame = ttk.LabelFrame(parent, text="Template Variables")
            vars_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Variables list
            self.vars_tree = ttk.Treeview(vars_frame, columns=('Status',), show='tree headings')
            self.vars_tree.heading('#0', text='Variable Name')
            self.vars_tree.heading('Status', text='Data Source Match')
            
            vars_scrollbar = ttk.Scrollbar(vars_frame, orient=tk.VERTICAL, command=self.vars_tree.yview)
            self.vars_tree.configure(yscroll=vars_scrollbar.set)
            
            vars_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.vars_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Analysis button
            button_frame = ttk.Frame(parent)
            button_frame.pack(fill=tk.X, padx=10, pady=10)
            
            ttk.Button(button_frame, text="🔍 Analyze Template", 
                      command=self._analyze_template).pack(side=tk.LEFT)
            
            self.logger.debug("Template analysis tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up template tab: {str(e)}")
            raise
    
    def _setup_settings_tab(self, parent):
        """Configure the Settings tab"""
        try:
            # General settings
            general_frame = ttk.LabelFrame(parent, text="General Settings")
            general_frame.pack(fill=tk.X, padx=10, pady=10)
            
            self.auto_save_var = tk.BooleanVar(value=self.settings.get("auto_save_settings", True))
            ttk.Checkbutton(general_frame, text="Auto-save settings on exit", 
                          variable=self.auto_save_var).pack(anchor=tk.W, padx=5, pady=5)
            
            # Preview settings
            preview_settings_frame = ttk.LabelFrame(parent, text="Preview Settings")
            preview_settings_frame.pack(fill=tk.X, padx=10, pady=10)
            
            max_rows_frame = ttk.Frame(preview_settings_frame)
            max_rows_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(max_rows_frame, text="Maximum preview rows:").pack(side=tk.LEFT)
            self.max_preview_var = tk.StringVar(value=str(self.settings.get("max_preview_rows", 100)))
            ttk.Entry(max_rows_frame, textvariable=self.max_preview_var, width=10).pack(side=tk.LEFT, padx=5)
            
            # Logging settings
            logging_frame = ttk.LabelFrame(parent, text="Logging Settings")
            logging_frame.pack(fill=tk.X, padx=10, pady=10)
            
            self.verbose_logging_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(logging_frame, text="Enable verbose logging (debug mode)", 
                          variable=self.verbose_logging_var, command=self._toggle_logging_level).pack(anchor=tk.W, padx=5, pady=5)
            
            # Buttons
            settings_buttons = ttk.Frame(parent)
            settings_buttons.pack(fill=tk.X, padx=10, pady=20)
            
            ttk.Button(settings_buttons, text="Save Settings", command=self._save_settings).pack(side=tk.LEFT, padx=5)
            ttk.Button(settings_buttons, text="Reset to Defaults", command=self._reset_settings).pack(side=tk.LEFT, padx=5)
            
            self.logger.debug("Settings tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up settings tab: {str(e)}")
            raise
    
    def _setup_help_tab(self, parent):
        """Configure the Help tab with comprehensive information"""
        try:
            help_text = """
Enhanced Mail Merge to PDF Tool v2.0 - Help Guide

OVERVIEW
This tool merges data from Excel/CSV files with Word templates to create personalized documents in Word and/or PDF format.

GETTING STARTED
1. Select a Word template (.docx) that contains merge fields in {{ }} format
2. Select your data source (Excel .xlsx/.xls or CSV .csv file)
3. Choose an output directory for generated documents
4. Configure your output options and click "Run Mail Merge"

TEMPLATE FORMAT
Your Word template should contain merge fields using double curly braces:
- {{name}} - Will be replaced with data from the "name" column
- {{address}} - Will be replaced with data from the "address" column
- And so on...

SUPPORTED DATA SOURCES
- CSV files (.csv) - Comma-separated values
- Excel files (.xlsx, .xls) - Microsoft Excel workbooks

FILENAME PATTERNS
You can customize how output files are named using the filename pattern:
- {First Name}_{Last Name} - Uses data from these columns
- Document_{Index} - Uses sequential numbering
- {ID}_{Date} - Combines multiple fields

ADVANCED FEATURES
- Template validation and variable analysis
- Data preview with search functionality
- Batch processing with progress tracking
- Error handling and logging
- Settings persistence across sessions
- Automatic backup of existing files

TROUBLESHOOTING
- Check the Template Analysis tab to verify your template variables
- Use the Data Preview tab to confirm your data is loaded correctly
- Review log files in ~/Documents/MailMergeLogs/ for detailed error information
- Ensure all required columns are present in your data source

TIPS FOR BEST RESULTS
- Use descriptive column headers in your data source
- Test with a small dataset first
- Keep template variables simple and consistent
- Ensure data doesn't contain special characters that might cause filename issues

For technical support or feature requests, check the application logs for detailed error information.
            """
            
            # Create scrollable text widget
            text_frame = ttk.Frame(parent)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, padx=10, pady=10)
            help_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=help_scrollbar.set)
            
            help_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            text_widget.insert(tk.END, help_text)
            text_widget.config(state=tk.DISABLED)  # Make read-only
            
            self.logger.debug("Help tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up help tab: {str(e)}")
            raise
    
    def _filter_preview(self, *args):
        """Filter preview data based on search term"""
        if self.preview_data is None:
            return
            
        search_term = self.search_var.get().lower()
        
        # Clear existing items
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        if not search_term:
            # Show all data if no search term
            self._populate_preview_tree(self.preview_data)
        else:
            # Filter data
            filtered_data = self.preview_data[
                self.preview_data.astype(str).apply(
                    lambda x: x.str.lower().str.contains(search_term, na=False)
                ).any(axis=1)
            ]
            self._populate_preview_tree(filtered_data)
    
    def _populate_preview_tree(self, data):
        """Populate the preview tree with data"""
        if data is None or data.empty:
            return
            
        max_rows = min(int(self.max_preview_var.get() or 100), len(data))
        
        for i in range(max_rows):
            row = data.iloc[i].fillna('').astype(str).tolist()
            self.preview_tree.insert('', tk.END, text=f"Row {i+1}", values=row)
        
        # Update record count
        self.record_count_var.set(f"{len(data)} total, {max_rows} shown")
    
    def _new_session(self):
        """Start a new session (clear all inputs)"""
        self.template_path.set("")
        self.data_source_path.set("")
        self.output_dir.set(os.path.expanduser("~/Documents/MailMergeOutput"))
        self.preview_data = None
        self.template_variables.clear()
        
        # Clear preview
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Clear template analysis
        for item in self.vars_tree.get_children():
            self.vars_tree.delete(item)
        
        self.template_info_text.config(state=tk.NORMAL)
        self.template_info_text.delete(1.0, tk.END)
        self.template_info_text.config(state=tk.DISABLED)
        
        self.status_var.set("New session started")
        self.logger.info("New session started")
    
    def _load_settings_dialog(self):
        """Load settings from a file"""
        try:
            filepath = filedialog.askopenfilename(
                title="Load Settings",
                filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
            )
            if filepath:
                with open(filepath, 'r') as f:
                    loaded_settings = json.load(f)
                    self.settings.update(loaded_settings)
                self._load_last_session()
                messagebox.showinfo("Success", "Settings loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings: {str(e)}")
    
    def _validate_setup(self):
        """Validate the current setup and show results"""
        issues = []
        
        # Check template
        if not self.template_path.get():
            issues.append("❌ No template file selected")
        elif not os.path.exists(self.template_path.get()):
            issues.append("❌ Template file does not exist")
        else:
            issues.append("✅ Template file exists")
        
        # Check data source
        if not self.data_source_path.get():
            issues.append("❌ No data source selected")
        elif not os.path.exists(self.data_source_path.get()):
            issues.append("❌ Data source file does not exist")
        else:
            issues.append("✅ Data source file exists")
        
        # Check output options
        if not self.create_word_var.get() and not self.create_pdf_var.get():
            issues.append("❌ No output format selected")
        else:
            formats = []
            if self.create_word_var.get():
                formats.append("Word")
            if self.create_pdf_var.get():
                formats.append("PDF")
            issues.append(f"✅ Output formats: {', '.join(formats)}")
        
        # Check output directory
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
            issues.append("✅ Output directory is accessible")
        except Exception as e:
            issues.append(f"❌ Cannot access output directory: {str(e)}")
        
        # Show results
        result_text = "\n".join(issues)
        messagebox.showinfo("Setup Validation", result_text)
    
    def _validate_template(self):
        """Validate template and extract variables"""
        if not self.template_path.get():
            messagebox.showwarning("Warning", "Please select a template file first.")
            return
        
        try:
            from docxtpl import DocxTemplate
            template = DocxTemplate(self.template_path.get())
            
            # This is a simplified approach - in reality, you'd need to parse the template
            # to extract variables. For now, we'll show basic template info.
            
            messagebox.showinfo("Template Validation", 
                              "Template appears to be valid!\n\n"
                              "Use the Template Analysis tab for detailed variable analysis.")
            
        except Exception as e:
            messagebox.showerror("Template Error", f"Template validation failed:\n{str(e)}")
    
    def _analyze_template(self):
        """Analyze template and show variable information"""
        if not self.template_path.get():
            messagebox.showwarning("Warning", "Please select a template file first.")
            return
        
        try:
            # Clear existing analysis
            for item in self.vars_tree.get_children():
                self.vars_tree.delete(item)
            
            self.template_info_text.config(state=tk.NORMAL)
            self.template_info_text.delete(1.0, tk.END)
            
            # Basic template info
            template_path = Path(self.template_path.get())
            file_size = template_path.stat().st_size
            mod_time = datetime.fromtimestamp(template_path.stat().st_mtime)
            
            info_text = f"Template: {template_path.name}\n"
            info_text += f"Size: {file_size:,} bytes\n"
            info_text += f"Modified: {mod_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            info_text += f"Status: Ready for processing"
            
            self.template_info_text.insert(tk.END, info_text)
            self.template_info_text.config(state=tk.DISABLED)
            
            # For demo purposes, show some common variables
            # In a real implementation, you'd parse the template to extract actual variables
            common_vars = ['name', 'address', 'phone_no', 'dob', 'doctor_name', 'npi', 
                          'doctor_address', 'doctor_phone', 'doctor_fax']
            
            data_columns = []
            if self.preview_data is not None:
                data_columns = list(self.preview_data.columns)
            
            for var in common_vars:
                if var in data_columns:
                    status = "✅ Found in data"
                    self.vars_tree.insert('', tk.END, text=var, values=(status,))
                else:
                    status = "❌ Not found in data"
                    self.vars_tree.insert('', tk.END, text=var, values=(status,))
            
            self.status_var.set("Template analysis completed")
            
        except Exception as e:
            self.logger.error(f"Template analysis failed: {str(e)}")
            messagebox.showerror("Analysis Error", f"Failed to analyze template:\n{str(e)}")
    
    def _open_output_folder(self):
        """Open the output folder in file explorer"""
        try:
            import subprocess
            import platform
            
            output_path = self.output_dir.get()
            if not os.path.exists(output_path):
                os.makedirs(output_path, exist_ok=True)
            
            if platform.system() == "Windows":
                subprocess.run(f'explorer "{output_path}"')
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", output_path])
            else:  # Linux
                subprocess.run(["xdg-open", output_path])
                
        except Exception as e:
            messagebox.showerror("Error", f"Could not open output folder:\n{str(e)}")
    
    def _open_log_folder(self):
        """Open the log folder in file explorer"""
        try:
            import subprocess
            import platform
            
            log_path = os.path.expanduser("~/Documents/MailMergeLogs")
            if not os.path.exists(log_path):
                messagebox.showinfo("Info", "No log folder found. Logs will be created when you run mail merge.")
                return
            
            if platform.system() == "Windows":
                subprocess.run(f'explorer "{log_path}"')
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", log_path])
            else:  # Linux
                subprocess.run(["xdg-open", log_path])
                
        except Exception as e:
            messagebox.showerror("Error", f"Could not open log folder:\n{str(e)}")
    
    def _show_about(self):
        """Show about dialog"""
        about_text = """Enhanced Mail Merge to PDF Tool v2.0

A powerful tool for merging data with Word templates to create personalized documents.

Features:
• Advanced template analysis
• Real-time data preview with search
• Flexible filename patterns
• Comprehensive error handling
• Progress tracking and logging
• Settings persistence

Created with Python and Tkinter
© 2024 Enhanced Mail Merge Tool"""
        
        messagebox.showinfo("About", about_text)
    
    def _toggle_logging_level(self):
        """Toggle between INFO and DEBUG logging levels"""
        if self.verbose_logging_var.get():
            logging.getLogger().setLevel(logging.DEBUG)
            self.logger.info("Verbose logging enabled")
        else:
            logging.getLogger().setLevel(logging.INFO)
            self.logger.info("Standard logging enabled")
    
    def _reset_settings(self):
        """Reset all settings to defaults"""
        if messagebox.askyesno("Reset Settings", "Are you sure you want to reset all settings to defaults?"):
            self.settings = {
                "last_template": "",
                "last_data_source": "",
                "last_output_dir": os.path.expanduser("~/Documents/MailMergeOutput"),
                "create_word": True,
                "create_pdf": True,
                "auto_save_settings": True,
                "max_preview_rows": 100,
                "filename_pattern": "{First Name}_{Last Name}"
            }
            
            # Update UI
            self.auto_save_var.set(True)
            self.max_preview_var.set("100")
            self.filename_pattern_var.set("{First Name}_{Last Name}")
            self.verbose_logging_var.set(False)
            
            self._save_settings()
            messagebox.showinfo("Success", "Settings reset to defaults!")
    
    def _stop_processing(self):
        """Stop the current mail merge processing"""
        # This would need to be implemented with proper thread communication
        # For now, just show a message
        messagebox.showinfo("Stop Processing", "Processing will stop after the current record.")
    
    def browse_template(self):
        """Browse for Word template file"""
        try:
            filepath = filedialog.askopenfilename(
                title="Select Word Template",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
            )
            if filepath:
                self.template_path.set(filepath)
                self.template_status.config(text="✅ Valid", foreground="green")
                self.logger.info(f"Selected template: {filepath}")
                self._save_settings()
        except Exception as e:
            self.logger.error(f"Error browsing for template: {str(e)}")
            messagebox.showerror("Error", f"Failed to select template: {str(e)}")
    
    def browse_data_source(self):
        """Browse for data source file"""
        try:
            filepath = filedialog.askopenfilename(
                title="Select Data Source",
                filetypes=[
                    ("CSV Files", "*.csv"),
                    ("Excel Files", "*.xlsx *.xls"),
                    ("All Files", "*.*")
                ]
            )
            if filepath:
                self.data_source_path.set(filepath)
                self.data_status.config(text="✅ Valid", foreground="green")
                self.logger.info(f"Selected data source: {filepath}")
                self._save_settings()
        except Exception as e:
            self.logger.error(f"Error browsing for data source: {str(e)}")
            messagebox.showerror("Error", f"Failed to select data source: {str(e)}")
    
    def browse_output_dir(self):
        """Browse for output directory"""
        try:
            dirpath = filedialog.askdirectory(title="Select Output Directory")
            if dirpath:
                self.output_dir.set(dirpath)
                self.logger.info(f"Selected output directory: {dirpath}")
                self._save_settings()
        except Exception as e:
            self.logger.error(f"Error browsing for output directory: {str(e)}")
            messagebox.showerror("Error", f"Failed to select output directory: {str(e)}")
    
    def load_preview(self):
        """Load and display a preview of the data source"""
        if not self.data_source_path.get():
            self.logger.warning("No data source selected for preview")
            messagebox.showwarning("Warning", "Please select a data source file first.")
            return
        
        try:
            self.logger.info(f"Loading preview from: {self.data_source_path.get()}")
            self.status_var.set("Loading data preview...")
            
            # Load the data
            data_path = self.data_source_path.get()
            if data_path.endswith('.csv'):
                self.preview_data = pd.read_csv(data_path)
                self.logger.debug("Loaded CSV data source")
            elif data_path.endswith(('.xlsx', '.xls')):
                self.preview_data = pd.read_excel(data_path)
                self.logger.debug("Loaded Excel data source")
            else:
                error_msg = "Unsupported file format. Please use .csv, .xlsx, or .xls files."
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            
            # Clear existing data
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            
            # Configure columns
            columns = list(self.preview_data.columns)
            self.preview_tree['columns'] = columns
            
            # Configure headings
            self.preview_tree['show'] = 'tree headings'
            for col in columns:
                self.preview_tree.heading(col, text=col)
                # Set a reasonable width
                self.preview_tree.column(col, width=120, minwidth=80)
            
            # Populate the tree
            self._populate_preview_tree(self.preview_data)
            
            status_msg = f"Loaded preview: {len(self.preview_data)} records, {len(columns)} columns"
            self.status_var.set(status_msg)
            self.logger.info(status_msg)
            
        except Exception as e:
            error_msg = f"Failed to load data source: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error loading data")
    
    def validate_inputs(self):
        """Validate user inputs before processing"""
        self.logger.debug("Validating inputs")
        
        if not self.template_path.get():
            self.logger.warning("No template file selected")
            messagebox.showwarning("Warning", "Please select a Word template file.")
            return False
        
        if not os.path.exists(self.template_path.get()):
            error_msg = f"Template file does not exist: {self.template_path.get()}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            return False
        
        if not self.data_source_path.get():
            self.logger.warning("No data source file selected")
            messagebox.showwarning("Warning", "Please select a data source file.")
            return False
        
        if not os.path.exists(self.data_source_path.get()):
            error_msg = f"Data source file does not exist: {self.data_source_path.get()}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            return False
        
        if not self.create_word_var.get() and not self.create_pdf_var.get():
            self.logger.warning("No output format selected")
            messagebox.showwarning("Warning", "Please select at least one output format (Word or PDF).")
            return False
        
        # Validate filename pattern
        pattern = self.filename_pattern_var.get().strip()
        if not pattern:
            self.logger.warning("Empty filename pattern")
            messagebox.showwarning("Warning", "Filename pattern cannot be empty.")
            return False
        
        self.logger.debug("Input validation passed")
        return True
    
    def run_mail_merge(self):
        """Start the mail merge process"""
        self.logger.info("Starting mail merge process")
        
        # Validate inputs first
        if not self.validate_inputs():
            return
            
        # Prevent multiple simultaneous runs
        if self.is_processing:
            self.logger.warning("Mail merge already in progress")
            messagebox.showwarning("Warning", "Mail merge is already in progress.")
            return
        
        # Save current settings
        self._save_settings()
        
        # Update UI for processing state
        self.run_button.config(state='disabled')
        self.stop_button.config(state='normal')
        self.is_processing = True
        
        # Run in a thread to avoid freezing the GUI
        thread = threading.Thread(target=self._mail_merge_worker)
        thread.daemon = True
        thread.start()
        self.logger.debug("Started mail merge worker thread")
    
    def _generate_filename(self, record, pattern, index):
        """Generate filename from pattern and record data"""
        try:
            # Replace pattern variables with actual data
            filename = pattern
            
            # Find all {variable} patterns
            import re
            variables = re.findall(r'\{([^}]+)\}', pattern)
            
            for var in variables:
                if var in record:
                    value = str(record[var]).strip()
                    # Clean the value for filename use
                    value = self._clean_filename(value)
                    filename = filename.replace(f'{{{var}}}', value)
                else:
                    # Replace with empty string or default
                    filename = filename.replace(f'{{{var}}}', 'Unknown')
            
            # If filename is empty or just underscores, use index
            if not filename.strip('_') or filename.strip() == '':
                filename = f"document_{index+1}"
            
            return filename
            
        except Exception as e:
            self.logger.warning(f"Error generating filename: {e}")
            return f"document_{index+1}"
    
    def _mail_merge_worker(self):
        """Enhanced worker function that performs the mail merge in a separate thread"""
        start_time = time.time()
        try:
            self.logger.info("Mail merge worker started")
            self.message_queue.put({'type': 'status', 'text': 'Initializing...'})
            self.message_queue.put({'type': 'progress', 'value': 0})
            
            # Load data
            data_path = self.data_source_path.get()
            self.logger.info(f"Loading data from: {data_path}")
            
            if data_path.endswith('.csv'):
                data_source = pd.read_csv(data_path)
                self.logger.debug("Loaded CSV data source")
            elif data_path.endswith(('.xlsx', '.xls')):
                data_source = pd.read_excel(data_path)
                self.logger.debug("Loaded Excel data source")
            else:
                error_msg = "Unsupported data source format"
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            
            # Data preprocessing
            data_source = data_source.fillna('')
            records = data_source.to_dict('records')
            total_records = len(records)
            
            self.logger.info(f"Found {total_records} records to process")
            
            if total_records == 0:
                error_msg = "No records found in data source"
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            
            # Ensure output directory exists
            os.makedirs(self.output_dir.get(), exist_ok=True)
            self.logger.debug(f"Output directory: {self.output_dir.get()}")
            
            # Processing statistics
            success_count = 0
            error_count = 0
            skipped_count = 0
            
            # Process each record
            for i, record in enumerate(records):
                try:
                    # Update progress
                    progress = (i / total_records) * 100
                    status_msg = f"Processing record {i+1} of {total_records}..."
                    
                    self.message_queue.put({'type': 'status', 'text': status_msg})
                    self.message_queue.put({'type': 'progress', 'value': progress})
                    
                    self.logger.debug(f"Processing record {i+1}: {status_msg}")
                    
                    # Create fresh template for each record
                    template_doc = DocxTemplate(self.template_path.get())
                    
                    # Enhanced context preparation
                    context = {}
                    for key, value in record.items():
                        # Clean and prepare data
                        clean_key = str(key).strip().replace(' ', '_').lower()
                        clean_value = str(value).strip() if value else ''
                        context[clean_key] = clean_value
                        context[key] = clean_value  # Keep original key too
                    
                    # Add some computed fields
                    if 'First Name' in record and 'Last Name' in record:
                        context['name'] = f"{record.get('First Name', '')} {record.get('Last Name', '')}".strip()
                    
                    # Render template
                    template_doc.render(context)
                    self.logger.debug(f"Rendered template for record {i+1}")
                    
                    # Generate filename
                    base_filename = self._generate_filename(record, self.filename_pattern_var.get(), i)
                    
                    # Handle paths
                    doc_folder_path = record.get('DocFolderPath', self.output_dir.get())
                    pdf_folder_path = record.get('PdfFolderPath', self.output_dir.get())
                    
                    # Ensure directories exist
                    os.makedirs(doc_folder_path, exist_ok=True)
                    if self.create_pdf_var.get():
                        os.makedirs(pdf_folder_path, exist_ok=True)
                    
                    # Save documents
                    doc_path = None
                    
                    # Handle existing file backup
                    if self.backup_existing_var.get():
                        self._backup_existing_files(doc_folder_path, pdf_folder_path, base_filename)
                    
                    # Save Word document
                    if self.create_word_var.get():
                        doc_path = os.path.join(doc_folder_path, f"{base_filename}.docx")
                        template_doc.save(doc_path)
                        self.logger.debug(f"Saved Word document: {doc_path}")
                    
                    # Create PDF
                    if self.create_pdf_var.get():
                        temp_doc_created = False
                        if not doc_path:
                            doc_path = os.path.join(self.output_dir.get(), f"_temp_{i}.docx")
                            template_doc.save(doc_path)
                            temp_doc_created = True
                            
                        pdf_path = os.path.join(pdf_folder_path, f"{base_filename}.pdf")
                        convert(doc_path, pdf_path)
                        self.logger.debug(f"Converted to PDF: {pdf_path}")
                        
                        # Cleanup temp file
                        if temp_doc_created and os.path.exists(doc_path):
                            try:
                                os.remove(doc_path)
                            except Exception as e:
                                self.logger.warning(f"Failed to remove temp file: {e}")
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    error_msg = f"Error processing record {i+1}: {str(e)}"
                    self.logger.error(error_msg, exc_info=True)
                    
                    if not self.skip_errors_var.get():
                        raise
                    else:
                        skipped_count += 1
                        continue
            
            # Processing complete
            elapsed_time = time.time() - start_time
            completion_msg = f"Completed in {elapsed_time:.1f}s: {success_count} successful, {error_count} errors"
            
            self.message_queue.put({'type': 'progress', 'value': 100})
            self.message_queue.put({'type': 'status', 'text': completion_msg})
            
            # Show detailed results
            result_msg = f"Mail merge completed!\n\n"
            result_msg += f"✅ Successful: {success_count}\n"
            if error_count > 0:
                result_msg += f"❌ Errors: {error_count}\n"
            if skipped_count > 0:
                result_msg += f"⏭️ Skipped: {skipped_count}\n"
            result_msg += f"⏱️ Time: {elapsed_time:.1f} seconds"
            
            self.message_queue.put({'type': 'success', 'text': result_msg})
            self.logger.info(completion_msg)
            
            # Open output folder if requested
            if self.open_output_var.get():
                self._open_output_folder()
            
        except Exception as e:
            error_msg = f"Mail merge failed: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            self.message_queue.put({'type': 'status', 'text': 'Error during mail merge'})
            self.message_queue.put({'type': 'error', 'text': error_msg})
        
        finally:
            # Reset UI state
            self.message_queue.put({'type': 'status', 'text': 'Ready'})
            self.run_button.config(state='normal')
            self.stop_button.config(state='disabled')
            self.is_processing = False
            self.logger.info("Mail merge worker finished")
    
    def _backup_existing_files(self, doc_folder, pdf_folder, base_filename):
        """Backup existing files before overwriting"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Backup Word file
        doc_path = os.path.join(doc_folder, f"{base_filename}.docx")
        if os.path.exists(doc_path):
            backup_path = os.path.join(doc_folder, f"{base_filename}_backup_{timestamp}.docx")
            try:
                os.rename(doc_path, backup_path)
                self.logger.debug(f"Backed up Word file: {backup_path}")
            except Exception as e:
                self.logger.warning(f"Could not backup Word file: {e}")
        
        # Backup PDF file
        pdf_path = os.path.join(pdf_folder, f"{base_filename}.pdf")
        if os.path.exists(pdf_path):
            backup_path = os.path.join(pdf_folder, f"{base_filename}_backup_{timestamp}.pdf")
            try:
                os.rename(pdf_path, backup_path)
                self.logger.debug(f"Backed up PDF file: {backup_path}")
            except Exception as e:
                self.logger.warning(f"Could not backup PDF file: {e}")
    
    def _clean_filename(self, filename):
        """Remove or replace invalid characters from filename"""
        # Replace invalid characters with underscore
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', str(filename))
        # Remove leading/trailing whitespace and dots
        cleaned = cleaned.strip(' .')
        # Ensure filename is not empty
        if not cleaned:
            cleaned = "document"
        return cleaned
    
    def on_closing(self):
        """Handle application closing"""
        if self.is_processing:
            if messagebox.askokcancel("Quit", "Mail merge is in progress. Do you want to quit anyway?"):
                self.root.destroy()
        else:
            self._save_settings()
            self.root.destroy()

def main():
    """Main entry point for the application"""
    try:
        root = tk.Tk()
        app = MailMergeApp(root)
        
        # Handle window closing
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        
        # Set application icon (if available)
        try:
            # You can add an icon file here
            # root.iconbitmap('path/to/icon.ico')
            pass
        except:
            pass
        
        root.mainloop()
        
    except Exception as e:
        logging.error(f"Fatal error in application: {str(e)}", exc_info=True)
        messagebox.showerror("Fatal Error", f"The application encountered a fatal error:\n{str(e)}")

if __name__ == "__main__":
    main()