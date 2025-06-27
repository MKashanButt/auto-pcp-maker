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

class MailMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Merge to PDF Tool")
        self.root.geometry("700x600")
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
        
        # Create the main UI
        self._create_ui()
    
    def _setup_logging(self):
        """Configure logging for the application"""
        log_dir = os.path.expanduser("~/Documents/MailMergeLogs")
        os.makedirs(log_dir, exist_ok=True)
        
        # Create a timestamped log file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"mailmerge_{timestamp}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
    
    def _create_ui(self):
        """Create the main user interface"""
        self.logger.debug("Creating UI components")
        try:
            # Create a notebook (tabbed interface)
            notebook = ttk.Notebook(self.root)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Tab 1: Configuration
            config_frame = ttk.Frame(notebook)
            notebook.add(config_frame, text="Configuration")
            
            # Tab 2: Preview
            preview_frame = ttk.Frame(notebook)
            notebook.add(preview_frame, text="Data Preview")
            
            # Tab 3: Help
            help_frame = ttk.Frame(notebook)
            notebook.add(help_frame, text="Help")
            
            # Setup each tab
            self._setup_config_tab(config_frame)
            self._setup_preview_tab(preview_frame)
            self._setup_help_tab(help_frame)
            
            # Status bar
            status_frame = ttk.Frame(self.root)
            status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
            
            self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, mode="determinate")
            self.progress_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
            
            status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
            status_label.pack(fill=tk.X, side=tk.LEFT)
            
            self.logger.info("UI created successfully")
        except Exception as e:
            self.logger.error(f"Error creating UI: {str(e)}")
            raise
    
    def _setup_config_tab(self, parent):
        """Configure the Configuration tab"""
        try:
            # File selection frame
            file_frame = ttk.LabelFrame(parent, text="File Selection")
            file_frame.pack(fill=tk.X, padx=10, pady=10, expand=False)
            
            # Template selection
            ttk.Label(file_frame, text="Word Template:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            ttk.Entry(file_frame, textvariable=self.template_path, width=50).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
            ttk.Button(file_frame, text="Browse...", command=self.browse_template).grid(row=0, column=2, padx=5, pady=5)
            
            # Data source selection
            ttk.Label(file_frame, text="Data Source:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
            ttk.Entry(file_frame, textvariable=self.data_source_path, width=50).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
            ttk.Button(file_frame, text="Browse...", command=self.browse_data_source).grid(row=1, column=2, padx=5, pady=5)
            
            # Output directory
            ttk.Label(file_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
            ttk.Entry(file_frame, textvariable=self.output_dir, width=50).grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
            ttk.Button(file_frame, text="Browse...", command=self.browse_output_dir).grid(row=2, column=2, padx=5, pady=5)
            
            file_frame.columnconfigure(1, weight=1)
            
            # Options frame
            options_frame = ttk.LabelFrame(parent, text="Options")
            options_frame.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)
            
            # Custom options
            self.create_word_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(options_frame, text="Create Word Documents", variable=self.create_word_var).pack(anchor=tk.W, padx=5, pady=5)
            
            self.create_pdf_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(options_frame, text="Create PDF Documents", variable=self.create_pdf_var).pack(anchor=tk.W, padx=5, pady=5)
            
            # Action buttons frame
            button_frame = ttk.Frame(parent)
            button_frame.pack(fill=tk.X, padx=10, pady=10)
            
            self.run_button = ttk.Button(button_frame, text="Run Mail Merge", command=self.run_mail_merge)
            self.run_button.pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="Load Preview", command=self.load_preview).pack(side=tk.RIGHT, padx=5)
            
            self.logger.debug("Configuration tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up configuration tab: {str(e)}")
            raise
    
    def _setup_preview_tab(self, parent):
        """Configure the Preview tab"""
        try:
            # Create a frame for the preview table
            preview_frame = ttk.Frame(parent)
            preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Create a Treeview widget for displaying data
            self.preview_tree = ttk.Treeview(preview_frame)
            
            # Add a scrollbar
            y_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
            x_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
            self.preview_tree.configure(yscroll=y_scrollbar.set, xscroll=x_scrollbar.set)
            
            # Pack the scrollbar and treeview
            y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Help text
            ttk.Label(parent, text="Click 'Load Preview' in the Configuration tab to see your data source.").pack(pady=5)
            
            self.logger.debug("Preview tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up preview tab: {str(e)}")
            raise
    
    def _setup_help_tab(self, parent):
        """Configure the Help tab"""
        try:
            help_text = """
            Mail Merge to PDF Tool Help
            
            [Help content remains the same as original]
            """
            
            text_widget = tk.Text(parent, wrap=tk.WORD, padx=10, pady=10)
            text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            text_widget.insert(tk.END, help_text)
            text_widget.config(state=tk.DISABLED)  # Make read-only
            
            self.logger.debug("Help tab setup completed")
        except Exception as e:
            self.logger.error(f"Error setting up help tab: {str(e)}")
            raise
    
    def browse_template(self):
        """Browse for Word template file"""
        try:
            filepath = filedialog.askopenfilename(
                title="Select Word Template",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
            )
            if filepath:
                self.template_path.set(filepath)
                self.logger.info(f"Selected template: {filepath}")
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
                self.logger.info(f"Selected data source: {filepath}")
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
            self.preview_tree['show'] = 'headings'  # Hide the first empty column
            for col in columns:
                self.preview_tree.heading(col, text=col)
                # Set a reasonable width
                self.preview_tree.column(col, width=100, minwidth=50)
            
            # Add data rows (limit to first 100 for performance)
            max_rows = min(100, len(self.preview_data))
            for i in range(max_rows):
                row = self.preview_data.iloc[i].fillna('').astype(str).tolist()  # Handle NaN values
                self.preview_tree.insert('', tk.END, values=row)
            
            status_msg = f"Loaded {max_rows} of {len(self.preview_data)} records for preview"
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
        
        # Disable the run button during processing
        self.run_button.config(state='disabled')
        self.is_processing = True
        
        # Run in a thread to avoid freezing the GUI
        thread = threading.Thread(target=self._mail_merge_worker)
        thread.daemon = True
        thread.start()
        self.logger.debug("Started mail merge worker thread")
    
    def _mail_merge_worker(self):
        """Worker function that performs the mail merge in a separate thread"""
        try:
            self.logger.info("Mail merge worker started")
            self.root.after(0, lambda: self.status_var.set("Loading data..."))
            self.root.after(0, lambda: self.progress_var.set(0))
            
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
            
            # Fill NaN values with empty strings
            data_source = data_source.fillna('')
            
            # Convert DataFrame to list of dictionaries
            records = data_source.to_dict('records')
            total_records = len(records)
            self.logger.info(f"Found {total_records} records to process")
            
            if total_records == 0:
                error_msg = "No records found in data source"
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            
            # Ensure base output directory exists
            os.makedirs(self.output_dir.get(), exist_ok=True)
            self.logger.debug(f"Output directory: {self.output_dir.get()}")
            
            # Process each record
            success_count = 0
            for i, record in enumerate(records):
                try:
                    # Update UI (thread-safe)
                    status_msg = f"Processing record {i+1} of {total_records}..."
                    self.root.after(0, lambda: self.status_var.set(status_msg))
                    self.root.after(0, lambda: self.progress_var.set((i / total_records) * 100))
                    self.logger.debug(status_msg)
                    
                    # Validate required fields
                    required_fields = ['First Name', 'Last Name', 'DOB', 'Address', 'Phone NO', 
                                    'Doc Name', 'NPI', 'DOC Address', 'Doc Phone no']
                    missing_fields = [field for field in required_fields if field not in record or not record[field]]
                    
                    if missing_fields:
                        error_msg = f"Record {i+1} missing required fields: {', '.join(missing_fields)}"
                        self.logger.warning(error_msg)
                        continue
                    
                    # Create fresh template for each record
                    template_doc = DocxTemplate(self.template_path.get())
                    
                    # Prepare data for template
                    context = {
                        'name': f"{record.get('First Name', '')} {record.get('Last Name', '')}",
                        'dob': record.get('DOB', ''),
                        'address': record.get('Address', ''),
                        'phone_no': record.get('Phone NO', ''),
                        'doctor_name': record.get('Doc Name', ''),
                        'npi': record.get('NPI', ''),
                        'doctor_address': record.get('DOC Address', ''),
                        'doctor_phone': record.get('Doc Phone no', ''),
                        'doctor_fax': record.get('Fax no', ''),
                    }
                    
                    # Render template with record data
                    template_doc.render(context)
                    self.logger.debug(f"Rendered template for record {i+1}")
                    
                    # Generate filename from patient name
                    first_name = record.get('First Name', '').strip()
                    last_name = record.get('Last Name', '').strip()
                    base_filename = f"{first_name}_{last_name}".replace(' ', '_')
                    
                    if not base_filename or base_filename == '_':
                        base_filename = f"document_{i+1}"
                        self.logger.warning(f"Using default filename for record {i+1}")
                    
                    # Determine paths
                    doc_folder_path = record.get('DocFolderPath', self.output_dir.get())
                    pdf_folder_path = record.get('PdfFolderPath', self.output_dir.get())
                    
                    # Ensure directories exist
                    try:
                        os.makedirs(doc_folder_path, exist_ok=True)
                        if self.create_pdf_var.get():
                            os.makedirs(pdf_folder_path, exist_ok=True)
                    except Exception as e:
                        self.logger.error(f"Failed to create directory: {str(e)}")
                        continue
                    
                    # Save Word document if option selected
                    doc_path = None
                    if self.create_word_var.get():
                        doc_path = os.path.join(doc_folder_path, f"{base_filename}.docx")
                        template_doc.save(doc_path)
                        self.logger.debug(f"Saved Word document: {doc_path}")
                    
                    # Create PDF if option selected
                    if self.create_pdf_var.get():
                        # If we didn't create a Word doc but need a PDF, create a temporary Word doc
                        temp_doc_created = False
                        if not doc_path:
                            doc_path = os.path.join(self.output_dir.get(), f"_temp_{i}.docx")
                            template_doc.save(doc_path)
                            temp_doc_created = True
                            self.logger.debug(f"Created temporary Word document: {doc_path}")
                            
                        pdf_path = os.path.join(pdf_folder_path, f"{base_filename}.pdf")
                        try:
                            convert(doc_path, pdf_path)
                            self.logger.debug(f"Converted to PDF: {pdf_path}")
                        except Exception as e:
                            self.logger.error(f"Failed to convert to PDF: {str(e)}", exc_info=True)
                            raise
                        
                        # Remove temporary Word document if we created one
                        if temp_doc_created and os.path.exists(doc_path):
                            try:
                                os.remove(doc_path)
                            except Exception as e:
                                self.logger.warning(f"Failed to remove temporary document: {str(e)}")
                    
                    success_count += 1
                    
                except Exception as e:
                    self.logger.error(f"Error processing record {i+1}: {str(e)}", exc_info=True)
                    continue
            
            # Complete
            self.root.after(0, lambda: self.progress_var.set(100))
            completion_msg = f"Completed processing {success_count}/{total_records} records successfully"
            self.root.after(0, lambda: self.status_var.set(completion_msg))
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Mail merge completed!\n{success_count} of {total_records} documents were processed successfully."))
            self.logger.info(completion_msg)
            
        except Exception as e:
            error_msg = f"Mail merge failed: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            self.root.after(0, lambda: self.status_var.set("Error during mail merge"))
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
        
        finally:
            self.root.after(0, lambda: self.run_button.config(state='normal'))
            self.is_processing = False
            self.logger.info("Mail merge worker finished")
    
    def _clean_filename(self, filename):
        """Remove or replace invalid characters from filename"""
        import re
        # Replace invalid characters with underscore
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', str(filename))
        # Remove leading/trailing whitespace and dots
        cleaned = cleaned.strip(' .')
        # Ensure filename is not empty
        if not cleaned:
            cleaned = "document"
        return cleaned

def main():
    """Main entry point for the application"""
    try:
        root = tk.Tk()
        app = MailMergeApp(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Fatal error in application: {str(e)}", exc_info=True)
        messagebox.showerror("Fatal Error", f"The application encountered a fatal error:\n{str(e)}")

if __name__ == "__main__":
    main()