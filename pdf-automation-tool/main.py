"""
PDF Title Block Automation Tool for Leviat Engineering Team
Main application entry point
"""

import json
import os
import threading
import tkinter as tk
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox

from pdf_utils import generate_preview_image, get_page_count
import pdf_logic


def parse_page_range(range_string: str) -> set:
    """
    Parse a page range string into a set of 0-based page indices.
    
    Examples:
        "1" -> {0}
        "1, 3" -> {0, 2}
        "1, 3-5" -> {0, 2, 3, 4}
        "2-4, 7, 9-10" -> {1, 2, 3, 6, 8, 9}
    
    Args:
        range_string: A string like "1, 3-5" (1-based page numbers)
        
    Returns:
        A set of 0-based page indices
    """
    if not range_string or not range_string.strip():
        return set()
    
    excluded_indices = set()
    
    # Split by commas
    parts = range_string.split(',')
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
        
        try:
            if '-' in part:
                # Handle range like "3-5"
                range_parts = part.split('-')
                if len(range_parts) == 2:
                    start = int(range_parts[0].strip())
                    end = int(range_parts[1].strip())
                    # Convert 1-based to 0-based and add all pages in range
                    for page in range(start, end + 1):
                        if page >= 1:  # Only valid page numbers
                            excluded_indices.add(page - 1)  # Convert to 0-based
            else:
                # Handle single number like "1"
                page = int(part)
                if page >= 1:  # Only valid page numbers
                    excluded_indices.add(page - 1)  # Convert to 0-based
        except ValueError:
            # Ignore invalid entries (non-numbers)
            print(f"Warning: Ignoring invalid page specification: '{part}'")
            continue
    
    return excluded_indices


class ConfigManager:
    """Handles loading and saving configuration settings"""
    
    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.config: Dict = {}
        self.load_config()
    
    def load_config(self) -> None:
        """Load configuration from JSON file"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                # Create default config if file doesn't exist
                self.config = {
                    "Company Name": "Leviat",
                    "Default Output Folder": "Desktop",
                    "Window": {
                        "Width": 1200,
                        "Height": 800,
                        "Min_Width": 900,
                        "Min_Height": 600
                    },
                    "Preview": {
                        "Background_Color": "#2B2B2B"
                    }
                }
                self.save_config()
        except Exception as e:
            print(f"Error loading config: {e}")
            self.config = {}
    
    def save_config(self) -> None:
        """Save configuration to JSON file"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def get(self, key: str, default=None):
        """Get configuration value by key (supports dot notation)"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k, default)
            else:
                return default
        return value if value is not None else default


class PDFAutomationApp(ctk.CTk):
    """Main application class for PDF Title Block Automation Tool"""
    
    def __init__(self):
        super().__init__()
        
        # Load configuration
        self.config = ConfigManager()
        
        # Application state
        self.selected_pdf_path: Optional[str] = None
        self.current_file_path: Optional[str] = None  # Alias for selected_pdf_path
        self.selected_files: List[str] = []  # List of selected file paths for batch processing
        self.template_path: Optional[str] = None  # Path to title block template
        self.preview_image: Optional[ctk.CTkImage] = None
        
        # Navigation state
        self.current_file_index: int = 0  # Index of current file being previewed
        self.current_page_index: int = 0  # Index of current page being previewed
        self.current_file_page_count: int = 0  # Total pages in current file
        self.project_data = {
            'project_name': '',
            'client_name': '',
            'date': '',
            'drawn_by': ''
        }
        
        # Setup window
        self.setup_window()
        
        # Setup UI components
        self.setup_ui()
        
        # Check for templates on startup (show warning after UI is ready)
        if not self.template_files:
            self.after(500, self._show_library_warning)
    
    def _show_library_warning(self) -> None:
        """Show warning about missing templates on startup"""
        app_dir = Path(__file__).parent
        library_path = app_dir / "library"
        
        messagebox.showwarning(
            "No Templates Found",
            f"The library folder is empty or missing.\n\n"
            f"Please add your title block template PDF files to:\n"
            f"{library_path}\n\n"
            f"Then restart the application to see them in the dropdown."
        )
        
    def setup_window(self) -> None:
        """Configure the main window properties"""
        self.title("PDF Title Block Automation - Leviat")
        
        width = self.config.get('Window.Width', 1200)
        height = self.config.get('Window.Height', 800)
        min_width = self.config.get('Window.Min_Width', 900)
        min_height = self.config.get('Window.Min_Height', 600)
        
        self.geometry(f"{width}x{height}")
        self.minsize(min_width, min_height)
        
        # Set appearance mode and color theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
    
    def setup_ui(self) -> None:
        """Create and arrange UI components"""
        # Main container with padding
        self.grid_columnconfigure(0, weight=0)  # Sidebar - fixed width
        self.grid_columnconfigure(1, weight=1)  # Main area - flexible
        self.grid_rowconfigure(0, weight=1)  # Content area - flexible
        self.grid_rowconfigure(1, weight=0)  # Bottom bar - fixed height
        
        # Left Sidebar
        self.create_sidebar()
        
        # Right Main Area (PDF Preview)
        self.create_preview_area()
        
        # Bottom Bar
        self.create_bottom_bar()
    
    def create_sidebar(self) -> None:
        """Create the left sidebar with input fields"""
        sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 5), pady=5)
        sidebar.grid_propagate(False)
        
        # Sidebar title
        title_label = ctk.CTkLabel(
            sidebar,
            text="Title Block Information",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=(20, 20))
        
        # Master Template Selection Dropdown
        template_label = ctk.CTkLabel(sidebar, text="Select Master Template:", anchor="w")
        template_label.pack(pady=(0, 5), padx=20, fill="x")
        
        # Scan library folder and populate dropdown
        self.template_files = self._scan_library_folder()
        template_display_names = [os.path.basename(f) for f in self.template_files] if self.template_files else ["No templates found"]
        
        self.template_dropdown = ctk.CTkComboBox(
            sidebar,
            values=template_display_names,
            height=35,
            font=ctk.CTkFont(size=13),
            dropdown_font=ctk.CTkFont(size=12),
            state="readonly" if self.template_files else "disabled",
            command=self._on_template_selected
        )
        self.template_dropdown.pack(pady=(0, 20), padx=20, fill="x")
        
        # Set default to first template if available
        if self.template_files:
            self.template_dropdown.set(template_display_names[0])
            self.template_path = self.template_files[0]
        
        # Select PDF Button
        self.select_pdf_btn = ctk.CTkButton(
            sidebar,
            text="Select Customer Drawing",
            command=self.select_pdf,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.select_pdf_btn.pack(pady=(0, 10), padx=20, fill="x")
        
        # Status label to show number of files selected
        self.status_label = ctk.CTkLabel(
            sidebar,
            text="No files selected",
            font=ctk.CTkFont(size=12),
            text_color="#888888"
        )
        self.status_label.pack(pady=(0, 20), padx=20)
        
        # Project Name
        project_label = ctk.CTkLabel(sidebar, text="Project Name:", anchor="w")
        project_label.pack(pady=(0, 5), padx=20, fill="x")
        
        self.project_name_entry = ctk.CTkEntry(
            sidebar,
            placeholder_text="Enter project name",
            height=35
        )
        self.project_name_entry.pack(pady=(0, 15), padx=20, fill="x")
        self.project_name_entry.bind('<KeyRelease>', self.update_project_data)
        
        # Client Name
        client_label = ctk.CTkLabel(sidebar, text="Client Name:", anchor="w")
        client_label.pack(pady=(0, 5), padx=20, fill="x")
        
        self.client_name_entry = ctk.CTkEntry(
            sidebar,
            placeholder_text="Enter client name",
            height=35
        )
        self.client_name_entry.pack(pady=(0, 15), padx=20, fill="x")
        self.client_name_entry.bind('<KeyRelease>', self.update_project_data)
        
        # Date
        date_label = ctk.CTkLabel(sidebar, text="Date:", anchor="w")
        date_label.pack(pady=(0, 5), padx=20, fill="x")
        
        self.date_entry = ctk.CTkEntry(
            sidebar,
            placeholder_text="MM/DD/YYYY",
            height=35
        )
        self.date_entry.pack(pady=(0, 15), padx=20, fill="x")
        self.date_entry.bind('<KeyRelease>', self.update_project_data)
        
        # Drawn By
        drawn_by_label = ctk.CTkLabel(sidebar, text="Drawn By:", anchor="w")
        drawn_by_label.pack(pady=(0, 5), padx=20, fill="x")
        
        self.drawn_by_entry = ctk.CTkEntry(
            sidebar,
            placeholder_text="Enter name",
            height=35
        )
        self.drawn_by_entry.pack(pady=(0, 15), padx=20, fill="x")
        self.drawn_by_entry.bind('<KeyRelease>', self.update_project_data)
        
        # Output Destination Section
        output_label = ctk.CTkLabel(sidebar, text="Output Destination:", anchor="w")
        output_label.pack(pady=(10, 5), padx=20, fill="x")
        
        # Frame to hold entry and browse button
        output_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        output_frame.pack(pady=(0, 15), padx=20, fill="x")
        
        # Default output path
        default_output = os.path.join(os.path.expanduser('~'), 'Desktop', 'Leviat_Output')
        self.output_folder = default_output
        
        self.output_path_entry = ctk.CTkEntry(
            output_frame,
            height=35,
            font=ctk.CTkFont(size=11)
        )
        self.output_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.output_path_entry.insert(0, default_output)
        
        self.browse_output_btn = ctk.CTkButton(
            output_frame,
            text="Browse...",
            command=self._browse_output_folder,
            width=80,
            height=35,
            font=ctk.CTkFont(size=12)
        )
        self.browse_output_btn.pack(side="right")
        
        # Store sidebar reference
        self.sidebar = sidebar
    
    def create_preview_area(self) -> None:
        """Create the PDF preview area with navigation controls"""
        preview_frame = ctk.CTkFrame(self, corner_radius=10)
        preview_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 5), pady=(5, 5))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(1, weight=1)  # Canvas row expands
        preview_frame.grid_rowconfigure(2, weight=0)  # Navigation bar fixed
        
        # Preview title label
        preview_title = ctk.CTkLabel(
            preview_frame,
            text="PDF Preview",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        preview_title.grid(row=0, column=0, pady=10, sticky="n")
        
        # Use a Canvas for proper image centering and display
        bg_color = self.config.get('Preview.Background_Color', '#2B2B2B')
        
        self.preview_canvas = tk.Canvas(
            preview_frame,
            bg=bg_color,
            highlightthickness=0,
            relief="flat"
        )
        self.preview_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 5))
        
        # Display initial "No PDF loaded" text
        self.preview_canvas.bind('<Configure>', self._on_canvas_configure)
        self._canvas_text_id = None
        self._canvas_image_id = None
        self._show_canvas_message("No PDF loaded")
        
        # Navigation Control Bar
        nav_bar = ctk.CTkFrame(preview_frame, height=50, fg_color="transparent")
        nav_bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        nav_bar.grid_columnconfigure(0, weight=1)  # Left spacer
        nav_bar.grid_columnconfigure(1, weight=0)  # File controls
        nav_bar.grid_columnconfigure(2, weight=0)  # Separator
        nav_bar.grid_columnconfigure(3, weight=0)  # Page controls
        nav_bar.grid_columnconfigure(4, weight=1)  # Right spacer
        
        # File Navigation Controls
        file_nav_frame = ctk.CTkFrame(nav_bar, fg_color="transparent")
        file_nav_frame.grid(row=0, column=1, padx=10)
        
        self.prev_file_btn = ctk.CTkButton(
            file_nav_frame,
            text="<< Prev File",
            command=self._prev_file,
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            state="disabled"
        )
        self.prev_file_btn.pack(side="left", padx=(0, 10))
        
        self.file_label = ctk.CTkLabel(
            file_nav_frame,
            text="File 0 of 0",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=80
        )
        self.file_label.pack(side="left", padx=5)
        
        self.next_file_btn = ctk.CTkButton(
            file_nav_frame,
            text="Next File >>",
            command=self._next_file,
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            state="disabled"
        )
        self.next_file_btn.pack(side="left", padx=(10, 0))
        
        # Separator
        separator = ctk.CTkLabel(
            nav_bar,
            text="|",
            font=ctk.CTkFont(size=16),
            text_color="#666666"
        )
        separator.grid(row=0, column=2, padx=15)
        
        # Page Navigation Controls
        page_nav_frame = ctk.CTkFrame(nav_bar, fg_color="transparent")
        page_nav_frame.grid(row=0, column=3, padx=10)
        
        self.prev_page_btn = ctk.CTkButton(
            page_nav_frame,
            text="< Prev Page",
            command=self._prev_page,
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            state="disabled"
        )
        self.prev_page_btn.pack(side="left", padx=(0, 10))
        
        self.page_label = ctk.CTkLabel(
            page_nav_frame,
            text="Page 0 of 0",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=90
        )
        self.page_label.pack(side="left", padx=5)
        
        self.next_page_btn = ctk.CTkButton(
            page_nav_frame,
            text="Next Page >",
            command=self._next_page,
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            state="disabled"
        )
        self.next_page_btn.pack(side="left", padx=(10, 0))
        
        # Store preview frame reference
        self.preview_frame = preview_frame
        
        # Store reference for the PhotoImage to prevent garbage collection
        self._photo_image = None
    
    def create_bottom_bar(self) -> None:
        """Create the bottom action bar"""
        bottom_bar = ctk.CTkFrame(self, height=120, corner_radius=0)
        bottom_bar.grid(row=1, column=1, sticky="ew", padx=(0, 5), pady=(0, 5))
        bottom_bar.grid_propagate(False)
        bottom_bar.grid_columnconfigure(0, weight=1)
        bottom_bar.grid_columnconfigure(1, weight=0)
        
        # Progress bar (hidden by default)
        self.progress_bar = ctk.CTkProgressBar(
            bottom_bar,
            width=400,
            height=15,
            mode="determinate"
        )
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, columnspan=2, pady=(10, 5), padx=20, sticky="ew")
        self.progress_bar.grid_remove()  # Hide by default
        
        # Progress status label (hidden by default)
        self.progress_status_label = ctk.CTkLabel(
            bottom_bar,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="#AAAAAA"
        )
        self.progress_status_label.grid(row=1, column=0, columnspan=2, pady=(0, 5), padx=20, sticky="w")
        self.progress_status_label.grid_remove()  # Hide by default
        
        # Bottom controls frame (Exclude Pages + Process Button)
        controls_frame = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        controls_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=10)
        controls_frame.grid_columnconfigure(0, weight=1)
        
        # Exclude Pages input (left side)
        exclude_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        exclude_frame.pack(side="left", fill="x", expand=True)
        
        exclude_label = ctk.CTkLabel(
            exclude_frame,
            text="Exclude Pages (e.g., 1, 3-5):",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        exclude_label.pack(side="left", padx=(0, 10))
        
        self.exclude_pages_entry = ctk.CTkEntry(
            exclude_frame,
            placeholder_text="Leave empty to include all",
            width=200,
            height=35
        )
        self.exclude_pages_entry.pack(side="left")
        
        # Process & Save Button (right side)
        self.save_export_btn = ctk.CTkButton(
            controls_frame,
            text="Process & Save",
            command=self.save_and_export,
            height=40,
            width=150,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.save_export_btn.pack(side="right")
        
        # Store bottom bar reference
        self.bottom_bar = bottom_bar
    
    def select_pdf(self) -> None:
        """Handle PDF file selection - supports multiple file selection"""
        file_paths = filedialog.askopenfilenames(
            title="Select PDF Files (Shift-Click for multiple)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_paths:
            # Store list of selected files for batch processing
            self.selected_files = list(file_paths)
            
            # Reset navigation indices to first file, first page
            self.current_file_index = 0
            self.current_page_index = 0
            
            # Get page count for first file
            self.current_file_page_count = get_page_count(self.selected_files[0])
            
            # Update legacy single-file references (use first file)
            self.selected_pdf_path = self.selected_files[0]
            self.current_file_path = self.selected_files[0]
            
            # Update status label with file count
            n = len(self.selected_files)
            self.status_label.configure(
                text=f"{n} file{'s' if n != 1 else ''} selected",
                text_color="#00CC66"  # Green color for success
            )
            
            # Update navigation labels and button states
            self._update_navigation_ui()
            
            # Show loading message
            self._show_canvas_message("Loading preview...")
            self.update()
            
            # Load preview of the first page of the first file
            self._load_current_preview()
    
    def _scan_library_folder(self) -> List[str]:
        """
        Scan the /library/ folder for PDF template files.
        
        Returns:
            List of full file paths to PDF files found in the library folder.
        """
        # Get the library folder path (relative to the app directory)
        app_dir = Path(__file__).parent
        library_path = app_dir / "library"
        
        # Check if library folder exists
        if not library_path.exists():
            # Try to create it
            try:
                library_path.mkdir(parents=True, exist_ok=True)
                print(f"Created library folder at: {library_path}")
            except Exception as e:
                print(f"Could not create library folder: {e}")
            return []
        
        # Scan for PDF files
        pdf_files = sorted(library_path.glob("*.pdf"))
        
        if not pdf_files:
            print(f"No PDF templates found in: {library_path}")
            return []
        
        # Convert to string paths
        file_paths = [str(f) for f in pdf_files]
        print(f"Found {len(file_paths)} template(s) in library: {[os.path.basename(f) for f in file_paths]}")
        
        return file_paths
    
    def _on_template_selected(self, selected_name: str) -> None:
        """
        Handle template selection from dropdown.
        
        Args:
            selected_name: The filename selected in the dropdown
        """
        # Find the full path for the selected template
        for file_path in self.template_files:
            if os.path.basename(file_path) == selected_name:
                self.template_path = file_path
                print(f"Selected template: {file_path}")
                break
    
    def _browse_output_folder(self) -> None:
        """Open folder browser dialog to select output destination"""
        folder_path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_path_entry.get() or os.path.expanduser('~')
        )
        
        if folder_path:
            # Update the entry field
            self.output_path_entry.delete(0, "end")
            self.output_path_entry.insert(0, folder_path)
            
            # Update the output folder variable
            self.output_folder = folder_path
            print(f"Output folder set to: {folder_path}")
    
    def _check_library_and_warn(self) -> bool:
        """
        Check if library folder has templates and show warning if not.
        
        Returns:
            True if templates are available, False otherwise.
        """
        if not self.template_files:
            app_dir = Path(__file__).parent
            library_path = app_dir / "library"
            
            messagebox.showwarning(
                "No Templates Found",
                f"No PDF templates found in the library folder.\n\n"
                f"Please add your title block template PDF files to:\n"
                f"{library_path}\n\n"
                f"Then restart the application."
            )
            return False
        return True
    
    # ==================== Navigation Methods ====================
    
    def _update_navigation_ui(self) -> None:
        """Update navigation labels and button states based on current indices"""
        total_files = len(self.selected_files)
        total_pages = self.current_file_page_count
        
        # Update file label
        if total_files > 0:
            self.file_label.configure(text=f"File {self.current_file_index + 1} of {total_files}")
        else:
            self.file_label.configure(text="File 0 of 0")
        
        # Update page label
        if total_pages > 0:
            self.page_label.configure(text=f"Page {self.current_page_index + 1} of {total_pages}")
        else:
            self.page_label.configure(text="Page 0 of 0")
        
        # Update file button states
        if total_files > 1:
            self.prev_file_btn.configure(state="normal" if self.current_file_index > 0 else "disabled")
            self.next_file_btn.configure(state="normal" if self.current_file_index < total_files - 1 else "disabled")
        else:
            self.prev_file_btn.configure(state="disabled")
            self.next_file_btn.configure(state="disabled")
        
        # Update page button states
        if total_pages > 1:
            self.prev_page_btn.configure(state="normal" if self.current_page_index > 0 else "disabled")
            self.next_page_btn.configure(state="normal" if self.current_page_index < total_pages - 1 else "disabled")
        else:
            self.prev_page_btn.configure(state="disabled")
            self.next_page_btn.configure(state="disabled")
    
    def _load_current_preview(self) -> None:
        """Load the preview for the current file and page indices"""
        if not self.selected_files:
            return
        
        file_path = self.selected_files[self.current_file_index]
        
        # Show loading message
        self._show_canvas_message("Loading preview...")
        self.update()
        
        # Load preview in a separate thread
        threading.Thread(
            target=self._generate_and_display_preview_page,
            args=(file_path, self.current_page_index),
            daemon=True
        ).start()
    
    def _generate_and_display_preview_page(self, file_path: str, page_number: int) -> None:
        """Generate preview for a specific page and update GUI (runs in separate thread)"""
        try:
            # Get preview area dimensions
            self.preview_canvas.update_idletasks()
            width = self.preview_canvas.winfo_width()
            height = self.preview_canvas.winfo_height()
            
            if width <= 1:
                width = 800
            if height <= 1:
                height = 600
            
            padding = 40
            max_width = max(width - padding, 400)
            max_height = max(height - padding, 300)
            
            # Generate preview image for specific page
            pil_image = generate_preview_image(file_path, (max_width, max_height), page_number=page_number)
            
            if pil_image is None:
                self.after(0, self._show_preview_error, "Failed to generate preview image")
                return
            
            # Update GUI in main thread
            self.after(0, self._display_preview, pil_image, os.path.basename(file_path))
            
        except Exception as e:
            error_msg = f"Error loading PDF page: {str(e)}"
            self.after(0, self._show_preview_error, error_msg)
    
    def _prev_file(self) -> None:
        """Navigate to previous file"""
        if self.current_file_index > 0:
            self.current_file_index -= 1
            self.current_page_index = 0  # Reset to first page
            
            # Update page count for new file
            self.current_file_page_count = get_page_count(self.selected_files[self.current_file_index])
            
            # Update UI and load preview
            self._update_navigation_ui()
            self._load_current_preview()
    
    def _next_file(self) -> None:
        """Navigate to next file"""
        if self.current_file_index < len(self.selected_files) - 1:
            self.current_file_index += 1
            self.current_page_index = 0  # Reset to first page
            
            # Update page count for new file
            self.current_file_page_count = get_page_count(self.selected_files[self.current_file_index])
            
            # Update UI and load preview
            self._update_navigation_ui()
            self._load_current_preview()
    
    def _prev_page(self) -> None:
        """Navigate to previous page"""
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self._update_navigation_ui()
            self._load_current_preview()
    
    def _next_page(self) -> None:
        """Navigate to next page"""
        if self.current_page_index < self.current_file_page_count - 1:
            self.current_page_index += 1
            self._update_navigation_ui()
            self._load_current_preview()
    
    # ==================== Preview Loading Methods ====================
    
    def load_preview_async(self, file_path: str) -> None:
        """Load PDF preview in a separate thread"""
        try:
            # Get preview area size
            self.after(0, self._get_preview_size_and_load, file_path)
        except Exception as e:
            self.after(0, self._show_preview_error, str(e))
    
    def _get_preview_size_and_load(self, file_path: str) -> None:
        """Get preview area dimensions and trigger preview loading"""
        # Get actual dimensions of preview canvas
        self.preview_canvas.update_idletasks()
        width = self.preview_canvas.winfo_width()
        height = self.preview_canvas.winfo_height()
        
        # Use default size if dimensions aren't available yet
        if width <= 1:
            width = 800
        if height <= 1:
            height = 600
        
        # The border is handled inside generate_preview_image, so we pass the full available size
        # Small padding to ensure image doesn't touch edges
        padding = 40
        max_width = max(width - padding, 400)
        max_height = max(height - padding, 300)
        
        # Start loading in thread
        threading.Thread(
            target=self._generate_and_display_preview,
            args=(file_path, (max_width, max_height)),
            daemon=True
        ).start()
    
    def _generate_and_display_preview(self, file_path: str, max_size: Tuple[int, int]) -> None:
        """Generate preview image and update GUI (runs in separate thread)"""
        try:
            # Generate preview image
            pil_image = generate_preview_image(file_path, max_size)
            
            if pil_image is None:
                self.after(0, self._show_preview_error, "Failed to generate preview image")
                return
            
            # Update GUI in main thread with the PIL image
            self.after(0, self._display_preview, pil_image, os.path.basename(file_path))
            
        except Exception as e:
            error_msg = f"Error loading PDF: {str(e)}"
            self.after(0, self._show_preview_error, error_msg)
    
    def _display_preview(self, pil_image: Image.Image, filename: str) -> None:
        """Display the preview image centered on canvas (called in main thread)"""
        try:
            # Convert PIL Image to PhotoImage for canvas
            self._photo_image = ImageTk.PhotoImage(pil_image)
            
            # Clear the canvas
            self.preview_canvas.delete("all")
            
            # Get canvas dimensions
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            # Calculate center position
            x = canvas_width // 2
            y = canvas_height // 2
            
            # Place image centered on canvas
            self._canvas_image_id = self.preview_canvas.create_image(
                x, y,
                image=self._photo_image,
                anchor="center"
            )
            
            print(f"Preview loaded: {filename} (size: {pil_image.size})")
            
        except Exception as e:
            self._show_preview_error(f"Error displaying preview: {str(e)}")
    
    def _show_canvas_message(self, message: str, color: str = "#FFFFFF") -> None:
        """Display a text message centered on the canvas"""
        self.preview_canvas.delete("all")
        
        # Get canvas dimensions (use defaults if not yet rendered)
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        
        if canvas_width <= 1:
            canvas_width = 400
        if canvas_height <= 1:
            canvas_height = 300
        
        # Create centered text
        x = canvas_width // 2
        y = canvas_height // 2
        
        self._canvas_text_id = self.preview_canvas.create_text(
            x, y,
            text=message,
            fill=color,
            font=("Arial", 14),
            anchor="center"
        )
    
    def _on_canvas_configure(self, event=None) -> None:
        """Handle canvas resize - recenter the image if one is loaded"""
        if self._photo_image and self._canvas_image_id:
            # Recenter the image when canvas is resized
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            x = canvas_width // 2
            y = canvas_height // 2
            self.preview_canvas.coords(self._canvas_image_id, x, y)
        elif self._canvas_text_id:
            # Recenter text message
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            x = canvas_width // 2
            y = canvas_height // 2
            self.preview_canvas.coords(self._canvas_text_id, x, y)
    
    def _show_preview_error(self, error_message: str) -> None:
        """Show error message in preview area and popup"""
        try:
            # Show error in preview area
            self._show_canvas_message(f"Error: {error_message}", color="#FF6B6B")
            
            # Show error popup
            messagebox.showerror("PDF Preview Error", error_message)
            
            print(f"Preview error: {error_message}")
            
        except Exception as e:
            print(f"Error showing error message: {e}")
    
    def update_project_data(self, event=None) -> None:
        """Update project data dictionary when inputs change"""
        self.project_data['project_name'] = self.project_name_entry.get()
        self.project_data['client_name'] = self.client_name_entry.get()
        self.project_data['date'] = self.date_entry.get()
        self.project_data['drawn_by'] = self.drawn_by_entry.get()
    
    def update_preview_message(self, message: str) -> None:
        """Display a message in the preview area"""
        self._show_canvas_message(message)
    
    def save_and_export(self) -> None:
        """Handle batch processing and export of all selected PDFs"""
        # Check if any files are selected
        if not self.selected_files:
            messagebox.showwarning("No Files Selected", "Please select one or more PDF files first.")
            return
        
        # Check if library has templates
        if not self._check_library_and_warn():
            return
        
        # Get currently selected template from dropdown
        selected_template_name = self.template_dropdown.get()
        for file_path in self.template_files:
            if os.path.basename(file_path) == selected_template_name:
                self.template_path = file_path
                break
        
        # Update project data from entries
        self.update_project_data()
        
        # Disable the button during processing
        self.save_export_btn.configure(state="disabled")
        
        # Run batch processing in a separate thread to keep GUI responsive
        threading.Thread(
            target=self._batch_process_files,
            daemon=True
        ).start()
    
    def _batch_process_files(self) -> None:
        """Process all selected files in a batch (runs in separate thread)"""
        total_files = len(self.selected_files)
        successful = 0
        failed = 0
        failed_files = []
        
        # Parse excluded pages from the input field
        exclude_string = self.exclude_pages_entry.get()
        excluded_pages = parse_page_range(exclude_string)
        if excluded_pages:
            print(f"Excluding pages (0-indexed): {sorted(excluded_pages)}")
        
        # Get output folder from entry field (allows manual paste or browse selection)
        output_folder_str = self.output_path_entry.get().strip()
        if not output_folder_str:
            # Fallback to default if empty
            output_folder_str = os.path.join(os.path.expanduser('~'), 'Desktop', 'Leviat_Output')
        
        # Create output folder if it doesn't exist
        output_folder_path = Path(output_folder_str)
        output_folder_path.mkdir(parents=True, exist_ok=True)
        print(f"Output folder: {output_folder_path}")
        
        # Show progress bar and status label
        self.after(0, self._show_progress_bar)
        
        for i, file_path in enumerate(self.selected_files, start=1):
            filename = os.path.basename(file_path)
            
            # Update progress bar and status
            progress_value = (i - 1) / total_files
            status_text = f"Processing file {i}/{total_files}: {filename}..."
            self.after(0, self._update_progress, progress_value, status_text)
            
            try:
                # Process the PDF file
                output_filename = f"processed_{filename}"
                file_output_path = output_folder_path / output_filename
                
                # Call pdf_logic.process_with_margins for the file
                pdf_logic.process_with_margins(
                    input_path=file_path,
                    template_path=self.template_path,
                    output_path=str(file_output_path),
                    project_data=self.project_data,
                    excluded_pages=excluded_pages
                )
                
                successful += 1
                print(f"Successfully processed: {filename}")
                
            except Exception as e:
                # Log error but continue with other files
                failed += 1
                failed_files.append(filename)
                print(f"Error processing {filename}: {str(e)}")
        
        # Update progress to 100%
        self.after(0, self._update_progress, 1.0, "Batch processing complete!")
        
        # Hide progress bar and show completion message
        self.after(500, self._hide_progress_bar)
        self.after(600, self._show_batch_complete_message, successful, failed, failed_files, str(output_folder_path))
        
        # Re-enable the button
        self.after(0, lambda: self.save_export_btn.configure(state="normal"))
    
    def _show_progress_bar(self) -> None:
        """Show the progress bar and status label"""
        self.progress_bar.grid()
        self.progress_status_label.grid()
        self.progress_bar.set(0)
    
    def _hide_progress_bar(self) -> None:
        """Hide the progress bar and status label"""
        self.progress_bar.grid_remove()
        self.progress_status_label.grid_remove()
    
    def _update_progress(self, value: float, status_text: str) -> None:
        """Update progress bar value and status text"""
        self.progress_bar.set(value)
        self.progress_status_label.configure(text=status_text)
        self.status_label.configure(text=status_text)
    
    def _show_batch_complete_message(self, successful: int, failed: int, failed_files: list, output_folder: str) -> None:
        """Show batch complete popup message"""
        if failed == 0:
            message = f"Batch Complete!\n\n{successful} file(s) processed successfully.\n\nOutput saved to:\n{output_folder}"
            messagebox.showinfo("Batch Complete", message)
        else:
            message = f"Batch Complete with Errors!\n\n{successful} file(s) processed successfully.\n{failed} file(s) failed.\n\nFailed files:\n"
            message += "\n".join(f"  • {f}" for f in failed_files)
            message += f"\n\nOutput saved to:\n{output_folder}"
            messagebox.showwarning("Batch Complete", message)
        
        # Reset status label
        n = len(self.selected_files)
        self.status_label.configure(
            text=f"{n} file{'s' if n != 1 else ''} selected",
            text_color="#00CC66"
        )


def main():
    """Application entry point"""
    app = PDFAutomationApp()
    app.mainloop()


if __name__ == "__main__":
    main()

