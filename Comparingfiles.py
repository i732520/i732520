import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar, IntVar
import os
import re

# Constants for chunk size and display limits
CHUNKSIZE = 50000
# Removed MAX_PREVIEW and MAX_DISPLAY as they will now be user-configurable or derived

def normalize_colname(name):
    """Normalizes column names by stripping whitespace and converting to lowercase."""
    return re.sub(r"\s+", " ", str(name).strip()).lower()

class ToolTip:
    """
    A simple tooltip class to display information when hovering over a widget.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        "Displays the tooltip."
        if self.tipwindow or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert") or (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 20
        y = y + cy + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True) # Removes window decorations
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 9))
        label.pack()

    def leave(self, event=None):
        "Hides the tooltip."
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

class ColorTreeview(ttk.Treeview):
    """
    Custom Treeview class (currently a placeholder, but can be extended for coloring).
    """
    pass

class MappingSearchSBSApp:
    """
    Main application class for the Side-by-Side File Comparison Tool.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Side-by-Side File Comparison Tool")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)

        # Configure Treeview style
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        style.configure("Treeview", font=('Arial', 10))

        # Initialize dataframes and headers
        self.df1 = None
        self.df2 = None
        self.headers1 = []
        self.headers2 = []
        
        # Default preview limit, can be adjusted if needed
        self.max_preview_rows = 1000 
        # User-defined display limit
        self.max_display_rows = IntVar(value=1000) 

        self.loaded_file1 = None
        self.loaded_file2 = None

        # --- FILE SELECTION SECTION ---
        file_frame = tk.LabelFrame(root, text="Step 1: Select Files", font=('Arial', 11, 'bold'))
        file_frame.pack(padx=14, pady=(12,6), fill="x")

        tk.Label(file_frame, text="File 1:", font=('Arial', 10)).grid(row=0, column=0, sticky="e", padx=4, pady=2)
        self.file1_entry = tk.Entry(file_frame, width=60, font=('Arial', 10))
        self.file1_entry.grid(row=0, column=1, padx=2, pady=2)
        tk.Button(file_frame, text="Browse", command=self.load_file1).grid(row=0, column=2, padx=4)

        tk.Label(file_frame, text="File 2:", font=('Arial', 10)).grid(row=1, column=0, sticky="e", padx=4, pady=2)
        self.file2_entry = tk.Entry(file_frame, width=60, font=('Arial', 10))
        self.file2_entry.grid(row=1, column=1, padx=2, pady=2)
        tk.Button(file_frame, text="Browse", command=self.load_file2).grid(row=1, column=2, padx=4)

        tk.Button(file_frame, text="Load & Map Columns", command=self.reload_headers).grid(row=0, column=3, rowspan=2, padx=14, pady=2, sticky="ns")

        # --- COLUMN MAPPING SECTION ---
        map_frame = tk.LabelFrame(root, text="Step 2: Map Columns for Matching (Composite Key Supported)", font=('Arial', 11, 'bold'))
        map_frame.pack(padx=14, pady=(0,7), fill="x")

        self.mapping_rows = [] # Stores tuples of (combo1, combo2, remove_button)
        self.map_frame_inner = tk.Frame(map_frame)
        self.map_frame_inner.pack(anchor="w", pady=2)

        self.add_mapping_btn = tk.Button(map_frame, text="Add Mapping", command=self.add_mapping_row, state="disabled")
        self.add_mapping_btn.pack(side="left", padx=2, pady=3)
        ToolTip(self.add_mapping_btn, "Map additional column pairs for composite key matching.")

        # --- SEARCH & OPTIONS SECTION ---
        options_frame = tk.Frame(root)
        options_frame.pack(fill="x")

        search_frame = tk.LabelFrame(options_frame, text="Step 3: Comparison & Search Options", font=('Arial', 11, 'bold'))
        search_frame.pack(side="left", padx=(0,14), pady=(0,8), fill="both", expand=True)

        tk.Label(search_frame, text="Search Type: ").grid(row=0, column=0, sticky="e")
        self.search_type = tk.StringVar(value="exact")
        ttk.Radiobutton(search_frame, text="Exact", variable=self.search_type, value="exact").grid(row=0, column=1)
        ttk.Radiobutton(search_frame, text="Contains", variable=self.search_type, value="contains").grid(row=0, column=2)
        ttk.Radiobutton(search_frame, text="Regex", variable=self.search_type, value="regex").grid(row=0, column=3)
        
        self.case_sensitive = tk.BooleanVar(value=False)
        tk.Checkbutton(search_frame, text="Case Sensitive", variable=self.case_sensitive).grid(row=0, column=4, padx=4)
        
        tk.Label(search_frame, text="Search Field: ").grid(row=1, column=0, sticky="e")
        self.mapfield_combo = ttk.Combobox(search_frame, state="readonly", width=30)
        self.mapfield_combo.grid(row=1, column=1, padx=2, pady=2, columnspan=2)
        
        tk.Label(search_frame, text="Value: ").grid(row=1, column=3, sticky="e")
        self.value_entry = tk.Entry(search_frame, width=20)
        self.value_entry.grid(row=1, column=4, padx=2, pady=2)
        
        tk.Button(search_frame, text="Search", command=self.do_search, width=12).grid(row=1, column=5, padx=8)
        tk.Button(search_frame, text="Clear Results", command=self.clear_results, width=12).grid(row=1, column=6, padx=2)
        ToolTip(self.value_entry, "Enter a value to restrict search to rows containing this value in the selected field.")

        # --- COUNT OPTIONS ---
        count_option_frame = tk.LabelFrame(options_frame, text="Match Counting & Display", font=('Arial', 11, 'bold'))
        count_option_frame.pack(side="left", padx=(0,14), pady=(0,8), fill="y")
        self.count_option = IntVar(value=1) # Default to "All matching pairs"
        ttk.Radiobutton(count_option_frame, text="All matching pairs", variable=self.count_option, value=1).pack(anchor="w", pady=1)
        ttk.Radiobutton(count_option_frame, text="Unique matched File 1 rows", variable=self.count_option, value=2).pack(anchor="w", pady=1)
        ttk.Radiobutton(count_option_frame, text="Unique matched File 2 rows", variable=self.count_option, value=3).pack(anchor="w", pady=1)
        ToolTip(count_option_frame, "Affects both match count and grid display.")

        # New: MAX_DISPLAY option
        tk.Label(count_option_frame, text="Max Display Rows:").pack(anchor="w", pady=(5,0))
        self.max_display_entry = tk.Entry(count_option_frame, textvariable=self.max_display_rows, width=10)
        self.max_display_entry.pack(anchor="w", padx=2, pady=1)
        ToolTip(self.max_display_entry, "Maximum number of rows to display in the results table.")


        self.match_count_label = tk.Label(search_frame, text="Matching: 0 | Non-matching: 0", font=('Arial', 10, 'bold'))
        self.match_count_label.grid(row=2, column=0, columnspan=7, pady=(4,0), sticky="w")

        # --- FILTERS ---
        filter_frame = tk.Frame(root)
        filter_frame.pack(padx=14, pady=(0,5), fill="x")
        self.show_matches = tk.BooleanVar(value=True)
        self.show_nonmatches = tk.BooleanVar(value=True)
        tk.Checkbutton(filter_frame, text="Show Matches", variable=self.show_matches, command=self.refresh_grid).pack(side="left")
        tk.Checkbutton(filter_frame, text="Show Non-matches", variable=self.show_nonmatches, command=self.refresh_grid).pack(side="left", padx=6)
        ToolTip(filter_frame, "Toggle the display of matched and unmatched rows in the results table.")

        # --- RESULTS TABLE ---
        grid_frame = tk.LabelFrame(root, text="Step 4: Results Table", font=('Arial', 11, 'bold'))
        grid_frame.pack(padx=14, pady=8, fill="both", expand=True)
        self.grid = ColorTreeview(grid_frame, show="headings", selectmode="browse")
        self.grid.pack(side="left", fill="both", expand=True)
        grid_scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.grid.yview)
        grid_scroll.pack(side="right", fill="y")
        self.grid.configure(yscrollcommand=grid_scroll.set)

        # --- EXPORTS ---
        export_frame = tk.Frame(root)
        export_frame.pack(padx=14, pady=(0,10), fill='x')
        tk.Button(export_frame, text="Export Matched to Excel", command=lambda: self.export_to_excel(only_matches=True), width=20).pack(side="left", padx=10)
        tk.Button(export_frame, text="Export Non-matched to Excel", command=lambda: self.export_to_excel(only_matches=False), width=22).pack(side="left", padx=10)
        tk.Button(export_frame, text="Load Full File for Export", command=self.load_full_files, width=23).pack(side="left", padx=16)

        self.grid_content = [] # Stores the data to be displayed in the grid
        self.grid_columns = [] # Stores the column headers for the grid

    def load_file1(self):
        """Opens a file dialog to select File 1 and updates the entry field."""
        path = filedialog.askopenfilename(title="Select File 1", filetypes=[("Excel/CSV/TXT", "*.xlsx *.csv *.txt"), ("All files", "*.*")])
        if path:
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, path)
            self.loaded_file1 = path

    def load_file2(self):
        """Opens a file dialog to select File 2 and updates the entry field."""
        path = filedialog.askopenfilename(title="Select File 2", filetypes=[("Excel/CSV/TXT", "*.xlsx *.csv *.txt"), ("All files", "*.*")])
        if path:
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, path)
            self.loaded_file2 = path

    def reload_headers(self):
        """
        Loads preview data from selected files, extracts headers, and initializes
        the column mapping section.
        """
        self.df1 = self.read_file(self.file1_entry.get())
        self.df2 = self.read_file(self.file2_entry.get())

        if self.df1 is None or self.df2 is None or self.df1.empty or self.df2.empty:
            messagebox.showerror("Error", "Both files must load successfully and contain data.")
            return

        self.headers1 = list(self.df1.columns)
        self.headers2 = list(self.df2.columns)

        self.clear_mapping_rows()
        self.add_mapping_btn.config(state="normal")
        self.mapfield_combo['values'] = [] # Clear search field options
        self.value_entry.delete(0, tk.END) # Clear search value
        self.clear_grid() # Clear previous results
        self.match_count_label.config(text="Matching: 0 | Non-matching: 0")

        # Attempt to auto-map columns with similar normalized names
        norm1 = {normalize_colname(h): h for h in self.headers1}
        norm2 = {normalize_colname(h): h for h in self.headers2}
        
        auto_mapped_count = 0
        for n1, h1 in norm1.items():
            if n1 in norm2:
                self.add_mapping_row(h1, norm2[n1])
                auto_mapped_count += 1
        
        # If no columns were auto-mapped, add at least one empty mapping row
        if auto_mapped_count == 0 and (self.headers1 and self.headers2):
            self.add_mapping_row()

        # Display preview mode notices if applicable
        if hasattr(self.df1, '__len__') and len(self.df1) == self.max_preview_rows:
            messagebox.showinfo("Notice", f"Preview mode: Only first {self.max_preview_rows} rows loaded from File 1.")
        if hasattr(self.df2, '__len__') and len(self.df2) == self.max_preview_rows:
            messagebox.showinfo("Notice", f"Preview mode: Only first {self.max_preview_rows} rows loaded from File 2.")

    def read_file(self, path):
        """
        Reads data from a given file path, handling CSV, Excel, and TXT formats.
        Applies a preview limit for large files.
        """
        if not path:
            return None
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".csv":
                # For very large CSVs, read in chunks for preview
                if os.path.getsize(path) > 100 * 1024 * 1024: # 100 MB
                    preview = []
                    row_count = 0
                    for chunk in pd.read_csv(path, chunksize=CHUNKSIZE, dtype=str):
                        preview.append(chunk)
                        row_count += len(chunk)
                        if row_count >= self.max_preview_rows:
                            break
                    if preview:
                        df = pd.concat(preview)[:self.max_preview_rows]
                        return df
                    else:
                        return pd.DataFrame() # Return empty DataFrame if no data
                else:
                    df = pd.read_csv(path, dtype=str)
                    return df
            elif ext == ".xlsx":
                # For very large Excels, read only nrows for preview
                if os.path.getsize(path) > 10 * 1024 * 1024: # 10 MB
                    df = pd.read_excel(path, dtype=str, nrows=self.max_preview_rows)
                    return df
                else:
                    df = pd.read_excel(path, dtype=str)
                    return df
            elif ext == ".txt":
                # Try reading as CSV, then as tab-separated if first fails
                try:
                    df = pd.read_csv(path, dtype=str, nrows=self.max_preview_rows)
                    return df
                except Exception:
                    df = pd.read_csv(path, sep="\t", dtype=str, nrows=self.max_preview_rows)
                    return df
            else:
                messagebox.showerror("Error", f"Unsupported file extension: {ext}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read {path}:\n{e}")
        return None

    def clear_mapping_rows(self):
        """Removes all dynamically added column mapping rows."""
        for row in self.mapping_rows:
            for widget in row:
                widget.destroy()
        self.mapping_rows.clear()

    def add_mapping_row(self, sel1=None, sel2=None):
        """
        Adds a new row for column mapping with two comboboxes and a remove button.
        Pre-selects values if provided (e.g., for auto-mapping).
        """
        row_idx = len(self.mapping_rows)
        # Set default values for comboboxes
        var1 = StringVar(value=sel1 or (self.headers1[0] if self.headers1 else ""))
        var2 = StringVar(value=sel2 or (self.headers2[0] if self.headers2 else ""))
        
        combo1 = ttk.Combobox(self.map_frame_inner, values=self.headers1, textvariable=var1, state="readonly", width=30)
        combo2 = ttk.Combobox(self.map_frame_inner, values=self.headers2, textvariable=var2, state="readonly", width=30)
        
        combo1.grid(row=row_idx, column=0, padx=2, pady=2)
        combo2.grid(row=row_idx, column=1, padx=2, pady=2)
        
        rm_btn = tk.Button(self.map_frame_inner, text="Remove", command=lambda: self.remove_mapping_row(row_idx))
        rm_btn.grid(row=row_idx, column=2, padx=2)
        
        self.mapping_rows.append((combo1, combo2, rm_btn))
        self.update_mapfield_combo()

    def remove_mapping_row(self, idx):
        """Removes a specific column mapping row by its index."""
        if idx < len(self.mapping_rows):
            for widget in self.mapping_rows[idx]:
                widget.destroy()
            self.mapping_rows.pop(idx)
            # Re-grid remaining rows to fill the gap
            for i, (c1, c2, rm) in enumerate(self.mapping_rows):
                c1.grid(row=i, column=0)
                c2.grid(row=i, column=1)
                rm.grid(row=i, column=2)
        self.update_mapfield_combo()

    def update_mapfield_combo(self):
        """Updates the 'Search Field' combobox with currently mapped columns from File 1."""
        mapped_names = [row[0].get() for row in self.mapping_rows if row[0].get()]
        self.mapfield_combo['values'] = mapped_names
        if mapped_names:
            self.mapfield_combo.current(0) # Select the first mapped column by default

    def do_search(self):
        """
        Performs the comparison and search operation based on selected files,
        column mappings, and search criteria. Populates the results grid.
        This version is optimized for memory by avoiding a full pandas merge.
        """
        if self.df1 is None or self.df2 is None:
            messagebox.showerror("Error", "Please load both files before searching.")
            return

        mapping_keys = []
        for combo1, combo2, _ in self.mapping_rows:
            k1, k2 = combo1.get(), combo2.get()
            if k1 and k2: # Ensure both columns are selected for a mapping
                mapping_keys.append((k1, k2))
        
        if not mapping_keys:
            messagebox.showerror("Error", "Please map at least one column for comparison.")
            return

        search_field = self.mapfield_combo.get()
        search_value = self.value_entry.get().strip()
        
        # If a search field is selected but no value is provided, raise an error
        if search_field and not search_value:
            messagebox.showerror("Error", "Please enter a value for the selected search field.")
            return
        
        search_type = self.search_type.get()
        case_sensitive = self.case_sensitive.get()

        cols1 = self.headers1
        cols2 = self.headers2

        # Retrieve user-defined MAX_DISPLAY_ROWS
        try:
            current_max_display = int(self.max_display_rows.get())
            if current_max_display <= 0:
                messagebox.showerror("Invalid Input", "Max Display Rows must be a positive integer.")
                return
        except ValueError:
            messagebox.showerror("Invalid Input", "Max Display Rows must be a valid integer.")
            return

        # Determine if a specific search filter is active (i.e., both field and value are provided)
        is_search_active = bool(search_field and search_value)
        print(f"DEBUG: do_search - search_field: '{search_field}', search_value: '{search_value}', is_search_active: {is_search_active}")


        # Helper function to apply search filter to a single cell value
        def apply_search_filter(cell_value, value_to_match, search_type, case_sensitive):
            cell_value_str = str(cell_value).strip()
            
            if not case_sensitive:
                cell_value_str = cell_value_str.lower()
                value_to_match = value_to_match.lower()

            result = False
            if search_type == "exact":
                result = cell_value_str == value_to_match
            elif search_type == "contains":
                result = value_to_match in cell_value_str
            elif search_type == "regex":
                try:
                    result = re.fullmatch(value_to_match, cell_value_str) is not None
                except re.error as e:
                    messagebox.showerror("Regex Error", f"Invalid regex pattern: {value_to_match}\nError: {e}")
                    result = False
            return result

        # --- Step 1: Pre-index df2 for efficient lookups ---
        # df2_key_to_rows will map composite keys to a list of (original_index, row_data) tuples
        df2_key_to_rows = {}
        for idx2, row2 in self.df2.iterrows():
            composite_key_values = []
            for _, k2 in mapping_keys:
                # Get value, handle missing columns and NaN, then strip
                val = row2.get(k2, '')
                composite_key_values.append(str(val).strip() if pd.notna(val) else '')
            composite_key = tuple(composite_key_values)
            df2_key_to_rows.setdefault(composite_key, []).append((idx2, row2))
        
        print(f"DEBUG: df2_key_to_rows created with {len(df2_key_to_rows)} unique keys.")

        # --- Step 2: Initialize results and tracking sets ---
        self.grid_content = []
        match_count = 0
        nonmatch_count = 0
        displayed_rows = 0

        # Sets to track original indices that have been processed to ensure uniqueness for counts/display
        processed_file1_original_indices = set()
        processed_file2_original_indices = set()
        
        # Store all potential rows to be displayed, then filter by current_max_display
        # This helps ensure counts are accurate even if not all rows are displayed.
        all_display_candidates = []

        # --- Step 3: Iterate through df1 to find matches and File 1 Only rows ---
        for idx1, row1 in self.df1.iterrows():
            composite_key_values = []
            for k1, _ in mapping_keys:
                val = row1.get(k1, '')
                composite_key_values.append(str(val).strip() if pd.notna(val) else '')
            composite_key = tuple(composite_key_values)

            # Determine if this File 1 row meets the search criteria
            row1_meets_search_criteria = True # Assume it meets criteria by default
            if is_search_active: # Only apply filter if a search is active
                if search_field in cols1: # ONLY apply filter if search_field exists in File 1 headers
                    if not apply_search_filter(row1.get(search_field, ''), search_value, search_type, case_sensitive):
                        row1_meets_search_criteria = False
                        # print(f"DEBUG: File 1 row {idx1} skipped - search filter mismatch on '{search_field}' (value: '{row1.get(search_field, '')}').")
                # else: If search_field is active but not in cols1, row1_meets_search_criteria remains True
                # This means the search filter doesn't apply to this side, so it's not filtered out.
            
            if not row1_meets_search_criteria:
                continue # Skip this File 1 row if it doesn't meet search criteria

            # Look for matches in df2
            df2_matches = df2_key_to_rows.get(composite_key, [])

            if df2_matches:
                # This File 1 row has at least one match in File 2
                if self.count_option.get() == 1: # All matching pairs
                    for idx2, row2 in df2_matches:
                        # Determine if this File 2 match meets the search criteria
                        row2_meets_search_criteria = True
                        if is_search_active: # Only apply filter if a search is active
                            if search_field in cols2: # ONLY apply filter if search_field exists in File 2 headers
                                if not apply_search_filter(row2.get(search_field, ''), search_value, search_type, case_sensitive):
                                    row2_meets_search_criteria = False
                                    # print(f"DEBUG: Matched File 2 row {idx2} skipped - search filter mismatch on '{search_field}' (value: '{row2.get(search_field, '')}').")
                            # else: If search_field is active but not in cols2, row2_meets_search_criteria remains True
                        
                        # A match is added only if BOTH sides meet their respective search criteria (or if filter doesn't apply)
                        if row1_meets_search_criteria and row2_meets_search_criteria:
                            all_display_candidates.append(("Match", row1, row2, True, idx1, idx2))
                            # print(f"DEBUG: Added Match (Option 1) for File 1 idx {idx1}, File 2 idx {idx2}.")
                            # Mark both indices as processed for uniqueness tracking
                            processed_file1_original_indices.add(idx1)
                            processed_file2_original_indices.add(idx2)
                
                elif self.count_option.get() == 2: # Unique matched File 1 rows
                    if idx1 not in processed_file1_original_indices:
                        # Find at least one matching row from df2 that also meets search criteria
                        valid_df2_match_found = False
                        idx2_to_display = None
                        row2_to_display = pd.Series(dtype=str)

                        for idx2_cand, row2_cand in df2_matches:
                            row2_cand_meets_criteria = True
                            if is_search_active:
                                if search_field in cols2:
                                    if not apply_search_filter(row2_cand.get(search_field, ''), search_value, search_type, case_sensitive):
                                        row2_cand_meets_criteria = False
                                else:
                                    # If search_field not in cols2, it passes for this side
                                    pass 
                            
                            if row2_cand_meets_criteria:
                                valid_df2_match_found = True
                                idx2_to_display = idx2_cand
                                row2_to_display = row2_cand
                                break # Found a valid match, no need to check others

                        if valid_df2_match_found:
                            all_display_candidates.append(("Match", row1, row2_to_display, True, idx1, idx2_to_display))
                            # print(f"DEBUG: Added Match (Option 2) for File 1 idx {idx1}, File 2 idx {idx2_to_display}.")
                            processed_file1_original_indices.add(idx1)
                            # Mark all associated df2 indices as processed to avoid them becoming "File 2 Only"
                            for df2_match_idx, _ in df2_matches:
                                processed_file2_original_indices.add(df2_match_idx)

                elif self.count_option.get() == 3: # Unique matched File 2 rows
                    # For this option, we need to iterate through df2_matches and add if df2_idx is unique
                    for idx2, row2 in df2_matches:
                        if idx2 not in processed_file2_original_indices:
                            # Determine if this File 2 match meets the search criteria
                            row2_meets_search_criteria = True
                            if is_search_active: # Only apply filter if a search is active
                                if search_field in cols2: # ONLY apply filter if search_field exists in File 2 headers
                                    if not apply_search_filter(row2.get(search_field, ''), search_value, search_type, case_sensitive):
                                        row2_meets_search_criteria = False
                                        # print(f"DEBUG: Unique matched File 2 row {idx2} skipped - search filter mismatch.")
                                else:
                                    # If search_field is active but not in cols2, row2_meets_search_criteria remains True
                                    pass
                            
                            # A match is added only if BOTH sides meet their respective search criteria (or if filter doesn't apply)
                            if row1_meets_search_criteria and row2_meets_search_criteria:
                                all_display_candidates.append(("Match", row1, row2, True, idx1, idx2))
                                # print(f"DEBUG: Added Match (Option 3) for File 1 idx {idx1}, File 2 idx {idx2}.")
                                processed_file1_original_indices.add(idx1) # Mark File 1 row as processed
                                processed_file2_original_indices.add(idx2) # Mark this specific File 2 row as processed

            else: # No match found for this File 1 row (it's a potential File 1 Only row)
                if idx1 not in processed_file1_original_indices: # Ensure it's not already handled
                    # The row1_meets_search_criteria is already checked above for File 1 Only rows
                    all_display_candidates.append(("File 1 Only", row1, pd.Series(dtype=str), False, idx1, None))
                    # print(f"DEBUG: Added File 1 Only for idx {idx1}.")
                    processed_file1_original_indices.add(idx1)


        # --- Step 4: Iterate through df2 to find File 2 Only rows ---
        for idx2, row2 in self.df2.iterrows():
            if idx2 not in processed_file2_original_indices:
                # This df2 row was not part of any match found in the df1 iteration
                
                # Determine if this File 2 Only row meets the search criteria
                row2_meets_search_criteria = True
                if is_search_active: # Only apply filter if a search is active
                    if search_field in cols2: # ONLY apply filter if search_field exists in File 2 headers
                        if not apply_search_filter(row2.get(search_field, ''), search_value, search_type, case_sensitive):
                            row2_meets_search_criteria = False
                            # print(f"DEBUG: File 2 Only row {idx2} skipped - search filter mismatch.")
                    else: # Search field not in File 2 headers, so it passes for this side
                        pass
                
                if row2_meets_search_criteria:
                    all_display_candidates.append(("File 2 Only", pd.Series(dtype=str), row2, False, None, idx2))
                    # print(f"DEBUG: Added File 2 Only for idx {idx2}.")
                    processed_file2_original_indices.add(idx2) # Mark as processed

        print(f"DEBUG: Total display candidates before current_max_display: {len(all_display_candidates)}")

        # --- Step 5: Finalize grid_content and counts based on current_max_display and filters ---
        # Reset counts for final calculation based on what will actually be displayed
        final_match_count = 0
        final_nonmatch_count = 0
        
        # Helper to format cell values, replacing NaN with empty string
        def format_cell_value_for_display(val):
            return str(val).strip() if pd.notna(val) else ''

        for item in all_display_candidates:
            source_tag, row1_data, row2_data, is_match, _, _ = item
            
            # Apply display filters (Show Matches/Show Non-matches checkboxes)
            if is_match and not self.show_matches.get():
                continue
            if not is_match and not self.show_nonmatches.get():
                continue
            
            if displayed_rows >= current_max_display: # Use user-defined limit
                break # Stop adding if current_max_display is reached

            # Apply the NaN removal here
            v_f1 = [format_cell_value_for_display(row1_data.get(h, '')) for h in cols1]
            v_f2 = [format_cell_value_for_display(row2_data.get(h, '')) for h in cols2]
            
            self.grid_content.append((source_tag, v_f1, v_f2, is_match, {}))
            displayed_rows += 1

            if is_match:
                final_match_count += 1
            else:
                final_nonmatch_count += 1

        print(f"DEBUG: Final match_count: {final_match_count}")
        print(f"DEBUG: Final nonmatch_count: {final_nonmatch_count}")
        print(f"DEBUG: Length of grid_content for display: {len(self.grid_content)}")

        # Configure grid columns
        self.grid_columns = ['Source'] + [f"File1_{h}" for h in self.headers1] + [f"File2_{h}" for h in self.headers2]
        print(f"DEBUG: Grid Columns: {self.grid_columns}")
        
        self.refresh_grid()
        self.match_count_label.config(text=f"Matching: {final_match_count} | Non-matching: {final_nonmatch_count}")

    def refresh_grid(self):
        """
        Clears the current grid and repopulates it with data from self.grid_content,
        applying the show_matches and show_nonmatches filters.
        """
        self.clear_grid()
        
        # Set up Treeview columns
        self.grid['columns'] = self.grid_columns
        self.grid.column("#0", width=0, stretch=tk.NO) # Hide the default first empty column

        for col in self.grid_columns:
            self.grid.heading(col, text=col)
            self.grid.column(col, width=120, anchor='w')

        inserted_count = 0
        # Use the user-defined max_display_rows for the actual grid insertion limit
        try:
            display_limit = int(self.max_display_rows.get())
            if display_limit <= 0:
                display_limit = 1 # Ensure at least 1 if invalid input
        except ValueError:
            display_limit = 1000 # Default to 1000 if invalid input

        for gc in self.grid_content:
            source, v_f1, v_f2, is_match, cell_diff_map = gc
            
            # Apply display filters (Show Matches/Show Non-matches checkboxes)
            # These are already applied when populating grid_content, but kept here for robustness
            if is_match and not self.show_matches.get():
                continue
            if not is_match and not self.show_nonmatches.get():
                continue
            
            values = [source] + list(v_f1) + list(v_f2) # Ensure values are in list format
            
            # Basic validation: ensure number of values matches number of columns
            if len(values) != len(self.grid_columns):
                print(f"WARNING: Row value count ({len(values)}) does not match column count ({len(self.grid_columns)}) for row: {values[:5]}...")
                continue # Skip this row to prevent Treeview errors

            self.grid.insert('', 'end', values=values)
            inserted_count += 1
            # Apply the display_limit here for the actual grid insertion
            if inserted_count >= display_limit:
                break
        print(f"DEBUG: Total rows inserted into grid: {inserted_count}")

    def clear_results(self):
        """Clears the search value, grid, and match count label."""
        self.value_entry.delete(0, tk.END)
        self.clear_grid()
        self.match_count_label.config(text="Matching: 0 | Non-matching: 0")
        self.grid_content = [] # Also clear the underlying data

    def clear_grid(self):
        """Removes all items from the Treeview grid."""
        for item in self.grid.get_children():
            self.grid.delete(item)

    def export_to_excel(self, only_matches=True):
        """
        Exports the current grid content (filtered by matches/non-matches) to an Excel file.
        """
        if not self.grid_content:
            messagebox.showerror("Export Error", "No data in grid to export.")
            return

        columns_to_export = self.grid_columns
        data_to_export = []

        # Helper to format cell values, replacing NaN with empty string for export
        def format_cell_value_for_export(val):
            return str(val).strip() if pd.notna(val) else ''

        for gc in self.grid_content:
            source, v_f1, v_f2, is_match, cell_diff_map = gc
            if only_matches and is_match:
                # Apply NaN removal for export as well
                formatted_v_f1 = [format_cell_value_for_export(val) for val in v_f1]
                formatted_v_f2 = [format_cell_value_for_export(val) for val in v_f2]
                data_to_export.append([source] + formatted_v_f1 + formatted_v_f2)
            elif not only_matches and not is_match:
                # Apply NaN removal for export as well
                formatted_v_f1 = [format_cell_value_for_export(val) for val in v_f1]
                formatted_v_f2 = [format_cell_value_for_export(val) for val in v_f2]
                data_to_export.append([source] + formatted_v_f1 + formatted_v_f2)
        
        if not data_to_export:
            messagebox.showinfo("Export", "No records to export based on current filters.")
            return
        
        df_export = pd.DataFrame(data_to_export, columns=columns_to_export)
        
        export_path = filedialog.asksaveasfilename(
            title="Export Grid to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if export_path:
            try:
                df_export.to_excel(export_path, index=False)
                messagebox.showinfo("Exported", f"Grid exported to {export_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {e}")

    def load_full_files(self):
        """
        Loads the entire content of selected files into memory (not just preview)
        for full export capabilities. Warns for very large files.
        """
        if not self.loaded_file1 or not self.loaded_file2:
            messagebox.showerror("Error", "Please select both files first.")
            return
        
        try:
            size1 = os.path.getsize(self.loaded_file1)
            size2 = os.path.getsize(self.loaded_file2)
            
            # Warn for extremely large files that might cause memory issues
            if size1 > 200*1024*1024 or size2 > 200*1024*1024: # 200 MB
                response = messagebox.askyesno(
                    "Memory Warning",
                    "One or both files are very large (>200MB). Loading full files may consume significant memory and could crash the application. Do you want to proceed?"
                )
                if not response:
                    return

            # Read full files based on extension
            ext1 = os.path.splitext(self.loaded_file1)[1].lower()
            ext2 = os.path.splitext(self.loaded_file2)[1].lower()

            if ext1 == ".csv":
                self.df1 = pd.read_csv(self.loaded_file1, dtype=str)
            elif ext1 == ".xlsx":
                self.df1 = pd.read_excel(self.loaded_file1, dtype=str)
            elif ext1 == ".txt":
                try:
                    self.df1 = pd.read_csv(self.loaded_file1, dtype=str)
                except Exception:
                    self.df1 = pd.read_csv(self.loaded_file1, sep="\t", dtype=str)
            else:
                messagebox.showerror("Error", f"Unsupported file extension for File 1: {ext1}")
                return

            if ext2 == ".csv":
                self.df2 = pd.read_csv(self.loaded_file2, dtype=str)
            elif ext2 == ".xlsx":
                self.df2 = pd.read_excel(self.loaded_file2, dtype=str)
            elif ext2 == ".txt":
                try:
                    self.df2 = pd.read_csv(self.loaded_file2, dtype=str)
                except Exception:
                    self.df2 = pd.read_csv(self.loaded_file2, sep="\t", dtype=str)
            else:
                messagebox.showerror("Error", f"Unsupported file extension for File 2: {ext2}")
                return
            
            self.headers1 = list(self.df1.columns)
            self.headers2 = list(self.df2.columns)
            messagebox.showinfo("Full Load Complete", "Full files loaded into memory. You may now run export for all rows.")
            
            # After loading full files, it's good practice to re-map columns
            # as headers might differ slightly if preview was truncated
            self.reload_headers() 

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load full files: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MappingSearchSBSApp(root)
    root.mainloop()
