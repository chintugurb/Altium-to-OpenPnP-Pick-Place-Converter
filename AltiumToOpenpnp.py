import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import csv
import sys # <-- NEW: Import sys

# --- NEW: Helper function for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# ------------------------------------------

class AltiumToOpenPnPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Altium to OpenPnP Pick & Place Converter")
        self.root.geometry("900x600")
        
        # --- NEW: Set the window icon ---
        try:
            icon_path = resource_path("favicon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            pass # If icon is missing, just ignore and use default
        # --------------------------------
        
        self.current_df = None
        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        # --- Top Frame (Controls) ---
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(fill=tk.X)

        self.btn_browse = tk.Button(top_frame, text="1. Browse Input File", command=self.load_file, font=("Arial", 10, "bold"))
        self.btn_browse.pack(side=tk.LEFT, padx=10)

        self.lbl_file = tk.Label(top_frame, text="No file selected...", fg="gray")
        self.lbl_file.pack(side=tk.LEFT, padx=10)

        self.btn_convert = tk.Button(top_frame, text="2. Convert & Save for OpenPnP", command=self.convert_and_save, font=("Arial", 10, "bold"), state=tk.DISABLED, bg="#d4edda")
        self.btn_convert.pack(side=tk.RIGHT, padx=10)

        # --- Middle Frame (Data Display) ---
        mid_frame = tk.Frame(self.root)
        mid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Scrollbars for the Treeview
        tree_scroll_y = tk.Scrollbar(mid_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = tk.Scrollbar(mid_frame, orient='horizontal')
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(mid_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Select Altium Pick & Place File",
            filetypes=(("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*"))
        )
        if not filepath:
            return

        self.file_path = filepath
        self.lbl_file.config(text=os.path.basename(filepath), fg="black")
        
        # 1. Determine file type and locate the actual table header
        ext = os.path.splitext(filepath)[1].lower()
        header_idx = 0
        
        try:
            if ext == '.csv':
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    for i, line in enumerate(f):
                        if line.strip().lower().startswith('designator'):
                            header_idx = i
                            break
                self.current_df = pd.read_csv(filepath, skiprows=header_idx)
                
            elif ext in ['.xls', '.xlsx']:
                temp_df = pd.read_excel(filepath, nrows=30)
                for i, row in temp_df.iterrows():
                    if str(row.iloc[0]).strip().lower().startswith('designator'):
                        header_idx = i + 1  
                        break
                self.current_df = pd.read_excel(filepath, skiprows=header_idx)
            else:
                messagebox.showerror("Error", "Unsupported file format!")
                return
                
            # Display data in UI
            self.display_dataframe(self.current_df)
            self.btn_convert.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error Reading File", str(e))

    def display_dataframe(self, df):
        # Clear existing Treeview
        self.tree.delete(*self.tree.get_children())
        
        if df is None or df.empty:
            return

        # Set up columns
        self.tree["column"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)

        # Insert rows
        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            # Replace NaNs with empty strings for display
            clean_row = ["" if pd.isna(x) else x for x in row]
            self.tree.insert("", "end", values=clean_row)

    def convert_and_save(self):
        if self.current_df is None:
            return

        df = self.current_df.copy()

        try:
            # 2. Rename Altium coordinate headers to OpenPnP recognizable ones
            col_mapping = {
                'Center-X(mm)': 'X (mm)', 'Center-Y(mm)': 'Y (mm)',
                'Center-X(mil)': 'X (mil)', 'Center-Y(mil)': 'Y (mil)',
                'Ref-X(mm)': 'X (mm)', 'Ref-Y(mm)': 'Y (mm)'
            }
            df.rename(columns=col_mapping, inplace=True)

            # 3. Clean and FIX coordinate data (including division by 1000)
            for col in ['X (mm)', 'Y (mm)', 'X (mil)', 'Y (mil)']:
                if col in df.columns:
                    # Fix comma decimals and text formatting
                    df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                    df[col] = df[col].str.replace('mm', '', regex=False).str.replace('mil', '', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    
                    # Apply scaling fix: Divide by 1000 if number is impossibly large (> 1000mm)
                    df[col] = df[col].apply(lambda x: x / 1000.0 if pd.notnull(x) and abs(x) > 1000 else x)

            # 4. Standardize Layer/Side to 'T' (Top) and 'B' (Bottom)
            if 'Layer' in df.columns:
                df['Layer'] = df['Layer'].astype(str).apply(
                    lambda x: 'T' if x.lower().startswith('t') else ('B' if x.lower().startswith('b') else x)
                )

            # Drop any trailing empty/invalid rows 
            df.dropna(subset=[df.columns[0]], inplace=True) 

            # 5. Arrange the columns according to the OpenPnP sequence
            openpnp_sequence = [
                ["Designator", "Part", "Component", "RefDes", "Ref"],
                ["Value", "Val", "Comment", "Comp_Value"],
                ["Footprint", "Package", "Pattern", "Comp_Package"],
                ["X", "X (mm)", "Ref X", "PosX", "Ref-X(mm)", "Ref-X(mil)", "Sym_X"],
                ["Y", "Y (mm)", "Ref Y", "PosY", "Ref-Y(mm)", "Ref-Y(mil)", "Sym_Y"],
                ["Rotation", "Rot", "Rotate", "Sym_Rotate"],
                ["Layer", "Side", "TB", "Sym_Mirror"],
                ["Height", "Height(mil)", "Height(mm)"]
            ]

            new_col_order = []
            
            # Extract standard columns
            for category in openpnp_sequence:
                for possible_name in category:
                    matching_cols = [c for c in df.columns if c.lower() == possible_name.lower()]
                    if matching_cols:
                        new_col_order.append(matching_cols[0])
                        break 

            # Push remaining un-categorized columns to the end
            for col in df.columns:
                if col not in new_col_order:
                    new_col_order.append(col)

            df = df[new_col_order]

            # Ask user where to save the new file
            initial_dir = os.path.dirname(self.file_path)
            base_name, _ = os.path.splitext(os.path.basename(self.file_path))
            default_out_name = f"{base_name}_openpnp.csv"

            save_path = filedialog.asksaveasfilename(
                initialdir=initial_dir,
                initialfile=default_out_name,
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv")],
                title="Save OpenPnP Compatible File"
            )

            if not save_path:
                return # User cancelled save

            # 6. Export to CSV
            df.to_csv(save_path, index=False, quoting=csv.QUOTE_ALL, sep=',')
            
            # Update the UI to show the converted Data
            self.display_dataframe(df)
            self.current_df = df # Update underlying df so they can save again if needed
            
            messagebox.showinfo("Success", f"File successfully converted and saved to:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Error During Conversion", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AltiumToOpenPnPApp(root)
    root.mainloop()