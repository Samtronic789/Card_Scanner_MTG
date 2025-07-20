import os
import re
import csv
import tkinter as tk
import urllib.parse
from tkinter import filedialog, ttk, messagebox, scrolledtext
from PIL import Image, ImageTk
import threading
from pathlib import Path
import sys
import json
import time

# Handle imports with try/except to identify missing dependencies
try:
    import pandas as pd
except ImportError:
    print("Error: pandas module not installed. Please install it using 'pip install pandas'")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("Error: requests module not installed. Please install it using 'pip install requests'")
    sys.exit(1)

# Try to import RapidOCR with detailed error handling
try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_OCR = True
except ImportError as e:
    print(f"Warning: RapidOCR module not installed or has issues: {str(e)}")
    print("OCR functionality will be disabled. Install RapidOCR using 'pip install rapidocr-onnxruntime'")
    HAS_OCR = False

class CardScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Card Scanner - OCR Tool")
        self.root.geometry("1000x750")
        self.root.configure(bg="#f0f0f0")
        
        # Initialize OCR if available
        self.ocr = None
        if HAS_OCR:
            try:
                self.ocr = RapidOCR()
                print("RapidOCR initialized successfully")
            except Exception as e:
                print(f"Error initializing RapidOCR: {str(e)}")
                messagebox.showwarning("OCR Warning", 
                    "Could not initialize RapidOCR. OCR functionality will be disabled.\n"
                    f"Error: {str(e)}")
        
        # Variables
        self.input_folder = tk.StringVar()
        self.output_csv = tk.StringVar(value="card_data.csv")
        self.current_image_path = None
        self.card_data = []
        self.processed_count = 0
        self.total_images = 0
        self.processing_active = False
        self.tk_image = None  # Keep reference to prevent garbage collection
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input folder selection
        input_frame = ttk.LabelFrame(main_frame, text="Input/Output Settings", padding=10)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Image Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_folder, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_folder).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Output CSV:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_csv, width=50).grid(row=1, column=1, sticky=tk.W+tk.E, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output_file).grid(row=1, column=2, padx=5, pady=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.process_btn = ttk.Button(
            button_frame, 
            text="Process Images", 
            command=self.start_processing
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(
            button_frame, 
            text="Stop Processing", 
            command=self.stop_processing,
            state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(
            button_frame, 
            text="Export to CSV", 
            command=self.export_to_csv,
            state=tk.DISABLED
        )
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress indicator
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(progress_frame, text="Progress:").pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.progress_label = ttk.Label(progress_frame, text="0/0 images processed")
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        # Content area with PanedWindow
        content_pane = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content_pane.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Left side - Image preview
        self.image_frame = ttk.LabelFrame(content_pane, text="Card Preview")
        content_pane.add(self.image_frame, weight=1)
        
        self.image_label = ttk.Label(self.image_frame)
        self.image_label.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Right side - Results
        right_pane = ttk.Frame(content_pane)
        content_pane.add(right_pane, weight=1)
        
        # Split right side into top (table) and bottom (OCR text)
        results_pane = ttk.PanedWindow(right_pane, orient=tk.VERTICAL)
        results_pane.pack(fill=tk.BOTH, expand=True)
        
        # Top-right: Results table
        table_frame = ttk.LabelFrame(results_pane, text="Extracted Card Data")
        results_pane.add(table_frame, weight=2)
        
        # Create treeview for results
        columns = ("filename", "title", "collector_number", "expansion", "status")
        self.results_tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        for col, width, text in [("filename", 150, "Filename"),
                                 ("title", 250, "Card Title"),
                                 ("collector_number", 120, "Collector Number"),
                                 ("expansion", 120, "Set/Expansion"),
                                 ("status", 100, "Status")]:
            self.results_tree.heading(col, text=text)
            self.results_tree.column(col, width=width)
        
        tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=tree_scroll.set)
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_tree.bind("<<TreeviewSelect>>", self.on_result_select)
        
        # Bottom-right: OCR Text
        text_frame = ttk.LabelFrame(results_pane, text="Raw OCR Text")
        results_pane.add(text_frame, weight=1)
        
        self.text_display = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.text_display.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Edit panel for correcting data
        edit_frame = ttk.LabelFrame(main_frame, text="Edit Card Data")
        edit_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(edit_frame, text="Title:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.title_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.title_var, width=40).grid(row=0, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        
        ttk.Label(edit_frame, text="Collector #:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        self.collector_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.collector_var, width=15).grid(row=0, column=3, sticky=tk.W+tk.E, pady=5, padx=5)
        
        ttk.Label(edit_frame, text="Set:").grid(row=0, column=4, sticky=tk.W, pady=5, padx=5)
        self.expansion_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.expansion_var, width=15).grid(row=0, column=5, sticky=tk.W+tk.E, pady=5, padx=5)
        
        self.update_btn = ttk.Button(edit_frame, text="Update Selected", command=self.update_selected_item, state=tk.DISABLED)
        self.update_btn.grid(row=0, column=6, padx=5, pady=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log")
        log_frame.pack(fill=tk.X, pady=5)
        
        self.log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9), height=5)
        self.log.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        
        # Show status message about OCR
        if not HAS_OCR or self.ocr is None:
            self.status_var.set("Warning: OCR functionality is disabled. Install rapidocr-onnxruntime package.")
            self.log_message("OCR IS DISABLED: The application will still allow you to view images and manage data manually.")
        else:
            self.status_var.set("Ready - OCR initialized successfully")

    def clean_expansion_code(self, expansion_code):
        """Clean expansion code by removing dots and limiting to 3 chars"""
        if not expansion_code or expansion_code == "Unknown":
            return expansion_code
        
        # Remove dots and limit to 3 characters
        cleaned = expansion_code.replace('.', '')
        if len(cleaned) > 3:
            cleaned = cleaned[:3]
        return cleaned

    def clean_collector_number(self, collector_number):
        """Clean collector number by taking first part before slash"""
        if not collector_number or collector_number == "Unknown":
            return collector_number
        
        # Take first part if slash exists
        if '/' in collector_number:
            collector_number = collector_number.split('/')[0]
        
        # Further clean non-digit characters if it makes sense
        if re.search(r'\d', collector_number):  # Only clean if it contains digits
            cleaned_number = re.sub(r'\D', '', collector_number)
            if cleaned_number:  # Only use cleaned version if not empty
                collector_number = cleaned_number
        
        return collector_number

    def browse_input_folder(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_folder.set(folder)

    def browse_output_file(self):
        file = filedialog.asksaveasfilename(title="Save CSV As", defaultextension=".csv",
                                            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file:
            self.output_csv.set(file)

    def start_processing(self):
        input_folder = self.input_folder.get().strip()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return
            
        if not HAS_OCR or self.ocr is None:
            messagebox.showwarning("OCR Not Available", 
                "OCR functionality is disabled. You can still manually process images, but text recognition won't be available.")
            
        self.clear_results()
        self.processing_active = True
        self.process_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_images, daemon=True).start()

    def stop_processing(self):
        self.processing_active = False
        self.status_var.set("Processing stopped by user")
        self.process_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        if self.card_data:
            self.export_btn.config(state=tk.NORMAL)

    def clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.text_display.delete(1.0, tk.END)
        self.log.delete(1.0, tk.END)
        self.card_data = []
        self.processed_count = 0
        self.progress_bar["value"] = 0
        self.progress_label.config(text="0/0 images processed")
        self.current_image_path = None
        self.image_label.config(image='')

    def process_images(self):
        try:
            input_folder = self.input_folder.get()
            extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp')
            files = [f for f in os.listdir(input_folder)
                    if os.path.isfile(os.path.join(input_folder, f)) and f.lower().endswith(extensions)]
            
            if not files:
                self.update_status("No image files found in the selected folder.")
                self.process_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
            self.total_images = len(files)
            self.update_status(f"Found {self.total_images} images to process")
            
            for i, fname in enumerate(files):
                if not self.processing_active:
                    break
                path = os.path.join(input_folder, fname)
                self.update_status(f"Processing {i+1}/{self.total_images}: {fname}")
                self.processed_count = i+1
                self.update_progress()
                self.process_single_image(path)
                
            if self.processing_active:
                self.update_status(f"Processing complete. {self.processed_count} images processed.")
                self.processing_active = False
                
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            
            if self.card_data:
                self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
                
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            print(error_msg)
            self.update_status(error_msg)
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))

    def process_single_image(self, image_path):
        try:
            filename = os.path.basename(image_path)
            lines = []
            
            # Attempt OCR if available
            if HAS_OCR and self.ocr is not None:
                try:
                    result, elapse = self.ocr(image_path)
                    if result:
                        lines = [line[1] for line in result]
                        full_text = " ".join(lines)
                        title, collector, expansion = self.parse_card_data(lines, full_text)
                    else:
                        title, collector, expansion = "No text detected", "Unknown", "Unknown"
                except Exception as ocr_error:
                    print(f"OCR error on {image_path}: {str(ocr_error)}")
                    self.log_message(f"OCR error on {filename}: {str(ocr_error)}")
                    title, collector, expansion = "OCR Error", "Unknown", "Unknown"
                    lines = [f"OCR Error: {str(ocr_error)}"]
            else:
                # If OCR not available, just use filename as title
                title = os.path.splitext(filename)[0]
                collector, expansion = "Unknown", "Unknown"
                lines = ["OCR not available"]
            
            # Clean the expansion code and collector number
            cleaned_expansion = self.clean_expansion_code(expansion)
            cleaned_collector = self.clean_collector_number(collector)
            
            item = {
                'filename': filename, 
                'title': title, 
                'collector_number': collector,  # Store original value
                'cleaned_collector': cleaned_collector,  # Store cleaned value
                'expansion': expansion,  # Store original value
                'cleaned_expansion': cleaned_expansion,  # Store cleaned value
                'image_path': image_path, 
                'text_lines': lines, 
                'status': "Processed"
            }
                
            self.card_data.append(item)
            self.add_to_results(item)
            
            # Display first image automatically
            if len(self.card_data) == 1:
                self.display_image(item['image_path'])
                self.update_text_display(item['text_lines'])
                
            return True
        except Exception as e:
            print(f"Error processing {image_path}: {e}")
            return False

    def parse_card_data(self, text_lines, full_text):
        title = "Unknown"
        collector_number = "Unknown"
        expansion = "Unknown"
        
        # Skip processing if empty
        if not text_lines or "Empty" in full_text:
            return title, collector_number, expansion
        
        # Extract title (first non-empty line)
        for line in text_lines:
            if line.strip():
                title = line.strip()
                break
        
        # 1. Look for ".EN" pattern for expansion extraction
        for line in text_lines:
            dot_en_match = re.search(r'([A-Z]{3})\.EN', line)
            if dot_en_match:
                # Take the three characters before the dot as expansion
                expansion = dot_en_match.group(1)
                break
                
        # 2. Look for "EN" pattern for expansion 
        if expansion == "Unknown":
            for line in text_lines:
                # Find EN in the line
                if "EN" in line:
                    # Take all characters from beginning of line up to "EN"
                    en_index = line.find("EN")
                    if en_index > 0:  # Make sure there are characters before "EN"
                        expansion = line[:en_index].strip()
                        break
                
        # 3. Fallback for expansion - look for uppercase sequence
        if expansion == "Unknown":
            exp_match = re.search(r'\b[A-Z]{3,}\b', full_text)
            if exp_match:
                expansion = exp_match.group(0)
        
        # 1. Look for "Inc." pattern followed by slash for collector number
        inc_found = False
        for i, line in enumerate(text_lines):
            if "Inc." in line and i < len(text_lines) - 1:
                inc_found = True
                # Check the next lines after "Inc."
                for j in range(i+1, min(i+3, len(text_lines))):
                    slash_match = re.search(r'(\d+/\d+[CUMLR]?)', text_lines[j])
                    if slash_match:
                        collector_number = slash_match.group(1)
                        break
                if collector_number != "Unknown":
                    break
        
        # 2. Look for rarity letters at beginning of line followed by 4 numbers
        if collector_number == "Unknown":
            for line in text_lines:
                # Match lines that start with C, L, U, M, or R followed by 4 digits
                rarity_start_match = re.match(r'^([CLUMR])(\d{4})', line.strip())
                if rarity_start_match:
                    # Capture the 4 digits after the rarity letter
                    collector_number = rarity_start_match.group(2)
                    break
                    
        # 3. Look for collector number with rarity indication and slash
        if collector_number == "Unknown":
            rarity_slash_pattern = r'(\d+)/(\d+)([CUMLR])'
            for line in text_lines:
                rarity_match = re.search(rarity_slash_pattern, line)
                if rarity_match:
                    collector_number = f"{rarity_match.group(1)}/{rarity_match.group(2)}{rarity_match.group(3)}"
                    break
        
        # 4. Look for regular slash pattern
        if collector_number == "Unknown":
            slash_pattern = r'(\d+)/(\d+)'
            for line in text_lines:
                slash_match = re.search(slash_pattern, line)
                if slash_match:
                    collector_number = f"{slash_match.group(1)}/{slash_match.group(2)}"
                    break
        
        # 5. Look for standalone numbers
        if collector_number == "Unknown":
            for line in text_lines:
                # Look for digits that are not part of other patterns
                standalone_digits = re.findall(r'\b\d+\b', line)
                if standalone_digits:
                    # Use the first sequence of digits that stands alone
                    collector_number = standalone_digits[0]
                    break
        
        return title, collector_number, expansion

    def add_to_results(self, item):
        """Add item to results tree with cleaned values for display"""
        self.root.after(0, lambda: self.results_tree.insert(
            "", tk.END,
            values=(item['filename'], item['title'], item['cleaned_collector'], 
                   item['cleaned_expansion'], item['status']),
            tags=(item['image_path'],)
        ))

    def on_result_select(self, event):
        sel = self.results_tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.results_tree.item(item, 'values')
        path = self.results_tree.item(item, 'tags')[0]
        self.current_image_path = path
        self.display_image(path)
        for card in self.card_data:
            if card['image_path'] == path:
                self.update_text_display(card['text_lines'])
                self.title_var.set(vals[1])
                self.collector_var.set(vals[2])  # This now uses the cleaned collector number
                self.expansion_var.set(vals[3])  # This now uses the cleaned expansion code
                self.update_btn.config(state=tk.NORMAL)
                break

    def display_image(self, image_path):
        try:
            img = Image.open(image_path)
            max_w, max_h = 350, 500
            w, h = img.size
            if w > max_w or h > max_h:
                r = min(max_w/w, max_h/h)
                img = img.resize((int(w*r), int(h*r)), Image.LANCZOS)
            self.tk_image = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.tk_image)
        except Exception as e:
            error_msg = f"Error displaying image: {str(e)}"
            print(error_msg)
            self.status_var.set(error_msg)
            self.tk_image = None
            self.image_label.config(image='')

    def update_text_display(self, lines):
        self.text_display.delete(1.0, tk.END)
        self.text_display.insert(tk.END, "\n".join(lines) if lines else "No text detected in this image.")

    def update_selected_item(self):
        sel = self.results_tree.selection()
        if not sel: return
        item = sel[0]
        fv = self.results_tree.item(item, 'values')
        
        # Get values from UI fields
        new_title = self.title_var.get()
        new_collector = self.collector_var.get()
        new_expansion = self.expansion_var.get()
        
        # Clean the new values
        cleaned_collector = self.clean_collector_number(new_collector)
        cleaned_expansion = self.clean_expansion_code(new_expansion)
        
        # Update the treeview with cleaned values
        new_vals = (fv[0], new_title, cleaned_collector, cleaned_expansion, "Updated")
        self.results_tree.item(item, values=new_vals)
        
        for card in self.card_data:
            if card['image_path'] == self.current_image_path:
                # Update both original and cleaned values
                card['title'] = new_title
                card['collector_number'] = new_collector
                card['cleaned_collector'] = cleaned_collector
                card['expansion'] = new_expansion
                card['cleaned_expansion'] = cleaned_expansion
                card['status'] = "Updated"
                break
                
        self.status_var.set(f"Updated card data for {fv[0]}")
        self.log_message(f"Updated card: {fv[0]}, Title: {new_title}, Collector: {cleaned_collector}, Set: {cleaned_expansion}")

    def export_to_csv(self):
        if not self.card_data:
            messagebox.showinfo("Export", "No data to export")
            return
        output = self.output_csv.get()
        if not output:
            return
        try:
            # Use cleaned values for CSV export
            df = pd.DataFrame([{
                'filename': c['filename'], 
                'title': c['title'],
                'collector_number': c['cleaned_collector'],  # Use cleaned value
                'expansion': c['cleaned_expansion'],  # Use cleaned value
                'status': c['status']
            } for c in self.card_data])
            df.to_csv(output, index=False)
            self.status_var.set(f"Data exported successfully to {output}")
            self.log_message(f"Exported {len(self.card_data)} cards to {output}")
            messagebox.showinfo("Export Complete", f"Card data has been exported to {output}")
        except Exception as e:
            error_msg = f"Error exporting data: {str(e)}"
            print(error_msg)
            self.status_var.set(error_msg)
            self.log_message(f"Export error: {str(e)}")
            messagebox.showerror("Export Error", error_msg)

    def log_message(self, message):
        """Add a message to the log"""
        timestamp = time.strftime("%H:%M:%S")
        self.root.after(0, lambda: self.log.insert(tk.END, f"[{timestamp}] {message}\n"))
        self.root.after(0, lambda: self.log.see(tk.END))  # Scroll to end

    def update_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def update_progress(self):
        self.root.after(0, lambda: (
            self.progress_bar.config(value=int((self.processed_count/self.total_images)*100) if self.total_images else 0),
            self.progress_label.config(text=f"{self.processed_count}/{self.total_images} images processed")
        ))


def main():
    try:
        root = tk.Tk()
        root.title("Card Scanner App")
        
        # Try to set a theme - handle potential errors
        try:
            style = ttk.Style(root)
            style.theme_use('default')  # Try a safe default theme
        except Exception as e:
            print(f"Warning: Could not set theme: {str(e)}")
        
        # Create the app
        app = CardScannerApp(root)
        
        # Add a safety try-except to handle errors when running mainloop
        try:
            root.mainloop()
        except Exception as e:
            print(f"Error in main loop: {str(e)}")
    except Exception as e:
        print(f"Fatal error initializing application: {str(e)}")
        
        # Try to show an error dialog if possible
        try:
            import tkinter.messagebox as mb
            mb.showerror("Application Error", f"Failed to start the application: {str(e)}")
        except:
            # If even the messagebox fails, print to console
            print("Could not display error dialog.")


if __name__ == "__main__":
    main()