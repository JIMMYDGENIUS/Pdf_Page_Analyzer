import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader
import csv
from datetime import datetime

class PDFAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF Page Size Analyzer")
        self.window.geometry("1024x768")
        
        # Configure window styling
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors and styles
        self.window.configure(bg='#f0f0f0')
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Custom.TButton', padding=5)
        
        # Store analysis results
        self.results = []
        
        self.create_gui()
        
    def create_gui(self):
        # Create main frame with padding
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights for responsive layout
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        
        ttk.Label(header_frame, text="PDF Page Size Analyzer", style='Header.TLabel').pack(side='left')
        
        # Button frame
        button_frame = ttk.Frame(header_frame)
        button_frame.pack(side='right')
        
        select_btn = ttk.Button(button_frame, text="Select PDF File", command=self.select_file, style='Custom.TButton')
        select_btn.pack(side='left', padx=5)
        
        export_btn = ttk.Button(button_frame, text="Export Results", command=self.export_results, style='Custom.TButton')
        export_btn.pack(side='left', padx=5)
        
        # Create table with improved styling
        table_frame = ttk.Frame(main_frame)
        table_frame.grid(row=1, column=0, sticky='nsew', pady=(0, 20))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        columns = ('Page', 'Width (mm)', 'Height (mm)', 'Size & Recommended Paper')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', style='Custom.Treeview')
        
        # Configure treeview styling
        self.style.configure('Custom.Treeview', rowheight=25)
        self.style.configure('Custom.Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Set column headings and widths
        column_widths = {
            'Page': 100,
            'Width (mm)': 150,
            'Height (mm)': 150,
            'Size & Recommended Paper': 400
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], minwidth=column_widths[col])
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Grid layout for table and scrollbars
        self.tree.grid(row=0, column=0, sticky='nsew')
        y_scrollbar.grid(row=0, column=1, sticky='ns')
        x_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Summary section with improved styling
        summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding="10")
        summary_frame.grid(row=2, column=0, sticky='ew', pady=(0, 20))
        summary_frame.grid_columnconfigure(0, weight=1)
        
        self.summary_text = tk.Text(summary_frame, height=6, width=60, font=('Segoe UI', 10),
                                   wrap=tk.WORD, relief='flat', padx=10, pady=10)
        self.summary_text.grid(row=0, column=0, sticky='ew')
        
        # Add status bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief='sunken', padding=(10, 5))
        status_bar.grid(row=3, column=0, sticky='ew')
        self.status_var.set("Ready")
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")],
            title="Select PDF File"
        )
        if file_path:
            self.status_var.set(f"Analyzing: {os.path.basename(file_path)}")
            self.window.update()
            self.analyze_pdf(file_path)
            self.status_var.set("Analysis complete")
    
    def analyze_pdf(self, file_path):
        try:
            # Clear previous results
            self.results.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            reader = PdfReader(file_path)
            size_count = {}
            
            for page_num, page in enumerate(reader.pages, 1):
                # Get page size in points
                width_pt = float(page.mediabox.width)
                height_pt = float(page.mediabox.height)
                
                # Convert to millimeters (1 pt = 0.352778 mm)
                width_mm = round(width_pt * 0.352778, 1)
                height_mm = round(height_pt * 0.352778, 1)
                
                # Determine page size
                result_size = self.determine_page_size(width_mm, height_mm)
                
                # Update size count
                size_count[result_size] = size_count.get(result_size, 0) + 1
                
                # Store result
                self.results.append({
                    'page': page_num,
                    'width': width_mm,
                    'height': height_mm,
                    'size': result_size
                })
                
                # Add to table
                self.tree.insert('', tk.END, values=(
                    page_num,
                    f"{width_mm:.1f}",
                    f"{height_mm:.1f}",
                    result_size
                ))
            
            # Update summary
            self.update_summary(size_count)
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Error analyzing PDF: {str(e)}")
    
    def determine_page_size(self, width_mm, height_mm):
        # Always use larger dimension as width
        width_mm, height_mm = max(width_mm, height_mm), min(width_mm, height_mm)
        
        # Standard page sizes (width × height in mm)
        sizes = {
            'A0': (841, 1189),
            'A1': (594, 841),
            'A2': (420, 594),
            'A3': (297, 420),
            'A4': (210, 297),
            'A5': (148, 210),
            'Letter': (215.9, 279.4),
            'Legal': (215.9, 355.6)
        }
        
        actual_size = None
        recommended_size = None
        
        # Find actual size with tolerance
        tolerance = 5
        for size_name, (std_width, std_height) in sizes.items():
            if (abs(width_mm - std_width) <= tolerance and 
                abs(height_mm - std_height) <= tolerance):
                actual_size = size_name
                break
        
        # Find recommended size (smallest size that fits the content)
        for size_name, (std_width, std_height) in sorted(sizes.items(), key=lambda x: x[1][0] * x[1][1]):
            if width_mm <= std_width + tolerance and height_mm <= std_height + tolerance:
                recommended_size = size_name
                break
        
        actual_size = actual_size or f"Custom ({width_mm:.1f}×{height_mm:.1f}mm)"
        recommended_size = recommended_size or "Custom (Too large for standard sizes)"
        
        return f"{actual_size} (Print on: {recommended_size})"
    
    def update_summary(self, size_count):
        summary = "Page Size Summary:\n\n"
        for size, count in size_count.items():
            summary += f"{size}: {count} pages\n"
        
        self.summary_text.delete('1.0', tk.END)
        self.summary_text.insert('1.0', summary)
    
    def export_results(self):
        if not self.results:
            self.status_var.set("Warning: No results to export")
            messagebox.showwarning("Warning", "No results to export!")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"pdf_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            title="Export Results"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=['page', 'width', 'height', 'size'])
                    writer.writeheader()
                    writer.writerows(self.results)
                self.status_var.set("Results exported successfully")
                messagebox.showinfo("Success", "Results exported successfully!")
            except Exception as e:
                self.status_var.set(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Error exporting results: {str(e)}")
    
    def on_closing(self):
        """Handle window closing event"""
        self.window.quit()
        self.window.destroy()
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = PDFAnalyzer()
    app.run() 