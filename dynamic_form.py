import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime
import json
import os
from document_generator import DocumentGenerator

class DynamicFormGUI:
    def __init__(self, root, template_file="eeg_template.json"):
        self.root = root
        self.template_file = template_file
        self.variables = {}
        self.widgets = {}
        self.doc_generator = DocumentGenerator()
        
        # Load template from JSON file
        self.load_template()
        
        # Create user interface
        self.create_interface()
    
    def load_template(self):
        """Load template from JSON file"""
        try:
            if not os.path.exists(self.template_file):
                self.create_default_template()
            
            with open(self.template_file, 'r', encoding='utf-8') as f:
                self.template = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Template loading error: {str(e)}")
            self.root.destroy()
            return
    
    def create_default_template(self):
        """Create default template file if it doesn't exist"""
        default_template = {
            "form_title": "🧠 EEG Score Checklist (ACNS Terminology)",
            "window_config": {"width": 800, "height": 900},
            "sections": [
                {
                    "id": "patient_info",
                    "title": "1. Patient and Study Information",
                    "fields": [
                        {"id": "patient_name", "label": "Patient Name:", "type": "text", "default": "", "required": True},
                        {"id": "patient_age", "label": "Age:", "type": "number", "default": "", "required": True}
                    ]
                }
            ],
            "buttons": [
                {"text": "Generate JSON", "action": "generate_json"},
                {"text": "Clear Form", "action": "clear_form"}
            ],
            "output_section": {"title": "📦 JSON Result", "height": 10}
        }
        
        with open(self.template_file, 'w', encoding='utf-8') as f:
            json.dump(default_template, f, ensure_ascii=False, indent=2)
    
    def create_interface(self):
        """Create interface based on template"""
        # Window configuration
        self.root.title(self.template.get('form_title', 'Dynamic Form'))
        window_config = self.template.get('window_config', {})
        width = window_config.get('width', 800)
        height = window_config.get('height', 900)
        self.root.geometry(f"{width}x{height}")
        
        # Create main frame with scrolling
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas and Scrollbar for scrolling
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create sections
        self.create_sections()
        
        # Create buttons
        self.create_buttons()
        
        # Create output section
        self.create_output_section()
        
        # Bind mouse wheel scrolling
        self.bind_mousewheel(canvas)
    
    def bind_mousewheel(self, canvas):
        """Bind mouse wheel scrolling"""
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def create_sections(self):
        """Create form sections based on template"""
        for section in self.template.get('sections', []):
            self.create_section(section)
    
    def create_section(self, section_config):
        """Create one form section"""
        frame = ttk.LabelFrame(self.scrollable_frame, text=section_config['title'], padding=10)
        frame.pack(fill=tk.X, pady=5)
        
        for i, field in enumerate(section_config.get('fields', [])):
            self.create_field(frame, field, i)
        
        frame.columnconfigure(1, weight=1)
    
    def create_field(self, parent, field_config, row):
        """Create form field based on configuration"""
        field_id = field_config['id']
        label_text = field_config['label']
        field_type = field_config['type']
        default_value = field_config.get('default', '')
        
        # Handle special value for date
        if default_value == "today":
            default_value = datetime.now().strftime("%Y-%m-%d")
        
        if field_type == 'checkbox':
            # Checkbox
            self.variables[field_id] = tk.BooleanVar(value=default_value)
            widget = ttk.Checkbutton(parent, text=label_text, 
                                   variable=self.variables[field_id])
            widget.grid(row=row, column=0, sticky="w", pady=2, columnspan=2)
            
        else:
            # Label
            ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=2)
            
            if field_type == 'text' or field_type == 'date' or field_type == 'number':
                # Text field
                self.variables[field_id] = tk.StringVar(value=str(default_value))
                widget = ttk.Entry(parent, textvariable=self.variables[field_id], width=40)
                widget.grid(row=row, column=1, sticky="ew", padx=(10,0))
                
            elif field_type == 'select':
                # Dropdown list
                self.variables[field_id] = tk.StringVar(value=str(default_value))
                options = field_config.get('options', [])
                widget = ttk.Combobox(parent, textvariable=self.variables[field_id],
                                    values=options, state="readonly", width=37)
                widget.grid(row=row, column=1, sticky="ew", padx=(10,0))
                
            elif field_type == 'textarea':
                # Multi-line text field
                height = field_config.get('height', 4)
                widget = scrolledtext.ScrolledText(parent, height=height, width=50)
                widget.grid(row=row, column=1, sticky="ew", padx=(10,0))
                widget.insert("1.0", str(default_value))
                # Store textarea widget separately
                self.widgets[field_id] = widget
        
        # Store widget for possible future use
        if field_id not in self.widgets:
            self.widgets[field_id] = widget
    
    def create_buttons(self):
        """Create buttons based on template"""
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill=tk.X, pady=10)
        
        for i, button_config in enumerate(self.template.get('buttons', [])):
            text = button_config['text']
            action = button_config['action']
            
            # Bind action to method
            command = getattr(self, action, None)
            if command:
                ttk.Button(frame, text=text, command=command).pack(side=tk.LEFT, padx=(0,10))
        
        # Add document generation buttons
        ttk.Button(frame, text="Save to Word", command=self.save_to_word).pack(side=tk.LEFT, padx=(0,10))
        ttk.Button(frame, text="Save to PDF", command=self.save_to_pdf).pack(side=tk.LEFT)
    
    def create_output_section(self):
        """Create output section for results"""
        output_config = self.template.get('output_section', {})
        title = output_config.get('title', 'Result')
        height = output_config.get('height', 10)
        
        frame = ttk.LabelFrame(self.scrollable_frame, text=title, padding=10)
        frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.output_text = scrolledtext.ScrolledText(frame, height=height, width=70)
        self.output_text.pack(fill=tk.BOTH, expand=True)
    
    def collect_data(self):
        """Collect data from all form fields"""
        data = {}
        
        # Collect data from regular variables
        for field_id, var in self.variables.items():
            if isinstance(var, tk.BooleanVar):
                data[field_id] = var.get()
            else:
                value = var.get()
                # Try to convert numeric values
                if self.is_number_field(field_id) and value:
                    try:
                        data[field_id] = int(value) if value.isdigit() else float(value)
                    except ValueError:
                        data[field_id] = value
                else:
                    data[field_id] = value
        
        # Collect data from textarea fields
        for field_id, widget in self.widgets.items():
            if isinstance(widget, scrolledtext.ScrolledText):
                data[field_id] = widget.get("1.0", tk.END).strip()
        
        return data
    
    def is_number_field(self, field_id):
        """Check if field is numeric"""
        for section in self.template.get('sections', []):
            for field in section.get('fields', []):
                if field['id'] == field_id:
                    return field['type'] == 'number'
        return False
    
    def generate_json(self):
        """Generate JSON from form data"""
        data = self.collect_data()
        
        # Format JSON
        json_output = json.dumps(data, ensure_ascii=False, indent=2)
        
        # Display result
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", json_output)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(json_output)
        
        messagebox.showinfo("Done", "JSON generated and copied to clipboard!")
    
    def clear_form(self):
        """Clear all form fields"""
        for section in self.template.get('sections', []):
            for field in section.get('fields', []):
                field_id = field['id']
                default_value = field.get('default', '')
                
                if field_id in self.variables:
                    var = self.variables[field_id]
                    if isinstance(var, tk.BooleanVar):
                        var.set(bool(default_value))
                    else:
                        if default_value == "today":
                            default_value = datetime.now().strftime("%Y-%m-%d")
                        var.set(str(default_value))
                
                # Clear textarea fields
                if field_id in self.widgets and isinstance(self.widgets[field_id], scrolledtext.ScrolledText):
                    widget = self.widgets[field_id]
                    widget.delete("1.0", tk.END)
                    if default_value:
                        widget.insert("1.0", str(default_value))
        
        # Clear output field
        self.output_text.delete("1.0", tk.END)
    
    def save_to_word(self):
        """Save form data to Word document"""
        try:
            data = self.collect_data()
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                title="Save Word document"
            )
            
            if filename:
                self.doc_generator.generate_word_document(data, filename)
                messagebox.showinfo("Success", f"Document saved to:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving Word document:\n{str(e)}")
    
    def save_to_pdf(self):
        """Save form data to PDF document"""
        try:
            data = self.collect_data()
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF documents", "*.pdf"), ("All files", "*.*")],
                title="Save PDF document"
            )
            
            if filename:
                self.doc_generator.generate_pdf_document(data, filename)
                messagebox.showinfo("Success", f"Document saved to:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving PDF document:\n{str(e)}")
    
    def load_template_file(self):
        """Load new template file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Select template file"
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    self.template = json.load(f)
                
                # Reload interface
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                
                self.variables.clear()
                self.widgets.clear()
                
                self.create_sections()
                self.create_buttons()
                self.create_output_section()
                
                messagebox.showinfo("Success", "Template loaded successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Template loading error: {str(e)}")

def main():
    root = tk.Tk()
    
    # Add menu for loading templates
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    
    # Create application
    app = DynamicFormGUI(root)
    
    # Add menu item for loading template
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Load Template...", command=app.load_template_file)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)
    
    root.mainloop()

if __name__ == "__main__":
    main()