import tkinter as tk
from tkinter import ttk, messagebox
import os
from maintenance_IMS_AUTO import main as run_automation

class IMSSelector:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IMS Selection")
        self.root.geometry("400x300")
        
        # Center the window
        self.center_window()
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add label
        ttk.Label(main_frame, text="Select IMS Template:", font=('Helvetica', 12)).grid(row=0, column=0, pady=10)
        
        # Create frame for buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, pady=10)
        
        # Get template options
        self.template_options = self.get_template_options()
        
        # Create buttons for each template
        for i, template in enumerate(self.template_options):
            ttk.Button(button_frame, 
                      text=template,
                      command=lambda t=template: self.select_template(t)).grid(
                          row=i, column=0, pady=5, padx=10, sticky='ew')
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def get_template_options(self):
        """Get list of template folders"""
        template_dir = "./IMS_TEMPLATE_COPIES"
        try:
            if not os.path.exists(template_dir):
                os.makedirs(template_dir)
            return [d for d in os.listdir(template_dir) 
                   if os.path.isdir(os.path.join(template_dir, d))]
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read template directory: {e}")
            return []
    
    def select_template(self, template):
        """Handle template selection"""
        self.root.withdraw()  # Hide the main window
        try:
            # Here you would call your main automation function
            # passing the selected template as parameter
            run_automation(template)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process template: {e}")
        finally:
            self.root.destroy()
    
    def run(self):
        """Start the UI"""
        if not self.template_options:
            messagebox.showerror("Error", "No templates found in IMS_TEMPLATE_COPIES directory")
            self.root.destroy()
            return
        self.root.mainloop()

if __name__ == "__main__":
    app = IMSSelector()
    app.run()
