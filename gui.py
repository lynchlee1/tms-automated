import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

import scrape
class ScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Update")
        self.root.geometry("800x700")
        self.root.resizable(False, False)
        
        # Configure style for better appearance
        self.setup_styles()
        
        # Center the window
        self.center_window()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="30")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Web Scraper", 
                               font=("Segoe UI", 18, "bold"), style="Title.TLabel")
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 30))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=25)
        
        # Run button
        self.run_button = ttk.Button(button_frame, text="실행", command=self.run_scraper, 
                                   style="Accent.TButton", width=10)
        self.run_button.pack(side=tk.LEFT, padx=(0, 15))
        
        # Headless mode button
        self.headless_var = tk.BooleanVar(value=False)
        self.headless_button = ttk.Button(button_frame, text="Chrome 숨기기 선택됨", 
                                          command=self.toggle_headless, 
                                          style="TButton", width=30)
        self.headless_button.pack(side=tk.LEFT, padx=(0, 15))
        
        # Exit button
        exit_button = ttk.Button(button_frame, text="종료", command=self.root.quit, 
                               style="TButton", width=10)
        exit_button.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', style="TProgressbar")
        self.progress.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="", font=("Segoe UI", 9))
        self.status_label.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def toggle_headless(self):
        """Toggle headless mode and update button appearance"""
        current_value = self.headless_var.get()
        self.headless_var.set(not current_value)
        
        # Update button appearance based on state
        if self.headless_var.get():
            self.headless_button.config(text="Chrome 숨기기 선택됨", style="Accent.TButton")
        else:
            self.headless_button.config(text="Chrome 열기 선택됨", style="TButton")
    
    def setup_styles(self):
        """Configure custom styles for better appearance"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Define custom colors
        blue = "#0d7efb"      # (13, 126, 251)
        grey = "#616365"      # (97, 99, 101)
        white = "#f5f5f5"     # (245, 245, 245)
        
        # Configure title style
        style.configure("Title.TLabel", 
                       font=("Segoe UI", 18, "bold"),
                       foreground=blue)
        
        # Configure label style
        style.configure("TLabel", 
                       font=("Segoe UI", 10),
                       foreground=grey,
                       background=white)
        
        # Configure entry style
        style.configure("TEntry", 
                       font=("Segoe UI", 10),
                       fieldbackground="white",
                       borderwidth=2,
                       relief="solid",
                       bordercolor=grey)
        
        # Configure button styles
        style.configure("Accent.TButton",
                       font=("Segoe UI", 11, "bold"),
                       foreground="white",
                       background=blue,
                       borderwidth=0,
                       focuscolor="none")
        
        style.map("Accent.TButton",
                 background=[("active", "#0a6bdf"),
                           ("pressed", "#0859c7")])
        
        # Configure regular button
        style.configure("TButton",
                       font=("Segoe UI", 10, "bold"),
                       foreground=grey,
                       background=white,
                       borderwidth=1,
                       bordercolor=grey,
                       focuscolor="none")
        
        style.map("TButton",
                 background=[("active", "#e8e8e8"),
                           ("pressed", "#dcdcdc")])
        
        # Configure progress bar
        style.configure("TProgressbar",
                       background=blue,
                       troughcolor=white,
                       borderwidth=0,
                       lightcolor=blue,
                       darkcolor=blue)
        
        # Configure frame background
        style.configure("TFrame",
                       background=white)
        
        # Configure LabelFrame style
        style.configure("TLabelframe",
                       background=white,
                       foreground=grey,
                       font=("Segoe UI", 9, "bold"))
        
        style.configure("TLabelframe.Label",
                       background=white,
                       foreground=blue,
                       font=("Segoe UI", 9, "bold"))
        
        # Configure Checkbutton style
        style.configure("TCheckbutton",
                       background=white,
                       foreground=grey,
                       font=("Segoe UI", 10))
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def run_scraper(self):
        """Run the scraper script"""
        # Disable run button and start progress
        self.run_button.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="실행 중...")
        
        # Get headless mode setting
        headless_mode = self.headless_var.get()
        
        # Run script in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.execute_scraper, args=(headless_mode,))
        thread.daemon = True
        thread.start()
    
    def execute_scraper(self, headless_mode):
        """Execute the scraper directly"""
        try:
            # Run the main scraper function
            scrape.scrape_once(headless=headless_mode)
            
            # Create a mock result object
            class MockResult:
                def __init__(self):
                    self.returncode = 0
                    self.stderr = ""
            
            result = MockResult()
            self.root.after(0, self.scraper_finished, result)
            
        except Exception as e:
            self.root.after(0, self.scraper_error, str(e))
    
    def scraper_finished(self, result):
        """Handle scraper completion"""
        self.progress.stop()
        self.run_button.config(state='normal')
        self.status_label.config(text="")
        
        if result.returncode == 0:
            print("실행 완료!")
            messagebox.showinfo("완료", "실행 완료")
        else:
            print("오류 발생")
            messagebox.showerror("오류", f"오류 발생:\n{result.stderr}")
    
    def scraper_error(self, error_msg):
        """Handle scraper error"""
        self.progress.stop()
        self.run_button.config(state='normal')
        self.status_label.config(text="")
        print("오류 발생")
        messagebox.showerror("오류", f"실행 실패:\n{error_msg}")

def main():
    root = tk.Tk()
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create and run the GUI
    app = ScraperGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

