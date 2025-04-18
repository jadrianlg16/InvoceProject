# main.py
# Main entry point for the Invoice PDF Processor application.

import tkinter as tk
from tkinter import messagebox

# --- Import the GUI class and PIL availability status ---
# Ensure gui.py is in the same directory or accessible via PYTHONPATH
try:
    from gui import InvoiceProcessorApp, PIL_AVAILABLE
except ImportError as e:
    print(f"Error importing GUI module: {e}")
    print("Make sure 'gui.py' is in the same directory.")
    # Show a simple Tkinter error if possible, otherwise just exit
    try:
        root_err = tk.Tk()
        root_err.withdraw()
        messagebox.showerror("Import Error", "Could not load the GUI module (gui.py).\nApplication cannot start.")
        root_err.destroy()
    except Exception:
        pass
    exit(1)


# --- Main Execution ---
if __name__ == "__main__":
    # Perform the initial Pillow check (optional, but good practice)
    if not PIL_AVAILABLE:
         # Use a temporary root window for the messagebox if the main one isn't created yet
         root_check = tk.Tk()
         root_check.withdraw() # Hide the temporary window
         messagebox.showwarning("Dependency Warning",
                                "Python Imaging Library (Pillow) not found.\n"
                                "PDF preview will be disabled.\n\n"
                                "Install it using:\npip install Pillow",
                                parent=None) # No parent needed here
         root_check.destroy() # Clean up the temporary window

    # Create the main application window
    root = tk.Tk()

    # Instantiate the application class (passing the root window)
    app = InvoiceProcessorApp(root)

    # Start the Tkinter event loop
    root.mainloop()
