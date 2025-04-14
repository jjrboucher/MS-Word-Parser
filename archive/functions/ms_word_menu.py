import tkinter as tk
from tkinter import ttk, filedialog
import time

# Variables for log files
timestamp = time.strftime("%Y%m%d_%H%M%S")
log_file = "DOCx_Parser_Log_" + timestamp + ".log"
error_log_file = "DOCx_Error_Log_" + timestamp + ".log"

def docx_menu():
    # Create the main application window (the parent window)
    root = tk.Tk()
    root.title("MS Word Parsing")

    # Define some style settings
    style = ttk.Style()
    style.configure("TFrame", background="#f0f0f0")
    style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
    style.configure("TButton", background="#d0d0d0", font=("Arial", 10))
    style.configure("TRadiobutton", background="#f0f0f0", font=("Arial", 10))
    style.configure("TCheckbutton", background="#f0f0f0", font=("Arial", 10))

    # Create a frame for the parent menu with padding and background color
    parent_frame = ttk.Frame(root, padding="10", style="TFrame")
    parent_frame.grid(row=0, column=0, sticky="NSEW")

    # Create a Label for the parent frame
    parent_label = ttk.Label(parent_frame, text="Main Menu", style="TLabel")
    parent_label.grid(row=0, column=0, sticky="W", pady=5)

    # Create the first child frame within the parent frame for parsing options
    parsing_frame = ttk.LabelFrame(parent_frame, text="Parsing Options", padding="10")
    parsing_frame.grid(row=1, column=0, sticky="W", pady=10)

    # Variables to be updated and returned
    excel_file = None
    docx_files = []
    radio_option = None
    hash_files = False
    clicked_button = ""

    # Create a StringVar to hold the selected radio button value
    radio_var = tk.StringVar(value="triage")

    # Create a BooleanVar to hold the state of the checkbox
    hash_var = tk.BooleanVar(value=False)

    # Add radio buttons for "Triage" and "Full" on the same row inside the child frame
    triage_radio = ttk.Radiobutton(parsing_frame, text="Triage", variable=radio_var, value="triage",
                                   style="TRadiobutton")
    triage_radio.grid(row=0, column=0, sticky="W", padx=5)

    full_radio = ttk.Radiobutton(parsing_frame, text="Full", variable=radio_var, value="full", style="TRadiobutton")
    full_radio.grid(row=0, column=1, sticky="W", padx=5)

    # Add a checkbox for "Hash files" inside the child frame
    hash_checkbox = ttk.Checkbutton(parsing_frame, text="Hash files", variable=hash_var, style="TCheckbutton")
    hash_checkbox.grid(row=1, column=0, columnspan=2, sticky="W", padx=5, pady=5)

    # Bind the radio buttons and checkbox to the print function
    triage_radio.config()
    full_radio.config()
    hash_checkbox.config()

    # Create the second child frame within the parent frame for Excel output file selection
    file_frame = ttk.LabelFrame(parent_frame, text="Excel Output File", padding="10")
    file_frame.grid(row=2, column=0, sticky="W", pady=10)

    # Function to open a file dialog to select an existing file or create a new one
    def select_or_create_file():
        nonlocal excel_file
        file_path = filedialog.asksaveasfilename(
            title="Select or Create Excel File",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
            confirmoverwrite=True  # This allows the user to select an existing file without overwriting it
        )
        if file_path:
            excel_file = file_path
            selected_file_var.set(file_path)
            update_process_button_state()

    # Create a StringVar to hold the selected file path
    selected_file_var = tk.StringVar(value="No file selected")

    # Add a button to either select an existing file or create a new one
    select_file_button = ttk.Button(file_frame, text="Select/Create File", command=select_or_create_file,
                                    style="TButton")
    select_file_button.grid(row=0, column=0, sticky="W", pady=5, padx=5)

    # Display the selected file path in the second child frame
    selected_file_label = ttk.Label(file_frame, textvariable=selected_file_var, style="TLabel")
    selected_file_label.grid(row=1, column=0, columnspan=2, sticky="W", pady=5, padx=5)

    # Create the third child frame within the parent frame for Log Files
    log_frame = ttk.LabelFrame(parent_frame, text="Log Files", padding="10")
    log_frame.grid(row=3, column=0, sticky="W", pady=10)

    # Display the log file path in the third child frame
    log_file_label = ttk.Label(log_frame, text="Log File:", style="TLabel")
    log_file_label.grid(row=0, column=0, sticky="W", padx=5)

    log_file_value_label = ttk.Label(log_frame, text=log_file, style="TLabel")
    log_file_value_label.grid(row=0, column=1, sticky="W", padx=5)

    # Display the error log file path in the third child frame
    error_log_file_label = ttk.Label(log_frame, text="Error Log File:", style="TLabel")
    error_log_file_label.grid(row=1, column=0, sticky="W", padx=5)

    error_log_file_value_label = ttk.Label(log_frame, text=error_log_file, style="TLabel")
    error_log_file_value_label.grid(row=1, column=1, sticky="W", padx=5)

    # Create a new frame to the right of the existing frames for DOCx file selection
    docx_frame = ttk.LabelFrame(parent_frame, text="DOCx File Selection", padding="10")
    docx_frame.grid(row=1, column=1, rowspan=3, sticky="NSEW", padx=10, pady=10)

    # Create a canvas and scrollbar for the scrollable area
    canvas = tk.Canvas(docx_frame, width=800, height=200, bg="#ffffff")  # Increased width to reduce wrapping
    canvas.grid(row=1, column=0, sticky="NSEW")

    scrollbar = ttk.Scrollbar(docx_frame, orient="vertical", command=canvas.yview)
    scrollbar.grid(row=1, column=1, sticky="NS")

    # Create a frame to contain the text inside the canvas
    scrollable_frame = ttk.Frame(canvas, padding="5")
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # Update the scrollbar and canvas when the size of the content changes
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", on_frame_configure)

    # Create a label to show the number of selected DOCx files
    num_files_label = ttk.Label(docx_frame, text="No files selected", foreground="blue", font=("Arial", 12, "bold"))
    num_files_label.grid(row=0, column=0, sticky="W", padx=5, pady=5)

    # Function to open a file dialog to select one or more DOCx files
    def select_docx_files():
        nonlocal docx_files
        file_paths = filedialog.askopenfilenames(
            title="Select DOCx Files",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*"))
        )
        if file_paths:
            docx_files = list(file_paths)
            num_files_label.config(text=f"{len(docx_files)} file(s) selected", foreground="green")

            for widget in scrollable_frame.winfo_children():
                widget.destroy()

            if len(docx_files) < 1001:  # if too many files, it causes a problem to display them.
                for file in docx_files:
                    file_label = ttk.Label(scrollable_frame, text=file, wraplength=780, anchor="w",
                                           justify="left")  # Adjusted wraplength
                    file_label.pack(fill="x", pady=2)
            else:
                file_label = ttk.Label(scrollable_frame, text="Too many files to list...", wraplength=780, anchor="w",
                                       justify="left")
                file_label.pack(fill="x", pady=2)
            update_process_button_state()

    # Add a button to select DOCx files
    select_docx_button = ttk.Button(docx_frame, text="Select DOCx Files", command=select_docx_files, style="TButton")
    select_docx_button.grid(row=2, column=0, sticky="W", pady=5, padx=5)

    # Configure grid weight for the frame and canvas
    docx_frame.grid_columnconfigure(0, weight=1)
    docx_frame.grid_rowconfigure(1, weight=1)

    # Function to handle button clicks and gather relevant information
    def button_clicked(button):
        nonlocal excel_file, docx_files, radio_option, hash_files, clicked_button
        radio_option = radio_var.get()
        hash_files = hash_var.get()
        clicked_button = button.cget('text')
        # Exit the application
        root.destroy()

    # Create and place "PROCESS" and "CANCEL" buttons at the bottom of the main frame
    process_button = tk.Button(parent_frame, text="PROCESS", bg="grey", fg="white", font=("Arial", 12, "bold"),
                               width=20, state="disabled", command=lambda: button_clicked(process_button))
    process_button.grid(row=4, column=0, padx=5, pady=10, sticky="EW")

    cancel_button = tk.Button(parent_frame, text="CANCEL", bg="red", fg="white", font=("Arial", 12, "bold"), width=20,
                              command=lambda: button_clicked(cancel_button))
    cancel_button.grid(row=4, column=1, padx=5, pady=10, sticky="E")

    # Configure the grid to center the buttons
    parent_frame.grid_columnconfigure(0, weight=1)
    parent_frame.grid_columnconfigure(1, weight=1)
    parent_frame.grid_rowconfigure(4, weight=0)

    # Function to update the state of the PROCESS button
    def update_process_button_state():
        excel_file_selected = selected_file_var.get() != "No file selected"
        docx_files_selected = num_files_label.cget("text") != "No files selected"
        if excel_file_selected and docx_files_selected:
            process_button.config(state="normal", bg="green")
        else:
            process_button.config(state="disabled", bg="grey")

    # Start the Tkinter event loop
    root.mainloop()

    return clicked_button, log_file, error_log_file, radio_option, hash_files, excel_file, docx_files

