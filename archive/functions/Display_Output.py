import tkinter as tk
from tkinter import ttk


def button_clicked(cancel_button):
    print("Exiting the application")
    cancel_button.quit()


def output_menu(log_file, error_log_file, folder, file_count, file_error_count, excel_file, start_time, end_time):
    # Create the main application window (the parent window)

    root = tk.Tk()
    root.title("MS Word Parsing")

    # Create a frame for the parent menu with padding and background color
    parent_frame = ttk.Frame(root, padding="10", style="TFrame")
    parent_frame.grid(row=0, column=0, sticky="NSEW")

    # Create a Label for the parent frame
    parent_label = ttk.Label(parent_frame, text="Processing Results", style="TLabel")
    parent_label.grid(row=0, column=0, sticky="W", pady=5)

    # Create the first child frame within the parent frame for parsing options
    parsing_frame = ttk.LabelFrame(parent_frame, text="Parsing Options", padding="10")
    parsing_frame.grid(row=1, column=0, sticky="W", pady=10)

    # Create the first child frame within the parent frame for Log Files
    input_frame = ttk.LabelFrame(parent_frame, text="Input", padding="10")
    input_frame.grid(row=1, column=0, sticky="W", pady=10)

    # Display the folder path in the first child frame
    folder_label = ttk.Label(input_frame, text="Path to files processed: ", style="TLabel")
    folder_label.grid(row=0, column=0, sticky="W", padx=5)

    folder_value_label = ttk.Label(input_frame, foreground="green", font="bold", text=folder, style="TLabel")
    folder_value_label.grid(row=0, column=1, sticky="W", padx=5)

    # Display the file processing summary log file path in the first child frame
    file_count_label = ttk.Label(input_frame, text="# of files submitted for processing:", style="TLabel")
    file_count_label.grid(row=1, column=0, sticky="W", padx=5)

    file_count_value_label = ttk.Label(input_frame, foreground="green", font=("Arial", 12, "bold"),
                                       text=file_count, style="TLabel")
    file_count_value_label.grid(row=1, column=1, sticky="W", padx=5)

    file_error_count_label = ttk.Label(input_frame, text="# of files not processed due to an error:", style="TLabel")
    file_error_count_label.grid(row=2, column=0, sticky="W", padx=5)

    file_error_count_value_label = ttk.Label(input_frame, foreground="red", font=("Arial", 12, "bold"),
                                             text=file_error_count, style="TLabel")
    file_error_count_value_label.grid(row=2, column=1, sticky="W", padx=5)

    results_label = ttk.Label(input_frame, text="Excel output file: ", style="TLabel")
    results_label.grid(row=3, column=0, sticky="W", padx=5)

    results_value_label = ttk.Label(input_frame, foreground="green", font="bold", text=excel_file, style="TLabel")
    results_value_label.grid(row=3, column=1, sticky="W", padx=5)

    # Create the second child frame within the parent frame for Log Files
    log_frame = ttk.LabelFrame(parent_frame, text="Log Files", padding="10")
    log_frame.grid(row=2, column=0, sticky="W", pady=10)

    # Display the log file path in the third child frame
    log_file_label = ttk.Label(log_frame, text="Log File:", style="TLabel")
    log_file_label.grid(row=0, column=0, sticky="W", padx=5)

    log_file_value_label = ttk.Label(log_frame, foreground="green", font="bold", text=log_file, style="TLabel")
    log_file_value_label.grid(row=0, column=1, sticky="W", padx=5)

    # Display the error log file path in the third child frame
    error_log_file_label = ttk.Label(log_frame, text="Error Log File (if applicable):", style="TLabel")
    error_log_file_label.grid(row=1, column=0, sticky="W", padx=5)

    error_log_file_value_label = ttk.Label(log_frame, foreground="red", font="bold",
                                           text=error_log_file, style="TLabel")
    error_log_file_value_label.grid(row=1, column=1, sticky="W", padx=5)

    # Create the third child frame within the parent frame for execution time
    execution_frame = ttk.LabelFrame(parent_frame, text="Execution Time", padding="10")
    execution_frame.grid(row=0, column=0, sticky="W", pady=10)

    # Create the third child frame within the parent frame for Log Files
    execution_frame = ttk.LabelFrame(parent_frame, text="Execution", padding="10")
    execution_frame.grid(row=3, column=0, sticky="W", pady=10)

    # Display the execution start time in the third child frame
    start_time_label = ttk.Label(execution_frame, text="Start Time:", style="TLabel")
    start_time_label.grid(row=0, column=0, sticky="W", padx=5)

    start_time_value_label = ttk.Label(execution_frame, foreground="green", text=start_time, style="TLabel")
    start_time_value_label.grid(row=0, column=1, sticky="W", padx=5)

    # Display the execution end time in the third child frame
    end_time_label = ttk.Label(execution_frame, text="End time:", style="TLabel")
    end_time_label.grid(row=1, column=0, sticky="W", padx=5)

    end_time_value_label = ttk.Label(execution_frame, foreground="green", text=end_time, style="TLabel")
    end_time_value_label.grid(row=1, column=1, sticky="W", padx=5)

    # Create and place "EXIT" button at the bottom of the main frame

    cancel_button = tk.Button(parent_frame, text="EXIT", bg="red", fg="white", font=("Arial", 12, "bold"), width=20,
                              command=lambda: button_clicked(cancel_button))
    cancel_button.grid(row=4, column=0, padx=5, pady=10, sticky="EW")

    # Start the Tkinter event loop
    root.mainloop()
