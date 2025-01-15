import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

def load_file(file_path):
    """Loads an Excel or CSV file into a DataFrame."""
    if file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file type. Please provide a .xlsx or .csv file.")

def save_file(data, output_file):
    """Saves a DataFrame to an Excel or CSV file."""
    if output_file.endswith('.xlsx'):
        data.to_excel(output_file, index=False)
    elif output_file.endswith('.csv'):
        data.to_csv(output_file, index=False)
    else:
        raise ValueError("Unsupported file type. Please provide a .xlsx or .csv file.")

def format_health_checks(df):
    """Formats the health checks DataFrame into the desired structure."""
    df.columns = df.iloc[0]
    df = df[1:]
    formatted_data = []
    for _, row in df.iterrows():
        desexed_status = "Desexed" if row["Desexed"].strip().lower() == "yes" else "ENTIRE"
        combined_entry = {
            "Customer & Pet": f"{row['Customer'][0]}, {row['Pet']}",
            "Check-In Date": row["Checkin"],
            "Check-Out Date": row["Checkout"],
            "Sex & Desexed": f"{row['Sex']}, {desexed_status}",
            "Breed": row["Breed"],
            "Short Checkin": row["Checkin"].split()[1] if ' ' in row["Checkin"] else row["Checkin"],
        }
        formatted_data.append(combined_entry)
    return pd.DataFrame(formatted_data)

def process_file(input_file, output_file):
    """Processes the file, formats it, and saves the result."""
    try:
        data = load_file(input_file)
        formatted_data = format_health_checks(data)
        save_file(formatted_data, output_file)
        messagebox.showinfo("Success", f"Formatted data has been saved to {output_file}.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_file():
    """Opens a file dialog for input file selection."""
    return filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])

def save_as_file():
    """Opens a file dialog for output file selection."""
    return filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])

def main_gui():
    """Main GUI for the Pet Bookings Formatter."""
    def format_action():
        input_file = open_file()
        if not input_file:
            return
        output_file = save_as_file()
        if not output_file:
            return
        process_file(input_file, output_file)

    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Pet Bookings Formatter")
    root.geometry("300x150")

    # Add a title label
    title_label = ctk.CTkLabel(root, text="Pet Bookings Formatter", font=("Helvetica", 20))
    title_label.pack(pady=20)

    # Add a format button
    format_button = ctk.CTkButton(root, text="Select and Format File", command=format_action, font=("Helvetica", 16))
    format_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main_gui()
