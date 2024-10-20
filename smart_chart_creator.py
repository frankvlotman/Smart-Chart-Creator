import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar, Listbox, MULTIPLE, END, Toplevel, Text
from PIL import Image

# Define the path for the blank icon
icon_path = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

# Create the blank ICO file
create_blank_ico(icon_path)



# Global variable to track chart size
small_chart = False

# Predefined pastel color options
pastel_color_options = ['Brown', 'Gold', 'Green', 'Blue', 'Orange']
color_map = {
    'Brown': '#B29287',
    'Gold': '#DFC57B',
    'Green': '#7DCC5D',
    'Blue': '#2596be',
    'Orange': '#F0B356'
}
selected_color = 'Blue'  # Default color

def toggle_chart_size():
    global small_chart
    small_chart = not small_chart
    size_text = "Use Normal Size" if small_chart else "Use Small Size"
    resize_button.config(text=size_text)

def update_selected_color(event):
    global selected_color
    selected_color = color_combobox.get()

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            load_data(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def load_data(file_path):
    global df
    df = pd.read_excel(file_path)
    
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%y', errors='coerce')
        df['YY-QQ'] = df['Date'].dt.strftime('%y') + '-Q' + df['Date'].dt.quarter.astype(str)

    populate_column_dropdowns()
    column2_dropdown.config(state="normal")
    column3_dropdown.config(state="normal")
    filter_product_dropdown.config(state="normal")
    chart_type_dropdown.config(state="normal")
    populate_button.config(state="normal")

def populate_column_dropdowns():
    columns = df.columns.tolist()
    column2_var.set('YY-QQ')
    column3_var.set('Qty')
    filter_product_var.set(columns[1])
    
    column2_dropdown['menu'].delete(0, 'end')
    column3_dropdown['menu'].delete(0, 'end')
    filter_product_dropdown['menu'].delete(0, 'end')
    
    for col in columns:
        column2_dropdown['menu'].add_command(label=col, command=lambda value=col: column2_var.set(value))
        column3_dropdown['menu'].add_command(label=col, command=lambda value=col: column3_var.set(value))
        filter_product_dropdown['menu'].add_command(label=col, command=lambda value=col: filter_product_var.set(value))

def populate_filter_dropdowns():
    try:
        selected_column = filter_product_var.get()
        unique_values = df[selected_column].dropna().unique().tolist()

        filter2_var.set(unique_values[0])
        filter2_dropdown.config(state="normal")
        filter2_dropdown['menu'].delete(0, 'end')

        for value in unique_values:
            filter2_dropdown['menu'].add_command(label=value, command=lambda val=value: filter2_var.set(val))

        yyqq_listbox.delete(0, END)
        unique_yyqq = df['YY-QQ'].dropna().unique().tolist()

        for value in unique_yyqq:
            yyqq_listbox.insert(END, value)

        yyqq_listbox.insert(END, 'ALL')
        generate_button.config(state="normal")

    except KeyError as e:
        messagebox.showerror("Error", f"Column selection failed: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def generate_chart():
    try:
        selected_value_col2 = filter2_var.get()
        selected_yyqq = [yyqq_listbox.get(idx) for idx in yyqq_listbox.curselection()]
        
        if 'ALL' in selected_yyqq:
            selected_yyqq = df['YY-QQ'].dropna().unique().tolist()
        
        df_filtered = df[(df['YY-QQ'].isin(selected_yyqq)) & (df[filter_product_var.get()] == selected_value_col2)]
        
        if df_filtered.empty:
            messagebox.showwarning("No Data", "The filtered data is empty.")
            return
        
        df_grouped = df_filtered.groupby('YY-QQ').agg({column3_var.get(): 'sum'}).reset_index()

        # Set the chart size based on the toggle state
        chart_size = (5, 3) if small_chart else (10, 6)

        plt.figure(figsize=chart_size)

        # Get the selected pastel color in hex from the color_map
        plot_color = color_map[selected_color]

        # Check selected chart type and plot accordingly
        if chart_type.get() == "Line":
            plt.plot(df_grouped['YY-QQ'], df_grouped[column3_var.get()], marker='o', color=plot_color, label=f'{column3_var.get()} over YY-QQ')

            # Add value labels without decimals
            for i, value in enumerate(df_grouped[column3_var.get()]):
                plt.text(df_grouped['YY-QQ'][i], value, f'{int(value)}', ha='center', va='bottom')

            # Show grid lines with a very light gray color for line chart
            plt.grid(True, color='#f0f0f0')

        elif chart_type.get() == "Bar":
            plt.bar(df_grouped['YY-QQ'], df_grouped[column3_var.get()], color=plot_color, label=f'{column3_var.get()} over YY-QQ')

            # Add value labels to the bar chart
            for i, value in enumerate(df_grouped[column3_var.get()]):
                plt.text(df_grouped['YY-QQ'][i], value, f'{int(value)}', ha='center', va='bottom')

            # Hide grid lines for bar chart
            plt.grid(False)

        plt.xlabel('YY-QQ')
        plt.ylabel(column3_var.get())
        plt.title(f'{column3_var.get()} for {selected_value_col2}')

        # Customize the grid lines to be lighter and less prominent
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)

        plt.xticks(rotation=45)
        plt.subplots_adjust(bottom=0.25)
        
        plt.legend()
        plt.show()

    except KeyError as e:
        messagebox.showerror("Error", f"Column selection failed: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def show_yyqq_formula():
    formula_window = Toplevel(root)
    formula_window.title("YY-QQ Formula")
    
    ttk.Label(formula_window, text="Use the following formula in Excel to create a YY-QQ column based on a date format of dd/mm/yy:", padding=10).pack()

    formula_text = Text(formula_window, height=2, width=60)
    formula_text.pack(padx=10, pady=10)
    
    formula = '=TEXT(A1, "yy") & "-Q" & ROUNDUP(MONTH(A1)/3, 0)'
    formula_text.insert(END, formula)
    
    formula_text.config(state='disabled')

    def copy_formula_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(formula)
        messagebox.showinfo("Copied", "The formula has been copied to the clipboard.")
    
    ttk.Button(formula_window, text="Copy to Clipboard", command=copy_formula_to_clipboard).pack(padx=10, pady=10)

# Initialize the main window
root = tk.Tk()
root.title("Smart Chart Creator")
root.iconbitmap(icon_path)  # Set the icon for the main window

# Initialize StringVars
column2_var = StringVar()
column3_var = StringVar()
filter_product_var = StringVar()
filter2_var = StringVar()
chart_type = StringVar()

# Apply button styles using ttk
style = ttk.Style()
style.theme_use('clam')

# Configure custom button styles
style.configure('Custom.TButton',
                background='#d0e8f1',   # Normal button color
                foreground='black',     # Text color
                borderwidth=1,
                focusthickness=3,
                focuscolor='none',
                font=('Arial', 10))

# Hover effect with a darker color
style.map('Custom.TButton',
          background=[('active', '#87CEFA')],  # Darker hover color
          foreground=[('active', 'black')])

# Main frame
main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))

# File Upload Button
ttk.Button(main_frame, text="Upload Excel File", command=upload_file, style='Custom.TButton').grid(row=0, column=0, columnspan=2, pady=10)

# Dropdowns and labels for filters and chart selection
ttk.Label(main_frame, text="Select Filter Column:").grid(row=1, column=0, sticky=tk.W, padx=5)
column2_dropdown = ttk.OptionMenu(main_frame, column2_var, '')
column2_dropdown.grid(row=1, column=1)
column2_dropdown.config(state="disabled")

ttk.Label(main_frame, text="Select Value Column:").grid(row=2, column=0, sticky=tk.W, padx=5)
column3_dropdown = ttk.OptionMenu(main_frame, column3_var, '')
column3_dropdown.grid(row=2, column=1)
column3_dropdown.config(state="disabled")

ttk.Label(main_frame, text="Filter by Product:").grid(row=3, column=0, sticky=tk.W, padx=5)
filter_product_dropdown = ttk.OptionMenu(main_frame, filter_product_var, '')
filter_product_dropdown.grid(row=3, column=1)
filter_product_dropdown.config(state="disabled")

ttk.Label(main_frame, text="Filter by Column Value:").grid(row=4, column=0, sticky=tk.W, padx=5)
filter2_dropdown = ttk.OptionMenu(main_frame, filter2_var, '')
filter2_dropdown.grid(row=4, column=1)
filter2_dropdown.config(state="disabled")

# Frame to hold Listbox and Scrollbar for YY-QQ Range
yyqq_frame = ttk.Frame(main_frame)
yyqq_frame.grid(row=5, column=1, padx=5, pady=5)

ttk.Label(main_frame, text="Select YY-QQ Range:").grid(row=5, column=0, sticky=tk.W, padx=5)

# Listbox for YY-QQ Range
yyqq_listbox = Listbox(yyqq_frame, selectmode=MULTIPLE, exportselection=False, height=6)
yyqq_listbox.grid(row=0, column=0)

# Scrollbar for the Listbox
yyqq_scrollbar = ttk.Scrollbar(yyqq_frame, orient="vertical", command=yyqq_listbox.yview)
yyqq_scrollbar.grid(row=0, column=1, sticky=tk.N + tk.S)
yyqq_listbox.config(yscrollcommand=yyqq_scrollbar.set)

# Dropdown for selecting chart type
ttk.Label(main_frame, text="Select Chart Type:").grid(row=6, column=0, sticky=tk.W, padx=5)
chart_type.set("Line")
chart_type_dropdown = ttk.OptionMenu(main_frame, chart_type, "Line", "Bar")
chart_type_dropdown.grid(row=6, column=1)
chart_type_dropdown.config(state="disabled")

# Combobox for selecting the line/bar color (Pastel Colors)
ttk.Label(main_frame, text="Select Line/Bar Color:").grid(row=7, column=0, sticky=tk.W, padx=5)
color_combobox = ttk.Combobox(main_frame, values=pastel_color_options, state="readonly")
color_combobox.set("Blue")  # Set default pastel color
color_combobox.grid(row=7, column=1)
color_combobox.bind("<<ComboboxSelected>>", update_selected_color)

# Buttons for populating, generating, resizing charts
populate_button = ttk.Button(main_frame, text="Populate Filter Options", command=populate_filter_dropdowns, style='Custom.TButton')
populate_button.grid(row=8, column=0, columnspan=2, pady=5)
populate_button.config(state="disabled")

generate_button = ttk.Button(main_frame, text="Generate Chart", command=generate_chart, style='Custom.TButton')
generate_button.grid(row=9, column=0, columnspan=2, pady=5)
generate_button.config(state="disabled")

resize_button = ttk.Button(main_frame, text="Use Small Size", command=toggle_chart_size, style='Custom.TButton')
resize_button.grid(row=10, column=0, columnspan=2, pady=5)

# Button to show YY-QQ formula
yyqq_formula_button = ttk.Button(main_frame, text="YY-QQ Formula", command=show_yyqq_formula, style='Custom.TButton')
yyqq_formula_button.grid(row=11, column=0, columnspan=2, pady=5)

# Start the GUI loop
root.mainloop()
