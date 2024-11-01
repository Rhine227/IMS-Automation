"""
IMS Automation Script

This script processes Excel-based Inspection Maintenance System (IMS) templates.
It extracts structured data from Excel worksheets and converts it to JSON format.
The script identifies categories, tasks, and input fields based on specific formatting rules.
"""

import json
import os
import logging
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

# Configure logging with timestamp, level, and message format
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_excel_data(file_path: str) -> list:
    """
    Extract and structure data from an Excel file into a hierarchical format.

    Args:
        file_path (str): Path to the Excel file to process.

    Returns:
        list: List of dictionaries containing structured worksheet data.
              Each dictionary contains sheet name, categories, tasks, and their inputs.
              Returns empty list on error.

    Structure:
    {
        'sheet': sheet_name,
        'categories': [
            {
                'category': category_name,
                'tasks': [
                    {
                        'task': task_name,
                        'description': task_description,
                        'inputs': {cell_coordinate: input_value}
                    }
                ]
            }
        ]
    }
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
    except InvalidFileException as e:
        logging.error(f"Invalid file: {file_path}. Error: {e}")
        return []
    except Exception as e:
        logging.error(f"Error loading workbook: {e}")
        return []

    data = []
    all_input_cells = []

    # Process each worksheet in the workbook
    for sheet in workbook.worksheets:
        sheet_data = {"sheet": sheet.title, "categories": []}
        current_category = None
        current_task = None
        input_columns = set()
        input_rows = set()
        comments_section = False

        # Identify input columns by searching for specific headers
        for col in sheet.iter_cols(values_only=False):
            for cell in col:
                if cell.value is not None:
                    logging.debug(f"Checking cell {cell.coordinate} with value: {cell.value}")
                # Look for columns that contain input field headers
                if cell.value in ["Date Inspected by who?", "OK", "OK?", "Date Inspected by who"]:
                    input_columns.add(openpyxl.utils.get_column_letter(cell.column))
        
        logging.info(f"Identified input columns in {sheet.title}: {sorted(input_columns)}")

        # Process rows to extract categories, tasks, and inputs
        for row in sheet.iter_rows(min_row=2, values_only=False):
            cell_value = row[0].value
            if cell_value is None:
                continue

            cell_style = row[0].font
            cell_fill = row[0].fill

            # Category identification (yellow background + bold text)
            if cell_fill.start_color and cell_fill.start_color.rgb == 'FFFFFF00' and cell_style.bold:
                current_category = cell_value
                sheet_data["categories"].append({"category": current_category, "tasks": []})
                current_task = None
                comments_section = False
            
            # Task identification (bold text within a category)
            elif cell_style.bold and current_category and not comments_section:
                current_task = cell_value
                task_row = row[0].row
                input_rows.add(task_row)
                logging.info(f"Task '{current_task}' identified in row {task_row}")
                sheet_data["categories"][-1]["tasks"].append({
                    "task": current_task,
                    "description": "",
                    "inputs": {}
                })
            
            # Description processing (non-bold text under a task)
            elif current_task and not cell_style.bold and not comments_section:
                if sheet_data["categories"][-1]["tasks"][-1]["description"]:
                    sheet_data["categories"][-1]["tasks"][-1]["description"] += " " + cell_value
                else:
                    sheet_data["categories"][-1]["tasks"][-1]["description"] = cell_value

            # Process input cells for the current task
            if current_category and current_task:
                for cell in row[1:sheet.max_column + 1]:
                    cell_coord = cell.coordinate
                    column_letter = openpyxl.utils.get_column_letter(cell.column)
                    if column_letter in input_columns and cell.row == task_row:
                        cell_value = cell.value if cell.value is not None else "no input"
                        sheet_data["categories"][-1]["tasks"][-1]["inputs"][cell_coord] = cell_value
                        logging.info(f"Processed input cell {cell_coord} with value: {cell_value}")
                        all_input_cells.append(cell_coord)

            # Comments section marker
            if cell_value == "Comments":
                comments_section = True

        data.append(sheet_data)
    
    logging.info(f"All input cells: {all_input_cells}")
    return data

def save_to_json(data: list, output_file: str) -> None:
    """
    Save extracted data to a JSON file with proper formatting.

    Args:
        data (list): Structured data to be saved
        output_file (str): Path to the output JSON file
    """
    try:
        with open(output_file, 'w') as json_file:
            json.dump(data, json_file, indent=4)
    except Exception as e:
        logging.error(f"Error saving to JSON file: {e}")

def main(template: str = None) -> None:
    """
    Main execution function that orchestrates the Excel data extraction process.

    Args:
        template (str, optional): Template name to process. If None, opens file dialog.
    """
    # Initialize hidden Tkinter root window for file dialog
    root = Tk()
    root.withdraw()
    
    # Handle template-based or manual file selection
    if template:
        file_path = os.path.join("IMS_TEMPLATE_COPIES", template, f"{template}.xlsx")
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"Template file not found: {file_path}")
            return
    else:
        file_path = askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            messagebox.showinfo("No file selected", "No file selected.")
            return
    
    # Process Excel file and save results
    data = get_excel_data(file_path)
    if not data:
        messagebox.showerror("Error", "Failed to extract data from the Excel file.")
        return
    
    output_file = file_path.rsplit('.', 1)[0] + ".json"
    save_to_json(data, output_file)
    print(f"Data extracted and saved to {output_file}")

if __name__ == "__main__":
    main()
