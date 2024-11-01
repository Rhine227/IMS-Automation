"""
IMS Automation Script
Processes Excel-based Inspection Maintenance System (IMS) templates and converts to JSON format.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set
from pathlib import Path
import json
import logging
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.exceptions import InvalidFileException

# Constants
INPUT_HEADERS = {"Date Inspected by who?", "OK", "OK?", "Date Inspected by who"}
CATEGORY_COLOR = 'FFFFFF00'  # Yellow background
DEFAULT_VALUE = "no input"

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class Task:
    """Represents a task with its properties."""
    name: str
    description: str = ""
    inputs: Dict[str, str] = field(default_factory=dict)

@dataclass
class Category:
    """Represents a category containing tasks."""
    name: str
    tasks: List[Task] = field(default_factory=list)

@dataclass
class SheetData:
    """Represents sheet data with categories."""
    name: str
    categories: List[Category] = field(default_factory=list)

class ExcelProcessor:
    """Handles Excel file processing and data extraction."""
    
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.workbook = None
        self.input_columns: Set[str] = set()
    
    def process_workbook(self) -> List[SheetData]:
        """Process the entire workbook and return structured data."""
        try:
            self.workbook = openpyxl.load_workbook(self.file_path)
            return [self._process_worksheet(sheet) for sheet in self.workbook.worksheets]
        except InvalidFileException as e:
            logger.error(f"Invalid Excel file: {e}")
            raise
        except Exception as e:
            logger.error(f"Error processing workbook: {e}")
            raise

    def _identify_input_columns(self, worksheet: Worksheet) -> Set[str]:
        """Identify columns containing input fields."""
        input_cols = set()
        for col in worksheet.iter_cols(values_only=False):
            for cell in col:
                if cell.value in INPUT_HEADERS:
                    input_cols.add(openpyxl.utils.get_column_letter(cell.column))
        return input_cols

    def _process_worksheet(self, worksheet: Worksheet) -> SheetData:
        """Process a single worksheet and extract structured data."""
        sheet_data = SheetData(name=worksheet.title)
        current_category = None
        current_task = None
        self.input_columns = self._identify_input_columns(worksheet)
        
        for row in worksheet.iter_rows(min_row=2, values_only=False):
            first_cell = row[0]
            if not first_cell.value:
                continue

            if self._is_category(first_cell):
                current_category = Category(name=first_cell.value)
                sheet_data.categories.append(current_category)
                current_task = None
            elif self._is_task(first_cell) and current_category:
                current_task = Task(name=first_cell.value)
                current_category.tasks.append(current_task)
                self._process_input_cells(row, current_task)
            elif current_task and not first_cell.font.bold:
                self._append_description(current_task, first_cell.value)

        return sheet_data

    @staticmethod
    def _is_category(cell) -> bool:
        """Check if cell represents a category."""
        return (cell.fill.start_color and 
                cell.fill.start_color.rgb == CATEGORY_COLOR and 
                cell.font.bold)

    @staticmethod
    def _is_task(cell) -> bool:
        """Check if cell represents a task."""
        return cell.font.bold

    def _process_input_cells(self, row, task: Task) -> None:
        """Process input cells for a task."""
        for cell in row[1:]:
            col_letter = openpyxl.utils.get_column_letter(cell.column)
            if col_letter in self.input_columns:
                task.inputs[cell.coordinate] = cell.value or DEFAULT_VALUE

    @staticmethod
    def _append_description(task: Task, text: str) -> None:
        """Append description text to a task."""
        task.description = f"{task.description} {text}".strip()

def save_to_json(data: List[SheetData], output_path: Path) -> None:
    """Save extracted data to JSON file."""
    try:
        output_path.write_text(
            json.dumps([sheet.__dict__ for sheet in data], indent=4),
            encoding='utf-8'
        )
        logger.info(f"Data saved to {output_path}")
    except Exception as e:
        logger.error(f"Failed to save JSON: {e}")
        raise

def main(template: Optional[str] = None) -> None:
    """Main execution function."""
    root = Tk()
    root.withdraw()

    try:
        if template:
            file_path = Path("IMS_TEMPLATE_COPIES") / template / f"{template}.xlsx"
            if not file_path.exists():
                raise FileNotFoundError(f"Template not found: {file_path}")
        else:
            file_path = Path(askopenfilename(
                title="Select Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            ))
            if not file_path.name:
                logger.info("No file selected")
                return

        processor = ExcelProcessor(file_path)
        data = processor.process_workbook()
        save_to_json(data, file_path.with_suffix('.json'))

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    main()
