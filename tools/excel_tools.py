import win32com.client
import pythoncom  # Add this import
import time
from typing import List, Dict
import json
import os

# Class to represent a cell edit, if needed elsewhere.
class CellEdit:
    def __init__(self, cell_address: str, value: str):
        self.cell_address = cell_address
        self.value = value

class ExcelTools:
    def __init__(self):
        self.excel_app = None
        self.open_workbooks = {}

    def _initialize_excel_app(self):
        """Initialize Excel application if not already initialized"""
        try:
            if self.excel_app is None:
                print("Initializing new Excel application...")
                # Initialize COM in this thread
                pythoncom.CoInitialize()
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                # Set visibility after dispatch
                try:
                    self.excel_app.Visible = True
                    self.excel_app.DisplayAlerts = False
                except Exception as e:
                    print(f"Warning: Could not set Excel visibility: {str(e)}")
        except Exception as e:
            print(f"Error initializing Excel: {str(e)}")
            raise

    def _activate_workbook(self, file_path: str):
        """
        Helper method to activate an existing workbook by file path.
        If the workbook is already open, it is activated; otherwise, it is opened.
        Returns: (workbook, is_new_workbook)
        """
        self._initialize_excel_app()
        normalized_path = os.path.abspath(file_path).lower()

        # Check if the workbook is already in our tracking dict
        if normalized_path in self.open_workbooks:
            try:
                self.open_workbooks[normalized_path].Activate()
                return self.open_workbooks[normalized_path], False
            except Exception as e:
                print(f"Warning: Could not activate tracked workbook: {str(e)}")
                del self.open_workbooks[normalized_path]

        # Check if the workbook is open in Excel
        try:
            for wb in self.excel_app.Workbooks:
                try:
                    if wb.FullName.lower() == normalized_path:
                        wb.Activate()
                        self.open_workbooks[normalized_path] = wb
                        return wb, False
                except:
                    continue
        except:
            pass

        # Open or create the workbook
        try:
            if os.path.exists(file_path):
                wb = self.excel_app.Workbooks.Open(file_path)
            else:
                wb = self.excel_app.Workbooks.Add()
                wb.SaveAs(file_path)
            
            # For new workbooks, adjust the view
            active_window = wb.Windows(1)
            active_window.Zoom = 200  # Set zoom to 100%
            
            self.open_workbooks[normalized_path] = wb
            return wb, True
        except Exception as e:
            print(f"Error opening workbook: {str(e)}")
            raise

    def _get_full_path(self, file_path: str) -> str:
        """Helper method to get the full path, with better error handling and logging"""
        try:
            if os.path.isabs(file_path):
                print(f"Using absolute path: {file_path}")
                return file_path
                
            # Use the datalake directory relative to the project root
            base_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'datalake')
            
            # Create the datalake directory if it doesn't exist
            if not os.path.exists(base_dir):
                os.makedirs(base_dir)
                print(f"Created datalake directory at: {base_dir}")
            
            full_path = os.path.join(base_dir, file_path)
            print(f"Converted to full path: {full_path}")
            return full_path
            
        except Exception as e:
            print(f"Error in path handling: {str(e)}")
            raise

    def read_excel(self, file_path: str) -> str:
        """
        Unified method to read a fixed range (A1:Z20) from the first worksheet of an Excel file.
        This method activates the target workbook and worksheet based on the provided file path.
        
        Args:
            file_path (str): The full path of the Excel file.
        
        Returns:
            str: A formatted string with each non-empty cell address and its value.
        """
        try:
            self._activate_workbook(file_path)
            state = []
            
            # Define the range to read ("A1:Z20").
            range_obj = self.worksheet.Range("A1:Z20")
            for row in range(1, 21):      # Rows 1 to 20.
                for col in range(1, 27):  # Columns 1 to 26 (A to Z).
                    cell = range_obj.Cells(row, col)
                    if cell.Text:
                        col_letter = chr(64 + col)  # Convert column number to letter.
                        cell_address = f"{col_letter}{row}"
                        state.append(f"{cell_address}: {cell.Text}")
            result = "\n".join(state)
            return result
        except Exception as e:
            error_msg = f"Error reading Excel file: {str(e)}"
            return error_msg

    def write_to_excel(self, file_path: str, data: dict) -> str:
        """
        Unified method to write data to cells in the first worksheet of an Excel file.
        After writing, the workbook is saved.
        This method activates the target workbook and worksheet based on the provided file path before performing any write operations.
        
        Args:
            file_path (str): The full path of the Excel file.
            data (dict): A dictionary mapping cell addresses (e.g., "A1") to values to write.
                         If data is nested under a sheet name (e.g., {"Sheet1": { "A1": "Valve Name", ... }})
                         or {"Sheet1": [ { ... } ]}), the method extracts the inner dictionary or processes the list payload.
                         Additionally, if the nested payload (e.g., under "cells") is a list of dictionaries each having
                         'cell' and 'value' keys, the method converts the list into a flat dictionary.
        
        Returns:
            str: A status message summarizing the write operations.
        """
        try:
            self._activate_workbook(file_path)
            # Try to make visible after activation
            try:
                self.excel_app.Visible = True
            except:
                pass
            print(f"Debug: Workbook and worksheet activated for writing using file {file_path}.")

            # Check for nested data payload under a single key.
            if isinstance(data, dict) and len(data) == 1:
                first_key = next(iter(data))
                inner_data = data[first_key]
                # If the inner data is already a dictionary, extract it.
                if isinstance(inner_data, dict):
                    print(f"Debug: Detected nested data payload under sheet name '{first_key}' (dictionary format). Extracting payload.")
                    data = inner_data
                # If the inner data is a list, determine its structure.
                elif isinstance(inner_data, list) and len(inner_data) > 0:
                    # Check if every item in the list is a dict with 'cell' and 'value' keys.
                    if all(isinstance(item, dict) and 'cell' in item and 'value' in item for item in inner_data):
                        # Convert the list into a flat dictionary with cell addresses as keys.
                        print(f"Debug: Detected list of cell-value dictionaries under key '{first_key}'. Converting list to flat dictionary.")
                        data = {item['cell']: item['value'] for item in inner_data}
                    else:
                        # Fallback: Extract the first element if it doesn't match the expected structure.
                        print(f"Debug: Detected nested data payload under sheet name '{first_key}' (list format) but not in cell-value format. Extracting first element as payload.")
                        data = inner_data[0]
                else:
                    print("Debug: Data payload nested under sheet name, but format is not recognized. Proceeding with original data.")

            results = []
            for cell_address, value in data.items():
                cell = self.worksheet.Range(cell_address)
                cell.Value2 = value
                print(f"Debug: Written {value} to {cell_address} in workbook {file_path}.")
                results.append(f"Wrote {value} to {cell_address}")
                time.sleep(0.1)  # Small delay to allow Excel to refresh.
            self.workbook.Save()
            return "\n".join(results)
        except Exception as e:
            return f"Error writing to Excel file: {str(e)}"

    def create_new_workbook(self, new_file_path: str) -> str:
        """
        Create a new Excel workbook, activate its first worksheet, and save it to the specified path.
        """
        try:
            self._initialize_excel_app()
            # Create new workbook without affecting others
            self.workbook = self.excel_app.Workbooks.Add()
            try:
                self.excel_app.Visible = True
            except:
                pass
            
            # Save immediately to establish the file
            self.workbook.SaveAs(new_file_path)
            self.worksheet = self.workbook.Worksheets(1)
            self.worksheet.Select()
            print(f"Debug: New workbook created and saved at {new_file_path}.")
            return f"New workbook created and saved at {new_file_path}"
        except Exception as e:
            return f"Error creating new workbook: {str(e)}"

    def close_workbook(self, file_path: str) -> str:
        """
        Close the workbook specified by file_path if it is open, saving changes before closing.
        
        Args:
            file_path (str): The full path of the Excel file to close.
        
        Returns:
            str: A status message indicating whether the workbook was closed.
        """
        try:
            if self.excel_app is None:
                return "Excel application is not initialized."
            normalized_path = file_path.lower()
            workbook_found = None
            for wb in self.excel_app.Workbooks:
                try:
                    if wb.FullName.lower() == normalized_path:
                        workbook_found = wb
                        break
                except Exception:
                    continue

            if workbook_found:
                workbook_found.Save()
                workbook_found.Close()
                print(f"Debug: Closed workbook at {file_path}.")
                if self.workbook and self.workbook.FullName.lower() == normalized_path:
                    self.workbook = None
                    self.worksheet = None
                return f"Workbook at {file_path} closed successfully."
            else:
                return f"No open workbook found at {file_path}."
        except Exception as e:
            return f"Error closing workbook: {str(e)}"

    def cleanup(self):
        """
        Clean up Excel resources by saving and closing all tracked workbooks
        and quitting the Excel application.
        """
        try:
            # Save and close all tracked workbooks
            for path, wb in self.open_workbooks.items():
                try:
                    wb.Save()
                    wb.Close()
                except Exception as e:
                    print(f"Error closing workbook {path}: {str(e)}")
            self.open_workbooks.clear()

            if self.excel_app:
                # Close any remaining open workbooks
                while self.excel_app.Workbooks.Count:
                    self.excel_app.Workbooks(1).Close()
                self.excel_app.Quit()
                self.excel_app = None
                print("Debug: Excel application quit.")
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Error during cleanup: {str(e)}")

    def process_excel(self, file_path, write_data=None):
        """Process Excel operations with better workbook state handling"""
        try:
            self._initialize_excel_app()
            
            # Convert to full path
            abs_path = self._get_full_path(file_path)
            if not abs_path.endswith('.xlsx'):
                abs_path += '.xlsx'
            
            print(f"Processing Excel file: {abs_path}")

            # Ensure we have a valid Excel instance
            try:
                _ = self.excel_app.Visible
            except:
                print("Reconnecting to Excel...")
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = True
                self.excel_app.DisplayAlerts = True

            # Make Excel visible
            self.excel_app.Visible = True

            # Get or create workbook
            workbook, is_new_workbook = self._activate_workbook(abs_path)
            worksheet = workbook.Worksheets(1)
            worksheet.Activate()

            # Only maximize if it's a new workbook
            if is_new_workbook:
                try:
                    window = workbook.Windows(1)
                    window.WindowState = -4137  # xlMaximized
                except Exception as e:
                    print(f"Warning: Could not maximize window: {str(e)}")

            if write_data:
                print(f"Writing data to workbook...")
                for cell, value in write_data.items():
                    try:
                        print(f"Writing {value} to {cell}")
                        cell_range = worksheet.Range(cell)
                        
                        # Store original color and format
                        original_color = cell_range.Interior.Color
                        original_pattern = cell_range.Interior.Pattern
                        
                        # Set value and highlight cell
                        cell_range.Value = value
                        cell_range.Select()
                        cell_range.Interior.Color = 0xFF9019
                        
                        # Force immediate update
                        self.excel_app.ScreenUpdating = True
                        
                        time.sleep(0.1)
                        
                        # Fade back to original color
                        cell_range.Interior.Color = original_color
                        cell_range.Interior.Pattern = original_pattern
                        
                        time.sleep(0.1)
                    except Exception as e:
                        print(f"Error writing to cell {cell}: {str(e)}")
                        continue

                try:
                    workbook.Save()
                    print("Workbook saved successfully")
                except Exception as e:
                    print(f"Error saving workbook: {str(e)}")
                    time.sleep(1)
                    workbook.Save()

                return f"Successfully wrote data to {abs_path}"
            else:
                print("Reading data from workbook...")
                data = {}
                used_range = worksheet.UsedRange
                for row in range(1, used_range.Rows.Count + 1):
                    for col in range(1, used_range.Columns.Count + 1):
                        cell = worksheet.Cells(row, col)
                        if cell.Value is not None:
                            col_letter = chr(64 + col) if col <= 26 else chr(64 + col//26) + chr(64 + col%26)
                            cell_ref = f"{col_letter}{row}"
                            data[cell_ref] = cell.Value
                return str(data)

        except Exception as e:
            print(f"Excel operation error: {str(e)}")
            try:
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = True
            except:
                pass
            return f"Error processing Excel file: {str(e)}"