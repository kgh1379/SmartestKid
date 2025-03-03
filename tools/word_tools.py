import win32com.client
import os
import time
import pythoncom

class WordTools:
    def __init__(self):
        self.word_app = None
        self.document = None
        self.base_directory = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'datalake')
        
        # Create the datalake directory if it doesn't exist
        if not os.path.exists(self.base_directory):
            os.makedirs(self.base_directory)
            print(f"Created datalake directory at: {self.base_directory}")

    def _initialize_word_app(self):
        """Initialize Microsoft Word application if not already running"""
        if self.word_app is None:
            # Initialize COM in this thread
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True
            print("Debug: Word application initialized.")

    def process_word(self, file_path: str, content: str = None) -> str:
        """
        Process a Word document - read current content and optionally write new content
        """
        try:
            # Convert relative path to full path if needed
            if not os.path.isabs(file_path):
                file_path = os.path.join(self.base_directory, file_path)
                print(f"Debug: Converting to full path: {file_path}")

            print("Debug: About to initialize Word...")
            self._initialize_word_app()
            print("Debug: Word initialized successfully")

            # Create new document if it doesn't exist
            if not os.path.exists(file_path):
                print("Debug: Creating new document...")
                self.document = self.word_app.Documents.Add()
                print("Debug: Document added, about to save...")
                try:
                    self.document.SaveAs(file_path)
                    print(f"Debug: Document saved successfully at {file_path}")
                except Exception as save_error:
                    print(f"Debug: Error saving document: {str(save_error)}")
                    raise
            else:
                print(f"Debug: Opening existing document at {file_path}")
                self.document = self.word_app.Documents.Open(file_path)

            # Read initial content
            initial_content = self.document.Content.Text
            
            # Write new content if provided
            if content:
                print("Debug: Writing new content...")
                self.document.Content.Text = content
                self.document.Save()
                time.sleep(0.1)  # Give Word time to process

            # Read final content
            final_content = self.document.Content.Text

            # Format result
            result_sections = [
                "Initial Content:",
                initial_content.strip() if initial_content.strip() else "(Empty document)",
                "\nWrite Operations:" if content else "",
                f"Wrote new content: {content}" if content else "",
                "\nFinal Content:",
                final_content.strip() if final_content.strip() else "(Empty document)"
            ]

            return "\n".join(filter(None, result_sections))

        except Exception as e:
            error_details = f"Error processing Word document: {str(e)}\nType: {type(e)}"
            print(f"Debug: {error_details}")
            return error_details

    def cleanup(self):
        """Clean up Word resources"""
        try:
            if self.document:
                self.document.Close()
            if self.word_app:
                self.word_app.Quit()
        except:
            pass 