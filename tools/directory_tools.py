import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

class DirectoryTools:
    def __init__(self):
        self.base_directory = os.getenv("BASE_DIRECTORY")
        if not self.base_directory:
            print("Warning: BASE_DIRECTORY not set in environment variables")

    def list_directory(self) -> str:
        """List all files and directories in the base directory"""
        try:
            if not self.base_directory:
                return "Error: BASE_DIRECTORY environment variable is not set"

            if not os.path.exists(self.base_directory):
                return f"Error: Directory {self.base_directory} does not exist"

            items = os.listdir(self.base_directory)
            if not items:
                return "Directory is empty"

            result = [f"Contents of directory ({self.base_directory}):"]
            for item in items:
                full_path = os.path.join(self.base_directory, item)
                item_type = "Directory" if os.path.isdir(full_path) else "File"
                result.append(f"{item_type}: {item}")

            return "\n".join(result)
        except Exception as e:
            return f"Error listing directory: {str(e)}" 