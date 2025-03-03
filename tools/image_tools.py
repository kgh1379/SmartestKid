import base64
import requests
import os
from typing import Optional
import PyPDF2
from dotenv import load_dotenv
import tkinter as tk
from PIL import Image, ImageTk
import time
import psutil
import subprocess
import win32api
load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

class ImageTools:
    def __init__(self):
        self.api_key = os.getenv("OPENAI_API_KEY")
        self.api_url = "https://api.openai.com/v1/chat/completions"
        self.base_directory = os.getenv("BASE_DIRECTORY")
        self.image_path = None
        self.pdf_path = None
        self.viewer = None
        self.pdf_process = None

    def show_file(self, file_path):
        """Display the file in a window"""
        if file_path.lower().endswith('.pdf'):
            print(f"Opening PDF: {file_path}")
            # Move mouse to main monitor center before opening PDF
            self.focus_main_monitor()
            # For PDFs, use system default viewer and track the process
            self.pdf_process = subprocess.Popen(['start', '', file_path], shell=True)
            # Wait 2 seconds then close
            time.sleep(2)
            print("Attempting to close PDF viewer...")
            self.close_pdf_viewer()
        else:
            # For images, use Tkinter
            self.viewer = tk.Toplevel()
            self.viewer.title("Image Analysis")
            self.viewer.attributes('-topmost', True)
            
            # Position window on main monitor
            main_screen_width = self.viewer.winfo_screenwidth()
            main_screen_height = self.viewer.winfo_screenheight()
            window_width = 800
            window_height = 600
            x = (main_screen_width - window_width) // 2
            y = (main_screen_height - window_height) // 2
            self.viewer.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            # Load and display image
            img = Image.open(file_path)
            # Resize if too large while maintaining aspect ratio
            max_size = (800, 600)
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            label = tk.Label(self.viewer, image=photo)
            label.image = photo  # Keep a reference
            label.pack()

    def focus_main_monitor(self):
        """Move mouse to center of main monitor to help focus windows there"""
        try:
            # Get main monitor resolution
            width = win32api.GetSystemMetrics(0)
            height = win32api.GetSystemMetrics(1)
            # Move mouse to center of main monitor
            win32api.SetCursorPos((width // 2, height // 2))
        except Exception as e:
            print(f"Could not move cursor to main monitor: {e}")

    def close_pdf_viewer(self):
        """Close any PDF viewer processes"""
        print("Searching for PDF viewer processes...")
        
        try:
            # Windows-specific: Use taskkill to close PDF viewer windows
            # This targets window titles containing "PDF" or specific viewer names
            subprocess.run([
                'taskkill', '/F', '/FI', 
                'WINDOWTITLE eq *PDF*'
            ], capture_output=True)
            
            # Also try to close specific PDF applications
            subprocess.run(['taskkill', '/F', '/IM', 'AcroRd32.exe'], capture_output=True)
            subprocess.run(['taskkill', '/F', '/IM', 'Acrobat.exe'], capture_output=True)
            subprocess.run(['taskkill', '/F', '/IM', 'SumatraPDF.exe'], capture_output=True)
            
            # For Edge PDF viewer specifically
            subprocess.run([
                'powershell', 
                '-command', 
                "Get-Process | Where-Object {$_.MainWindowTitle -like '*PDF*'} | Stop-Process -Force"
            ], capture_output=True)
            
        except Exception as e:
            print(f"Error during PDF viewer cleanup: {e}")

    def close_viewer(self):
        """Close the viewer window if it exists"""
        if self.viewer:
            print("Closing image viewer...")
            self.viewer.destroy()
            self.viewer = None
        print("Ensuring PDF viewer is closed...")
        self.close_pdf_viewer()

    def analyze_file(self, question: str, file_name: Optional[str] = None) -> str:
        """Analyze either an image or PDF file and answer questions about it"""
        try:
            # Always use the specified file if provided
            if file_name:
                file_path = os.path.join(self.base_directory, file_name)
                print(f"Debug: Using specified file: {file_name}")
            else:
                # Search for first image or PDF if no specific file was requested
                for file in os.listdir(self.base_directory):
                    if file.lower().endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                        file_path = os.path.join(self.base_directory, file)
                        print(f"Debug: Found first file: {file}")
                        break
                else:
                    return "Error: No image or PDF file found in the datalake directory"

            if not os.path.exists(file_path):
                return f"Error: File not found at {file_path}"

            print(f"Debug: Attempting to analyze file at: {file_path}")
            
            # Show the file before analysis
            self.show_file(file_path)

            # Determine file type and process accordingly
            if file_path.lower().endswith('.pdf'):
                # Process PDF
                pdf_text = ""
                with open(file_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page in pdf_reader.pages:
                        pdf_text += page.extract_text() + "\n"

                payload = {
                    "model": "gpt-4o",
                    "messages": [
                        {
                            "role": "user",
                            "content": f"PDF Content:\n{pdf_text}\n\nQuestion: {question}"
                        }
                    ],
                    "max_tokens": 500
                }
            else:
                # Process Image
                with open(file_path, "rb") as image_file:
                    base64_image = base64.b64encode(image_file.read()).decode('utf-8')

                payload = {
                    "model": "gpt-4o",
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": question
                                },
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:image/png;base64,{base64_image}",
                                        # "detail": "high"  # Can be "low", "high", or "auto"
                                    }
                                }
                            ]
                        }
                    ],
                    "max_tokens": 3000
                }

            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }

            response = requests.post(self.api_url, headers=headers, json=payload)
            
            if response.status_code != 200:
                print(f"\nDebug: API Error: {response.text}", flush=True)
                return f"Error: API call failed with status code: {response.status_code}"

            result = response.json()
            
            # Close the viewer after analysis
            self.close_viewer()
            
            print("RETURNED FROM IMAGE TOOL", result['choices'][0]['message']['content'])

            return result['choices'][0]['message']['content'].strip()

        except Exception as e:
            print(f"\nDebug: Exception details: {str(e)}", flush=True)
            self.close_viewer()  # Ensure viewer is closed even if there's an error
            return f"Error analyzing file: {str(e)}" 