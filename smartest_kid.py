import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageSequence
import os
import wave
import threading
import time
import pyaudio
import audioop  # For RMS calculations
import queue
from io import BytesIO
from chat_interface import ChatInterface

# Import the new ChatCompletion-based assistant.
from ai_assistant import ChatCompletionAssistant

# ------------------ Recording with VAD ------------------

def transcribe_file_with_whisper(audio_filename, app):
    """
    Uses OpenAI's Whisper API to transcribe a local audio file.
    Returns only the final transcript string.
    """
    try:
        transcribe_start = time.time()  # Start timing transcription
        from openai import OpenAI
        
        print("Calling Whisper API...")
        app.is_transcribing = True
        
        client = OpenAI(api_key=app.assistant_api_key)
        
        with open(audio_filename, "rb") as audio_file:
            try:
                response = client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )
                print("[WHISPER] Successfully transcribed audio")
            except Exception as whisper_error:
                print(f"[WHISPER] API Error: {str(whisper_error)}")
                raise
        
        transcript_str = response.strip()
        transcribe_duration = time.time() - transcribe_start
        print(f"[TIMING] Transcription completed in {transcribe_duration:.2f}s")
        
        with open("transcription.txt", "w") as f:
            f.write(transcript_str)
        
        return transcript_str
        
    except Exception as e:
        error_msg = f"Error during Whisper transcription: {str(e)}"
        app.log(error_msg)
        print(error_msg)
        return None
    finally:
        app.is_transcribing = False


def record_and_transcribe_vad(app):
    """
    Records audio until silence is detected using VAD.
    Saves the audio to a WAV file and transcribes it via Whisper.
    The transcript is then added to the AI processing queue.
    """
    audio_filename = "temp_recording.wav"
    transcript_filename = "transcription.txt"
    
    try:
        chunk = 1024
        sample_format = pyaudio.paInt16
        channels = 1
        rate = 16000
        silence_threshold = 200
        silence_chunks = 0

        pause_duration_sec = 2.5            
        max_silence_chunks = int(pause_duration_sec / (chunk / rate))
        min_chunks = int(0.5 / (chunk / rate))
        frames = []

        p = pyaudio.PyAudio()
        stream = p.open(format=sample_format, channels=channels, rate=rate,
                        input=True, frames_per_buffer=chunk)
        app.log("Recording started (VAD enabled). Speak now...")
        app.is_listening = False
        voiced = False

        while True:
            if app.is_paused:
                # Always process what we have when muting
                break

            data = stream.read(chunk, exception_on_overflow=False)
            frames.append(data)
            rms = audioop.rms(data, 2)
            if rms > silence_threshold:
                silence_chunks = 0
                voiced = True
                app.is_listening = True
            else:
                if voiced:
                    silence_chunks += 1
            if voiced and silence_chunks > max_silence_chunks and len(frames) > min_chunks:
                app.log("Silence detected. Finishing recording.")
                break

        # Close the stream after breaking from the loop
        stream.stop_stream()
        stream.close()
        p.terminate()

        # Always process the recording if we have enough voiced frames
        if voiced and len(frames) > min_chunks:
            wf = wave.open(audio_filename, 'wb')
            wf.setnchannels(channels)
            wf.setsampwidth(p.get_sample_size(sample_format))
            wf.setframerate(rate)
            wf.writeframes(b''.join(frames))
            wf.close()
            app.log("Recording finished. Audio saved to " + audio_filename)

            # Remove the pause check here - always process the final chunk
            transcript = transcribe_file_with_whisper(audio_filename, app)
            if transcript:
                app.log("Transcription: " + transcript)
                print("Transcription:", transcript)
                app.message_queue.put(transcript)
            else:
                app.log("Whisper transcription failed.")
    
    finally:
        try:
            if os.path.exists(audio_filename):
                os.remove(audio_filename)
                print(f"Cleaned up {audio_filename}")
            if os.path.exists(transcript_filename):
                os.remove(transcript_filename)
                print(f"Cleaned up {transcript_filename}")
        except Exception as e:
            print(f"Error cleaning up temporary files: {e}")

def continuous_record_and_transcribe(app):
    while True:
        record_and_transcribe_vad(app)
        # Optionally, add a delay here if desired.

# ------------------ Animated Tkinter Application ------------------

class AnimatedCharacter:
    def __init__(self, canvas, x, y):
        self.canvas = canvas
        self.x = x
        self.y = y
        self.is_animated = False
        
        static_img = Image.open("assets/static_assistant.png")
        if static_img.mode != 'RGBA':
            static_img = static_img.convert('RGBA')
        data = static_img.getdata()
        new_data = []
        for item in data:
            if item[0] > 240 and item[1] > 240 and item[2] > 240:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        static_img.putdata(new_data)
        self.static_image = ImageTk.PhotoImage(static_img)
        
        gif = Image.open("assets/animated_assistant.gif")
        self.animated_frames = []
        for frame in ImageSequence.Iterator(gif):
            frame = frame.convert('RGBA')
            data = frame.getdata()
            new_data = []
            for item in data:
                if item[0] < 15 and item[1] < 15 and item[2] < 15:
                    new_data.append((0, 0, 0, 0))
                else:
                    new_data.append(item)
            frame.putdata(new_data)
            self.animated_frames.append(ImageTk.PhotoImage(frame))
        
        self.image_id = canvas.create_image(x, y, image=self.static_image, anchor='center', tags='character')
        self.current_frame = 0
        
    def set_animated(self, animated):
        self.is_animated = animated
        if not animated:
            self.canvas.itemconfig(self.image_id, image=self.static_image)
            
    def update(self):
        if self.is_animated and self.animated_frames:
            self.current_frame = (self.current_frame + 1) % len(self.animated_frames)
            self.canvas.itemconfig(self.image_id, image=self.animated_frames[self.current_frame])

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Animated AI Avatar - Transparent Robot Listening")
        self.geometry("1800x1600")
        self.resizable(False, False)
        
        self.overrideredirect(True)
        self.attributes('-alpha', 0.0)
        self.wm_attributes("-transparentcolor", "SystemButtonFace")
        self.wm_attributes("-topmost", True)
        
        self.canvas = tk.Canvas(
            self, 
            bg="SystemButtonFace",
            width=1800, 
            height=1600, 
            highlightthickness=0
        )
        self.canvas.pack()
        self.after(100, lambda: self.attributes('-alpha', 1.0))
        
        self._drag_data = {"x": 0, "y": 0, "item": None, "start_time": 0, "start_x": 0, "start_y": 0}
        self.bind("<Escape>", lambda e: self.cleanup())
        
        self.is_listening = False
        self.is_transcribing = False
        self.is_ai_processing = False
        self.red_dot_id = None
        self.dot_visible = False
        self.is_paused = True

        self.chat_interface = None
        self.chat_window = None

        self.canvas.tag_bind("character", "<Button-1>", self.on_drag_start)
        self.canvas.tag_bind("character", "<ButtonRelease-1>", self.on_drag_stop)
        self.canvas.tag_bind("character", "<B1-Motion>", self.on_drag_motion)
        
        self.character = AnimatedCharacter(self.canvas, 400, 300)
        
        self.log_widget = tk.Text(self, height=5, width=100)
        self.log_widget.pack(padx=10, pady=10)

        # Load API key and initialize ChatCompletionAssistant.
        from os import getenv
        from dotenv import load_dotenv
        load_dotenv()
        self.assistant_api_key = getenv('OPENAI_API_KEY')
        if not self.assistant_api_key:
            print("Please set your OPENAI_API_KEY environment variable")
            exit(1)
        self.assistant = ChatCompletionAssistant(self.assistant_api_key)
        
        self.message_queue = queue.Queue()
        self._ai_lock = threading.Lock()
        
        threading.Thread(target=continuous_record_and_transcribe, args=(self,), daemon=True).start()
        self.animate_robot()
        threading.Thread(target=self.process_ai_messages, daemon=True).start()
        self.create_mode_toggles()
        self.after(100, self.toggle_messages)

    def log(self, message):
        self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
    
    def animate_robot(self):
        # Show red dot only when listening (not when processing)
        if self.is_listening and not self.is_ai_processing:
            self.toggle_red_dot()
        elif self.red_dot_id is not None:
            self.canvas.delete(self.red_dot_id)
            self.red_dot_id = None
            self.dot_visible = False
        
        # Animate robot when AI is processing (both voice and text)
        if self.is_ai_processing:
            if not self.character.is_animated:
                self.character.set_animated(True)
            self.character.update()
        else:
            if self.character.is_animated:
                self.character.set_animated(False)
        
        self.after(150, self.animate_robot)

    def on_drag_start(self, event):
        self._drag_data["x"] = event.x_root
        self._drag_data["y"] = event.y_root
        self._drag_data["start_time"] = time.time()
        self._drag_data["start_x"] = event.x_root
        self._drag_data["start_y"] = event.y_root

    def on_drag_stop(self, event):
        self._drag_data.update({"x": 0, "y": 0, "item": None, "start_time": 0, "start_x": 0, "start_y": 0})

    def on_drag_motion(self, event):
        delta_x = event.x_root - self._drag_data["x"]
        delta_y = event.y_root - self._drag_data["y"]
        x = self.winfo_x() + delta_x
        y = self.winfo_y() + delta_y
        self.geometry(f"+{x}+{y}")
        self._drag_data["x"] = event.x_root
        self._drag_data["y"] = event.y_root

    def toggle_mute(self):
        self.is_paused = not self.is_paused
        try:
            self.mic_button.configure(image=self.mic_icon_muted if self.is_paused else self.mic_icon_active)
        except Exception as e:
            print(f"Error in toggle_mute: {e}")

    def process_ai_messages(self):
        while True:
            message = self.message_queue.get()
            if message:
                try:
                    print(f"Processing with AI: {message}")
                    self.log("Processing with AI: " + message)
                    
                    # Add user message to chat interface
                    if self.chat_interface:
                        self.chat_interface.add_user_message(message)
                    
                    # Set AI processing state and animate
                    self.is_ai_processing = True
                    self.is_transcribing = False  # Ensure transcribing is off
                    self.is_listening = False     # Ensure listening is off
                    
                    def single_callback(text_fragment):
                        if self.chat_interface:
                            self.chat_interface.add_assistant_message(text_fragment)
                            # Keep AI processing true while streaming response
                            self.is_ai_processing = True
                            self.character.set_animated(True)
                    
                    # Use the new ChatCompletion-based assistant
                    with self._ai_lock:
                        self.assistant.send_message(message, ui_callback=single_callback)
                except Exception as e:
                    self.log(f"AI Processing error: {str(e)}")
                finally:
                    # Reset all states after complete response
                    self.is_ai_processing = False
                    self.character.set_animated(False)
                    self.message_queue.task_done()

    def process_new_message(self, user_input):
        start_time = time.time()
        print(f"\n[{time.strftime('%H:%M:%S')}] Processing new message: {user_input}")
        def single_callback(text_fragment):
            self.chat_interface and self.chat_interface.add_assistant_message(text_fragment)
        with self._ai_lock:
            try:
                self.is_transcribing = True
                if self.chat_interface:
                    if not hasattr(self, '_from_chat_interface') or not self._from_chat_interface:
                        self.chat_interface.add_user_message(user_input)
                    self._from_chat_interface = False
                response = self.assistant.send_message(user_input, ui_callback=single_callback)
                print(f"[TIMING] Total response completed in {time.time()-start_time:.2f}s")
            except Exception as e:
                print(f"[ERROR] after {time.time()-start_time:.2f}s: {str(e)}")
            finally:
                self.is_transcribing = False

    def cleanup(self):
        self.destroy()

    def create_mode_toggles(self):
        button_size = 64
        def create_button(icon_file, active=True, use_green=True):
            button = Image.new('RGBA', (button_size, button_size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(button)
            padding = 8
            circle_color = (220, 255, 220, 255) if active else (255, 220, 220, 255)
            draw.ellipse([padding, padding, button_size-padding, button_size-padding], fill=circle_color)
            icon = Image.open(icon_file).convert('RGBA')
            icon = icon.resize((28, 28), Image.Resampling.LANCZOS)
            icon_x = (button_size - icon.width) // 2
            icon_y = (button_size - icon.height) // 2
            final_button = Image.new('RGBA', (button_size, button_size), (0, 0, 0, 0))
            final_button.paste(button, (0, 0), button)
            final_button.paste(icon, (icon_x, icon_y), icon)
            return ImageTk.PhotoImage(final_button)

        self.mic_icon_active = create_button("assets/mic.png", True, use_green=True)
        self.mic_icon_muted = create_button("assets/mic-off.png", False, use_green=True)
        self.messages_icon_active = create_button("assets/messages-square.png", True, use_green=False)
        self.messages_icon_muted = create_button("assets/messages-square.png", False, use_green=False)

        self.mic_button = tk.Button(
            self.canvas,
            image=self.mic_icon_muted,
            command=self.toggle_mute,
            relief='flat',
            bg='SystemButtonFace',
            activebackground='SystemButtonFace',
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            width=64,
            height=64
        )

        self.messages_button = tk.Button(
            self.canvas,
            image=self.messages_icon_active,
            command=self.toggle_messages,
            relief='flat',
            bg='SystemButtonFace',
            activebackground='SystemButtonFace',
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            width=64,
            height=64
        )

        self.mic_button_window = self.canvas.create_window(
            self.character.x - 32,
            self.character.y - 100,
            window=self.mic_button,
            anchor='center'
        )

        self.messages_button_window = self.canvas.create_window(
            self.character.x + 32,
            self.character.y - 100,
            window=self.messages_button,
            anchor='center'
        )

    def toggle_messages(self):
        if not self.chat_interface:
            self.chat_interface = ChatInterface(
                self.canvas, 
                self.assistant,
                width=400,
                height=500,
                message_queue=self.message_queue
            )
            self.chat_window = self.canvas.create_window(
                self.character.x + 120,
                self.character.y - 200,
                window=self.chat_interface,
                anchor='nw',
                tags='chat'
            )
            self.chat_interface.chat_window_id = self.chat_window
            self.canvas.itemconfig(self.chat_window, state='normal')
            self.messages_button.configure(image=self.messages_icon_active)
        else:
            current_state = self.canvas.itemcget(self.chat_window, 'state')
            new_state = 'hidden' if current_state == 'normal' else 'normal'
            self.canvas.itemconfig(self.chat_window, state=new_state)
            new_icon = self.messages_icon_muted if new_state == 'hidden' else self.messages_icon_active
            self.messages_button.configure(image=new_icon)

    def toggle_red_dot(self):
        if not self.dot_visible:
            x = self.character.x + 50
            y = self.character.y - 50
            self.red_dot_id = self.canvas.create_oval(
                x-5, y-5, x+5, y+5,
                fill='red',
                outline='darkred',
                tags='red_dot'
            )
            self.dot_visible = True
        else:
            if self.red_dot_id:
                self.canvas.delete(self.red_dot_id)
                self.red_dot_id = None
            self.dot_visible = False

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()