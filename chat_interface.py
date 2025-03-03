import tkinter as tk
from tkinter import ttk
import time

class ChatInterface(tk.Frame):
    def __init__(self, parent, assistant, width=400, height=500, message_queue=None):
        super().__init__(parent)
        self.assistant = assistant
        self.message_queue = message_queue
        self.chat_window_id = None
        self.receiving_assistant_message = False
        
        self.configure(width=width, height=height, bg='#ECE9D8')
        self.grid_propagate(False)
        self.pack_propagate(False)
        
        main_container = tk.Frame(self, bg='#ECE9D8')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        header_frame = tk.Frame(main_container, bg='#2196F3', height=32)
        header_frame.pack(fill=tk.X)
        
        name_label = tk.Label(header_frame, text="SmartestLad", bg='#2196F3', fg='white', font=('Segoe UI', 10, 'bold'))
        name_label.pack(side=tk.LEFT, padx=12, pady=6)
        
        header_frame.bind("<Button-1>", self.start_move)
        header_frame.bind("<B1-Motion>", self.do_move)
        name_label.bind("<Button-1>", self.start_move)
        name_label.bind("<B1-Motion>", self.do_move)
        
        chat_height = int(height * 0.7)
        history_frame = tk.Frame(main_container, height=chat_height, bg='#ECE9D8')
        history_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        history_frame.pack_propagate(False)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Vertical.TScrollbar", troughcolor="#FFFFFF", background="#D4D0C8", bordercolor="#D4D0C8")

        scrollbar = ttk.Scrollbar(history_frame, style="Vertical.TScrollbar", orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.history = tk.Text(history_frame, wrap=tk.WORD, bg='white', font=('Segoe UI', 10), yscrollcommand=scrollbar.set)
        self.history.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        scrollbar.config(command=self.history.yview)
        
        input_height = int(height * 0.3)
        input_frame = tk.Frame(main_container, height=input_height, bg='#ECE9D8')
        input_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=2, pady=2)
        input_frame.pack_propagate(False)
        
        self.input_field = tk.Text(input_frame, wrap=tk.WORD, bg='white', font=('Segoe UI', 10), height=5)
        self.input_field.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.input_field.insert('1.0', 'Enter your thoughts here...')
        self.input_field.tag_configure('placeholder', foreground='grey')
        self.input_field.tag_add('placeholder', '1.0', 'end')
        
        self.input_field.bind('<FocusIn>', self._on_focus_in)
        self.input_field.bind('<FocusOut>', self._on_focus_out)
        
        self._configure_styles()
        
        def handle_enter(event):
            if not event.state & 0x1:
                self.send_message()
                return 'break'
        
        self.input_field.bind('<Return>', handle_enter)
    
    def start_move(self, event):
        self._drag_data = {"x": event.x_root, "y": event.y_root}

    def do_move(self, event):
        if not hasattr(self, '_drag_data'): return
        dx = event.x_root - self._drag_data["x"]
        dy = event.y_root - self._drag_data["y"]
        self._drag_data.update({"x": event.x_root, "y": event.y_root})
        if self.chat_window_id:
            coords = self.master.coords(self.chat_window_id)
            self.master.coords(self.chat_window_id, coords[0] + dx, coords[1] + dy)
    
    def add_message(self, message, sender_type='user'):
        # Only add newline if this is a new message (not a continuation)
        if not self.receiving_assistant_message:
            if self.history.get('1.0', tk.END).strip():
                self.history.insert(tk.END, "\n")
            timestamp = time.strftime("(%I:%M:%S %p)")
            sender = "uberushaximus" if sender_type == 'user' else "SmartestLad"
            tag = 'user_message' if sender_type == 'user' else 'assistant_message'
            
            self.history.insert(tk.END, f"{timestamp} ", 'timestamp')
            self.history.insert(tk.END, f"{sender}: ", tag)
            self.receiving_assistant_message = (sender_type == 'assistant')

        # Add the message fragment
        if isinstance(message, dict) and 'end_of_message' in message:
            self.receiving_assistant_message = False
        else:
            self.history.insert(tk.END, str(message))
            self.history.see(tk.END)

    def add_user_message(self, message):
        self.add_message(message, sender_type='user')

    def add_assistant_message(self, message):
        self.add_message(message, sender_type='assistant')

    def _on_focus_in(self, event):
        self._handle_placeholder(event)

    def _on_focus_out(self, event):
        self._handle_placeholder(event, removing=False)

    def _handle_placeholder(self, event, removing=True):
        text = self.input_field.get('1.0', 'end-1c')
        if removing and text == 'Enter your thoughts here...':
            self.input_field.delete('1.0', tk.END)
            self.input_field.tag_remove('placeholder', '1.0', 'end')
        elif not removing and not text.strip():
            self.input_field.insert('1.0', 'Enter your thoughts here...')
            self.input_field.tag_add('placeholder', '1.0', 'end')

    def send_message(self):
        message = self.input_field.get('1.0', tk.END).strip()
        if message and message != 'Enter your thoughts here...':
            # Clear input field first
            self.input_field.delete('1.0', tk.END)
            
            # Send to message queue if available
            if hasattr(self, 'message_queue') and self.message_queue:
                self._from_chat_interface = True
                self.message_queue.put(message)
                # Don't add the message here - let process_ai_messages handle it
                return

            # Only add directly if no message queue (fallback case)
            self.add_message(message)

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Vertical.TScrollbar", 
                       troughcolor="#FFFFFF", 
                       background="#D4D0C8", 
                       bordercolor="#D4D0C8")
        
        self.history.tag_configure('user_message', 
                                 foreground='#FF0000',
                                 font=('Segoe UI', 10))
        self.history.tag_configure('assistant_message', 
                                 foreground='#2196F3', 
                                 font=('Segoe UI', 10))
        self.history.tag_configure('timestamp', 
                                 foreground='#9E9E9E', 
                                 font=('Segoe UI', 9))