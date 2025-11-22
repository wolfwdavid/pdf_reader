import pypdf
import pyttsx3
import os
import threading
import time 
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple

# --- CONFIGURATION ---
DEFAULT_PDF_PATH = "" 
MAX_WORKERS = os.cpu_count() * 2
# Add a short delay for thread-safe polling
TTS_POLL_DELAY = 0.1 
DEFAULT_RATE = 180 # Define default rate globally for clarity

class AudioBookApp:
    def __init__(self, master):
        self.master = master
        master.title("Advanced Concurrent PDF Audiobook Reader")
        master.geometry("700x500")

        # Application State & Controls
        self.pdf_file_path = DEFAULT_PDF_PATH
        self.is_reading = False
        self.is_paused = False
        self.stop_flag = threading.Event() 
        self.current_content = [] 

        # 🛑 FIX: Initialize GUI attributes to None to prevent AttributeError 
        self.status_bar = None
        self.start_button = None
        self.pause_button = None
        self.stop_button = None
        self.display_text = None
        self.path_label = None
        self.speed_label = None # ✨ NEW: Speed label attribute
        # End FIX

        # TTS Engine Setup
        self.tts_engine = pyttsx3.init()
        self.tts_engine_lock = threading.Lock()
        self.tts_engine.connect('started-word', self._on_tts_word)
        self.current_page_text = ""

        # Setup order remains critical: TTS first, then GUI
        self.setup_tts_voice()
        self.setup_gui() 
        
        # 🛑 FINAL FIX 1: Add a protocol handler to stop the TTS engine when the window closes
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing) 

    # --- New Method to Handle Clean Shutdown ---
    def on_closing(self):
        """Called when the user closes the window. Stops TTS engine before quitting."""
        try:
            # 1. Ensure any ongoing reading process stops
            self.stop_reading() 
            
            # 2. Explicitly shut down the pyttsx3 engine
            if self.tts_engine:
                # Use a small delay to ensure TTS thread recognizes the stop before the app is destroyed
                time.sleep(TTS_POLL_DELAY) 
                self.tts_engine.stop()
                
        except Exception:
            # Ignore errors during shutdown
            pass
            
        # 3. Destroy the Tkinter root window
        self.master.destroy()
        
    def setup_tts_voice(self):
        """Configures the TTS engine for a female voice and optimal rate."""
        voices = self.tts_engine.getProperty('voices')
        female_voice_id = next((v.id for v in voices if 'female' in v.name.lower() or 'zira' in v.name.lower()), None)
        if female_voice_id:
            self.tts_engine.setProperty('voice', female_voice_id)
            
        self.tts_engine.setProperty('rate', DEFAULT_RATE) # Use DEFAULT_RATE

    def setup_gui(self):
        """Lays out all GUI elements including new controls."""
        
        # --- File Selection Frame ---
        file_frame = ttk.Frame(self.master, padding="10")
        file_frame.pack(fill='x')
        self.path_label = ttk.Label(file_frame, text="No PDF selected.")
        self.path_label.pack(side='left', fill='x', expand=True)
        ttk.Button(file_frame, text="Select PDF", command=self.select_pdf_file).pack(side='right')

        # --- Control Frame ---
        control_frame = ttk.Frame(self.master, padding="10 0 10 10")
        control_frame.pack(fill='x')
        
        # Start/Pause/Stop Buttons
        self.start_button = ttk.Button(control_frame, text="▶️ Start", command=self.start_reading, state=tk.DISABLED)
        self.start_button.pack(side='left', padx=5)
        self.pause_button = ttk.Button(control_frame, text="⏸️ Pause", command=self.pause_resume_reading, state=tk.DISABLED)
        self.pause_button.pack(side='left', padx=5)
        self.stop_button = ttk.Button(control_frame, text="⏹️ Stop", command=self.stop_reading, state=tk.DISABLED)
        self.stop_button.pack(side='left', padx=5)

        # Speed Slider & Label (Updated Layout)
        ttk.Label(control_frame, text="Speed:").pack(side='left', padx=(20, 5))
        
        self.speed_slider = ttk.Scale(control_frame, from_=100, to=300, orient=tk.HORIZONTAL, command=self._set_speed_from_slider)
        self.speed_slider.set(DEFAULT_RATE) # Default speed
        self.speed_slider.pack(side='left', expand=True, fill='x')
        
        # ✨ NEW: Label to show the current speed in WPM
        self.speed_label = ttk.Label(control_frame, text=f"{DEFAULT_RATE} WPM")
        self.speed_label.pack(side='left', padx=(5, 0)) # Positioned right after the slider
        
        # --- Display Frame ---
        display_frame = ttk.Frame(self.master, padding="10 0 10 10")
        display_frame.pack(fill='both', expand=True)
        
        ttk.Label(display_frame, text="Currently Reading (Highlighted Word):").pack(fill='x')
        
        # ScrolledText for display and highlighting
        self.display_text = scrolledtext.ScrolledText(display_frame, wrap=tk.WORD, state=tk.DISABLED, font=('Arial', 14))
        self.display_text.pack(fill='both', expand=True)
        
        # Define a unique tag for highlighting
        self.display_text.tag_config('highlight', background='yellow', foreground='black')
        
        # Status Bar
        self.status_bar = ttk.Label(self.master, text="Ready.", relief=tk.SUNKEN, anchor='w')
        self.status_bar.pack(fill='x')
        
    # --- GUI CALLBACKS ---

    def select_pdf_file(self):
        """Opens file dialog and triggers content extraction."""
        filepath = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            self.pdf_file_path = filepath
            self.path_label.config(text=f"Selected: {os.path.basename(filepath)}")
            self.start_button.config(state=tk.DISABLED)
            self.status_bar.config(text="Extracting content...")
            
            # Start extraction in background
            threading.Thread(target=self._initial_extraction, daemon=True).start()

    def _initial_extraction(self):
        """Handles the initial, concurrent text extraction."""
        content = self._get_extracted_content_concurrently()
        self.master.after(0, lambda: self._post_extraction_update(content)) 

    def _post_extraction_update(self, content):
        """Updates GUI after extraction finishes."""
        if content and content[0][0] == 0 and content[0][1].startswith("EXTRACTION_ERROR"):
            error_message = content[0][1].replace("EXTRACTION_ERROR: ", "")
            self.status_bar.config(text=f"❌ Extraction Failed: {error_message}")
            self.current_content = []
            self.start_button.config(state=tk.DISABLED)
            self.update_display(f"Extraction failed. Please check your PDF file.\nError: {error_message}")
        elif content:
            self.current_content = content
            self.start_button.config(state=tk.NORMAL)
            self.status_bar.config(text=f"Extraction complete. {len(self.current_content)} pages loaded. Ready to read.")
        else:
            self.status_bar.config(text="Extraction failed or PDF is empty.")
            self.start_button.config(state=tk.DISABLED)
            self.current_content = []

    def _set_speed_from_slider(self, value):
        """Updates the TTS engine reading speed and the speed label."""
        speed = int(float(value))
        
        # Update TTS engine property
        self.tts_engine.setProperty('rate', speed)
        
        # ✨ NEW: Update the dedicated speed label
        if self.speed_label:
             self.speed_label.config(text=f"{speed} WPM")
             
        # Optional: update status bar less verbosely, or remove this line
        if self.status_bar:
            self.status_bar.config(text=f"Speed set to {speed} WPM.")

    # --- WORD HIGHLIGHTING ---

    def _on_tts_word(self, name, location, length):
        """Callback triggered by pyttsx3 when a new word is about to be spoken."""
        self.master.after(0, lambda: self._highlight_word(location, length))

    def _highlight_word(self, location, length):
        """Removes old highlight and applies new highlight to the current word."""
        try:
            if not self.display_text: return
            
            self.display_text.config(state=tk.NORMAL)
            self.display_text.tag_remove('highlight', 1.0, tk.END)
            
            start_index = f"1.{location}"
            end_index = f"1.{location + length}"
            
            self.display_text.tag_add('highlight', start_index, end_index)
            self.display_text.see(end_index)
            
            self.display_text.config(state=tk.DISABLED)
            
        except Exception:
            pass 

    # --- READING CONTROLS ---

    def start_reading(self):
        """Starts the main reading thread."""
        if self.is_reading: return

        self.is_reading = True
        self.stop_flag.clear()
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(text="⏸️ Pause", state=tk.NORMAL)
        self.stop_button.config(state=tk.NORMAL)
        self.status_bar.config(text="Reading started...")

        threading.Thread(target=self._reading_process, daemon=True).start()

    def pause_resume_reading(self):
        """Toggles the pause state."""
        if not self.is_reading: return

        with self.tts_engine_lock:
            if self.is_paused:
                self.is_paused = False
                self.pause_button.config(text="⏸️ Pause")
                self.status_bar.config(text="Reading resumed.")
            else:
                self.is_paused = True
                self.pause_button.config(text="▶️ Resume")
                self.status_bar.config(text="Reading paused.")

    def stop_reading(self):
        """Sets the flag to stop reading and stops the TTS engine immediately."""
        self.stop_flag.set()
        with self.tts_engine_lock:
            # Safely stop the engine and flush the queue
            self.tts_engine.stop()
        self.is_reading = False
        self._cleanup_buttons()
        self.status_bar.config(text="Reading stopped.")

    # --- READING PROCESS ---

    def _get_extracted_content_concurrently(self) -> List[Tuple[int, str]]:
        """
        Handles concurrent text extraction.
        INCLUDES ROBUST THREAD-LEVEL EXCEPTION CAPTURE.
        """
        try:
            # Check for file existence before loading pypdf
            if not os.path.exists(self.pdf_file_path):
                 raise FileNotFoundError(f"File not found: {self.pdf_file_path}")

            reader = pypdf.PdfReader(self.pdf_file_path)
            num_pages = len(reader.pages)
            extracted_content: List[Tuple[int, str]] = [None] * num_pages
            
            def extract_page_text(page_index: int):
                return page_index, reader.pages[page_index].extract_text()

            error_during_extraction = None
            
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                future_to_page = {
                    executor.submit(extract_page_text, i): i 
                    for i in range(num_pages)
                }
                for future in as_completed(future_to_page):
                    if future.exception():
                        if error_during_extraction is None:
                            error_during_extraction = future.exception()
                    else:
                        page_index, text = future.result()
                        extracted_content[page_index] = (page_index + 1, text) 
            
            if error_during_extraction:
                raise error_during_extraction
            
            return [item for item in extracted_content if item is not None]
        
        except Exception as e:
            return [(0, f"EXTRACTION_ERROR: {e.__class__.__name__}: {str(e)}")]


    def _reading_process(self):
        """
        The main background thread function for speaking and display.
        The lock is released immediately after .say() to prevent GUI freeze.
        """
        
        try:
            for page_num, raw_text in self.current_content:
                if self.stop_flag.is_set(): break
                
                if not raw_text.strip(): continue

                self.current_page_text = f"PAGE {page_num}\n\n{raw_text}"
                self.master.after(0, lambda: self.update_display(self.current_page_text))
                
                # Only acquire the lock for the instantaneous .say() call
                with self.tts_engine_lock:
                    self.tts_engine.say(raw_text)
                    
                # The lock is now released. Thread monitors isBusy() status.
                while self.tts_engine.isBusy() and not self.stop_flag.is_set():
                    # Check for pause state (which is safely set by the main thread)
                    while self.is_paused and not self.stop_flag.is_set():
                        # The background thread sleeps when paused.
                        time.sleep(TTS_POLL_DELAY)
                    
                    # Wait a small amount of time to check busy status
                    time.sleep(TTS_POLL_DELAY) 
                    
                # Exit cleanup if stop was requested while busy
                if self.stop_flag.is_set():
                    self.master.after(0, lambda: self.display_text.tag_remove('highlight', 1.0, tk.END))
                    break
                
                # Cleanup highlight after page is finished speaking
                self.master.after(0, lambda: self.display_text.tag_remove('highlight', 1.0, tk.END))

            # Normal cleanup if loop completes or breaks due to stop_flag
            self.master.after(0, self._cleanup_buttons)
            self.master.after(0, lambda: self.status_bar.config(text="Reading finished." if not self.stop_flag.is_set() else "Reading stopped."))
            self.is_reading = False

        except Exception as e:
            # Log any unexpected crash for debugging
            print(f"CRASH INSIDE READING PROCESS THREAD: {e.__class__.__name__}: {e}")
            self.master.after(0, lambda: self.status_bar.config(text=f"❌ Reading Process Crash: {e.__class__.__name__}"))
            self.master.after(0, self.stop_reading) 


    def update_display(self, text):
        """Updates the ScrolledText box from a thread-safe manner."""
        if not self.display_text: return 
        
        self.display_text.config(state=tk.NORMAL)
        self.display_text.delete(1.0, tk.END)
        self.display_text.insert(tk.END, text)
        self.display_text.config(state=tk.DISABLED)

    def _cleanup_buttons(self):
        """Resets buttons to post-reading state."""
        self.start_button.config(state=tk.NORMAL if self.pdf_file_path else tk.DISABLED)
        self.pause_button.config(state=tk.DISABLED, text="⏸️ Pause")
        self.stop_button.config(state=tk.DISABLED)

# --- EXECUTION ---
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AudioBookApp(root)
        root.mainloop()
    except Exception as e:
        print(f"An unexpected error occurred during startup: {e}")