import tkinter as tk
from tkinter import Toplevel, Menu
import pyautogui
import time
import win32gui  # from pywin32
import threading

# -------------------------------------------------------------------
# Global State
# -------------------------------------------------------------------
is_typing = False        # Are we in a typing session?
is_paused = False        # Has the user manually paused typing?
current_index = 0        # Which character in user_text will we type next?
user_text = ""           # Holds the entire text to type
thread = None            # The typing thread
target_handle = None     # The handle of the window we originally started typing into

# We'll store the WPM in a Tk variable so it can be changed in the settings
wpm_var = None

# -------------------------------------------------------------------
# Helper to highlight the “next” character in the Tk Text widget
# -------------------------------------------------------------------
def highlight_next_char(index):
    """
    Schedule a call to highlight the next character (index)
    in the main Text widget. Must be done via .after(),
    because you can’t safely update Tk widgets from a worker thread.
    """
    window.after(0, lambda: do_highlight(index))

def do_highlight(index):
    """
    Actually perform the highlighting (runs in the main thread).
    Clears old highlight and highlights the character at 'index'.
    """
    global text_widget, user_text

    # Remove any existing highlight
    text_widget.tag_remove("highlight", "1.0", tk.END)

    # If index is past the last character, do nothing
    if index >= len(user_text):
        return

    # Calculate the positions for the character
    start_pos = f"1.0+{index}c"
    end_pos = f"1.0+{index+1}c"

    text_widget.tag_add("highlight", start_pos, end_pos)

# -------------------------------------------------------------------
# Main Typing Loop (Runs in a Thread)
# -------------------------------------------------------------------
def typing_loop():
    global is_typing, is_paused, current_index, user_text, target_handle

    # Convert WPM to delay per character (assuming 5 chars per word)
    wpm = wpm_var.get()  # e.g. 200
    delay_per_char = 60.0 / (wpm * 5)

    # Wait 3 seconds so user can switch to the target window
    time.sleep(3)

    # Capture the active window handle
    target_handle = win32gui.GetForegroundWindow()

    # While we still have characters to type and are in typing mode
    while is_typing and current_index < len(user_text):
        # 1) If the user is paused, or if the active window changed, wait
        if is_paused or (win32gui.GetForegroundWindow() != target_handle):
            while (is_paused or (win32gui.GetForegroundWindow() != target_handle)) and is_typing:
                time.sleep(0.1)
            if not is_typing:
                break  # The user hit Stop while we were paused

        # 2) Highlight the next character
        highlight_next_char(current_index)

        # 3) Type one character
        pyautogui.typewrite(user_text[current_index])

        # 4) Increment index
        current_index += 1

        # 5) Delay
        time.sleep(delay_per_char)

    # When the loop exits, we stop typing
    is_typing = False

# -------------------------------------------------------------------
# Button Handlers
# -------------------------------------------------------------------
def start_typing():
    """
    Start (or resume) typing:
      - If we’re already in a typing session, this “unpauses” if paused.
      - If not typing, starts from current_index (wherever we left off).
    """
    global is_typing, is_paused, thread, user_text, current_index

    if is_typing:
        # If we're already typing, just unpause
        is_paused = False
        return

    # If not typing, get the text from the widget. 
    # (If you truly want to pick up exactly where you left off 
    #  without re-reading the box, you could remove this.)
    user_text = text_widget.get("1.0", tk.END).rstrip("\n")

    # Start or resume from current_index
    is_typing = True
    is_paused = False

    # Spawn the worker thread
    thread = threading.Thread(target=typing_loop, daemon=True)
    thread.start()

def pause_typing():
    """
    Pauses typing by setting is_paused to True.
    """
    global is_paused
    if is_typing and not is_paused:
        is_paused = True

def stop_typing():
    """
    Fully stop typing and reset everything so if we start again
    it goes from the beginning.
    """
    global is_typing, is_paused, current_index, user_text
    is_typing = False
    is_paused = False
    current_index = 0
    # Remove highlight
    text_widget.tag_remove("highlight", "1.0", tk.END)

# -------------------------------------------------------------------
# Settings for WPM
# -------------------------------------------------------------------
def open_settings():
    global wpm_var
    settings_window = Toplevel(window)
    settings_window.title("Settings")
    settings_window.attributes('-topmost', True)  # Keep on top as well

    label = tk.Label(settings_window, text="Adjust typing speed (Words per Minute):")
    label.pack(padx=10, pady=5)

    scale = tk.Scale(
        settings_window,
        from_=200,
        to=1000,
        orient=tk.HORIZONTAL,
        variable=wpm_var,
        length=300
    )
    scale.pack(padx=10, pady=5)

def create_menu():
    menubar = Menu(window)
    settings_menu = Menu(menubar, tearoff=0)
    settings_menu.add_command(label="Typing Speed...", command=open_settings)
    menubar.add_cascade(label="Settings", menu=settings_menu)
    window.config(menu=menubar)

# -------------------------------------------------------------------
# Main UI Setup
# -------------------------------------------------------------------
window = tk.Tk()
window.title("Human‐Like Typer")
window.attributes('-topmost', True)  # Keep the main window always on top

wpm_var = tk.IntVar(value=200)  # Default WPM = 200

label = tk.Label(window, text="Type or paste the text you want to send below:")
label.pack(pady=10)

text_widget = tk.Text(window, height=10, width=50)
text_widget.pack(padx=10, pady=5)

# Configure a tag for highlight style
text_widget.tag_config("highlight", background="yellow")

# Create buttons
button_frame = tk.Frame(window)
button_frame.pack(pady=5)

start_button = tk.Button(button_frame, text="Start/Resume", command=start_typing)
start_button.pack(side=tk.LEFT, padx=5)

pause_button = tk.Button(button_frame, text="Pause", command=pause_typing)
pause_button.pack(side=tk.LEFT, padx=5)

stop_button = tk.Button(button_frame, text="Stop", command=stop_typing)
stop_button.pack(side=tk.LEFT, padx=5)

create_menu()

window.mainloop()
