import os
import sys
import winreg
import ctypes
import ctypes.wintypes
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import mss 

# Try to import Tesseract OCR
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

    
class ScreenCapture:
    def __init__(self, root, source_file_path=None):
        self.root = root
        self.source_file_path = source_file_path 
        self.start_x = None
        self.start_y = None
        self.end_x = None
        self.end_y = None
        self.rect = None
        self.canvas = None
        self.screenshot = None
        self.monitor_bbox = None       
     
    @staticmethod
    def get_current_monitor_bbox():
        """Return the bounding box (x1, y1, x2, y2) of the monitor where the mouse is located"""
        user32 = ctypes.windll.user32
        pt = ctypes.wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        mouse_x, mouse_y = pt.x, pt.y

        monitors = []

        def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
            r = ctypes.cast(lprcMonitor, ctypes.POINTER(ctypes.wintypes.RECT)).contents
            monitors.append((r.left, r.top, r.right, r.bottom))
            return 1

        MONITORENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong,
                                            ctypes.POINTER(ctypes.wintypes.RECT), ctypes.c_double)
        user32.EnumDisplayMonitors(0, 0, MONITORENUMPROC(callback), 0)

        # Find which monitor contains the mouse
        for (x1, y1, x2, y2) in monitors:
            if x1 <= mouse_x < x2 and y1 <= mouse_y < y2:
                return (x1, y1, x2, y2)

        # Fallback: return virtual screen
        return (0, 0, user32.GetSystemMetrics(78), user32.GetSystemMetrics(79))
        
    def capture_region(self):
        """Capture a region of the screen selected by the user (multi-monitor safe with mss)"""
        # Get monitor where mouse is located
        self.monitor_bbox = self.get_current_monitor_bbox()

        # Use mss to capture only that monitor
        with mss.mss() as sct:
            mon = {
                "left": self.monitor_bbox[0],
                "top": self.monitor_bbox[1],
                "width": self.monitor_bbox[2] - self.monitor_bbox[0],
                "height": self.monitor_bbox[3] - self.monitor_bbox[1],
            }
            sct_img = sct.grab(mon)
            self.screenshot = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

        # Configure overlay window
        self.root.deiconify()
        self.root.attributes('-fullscreen', False, '-alpha', 0.3, '-topmost', True)

        monitor_x, monitor_y, monitor_width, monitor_height = (
            self.monitor_bbox[0],
            self.monitor_bbox[1],
            self.monitor_bbox[2] - self.monitor_bbox[0],
            self.monitor_bbox[3] - self.monitor_bbox[1],
        )

        self.root.geometry(f"{monitor_width}x{monitor_height}+{monitor_x}+{monitor_y}")
        self.root.overrideredirect(True)
        self.root.lift()
        self.root.focus_force()
        self.root.configure(bg='black')

        # Canvas overlay
        self.canvas = tk.Canvas(self.root, highlightthickness=0, cursor="cross", bg='black')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Show screenshot of current monitor
        self.photo = ImageTk.PhotoImage(self.screenshot)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

        # Mouse bindings
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.root.bind("<Escape>", self.cancel_capture)

        self.canvas.update()
                        
    def on_click(self, event):
        self.start_x = event.x
        self.start_y = event.y
        
    def on_drag(self, event):
        # Remove previous selection
        if self.rect:
            self.canvas.delete(self.rect)
            
        # Draw new selection rectangle
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline="red", width=3, fill=""
        )
        
    def on_release(self, event):
        self.end_x = event.x
        self.end_y = event.y

        x1 = min(self.start_x, self.end_x)
        y1 = min(self.start_y, self.end_y)
        x2 = max(self.start_x, self.end_x)
        y2 = max(self.start_y, self.end_y)

        if abs(x2 - x1) > 10 and abs(y2 - y1) > 10:
            # Crop directly with canvas-relative coords (already matches screenshot)
            cropped = self.screenshot.crop((x1, y1, x2, y2))
            self.close_capture()
            self.perform_ocr(cropped)
        else:
            self.close_capture()
             
    def cancel_capture(self, event):
        self.close_capture()
        
    def close_capture(self):
        self.root.overrideredirect(False)  # Restore window decorations
        self.root.attributes('-fullscreen', False)
        self.root.withdraw()
        
    def perform_ocr(self, image):
        """Perform OCR using Tesseract and show results"""
        if not TESSERACT_AVAILABLE:
            text = (
                "Tesseract OCR not installed!\n\n"
                "To enable OCR functionality:\n"
                "1. pip install pytesseract\n"
                "2. Install Tesseract OCR engine:\n"
                "   https://github.com/UB-Mannheim/tesseract/wiki\n"
            )
            self.show_text_editor(text, image)
            return

        # Verify Tesseract executable
        try:
            pytesseract.get_tesseract_version()
        except pytesseract.TesseractNotFoundError:
            text = (
                "Tesseract executable not found!\n\n"
                "Solutions:\n"
                "1. Add Tesseract to PATH\n"
                "2. Or set path manually in the script:\n"
                "   pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'\n"
            )
            self.show_text_editor(text, image)
            return

        # Try OCR with different PSM modes
        text = ""
        for psm in [6, 7, 8, 3]:
            try:
                config = f'--oem 3 --psm {psm}'
                result = pytesseract.image_to_string(image, config=config)
                if result.strip():
                    text = result
                    break
            except Exception:
                continue
        
        # Try OCR
        # text = pytesseract.image_to_string(image, config="--oem 3 --psm 6").strip()
        
        # Fallback if no text found
        if not text.strip():
            text = pytesseract.image_to_string(image).strip()

        if not text:
            text = (
                "No text detected in the selected region.\n\n"
                "Tips:\n• Ensure good contrast\n• Avoid rotated text\n• Try smaller regions"
            )

        # Show results in editor
        self.show_text_editor(text, image)

    
        
    def show_text_editor(self, text, image=None):
        """Minimal text editor for OCR results, with screenshot preview and rename button"""
        editor = tk.Toplevel(self.root)
        editor.title("OCR Result")
        editor.wm_minsize(500, 300)
        
        frame = tk.Frame(editor)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Show screenshot preview
        if image is not None:
            img_preview = image.copy()
            img_preview.thumbnail((680, 250))
            photo = ImageTk.PhotoImage(img_preview)
            label = tk.Label(frame, image=photo)
            label.image = photo  # Prevent garbage collection
            label.pack(pady=5)
        
        # File management section
        has_source_file = self.source_file_path and os.path.exists(self.source_file_path)
        if has_source_file:
            self._create_file_controls(frame, editor)
        
        # Text editor section
        text_box = self._create_text_editor(frame, text)
        
        # Bind preview updates if file controls exist
        if has_source_file:
            self._bind_preview_updates(text_box, editor)

    def _create_file_controls(self, frame, editor):
        """Create file rename controls and buttons"""
        button_frame = tk.Frame(frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Current filename display
        current_filename = os.path.basename(self.source_file_path)
        tk.Label(button_frame, text=f"Source file: {current_filename}", 
                font=("Arial", 9), fg="gray").pack(side=tk.TOP, anchor=tk.W)
        
        # Preview label
        self.preview_label = tk.Label(button_frame, text="Preview: ", 
                                    font=("Arial", 9), fg="blue")
        self.preview_label.pack(side=tk.TOP, anchor=tk.W)
        
        # Rename button
        tk.Button(button_frame, text="Rename File", command=lambda: self._rename_file(editor),
                bg="#4CAF50", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=5, pady=5)
        
        # Prefix buttons
        context_handler = ContextMenuHandler()
        for prefix in context_handler.prefix_options:
            label = prefix.strip("+").split("+")[0]
            tk.Button(button_frame, text=label, command=lambda p=prefix: self._add_prefix(p),
                    bg="#2196F3", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=5, pady=5)
            
        # Suffix buttons
        for suffix in context_handler.suffix_options:
            label = suffix.strip("+").split("+")[0]
            tk.Button(button_frame, text=label, command=lambda p=suffix: self._add_suffix(p),
                    bg="#2196F3", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=5, pady=5)

    def _create_text_editor(self, frame, text):
        """Create the main text editor with scrollbar"""
        text_frame = tk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_box = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 11), 
                        yscrollcommand=scrollbar.set, undo=True)
        text_box.insert(tk.END, self._text_to_filename(text))
        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Undo/Redo shortcuts
        text_box.bind("<Control-z>", lambda e: self._safe_undo_redo(text_box.edit_undo))
        text_box.bind("<Control-y>", lambda e: self._safe_undo_redo(text_box.edit_redo))
        text_box.edit_reset()
        
        scrollbar.config(command=text_box.yview)
        self.text_box = text_box  # Store reference for other methods
        return text_box

    def _safe_undo_redo(self, action):
        """Safely perform undo/redo operations"""
        try:
            action()
        except tk.TclError:
            pass
        return "break"

    def _bind_preview_updates(self, text_box, editor):
        """Bind events to update filename preview"""
        for event in ['<KeyRelease>', '<Button-1>', '<FocusOut>']:
            text_box.bind(event, lambda e: self._update_preview())
        editor.after(100, self._update_preview)

    def _update_preview(self):
        """Update the filename preview label"""
        try:
            new_name = self.text_box.get(1.0, tk.END).strip()
            if not new_name:
                self.preview_label.config(text="Preview: ")
                return
            
            # Clean and validate filename
            cleaned_name = self._clean_filename(new_name)
            if not cleaned_name:
                self.preview_label.config(text="Preview: ")
                return
            
            # Add extension and check for duplicates
            _, ext = os.path.splitext(self.source_file_path)
            if not cleaned_name.endswith(ext):
                cleaned_name += ext
            
            file_dir = os.path.dirname(self.source_file_path)
            new_file_path = os.path.join(file_dir, cleaned_name)
            
            # Show potential duplicate handling
            display_name = cleaned_name
            if os.path.exists(new_file_path):
                name_without_ext = os.path.splitext(cleaned_name)[0]
                display_name = f"{name_without_ext} (1){ext}"
            
            self.preview_label.config(text=f"Preview: {display_name}")
            
        except Exception:
            self.preview_label.config(text="Preview: ")

    def _clean_filename(self, filename):
        """Clean filename by removing invalid characters and whitespace"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, ' ')
        return ' '.join(filename.split())  # Remove excessive whitespace

    def _add_prefix(self, prefix):
        """Add prefix to the beginning of text"""
        self.text_box.edit_separator()
        self.text_box.insert("1.0", prefix)
        self._update_preview()
        
    def _add_suffix(self, suffix):
        """Add suffix to the end of text"""
        self.text_box.edit_separator()
        self.text_box.insert(tk.END, suffix)
        self._update_preview()

    def _rename_file(self, editor):
        """Rename the source file using text from editor"""
        new_name = self.text_box.get(1.0, tk.END).strip()
        if not new_name:
            messagebox.showerror("Error", "Please enter a filename in the text box.")
            return
        
        cleaned_name = self._clean_filename(new_name)
        if not cleaned_name:
            messagebox.showerror("Error", "Filename cannot be empty after cleaning.")
            return
        
        # Add extension
        _, ext = os.path.splitext(self.source_file_path)
        if not cleaned_name.endswith(ext):
            cleaned_name += ext
        
        # Handle duplicates
        file_dir = os.path.dirname(self.source_file_path)
        new_file_path = self._get_unique_filepath(os.path.join(file_dir, cleaned_name))
        
        try:
            os.rename(self.source_file_path, new_file_path)
            messagebox.showinfo("Success", f"File renamed to:\n{os.path.basename(new_file_path)}")
            editor.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file:\n{str(e)}")

    def _get_unique_filepath(self, filepath):
        """Generate unique filepath by adding counter if file exists"""
        if not os.path.exists(filepath):
            return filepath
        
        counter = 1
        name_without_ext, ext = os.path.splitext(filepath)
        while os.path.exists(filepath):
            filepath = f"{name_without_ext} ({counter}){ext}"
            counter += 1
        return filepath
                       
    def _text_to_filename(self, text):
        """Converts text to a clean single sentence suitable for filenames"""
        if not text:
            return ""     

        invalid_chars = '<>:"/\\|?*'
        clean_text = text.strip()
        
        for char in invalid_chars:
            clean_text = clean_text.replace(char, ' ')
        
        clean_text = ' '.join(clean_text.split()) # Remove newlines and excessive whitespace  
        clean_text = clean_text.lower().title() # Convert to lowercase and capitalize each word
        
        return clean_text


class ContextMenuHandler:
    def __init__(self):
        self.prefix_options = ["+Book+year+", "+Paper+year+", "+Thesis+year+", "+Report+year+", 
                               "+Slides+year+", "+Presentation+year+", "+Draft+year+"]
        self.suffix_options = ["+authors"]  
        self.menu_name = "Add Prefix-Suffix and OCR"
        self.ocr_menu_name = "Tesseract OCR"
        self.python_executable = sys.executable
        self.script_path = os.path.abspath(__file__)
        self.pythonw_executable = self.python_executable.replace("python.exe", "pythonw.exe")
    
    def install(self):
        """Install the context menu entries in Windows Registry"""
        try:
            if self.install_file_menu():
                print("Context menu entries installed successfully!")
                print("- Right-click on any file -> 'Tesseract OCR' and 'Add Prefix-Suffix'\n")

                if TESSERACT_AVAILABLE:
                    try:
                        version = pytesseract.get_tesseract_version()
                        print(f"✅ Tesseract OCR is ready (version {version})")
                    except Exception:
                        print("⚠️ Tesseract package installed but executable not found")
                        print("   Install Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki")
                else:
                    print("⚠️ pytesseract not installed. Install with: pip install pytesseract")
                return True
            else:
                print("Some context menu entries could not be installed.")
                return False
        except Exception as e:
            print(f"Error installing context menu: {e}")
            return False
   
    def install_file_menu(self):
        """Install file-based context menu entries"""
        try:
            # Create main menu entry for files
            main_key_path = rf"*\shell\{self.menu_name}"
            with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, main_key_path) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, self.menu_name)
                winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
            
            # Create submenu entries in command store
            shell_commands_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
            subcommands = []
            
            # Helper function to create command entries
            def create_command(command_key, display_text, command_args):
                command_key_path = os.path.join(shell_commands_path, command_key)
                
                with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key_path) as key:
                    winreg.SetValueEx(key, None, 0, winreg.REG_SZ, display_text)
                
                command_path = os.path.join(command_key_path, "command")
                with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_path) as cmd_key:
                    command_str = f'"{self.pythonw_executable}" "{self.script_path}" {command_args} "%1"'
                    winreg.SetValueEx(cmd_key, None, 0, winreg.REG_SZ, command_str)
                
                subcommands.append(command_key)
            
            # Create OCR command
            create_command("ocr.tesseract", "Tesseract OCR (Select region. ESC to cancel)", "ocr")
            
            # Create prefix commands
            for prefix in self.prefix_options:
                create_command(f"prefix.{prefix}", f"Add prefix '{prefix}'", f'prefix "{prefix}"')
            
            # Create suffix commands  
            for suffix in self.suffix_options:
                create_command(f"suffix.{suffix}", f"Add suffix '{suffix}'", f'suffix "{suffix}"')
            
            # Connect subcommands to main menu
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, main_key_path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, ";".join(subcommands))
            
            return True
            
        except Exception as e:
            print(f"Error installing file menu: {e}")
            return False
        
    def safe_delete_key(self, root_key, key_path):
        """Safely delete a registry key and all its subkeys"""
        try:
            try:
                winreg.OpenKey(root_key, key_path)
            except FileNotFoundError:
                return True
            
            handle = winreg.OpenKey(root_key, key_path, 0, winreg.KEY_ALL_ACCESS)
            
            try:
                info = winreg.QueryInfoKey(handle)
                num_subkeys = info[0]
                
                for i in range(num_subkeys):
                    subkey_name = winreg.EnumKey(handle, 0)
                    subkey_path = f"{key_path}\\{subkey_name}"
                    self.safe_delete_key(root_key, subkey_path)
                
                winreg.CloseKey(handle)
                winreg.DeleteKey(root_key, key_path)
                return True
                
            except Exception as e:
                winreg.CloseKey(handle)
                print(f"Failed to delete {key_path}: {e}")
                return False
                
        except Exception as e:
            print(f"Error accessing {key_path}: {e}")
            return False
    
    def uninstall(self):
        """Remove the context menu entries from Windows Registry"""
        try:
            success = True
            
            # Remove file menu entries
            key_path = r"*\shell\{}".format(self.menu_name)
            if not self.safe_delete_key(winreg.HKEY_CLASSES_ROOT, key_path):
                success = False
            
            # Remove submenu entries from HKEY_LOCAL_MACHINE
            shell_commands_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
            
            # Define all keys to remove
            keys_to_remove = ["ocr.tesseract"]
            keys_to_remove.extend(f"prefix.{opt}" for opt in self.prefix_options)
            keys_to_remove.extend(f"suffix.{opt}" for opt in self.suffix_options)
            
            # Remove all command keys
            for key in keys_to_remove:
                if not self.safe_delete_key(winreg.HKEY_LOCAL_MACHINE, os.path.join(shell_commands_path, key)):
                    success = False
            
            if success:
                print("All context menu entries removed successfully!")
            else:
                print("Some registry keys could not be removed. You may need to delete them manually.")
            
            return success
            
        except Exception as e:
            print(f"Error uninstalling context menu: {e}")
            return False
     
    def add_prefix(self, prefix, file_path):
        """Add a prefix to the selected file"""
        if not os.path.exists(file_path):
            return False
            
        file_dir = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        new_file_path = os.path.join(file_dir, f"{prefix}{file_name}")
        
        counter = 1
        while os.path.exists(new_file_path):
            base_name, extension = os.path.splitext(file_name)
            new_file_path = os.path.join(file_dir, f"{prefix}{base_name} ({counter}){extension}")
            counter += 1
        
        os.rename(file_path, new_file_path)
        return True
                
    def add_suffix(self, suffix, file_path):
        """Add a suffix to the selected file"""
        if not os.path.exists(file_path):
            return False
            
        file_dir = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        base_name, extension = os.path.splitext(file_name)
        new_file_path = os.path.join(file_dir, f"{base_name}{suffix}{extension}")
        
        counter = 1
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(file_dir, f"{base_name}{suffix} ({counter}){extension}")
            counter += 1
        
        os.rename(file_path, new_file_path)
        return True
            
    def start_ocr(self, source_file_path=None):
        """Start the OCR region selection process"""
        root = tk.Tk()
        root.withdraw()
        capture = ScreenCapture(root, source_file_path)
        capture.capture_region()
        root.mainloop()    

def check_dependencies():
    """Check if required packages are installed"""
    missing_packages = []
    
    try:
        import PIL
    except ImportError:
        missing_packages.append("pillow")
            
    try:
        import mss
    except ImportError:
        missing_packages.append("mss")
    
    if missing_packages:
        print("Missing required packages:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\nInstall them with:")
        print(f"  pip install {' '.join(missing_packages)}")
        return False
    
    # Check Tesseract separately (optional but recommended)
    if not TESSERACT_AVAILABLE:
        print("⚠️  pytesseract not installed ")
        print("    Install with: pip install pytesseract")
        print("    Also need Tesseract executable: https://github.com/UB-Mannheim/tesseract/wiki")
        return True 
    else:
        try:
            pytesseract.get_tesseract_version()
            print("✅ Tesseract OCR is fully functional")
        except Exception as e:
            print(f"⚠️  Tesseract executable not found: {e}")
            print("    Install from: https://github.com/UB-Mannheim/tesseract/wiki")
    
    return True

def main():
    handler = ContextMenuHandler()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        
        if command == "install":
            if not check_dependencies():
                input("Press Enter to exit...")
                return           
            handler.install()
            
        elif command == "uninstall":
            handler.uninstall()
            
        elif command == "ocr":
            if not check_dependencies():
                return
            source_file_path = sys.argv[2] if len(sys.argv) > 2 else None
            handler.start_ocr(source_file_path)         
        
        elif command == "prefix":
            if len(sys.argv) > 3:
                prefix_text = sys.argv[2]
                file_path = sys.argv[3]
                handler.add_prefix(prefix_text, file_path)
                
        elif command == "suffix":
            if len(sys.argv) > 3:
                suffix_text = sys.argv[2]
                file_path = sys.argv[3]
                handler.add_suffix(suffix_text, file_path)
                                      
        else:
            print("Usage:")
            print("  - Install:         python presuffix.py install")
            print("  - Uninstall:       python presuffix.py uninstall")
            input("Press Enter to exit...")
    else:
        print("Prefix-Suffix + Tesseract OCR renamer Context Menu Tool")
        print("=======================================================")
        print("Features:")
        print("\nUsage:")
        print("  - Install:         python presuffix.py install")
        print("  - Uninstall:       python presuffix.py uninstall")
        print("\nRequired packages:")
        print("  pip install pillow pytesseract")
        print("\nTesseract OCR Engine:")
        print("  Download: https://github.com/UB-Mannheim/tesseract/wiki")
        
        # Show current status
        if TESSERACT_AVAILABLE:
            try:
                version = pytesseract.get_tesseract_version()
                print(f"✅ Tesseract Status: Ready (version {version})")
            except Exception as e:
                print(f"⚠️  Tesseract Status: Package installed, executable not found")
                print(f"   Error: {e}")
        else:
            print("❌ Tesseract Status: pytesseract package not installed")
        
        input("Press Enter to exit...")
        
if __name__ == "__main__":
    main()