import customtkinter as ctk
from tkinter import filedialog, messagebox, Toplevel, _tkinter
from PIL import Image, ImageTk, ImageGrab, PngImagePlugin, ExifTags, TiffImagePlugin
import numpy as np
import os
import threading
import queue
from pathlib import Path
from tkinterdnd2 import DND_FILES, TkinterDnD
import json
import traceback
import io

try:
    import tifffile
except ImportError:
    messagebox.showerror("Dependency Missing", "The 'tifffile' library is required. Please run: pip install tifffile")
    exit()

try:
    import cv2
except ImportError:
    messagebox.showerror("Dependency Missing", "The 'opencv-python' library is required. Please run: pip install opencv-python-headless")
    exit()

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIF_SUPPORT = True
except ImportError:
    HEIF_SUPPORT = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        if self.tooltip_window or not self.text: return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = ctk.CTkLabel(self.tooltip_window, text=self.text, corner_radius=5,
                             fg_color=("#333333", "#444444"), text_color="white",
                             wraplength=300, justify="left", font=ctk.CTkFont(size=12))
        label.pack(ipadx=5, ipady=3)

    def hide_tooltip(self, event):
        if self.tooltip_window: self.tooltip_window.destroy()
        self.tooltip_window = None

class AdvancedExportDialog(ctk.CTkToplevel):
    def __init__(self, master, original_info):
        super().__init__(master)
        self.transient(master)
        self.title("Advanced Export")
        self.geometry("550x720")
        self.resizable(False, False)
        self.result = None
        self.original_info = original_info
        self.original_aspect_ratio = original_info.get('size', (1,1))[0] / original_info.get('size', (1,1))[1] if original_info.get('size', (1,1))[1] != 0 else 1
        self.icc_profiles = self.scan_for_icc_profiles()
        self.setup_ui()
        self.after(10, self.update_ui_for_format)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.grab_set()
        self.master.wait_window(self)

    def scan_for_icc_profiles(self):
        profile_paths = {}
        system_folder = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32" / "spool" / "drivers" / "color"
        if system_folder.is_dir():
            for ext in ("*.icc", "*.icm"):
                for profile_file in system_folder.glob(ext):
                    profile_paths[profile_file.stem] = str(profile_file)
        if not profile_paths:
            profile_paths["sRGB (fallback)"] = None
        return profile_paths

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        basic_format_frame = ctk.CTkFrame(self)
        basic_format_frame.grid(row=0, column=0, padx=15, pady=10, sticky="ew")
        basic_format_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(basic_format_frame, text="Format Options", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        ctk.CTkLabel(basic_format_frame, text="File Format:", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        file_formats = [".png", ".tiff", ".jpeg", ".webp", ".bmp"]
        if HEIF_SUPPORT: file_formats.append(".heic")
        self.format_var = ctk.StringVar(value=".png")
        self.format_menu = ctk.CTkOptionMenu(basic_format_frame, variable=self.format_var, values=file_formats, command=self.update_ui_for_format)
        self.format_menu.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(basic_format_frame, text="Bit Depth:", anchor="w").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.bit_depth_var = ctk.StringVar(value=f"{self.original_info.get('bit_depth', 8)}-bit")
        self.bit_depth_menu = ctk.CTkOptionMenu(basic_format_frame, variable=self.bit_depth_var, values=["8-bit", "16-bit"])
        self.bit_depth_menu.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.specific_options_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.specific_options_frame.grid(row=1, column=0, padx=15, pady=0, sticky="new")
        self.specific_options_frame.grid_columnconfigure(0, weight=1)
        self.quality_label = ctk.CTkLabel(self.specific_options_frame, text="Quality (95%):", anchor="w")
        self.quality_slider = ctk.CTkSlider(self.specific_options_frame, from_=0, to=100, command=lambda v: self.quality_label.configure(text=f"Quality ({int(v)}%):"))
        self.quality_slider.set(95)
        self.subsampling_label = ctk.CTkLabel(self.specific_options_frame, text="Chroma Subsampling:", anchor="w")
        self.subsampling_var = ctk.StringVar(value="4:4:4 (Best)")
        self.subsampling_menu = ctk.CTkOptionMenu(self.specific_options_frame, variable=self.subsampling_var, values=["4:4:4 (Best)", "4:2:2 (High)", "4:2:0 (Standard)"])
        dims_frame = ctk.CTkFrame(self)
        dims_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        dims_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(dims_frame, text="Dimensions", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky="w")
        ctk.CTkLabel(dims_frame, text="Resolution (WxH):", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        res_frame = ctk.CTkFrame(dims_frame, fg_color="transparent")
        res_frame.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        w, h = self.original_info.get('size', (0,0))
        self.width_var = ctk.StringVar(value=str(w))
        self.height_var = ctk.StringVar(value=str(h))
        self.width_var.trace_add("write", self.on_width_change)
        self.height_var.trace_add("write", self.on_height_change)
        self.width_entry = ctk.CTkEntry(res_frame, textvariable=self.width_var)
        self.width_entry.pack(side="left", expand=True, fill="x", padx=(0,5))
        self.height_entry = ctk.CTkEntry(res_frame, textvariable=self.height_var)
        self.height_entry.pack(side="left", expand=True, fill="x")
        self.aspect_lock_var = ctk.BooleanVar(value=True)
        self.aspect_lock_button = ctk.CTkCheckBox(dims_frame, text="🔒", variable=self.aspect_lock_var, checkbox_width=20, checkbox_height=20)
        self.aspect_lock_button.grid(row=1, column=2, padx=5)
        ctk.CTkLabel(dims_frame, text="DPI:", anchor="w").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.dpi_entry = ctk.CTkEntry(dims_frame, placeholder_text="e.g., 300")
        if self.original_info.get('dpi'):
            dpi_val = self.original_info['dpi']
            dpi_to_set = dpi_val[0] if isinstance(dpi_val, (tuple, list)) else dpi_val
            self.dpi_entry.insert(0, str(int(dpi_to_set)))
        self.dpi_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        meta_frame = ctk.CTkFrame(self)
        meta_frame.grid(row=3, column=0, padx=15, pady=10, sticky="ew")
        meta_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(meta_frame, text="Color & Metadata", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        ctk.CTkLabel(meta_frame, text="Color Profile:", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.color_space_var = ctk.StringVar(value="sRGB Color Space Profile")
        self.color_space_menu = ctk.CTkComboBox(meta_frame, variable=self.color_space_var, values=list(self.icc_profiles.keys()))
        self.color_space_menu.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.alpha_var = ctk.BooleanVar(value=True)
        self.alpha_check = ctk.CTkCheckBox(meta_frame, text="Preserve Transparency (Alpha Channel)", variable=self.alpha_var)
        self.alpha_check.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.strip_metadata_var = ctk.BooleanVar(value=False)
        self.strip_metadata_check = ctk.CTkCheckBox(meta_frame, text="Strip all metadata (EXIF, etc.)", variable=self.strip_metadata_var)
        self.strip_metadata_check.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        preset_frame = ctk.CTkFrame(self)
        preset_frame.grid(row=4, column=0, padx=15, pady=10, sticky="ew")
        ctk.CTkButton(preset_frame, text="Save Preset", command=self.save_preset).pack(side="left", expand=True, padx=5)
        ctk.CTkButton(preset_frame, text="Load Preset", command=self.load_preset).pack(side="left", expand=True, padx=5)
        ctk.CTkButton(self, text="Export", command=self.on_export, height=40, font=ctk.CTkFont(weight="bold")).grid(row=5, column=0, padx=15, pady=(10, 15), sticky="ew")
        
    def update_ui_for_format(self, selected_format=None):
        widgets_to_forget = [self.quality_label, self.quality_slider, self.subsampling_label, self.subsampling_menu]
        for widget in widgets_to_forget:
            try: widget.grid_forget()
            except _tkinter.TclError: pass 
        self.specific_options_frame.configure(fg_color="transparent")
        fmt = self.format_var.get()
        row = 0
        if fmt in [".jpeg", ".webp", ".heic"]:
            self.quality_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            self.quality_slider.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
            row += 1
        if fmt in [".jpeg", ".heic"]:
            self.subsampling_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            self.subsampling_menu.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
            row += 1
        if row > 0:
            self.specific_options_frame.configure(fg_color=("gray92", "gray14"), corner_radius=5)
            self.specific_options_frame.grid_columnconfigure(1, weight=1)
        is_16bit_supported = fmt in [".png", ".tiff"]
        if not is_16bit_supported and self.bit_depth_var.get() == "16-bit": self.bit_depth_var.set("8-bit")
        self.bit_depth_menu.configure(values=["8-bit", "16-bit"] if is_16bit_supported else ["8-bit"])
        has_alpha_support = fmt in [".png", ".tiff", ".webp", ".heic"]
        self.alpha_check.configure(state="normal" if has_alpha_support else "disabled")
        if not has_alpha_support: self.alpha_var.set(False)

    def on_width_change(self, *args):
        if self.aspect_lock_var.get() and self.width_entry.focus_get() == self.width_entry:
            try:
                new_width = int(self.width_var.get())
                new_height = int(new_width / self.original_aspect_ratio)
                self.height_var.set(str(new_height))
            except (ValueError, ZeroDivisionError): pass
    
    def on_height_change(self, *args):
        if self.aspect_lock_var.get() and self.height_entry.focus_get() == self.height_entry:
            try:
                new_height = int(self.height_var.get())
                new_width = int(new_height * self.original_aspect_ratio)
                self.width_var.set(str(new_width))
            except (ValueError, ZeroDivisionError): pass

    def get_settings(self):
        w = int(self.width_var.get()) if self.width_var.get().isdigit() else self.original_info['size'][0]
        h = int(self.height_var.get()) if self.height_var.get().isdigit() else self.original_info['size'][1]
        dpi = int(self.dpi_entry.get()) if self.dpi_entry.get().isdigit() else None
        profile_path = self.icc_profiles.get(self.color_space_var.get())
        fmt = self.format_var.get()
        settings = {"format": fmt, "bit_depth": int(self.bit_depth_var.get().replace('-bit', '')), "size": (w, h), "dpi": dpi, "icc_profile_path": profile_path, "preserve_alpha": self.alpha_var.get(), "strip_metadata": self.strip_metadata_var.get()}
        if fmt in [".jpeg", ".webp", ".heic"]: settings['quality'] = int(self.quality_slider.get())
        if fmt in [".jpeg", ".heic"]:
            subsampling_map = {"4:4:4 (Best)": 0, "4:2:2 (High)": 1, "4:2:0 (Standard)": 2}
            settings['subsampling'] = subsampling_map.get(self.subsampling_var.get(), 0)
        return settings

    def set_settings(self, settings):
        self.format_var.set(settings.get("format", ".png"))
        self.bit_depth_var.set(f'{settings.get("bit_depth", 8)}-bit')
        self.quality_slider.set(settings.get("quality", 95))
        subsampling_rev_map = {0: "4:4:4 (Best)", 1: "4:2:2 (High)", 2: "4:2:0 (Standard)"}
        self.subsampling_var.set(subsampling_rev_map.get(settings.get("subsampling", 0)))
        w, h = settings.get("size", self.original_info['size'])
        self.width_var.set(str(w))
        self.height_var.set(str(h))
        self.dpi_entry.delete(0, "end")
        if settings.get("dpi"): self.dpi_entry.insert(0, str(settings.get("dpi")))
        profile_name_to_set = "sRGB Color Space Profile"
        if settings.get("icc_profile_path"):
            for name, path in self.icc_profiles.items():
                if path == settings["icc_profile_path"]:
                    profile_name_to_set = name; break
        self.color_space_var.set(profile_name_to_set)
        self.alpha_var.set(settings.get("preserve_alpha", True))
        self.strip_metadata_var.set(settings.get("strip_metadata", False))
        self.update_ui_for_format()

    def save_preset(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Preset", "*.json")])
        if filepath:
            with open(filepath, 'w') as f: json.dump(self.get_settings(), f, indent=4)
            messagebox.showinfo("Success", "Preset saved.", parent=self)

    def load_preset(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Preset", "*.json")])
        if filepath:
            with open(filepath, 'r') as f: settings = json.load(f)
            self.set_settings(settings)
            messagebox.showinfo("Success", "Preset loaded.", parent=self)
            
    def on_export(self):
        self.result = self.get_settings()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

class MainApplication(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.original_image = None
        self.preview_data = None
        self.original_info = {}
        self.batch_files = []
        self.task_queue = queue.Queue()
        self.result_queue = queue.Queue()
        threading.Thread(target=self.worker_loop, daemon=True).start()
        self.pack(fill="both", expand=True)
        self.setup_ui()
        self.process_results()
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind('<<Drop>>', self.on_drop)
        
    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.tab_view = ctk.CTkTabview(self, anchor="w")
        self.tab_view.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        single_tab = self.tab_view.add("Single Conversion")
        batch_tab = self.tab_view.add("Batch Processing")
        self.setup_single_conversion_tab(single_tab)
        self.setup_batch_processing_tab(batch_tab)
    
    def setup_single_conversion_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)
        control_frame = ctk.CTkFrame(tab, fg_color="transparent")
        control_frame.grid(row=0, column=0, pady=(0, 10), sticky="ew")
        file_ops_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        file_ops_frame.pack(side="left", padx=(0, 20))
        ctk.CTkButton(file_ops_frame, text="📁 Load Image", command=self.load_image).pack(side="left", padx=(0, 10))
        ctk.CTkButton(file_ops_frame, text="📋 From Clipboard", command=self.load_from_clipboard).pack(side="left", padx=(0, 10))
        self.export_button = ctk.CTkButton(file_ops_frame, text="💾 Export...", command=self.export_image, state="disabled")
        self.export_button.pack(side="left")
        mode_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        mode_frame.pack(side="left")
        ctk.CTkLabel(mode_frame, text="Conversion Mode:").pack(side="left", padx=(0, 10))
        self.conversion_mode_var = ctk.StringVar(value="Rec. 709")
        modes = ["L*a*b* (L*)", "Gamma", "Rec. 709", "HSL (Lightness)", "HSV (Value)", "Rec. 601", "Rec. 2100"]
        self.mode_menu = ctk.CTkSegmentedButton(mode_frame, values=modes, variable=self.conversion_mode_var, command=self.update_preview)
        self.mode_menu.pack(side="left")
        info_frame = ctk.CTkFrame(tab, fg_color="transparent")
        info_frame.grid(row=1, column=0, pady=(0, 10), sticky="ew")
        self.info_var = ctk.StringVar(value="Image Info: No image loaded.")
        ctk.CTkLabel(info_frame, textvariable=self.info_var, font=ctk.CTkFont(size=12), text_color="#FFD700").pack(side="left")
        display_frame = ctk.CTkFrame(tab, fg_color="transparent")
        display_frame.grid(row=2, column=0, sticky="nsew")
        display_frame.grid_columnconfigure((0, 1), weight=1)
        display_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(display_frame, text="Original Image", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, pady=(0, 10))
        self.original_canvas = ctk.CTkCanvas(display_frame, bg="#1e1e1e", highlightthickness=0)
        self.original_canvas.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        ctk.CTkLabel(display_frame, text="Enhanced Grayscale Preview", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=1, pady=(0, 10))
        self.preview_canvas = ctk.CTkCanvas(display_frame, bg="#1e1e1e", highlightthickness=0)
        self.preview_canvas.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        status_frame = ctk.CTkFrame(tab, fg_color="transparent")
        status_frame.grid(row=3, column=0, pady=(10, 0), sticky="ew")
        status_frame.grid_columnconfigure(0, weight=1)
        self.status_var = ctk.StringVar(value="Ready")
        ctk.CTkLabel(status_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress_bar = ctk.CTkProgressBar(status_frame, mode='indeterminate')
        self.original_canvas.bind("<Configure>", lambda e: self.request_display_update('original'))
        self.preview_canvas.bind("<Configure>", lambda e: self.request_display_update('preview'))
    
    def setup_batch_processing_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        batch_control_frame = ctk.CTkFrame(tab)
        batch_control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(batch_control_frame, text="Add Files...", command=self.add_batch_files).pack(side="left", padx=5)
        ctk.CTkButton(batch_control_frame, text="Clear List", command=self.clear_batch_list).pack(side="left", padx=5)
        self.batch_list_frame = ctk.CTkScrollableFrame(tab, label_text="Drag & Drop Files Here")
        self.batch_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        batch_action_frame = ctk.CTkFrame(tab)
        batch_action_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        batch_action_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(batch_action_frame, text="Output Folder:").pack(side="left", padx=5)
        self.output_folder_var = ctk.StringVar(value="")
        ctk.CTkEntry(batch_action_frame, textvariable=self.output_folder_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkButton(batch_action_frame, text="Browse...", command=self.select_output_folder).pack(side="left", padx=5)
        self.batch_export_button = ctk.CTkButton(tab, text="Start Batch Processing", command=self.start_batch_processing, height=40, state="disabled")
        self.batch_export_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        self.batch_progress = ctk.CTkProgressBar(tab)
        self.batch_progress.set(0)

    def worker_loop(self):
        while True:
            task_type, data = self.task_queue.get()
            try:
                if task_type == 'load': self.result_queue.put(('load_success', self._perform_load(data)))
                elif task_type == 'convert': self.result_queue.put(('convert_success', self.convert_to_enhanced_grayscale(*data)))
                elif task_type == 'resize_display': self.result_queue.put(('display_ready', (data[0], self._perform_resize_for_display(*data))))
                elif task_type == 'save':
                    self._perform_save(*data)
                    self.result_queue.put(('save_success', data[2]))
                elif task_type == 'batch_process':
                    in_path, out_path, settings = data
                    img_obj, info = self._perform_load(in_path)
                    gray_array, alpha_img = self.convert_to_enhanced_grayscale(img_obj, settings['conversion_mode'], settings.get('bit_depth', info['bit_depth']))
                    self._perform_save(gray_array, alpha_img, out_path, settings, info)
                    self.result_queue.put(('batch_item_success', in_path))
            except Exception as e:
                self.result_queue.put(('task_failed', (data, traceback.format_exc(), e)))

    def process_results(self):
        try:
            while not self.result_queue.empty():
                result_type, data = self.result_queue.get_nowait()
                if result_type == 'load_success': 
                    self.original_image, self.original_info = data
                    self._handle_load_success()
                elif result_type == 'convert_success':
                    self.preview_data, _ = data
                    self._handle_convert_success()
                elif result_type == 'display_ready': 
                    self._update_canvas_image(*data)
                elif result_type == 'save_success':
                    self.stop_processing_indicator(f"Successfully exported: {os.path.basename(data)}")
                    messagebox.showinfo("Success", f"Image saved successfully to:\n{data}")
                elif result_type == 'batch_item_success':
                    self._update_batch_item_status(data, "✅ Done", "green")
                    self.batch_progress.set(self.batch_progress.get() + (1/len(self.batch_files)))
                elif result_type == 'task_failed':
                    task_data, error_msg, exception = data
                    print(f"Task failed.\nData: {task_data}\nError: {error_msg}")
                    self.stop_processing_indicator("An error occurred.", "red")
                    messagebox.showerror("Processing Error", f"Task failed:\n{exception}")
        finally: 
            self.after(100, self.process_results)

    def to_linear(self, c):
        return np.where(c <= 0.04045, c / 12.92, ((c + 0.055) / 1.055) ** 2.4)

    def to_srgb(self, c):
        return np.where(c <= 0.0031308, c * 12.92, 1.055 * (c ** (1/2.4)) - 0.055)

    def convert_to_enhanced_grayscale(self, image: Image.Image, mode: str, target_bit_depth: int):
        if image.mode == 'RGBA': cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGBA2BGRA)
        elif image.mode == 'RGB': cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        elif image.mode == 'LA':
            L, A_pil = image.split()
            cv_image = cv2.cvtColor(np.array(L), cv2.COLOR_GRAY2BGRA)
            cv_image[:, :, 3] = np.array(A_pil)
        elif image.mode == 'L': cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_GRAY2BGR)
        else:
            rgba_image = image.convert("RGBA")
            cv_image = cv2.cvtColor(np.array(rgba_image), cv2.COLOR_RGBA2BGRA)
        has_alpha = cv_image.shape[2] == 4
        alpha_channel_pil = Image.fromarray(cv_image[:, :, 3]) if has_alpha else None
        if cv_image.dtype == np.uint8: cv_image = cv_image.astype(np.uint16) * 257
        elif cv_image.dtype != np.uint16: raise ValueError(f"Unsupported image dtype {cv_image.dtype}")
        B, G, R = cv2.split(cv_image)[:3]
        Rf, Gf, Bf = R.astype(np.float64)/65535.0, G.astype(np.float64)/65535.0, B.astype(np.float64)/65535.0
        mode_map = {"Rec. 601": '601', "Rec. 709": '709', "Rec. 2100": '2100', "Gamma": 'gamma'}
        script_mode = mode_map.get(mode, '709')
        if mode in ["L*a*b* (L*)", "HSV (Value)", "HSL (Lightness)"]:
            rgb_image_pil = image.convert('RGB')
            if mode == "L*a*b* (L*)":
                l, _, _ = rgb_image_pil.convert('LAB').split()
                gray_float = np.array(l, dtype=np.float64) / 255.0
            elif mode == "HSV (Value)":
                _, _, v = rgb_image_pil.convert('HSV').split()
                gray_float = np.array(v, dtype=np.float64) / 255.0
            else:
                rgb_array_float = np.array(rgb_image_pil, dtype=np.float64) / 255.0
                cmax, cmin = np.maximum.reduce(rgb_array_float, axis=-1), np.minimum.reduce(rgb_array_float, axis=-1)
                gray_float = (cmax + cmin) / 2.0
        elif script_mode == 'gamma':
            Rl, Gl, Bl = self.to_linear(Rf), self.to_linear(Gf), self.to_linear(Bf)
            Yl = 0.2126 * Rl + 0.7152 * Gl + 0.0722 * Bl
            gray_float = self.to_srgb(Yl)
        else:
            weights = {'601':(0.299,0.587,0.114),'709':(0.2126,0.7152,0.0722),'2100':(0.2627,0.6780,0.0593)}
            wR, wG, wB = weights[script_mode]
            gray_float = wR * Rf + wG * Gf + wB * Bf
        gray_float = np.clip(gray_float, 0, 1)
        if target_bit_depth == 16: multiplier, dtype = 65535, np.uint16
        else: multiplier, dtype = 255, np.uint8
        final_array = np.round(gray_float * multiplier).astype(dtype)
        return final_array, alpha_channel_pil

    def _perform_save(self, gray_array, alpha_image, filepath, settings, original_info):
        file_ext = Path(filepath).suffix.lower()
        is_high_bit_depth = settings["bit_depth"] > 8
        has_alpha = alpha_image and settings["preserve_alpha"]
        if file_ext == ".tiff" and is_high_bit_depth and has_alpha:
            alpha_8bit_np = np.array(alpha_image.convert("L"))
            A16 = (alpha_8bit_np.astype(np.uint16)) * 257
            stacked = np.stack([gray_array, A16], axis=-1)
            tifffile.imwrite(filepath, stacked, photometric="minisblack", extrasamples=["unassalpha"])
            return
        if file_ext == ".png" and is_high_bit_depth and has_alpha:
            Y16 = gray_array
            alpha_8bit_np = np.array(alpha_image.convert("L"))
            A16 = (alpha_8bit_np.astype(np.uint16)) * 257
            out_cv = cv2.merge([Y16, Y16, Y16, A16])
            success, buffer = cv2.imencode(file_ext, out_cv)
            if not success: raise IOError("Failed to encode 16-bit PNG with alpha.")
            with open(filepath, 'wb') as f: f.write(buffer)
            return
        target_mode = "I;16" if is_high_bit_depth else "L"
        final_image = Image.fromarray(gray_array, mode=target_mode)
        if settings.get("size") != original_info['size']:
            final_image = final_image.resize(settings["size"], Image.Resampling.LANCZOS)
        if has_alpha:
            final_image = Image.merge("LA", (final_image.convert("L"), alpha_image.convert("L")))
        save_kwargs = {}
        if not settings.get("strip_metadata", False):
            icc_profile_path = settings.get("icc_profile_path")
            if icc_profile_path:
                with open(icc_profile_path, 'rb') as f: save_kwargs['icc_profile'] = f.read()
            elif original_info.get("icc_profile"): save_kwargs['icc_profile'] = original_info.get("icc_profile")
            dpi = settings.get("dpi")
            if dpi: save_kwargs['dpi'] = (dpi, dpi)
        format_map = {".jpeg": "JPEG", ".jpg": "JPEG", ".png": "PNG", ".tiff": "TIFF", ".webp": "WEBP", ".bmp": "BMP", ".heic": "HEIF"}
        file_format = format_map.get(file_ext, "PNG")
        if file_ext in [".jpg", ".jpeg"]:
            final_image = final_image.convert("L")
            save_kwargs.update({"quality": settings.get("quality", 95), "subsampling": settings.get("subsampling", 0)})
        elif file_ext == ".heic":
             save_kwargs.update({"quality": settings.get("quality", 95), "chroma": settings.get("subsampling", 0)})
        final_image.save(filepath, format=file_format, **save_kwargs)

    def on_drop(self, event):
        paths = self.master.tk.splitlist(event.data)
        if self.tab_view.get() == "Single Conversion":
            if paths: self.load_image(paths[0])
        else: self.add_batch_files(paths)

    def load_image(self, filepath=None):
        if not filepath:
            filetypes = [("All Images", "*.*")]
            if HEIF_SUPPORT: filetypes.insert(0, ("HEIF/HEIC", "*.heic"))
            filepath = filedialog.askopenfilename(filetypes=filetypes)
        if filepath:
            self.start_processing_indicator(f"Loading: {os.path.basename(filepath)}...")
            self.task_queue.put(('load', filepath))
            
    def load_from_clipboard(self):
        try:
            image = ImageGrab.grabclipboard()
            if isinstance(image, Image.Image):
                 self.start_processing_indicator("Loading from clipboard...")
                 self.task_queue.put(('load', image))
            else: messagebox.showwarning("Clipboard Error", "No valid image on clipboard.")
        except: messagebox.showerror("Clipboard Error", "Could not access clipboard.")

    def export_image(self):
        if self.original_image is None: messagebox.showwarning("Warning", "No image to export."); return
        dialog = AdvancedExportDialog(self, self.original_info)
        settings = dialog.result
        if not settings: return
        default_name = Path(self.original_info.get('filepath', 'image.png')).stem + "_grayscale"
        file_ext = settings['format']
        filepath = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=file_ext, filetypes=[(f"{file_ext.upper()[1:]} files", f"*{file_ext}")])
        if not filepath: return
        self.start_processing_indicator("Exporting image...")
        gray_array, alpha_img = self.convert_to_enhanced_grayscale(self.original_image, self.conversion_mode_var.get(), settings['bit_depth'])
        self.task_queue.put(('save', (gray_array, alpha_img, filepath, settings, self.original_info)))

    def _perform_load(self, source):
        pil_image = None
        if isinstance(source, str): pil_image = Image.open(source)
        elif isinstance(source, Image.Image): pil_image = source
        pil_image.load()
        info = self.analyze_image_properties(pil_image)
        return pil_image, info
        
    def add_batch_files(self, filepaths=None):
        if not filepaths: filepaths = filedialog.askopenfilenames(title="Select files for batch processing")
        for path in filepaths:
            if path not in [f['path'] for f in self.batch_files]:
                frame = ctk.CTkFrame(self.batch_list_frame, fg_color=("gray85", "gray28"))
                frame.pack(fill="x", pady=2, padx=2)
                label = ctk.CTkLabel(frame, text=os.path.basename(path), anchor="w")
                label.pack(side="left", expand=True, fill="x", padx=5)
                status_label = ctk.CTkLabel(frame, text="⏳ Pending", width=120, anchor="e")
                status_label.pack(side="right", padx=5)
                self.batch_files.append({"path": path, "frame": frame, "status_label": status_label})
        self.batch_export_button.configure(state="normal" if self.batch_files and self.output_folder_var.get() else "disabled")

    def clear_batch_list(self):
        for item in self.batch_files: item['frame'].destroy()
        self.batch_files.clear()
        self.batch_export_button.configure(state="disabled")

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder_var.set(folder)
            self.batch_export_button.configure(state="normal" if self.batch_files else "disabled")

    def start_batch_processing(self):
        output_folder = self.output_folder_var.get()
        if not output_folder or not os.path.isdir(output_folder): messagebox.showerror("Error", "Please select a valid output folder."); return
        dialog = AdvancedExportDialog(self, self.original_info or {'bit_depth': 8, 'size': (0,0), 'dpi': (72,72)})
        export_settings = dialog.result
        if not export_settings: return
        export_settings['conversion_mode'] = self.conversion_mode_var.get()
        self.batch_progress.grid(row=4, column=0, padx=10, pady=(0,10), sticky="ew")
        self.batch_progress.set(0)
        for item in self.batch_files:
            self._update_batch_item_status(item['path'], "Queued", "#cccccc")
            in_path = item['path']
            out_name = Path(in_path).stem + "_grayscale" + export_settings['format']
            out_path = os.path.join(output_folder, out_name)
            self.task_queue.put(('batch_process', (in_path, out_path, export_settings)))
    
    def _update_batch_item_status(self, path, text, color):
        for item in self.batch_files:
            if item['path'] == path: item['status_label'].configure(text=text, text_color=color); return

    def _perform_resize_for_display(self, canvas, image: Image.Image):
        # CORRECTED FUNCTION
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w <= 1 or h <= 1: return None
        
        # This function should NOT do any color conversion. It just resizes.
        # It needs to handle color images for the 'original' canvas.
        display_image = image
        if display_image.mode not in ['RGB', 'RGBA', 'L', 'LA']:
            display_image = image.convert('RGBA')

        img_w, img_h = display_image.size
        if img_w == 0 or img_h == 0: return None
        scale = min(w / img_w, h / img_h, 1.0)
        new_size = (int(img_w * scale), int(img_h * scale))
        
        if new_size[0] < 1 or new_size[1] < 1: return None
            
        resized_image = display_image.resize(new_size, Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(resized_image)
        
    def update_preview(self, _=None):
        if self.original_image:
            self.start_processing_indicator("Applying grayscale effect...")
            self.task_queue.put(('convert', (self.original_image, self.conversion_mode_var.get(), self.original_info.get('bit_depth', 8))))
            
    def _handle_load_success(self):
        self.stop_processing_indicator(f"Loaded: {os.path.basename(self.original_info.get('filepath', 'clipboard'))}")
        self.info_var.set(self.original_info['display_text'])
        self.export_button.configure(state="disabled")
        self.request_display_update('original') # THIS IS KEY - always show original first
        self.update_preview()

    def _handle_convert_success(self):
        self.stop_processing_indicator("Enhanced preview ready.")
        self.export_button.configure(state="normal")
        self.request_display_update('preview') # THEN update the preview

    def request_display_update(self, canvas_name):
        # CORRECTED FUNCTION
        image_to_display = None
        if canvas_name == 'original' and self.original_image:
            # For 'original', always use the original color image
            image_to_display = self.original_image
        elif canvas_name == 'preview' and self.preview_data is not None:
            # For 'preview', use the grayscale numpy array
            mode = 'I;16' if self.preview_data.dtype == np.uint16 else 'L'
            image_to_display = Image.fromarray(self.preview_data, mode=mode)

        canvas = self.original_canvas if canvas_name == 'original' else self.preview_canvas
        if image_to_display:
            self.task_queue.put(('resize_display', (canvas, image_to_display)))
            
    def start_processing_indicator(self, message):
        self.status_var.set(message)
        self.progress_bar.grid(row=0, column=1, sticky="e", pady=(5, 0))
        self.progress_bar.start()

    def stop_processing_indicator(self, message, color=None):
        self.status_var.set(message)
        self.progress_bar.stop()
        self.progress_bar.grid_forget()

    def analyze_image_properties(self, image):
        info = {'filepath': getattr(image, 'filename', 'clipboard'), 'size': image.size, 'mode': image.mode}
        info['exif'] = image.info.get('exif')
        info['icc_profile'] = image.info.get('icc_profile')
        info['dpi'] = image.info.get('dpi')
        try:
            dtype_str = str(np.array(image).dtype)
            if '16' in dtype_str: info['bit_depth'] = 16
            elif '32' in dtype_str: info['bit_depth'] = 32
            else: info['bit_depth'] = 8
        except Exception:
            if image.mode in ('I;16', 'I;16B', 'I;16L', 'I;16N', 'I;16LA') or (image.mode == 'LA' and image.getextrema()[0][1] > 255): info['bit_depth'] = 16
            elif image.mode in ('I', 'F'): info['bit_depth'] = 32
            else: info['bit_depth'] = 8
        info['display_text'] = f"Size: {image.size[0]}×{image.size[1]} | Mode: {image.mode} | Bit Depth: {info['bit_depth']}-bit | {'ICC' if info['icc_profile'] else 'No ICC'}"
        return info

    def _update_canvas_image(self, canvas, photo_image):
        if not photo_image: return
        canvas.delete("all")
        canvas.photo = photo_image
        x = (canvas.winfo_width() - photo_image.width()) // 2
        y = (canvas.winfo_height() - photo_image.height()) // 2
        canvas.create_image(x, y, anchor='nw', image=photo_image)

if __name__ == "__main__":
    if not HEIF_SUPPORT: print("WARNING: HEIF/HEIC support is not available. Please run 'pip install pillow-heif'.")
    root = TkinterDnD.Tk()
    root.withdraw()
    app_window = ctk.CTkToplevel()
    app_window.title("Enhanced Precision Grayscale Converter")
    app_window.geometry("1400x900")
    app_window.minsize(1200, 800)
    app = MainApplication(master=app_window)
    app_window.protocol("WM_DELETE_WINDOW", root.destroy)
    app_window.mainloop()