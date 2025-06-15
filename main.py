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
    
# Keep imageio for other formats if needed, or rely on Pillow
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
        self.geometry("550x680")
        self.resizable(False, False)
        self.result = None
        self._is_updating_dims = False
        
        self.original_info = original_info
        self.original_aspect_ratio = original_info.get('size', (1,1))[0] / original_info.get('size', (1,1))[1] if original_info.get('size', (1,1))[1] != 0 else 1
        self.icc_profiles = self.scan_for_icc_profiles()
        
        self.setup_ui()
        self.after(10, self.update_ui_for_format)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.grab_set()
        self.master.wait_window(self)

    def scan_for_icc_profiles(self):
        profile_paths = {"sRGB (Built-in)": None}
        system_folder = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32" / "spool" / "drivers" / "color"
        if system_folder.is_dir():
            for ext in ("*.icc", "*.icm"):
                for profile_file in system_folder.glob(ext):
                    profile_paths[profile_file.stem] = str(profile_file)
        return profile_paths

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        basic_format_frame = ctk.CTkFrame(self)
        basic_format_frame.grid(row=0, column=0, padx=15, pady=10, sticky="ew")
        basic_format_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(basic_format_frame, text="Format Options", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        
        ctk.CTkLabel(basic_format_frame, text="File Format:", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        file_formats = [".png", ".tiff", ".jpeg", ".webp", ".bmp"]
        if HEIF_SUPPORT: file_formats.insert(2, ".heic")
        self.format_var = ctk.StringVar(value=".png")
        self.format_menu = ctk.CTkOptionMenu(basic_format_frame, variable=self.format_var, values=file_formats, command=self.update_ui_for_format)
        self.format_menu.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(basic_format_frame, text="Bit Depth:", anchor="w").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.bit_depth_var = ctk.StringVar(value="8-bit")
        self.bit_depth_menu = ctk.CTkOptionMenu(basic_format_frame, variable=self.bit_depth_var, values=["8-bit", "10-bit", "16-bit"], command=self.update_ui_for_format)
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
        self.color_range_label = ctk.CTkLabel(self.specific_options_frame, text="Color Range:", anchor="w")
        self.color_range_var = ctk.StringVar(value="Preserve Original")
        self.color_range_menu = ctk.CTkOptionMenu(self.specific_options_frame, variable=self.color_range_var, values=["Preserve Original", "Full", "Limited"])

        dims_frame = ctk.CTkFrame(self)
        dims_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        dims_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(dims_frame, text="Dimensions & Resolution", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=3, pady=(5, 10), sticky="w")
        ctk.CTkLabel(dims_frame, text="Size (WxH):", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        res_frame = ctk.CTkFrame(dims_frame, fg_color="transparent")
        res_frame.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        w, h = self.original_info.get('size', (0,0))
        self.width_var = ctk.StringVar(value=str(w))
        self.height_var = ctk.StringVar(value=str(h))
        self.width_var.trace_add("write", self.on_width_change)
        self.height_var.trace_add("write", self.on_height_change)
        self.width_entry = ctk.CTkEntry(res_frame, textvariable=self.width_var, width=80)
        self.width_entry.pack(side="left", expand=True, fill="x", padx=(0,5))
        self.height_entry = ctk.CTkEntry(res_frame, textvariable=self.height_var, width=80)
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
        ctk.CTkLabel(meta_frame, text="Color & Metadata", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky="w")
        ctk.CTkLabel(meta_frame, text="Color Profile:", anchor="w").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.color_space_var = ctk.StringVar(value="sRGB (Built-in)")
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
        widgets_to_forget = [
            self.quality_label, self.quality_slider,
            self.subsampling_label, self.subsampling_menu,
            self.color_range_label, self.color_range_menu
        ]
        for widget in widgets_to_forget:
            try:
                widget.grid_forget()
            except _tkinter.TclError:
                pass 
        self.specific_options_frame.configure(fg_color="transparent")
        fmt = self.format_var.get()
        caps = self.FORMAT_CAPABILITIES.get(fmt)
        if not caps: return

        self.bit_depth_menu.configure(values=caps["bit_depths"])
        if self.bit_depth_var.get() not in caps["bit_depths"]:
            self.bit_depth_var.set(caps["bit_depths"][0])
        current_bit_depth = self.bit_depth_var.get()
        
        alpha_supported = caps["alpha"](current_bit_depth) if callable(caps["alpha"]) else caps["alpha"]
        self.alpha_check.configure(state="normal" if alpha_supported else "disabled")
        if not alpha_supported:
            self.alpha_var.set(False)

        row = 0
        
        if fmt in [".jpeg", ".webp", ".heic"]:
            self.quality_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            self.quality_slider.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
            row += 1
        
        if fmt == ".jpeg":
            self.subsampling_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            self.subsampling_menu.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
            row += 1

        if fmt in [".jpeg", ".png"]:
            self.color_range_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            self.color_range_menu.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
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
        if self.aspect_lock_var.get() and self.width_entry.focus_get() == self.width_entry and not self._is_updating_dims:
            self._is_updating_dims = True
            try:
                new_width = int(self.width_var.get())
                new_height = int(new_width / self.original_aspect_ratio)
                self.height_var.set(str(new_height))
            except (ValueError, ZeroDivisionError): pass
            self._is_updating_dims = False
    
    def on_height_change(self, *args):
        if self.aspect_lock_var.get() and self.height_entry.focus_get() == self.height_entry and not self._is_updating_dims:
            self._is_updating_dims = True
            try:
                new_height = int(self.height_var.get())
                new_width = int(new_height * self.original_aspect_ratio)
                self.width_var.set(str(new_width))
            except (ValueError, ZeroDivisionError): pass
            self._is_updating_dims = False

    def get_settings(self):
        w = int(self.width_var.get()) if self.width_var.get().isdigit() else self.original_info['size'][0]
        h = int(self.height_var.get()) if self.height_var.get().isdigit() else self.original_info['size'][1]
        dpi = int(self.dpi_entry.get()) if self.dpi_entry.get().isdigit() else None
        profile_path = self.icc_profiles.get(self.color_space_var.get())
        fmt = self.format_var.get()

        settings = {
            "format": fmt,
            "bit_depth": int(self.bit_depth_var.get().replace('-bit', '')),
            "size": (w, h), "dpi": dpi, "icc_profile_path": profile_path,
            "preserve_alpha": self.alpha_var.get(), "strip_metadata": self.strip_metadata_var.get()
        }

        if fmt in [".jpeg", ".webp", ".heic"]:
            settings['quality'] = int(self.quality_slider.get())
        if fmt == ".jpeg":
            subsampling_map = {"4:4:4 (Best)": 0, "4:2:2 (High)": 1, "4:2:0 (Standard)": 2}
            settings['subsampling_val'] = subsampling_map.get(self.subsampling_var.get(), 0)
            settings['subsampling_str'] = self.subsampling_var.get()
        if fmt in [".jpeg", ".png"]:
            settings['color_range'] = self.color_range_var.get()

        return settings

    def set_settings(self, settings):
        self.format_var.set(settings.get("format", ".png"))
        self.update_ui_for_format()
        self.bit_depth_var.set(f'{settings.get("bit_depth", 8)}-bit')
        
        self.quality_slider.set(settings.get("quality", 95))
        
        subsampling_rev_map = {0: "4:4:4 (Best)", 1: "4:2:2 (High)", 2: "4:2:0 (Standard)"}
        self.subsampling_var.set(subsampling_rev_map.get(settings.get("subsampling_val", 0)))
        
        self.color_range_var.set(settings.get("color_range", "Preserve Original"))
        
        w, h = settings.get("size", self.original_info['size'])
        self.width_var.set(str(w)); self.height_var.set(str(h))
        self.dpi_entry.delete(0, "end")
        if settings.get("dpi"): self.dpi_entry.insert(0, str(settings.get("dpi")))
        
        profile_name_to_set = "sRGB Color Space Profile"
        if settings.get("icc_profile_path"):
            for name, path in self.icc_profiles.items():
                if path == settings["icc_profile_path"]:
                    profile_name_to_set = name
                    break
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
            try:
                with open(filepath, 'r') as f: settings = json.load(f)
                self.set_settings(settings)
                messagebox.showinfo("Success", "Preset loaded.", parent=self)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load preset:\n{e}", parent=self)
            
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
        ctk.CTkLabel(info_frame, textvariable=self.info_var, font=ctk.CTkFont(size=12)).pack(side="left")
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
                elif task_type == 'save': self._perform_save(*data); self.result_queue.put(('save_success', data[2]))
                elif task_type == 'batch_process':
                    in_path, out_path, settings = data
                    img_obj, info = self._perform_load(in_path)
                    gray_array, alpha_img = self.convert_to_enhanced_grayscale(img_obj, settings['conversion_mode'], settings['bit_depth'])
                    self._perform_save(gray_array, alpha_img, out_path, settings, info)
                    self.result_queue.put(('batch_item_success', in_path))
            except Exception:
                self.result_queue.put(('task_failed', (data, traceback.format_exc())))
    
    def process_results(self):
        try:
            while not self.result_queue.empty():
                result_type, data = self.result_queue.get_nowait()
                if result_type == 'load_success': self._handle_load_success(data)
                elif result_type == 'convert_success': self._handle_convert_success(data)
                elif result_type == 'display_ready': self._update_canvas_image(*data)
                elif result_type == 'save_success':
                    self.stop_processing_indicator(f"Successfully exported: {os.path.basename(data)}")
                    messagebox.showinfo("Success", f"Image saved successfully to:\n{data}")
                elif result_type == 'batch_item_success':
                    self._update_batch_item_status(data, "✅ Done", "green")
                    self.batch_progress.set(self.batch_progress.get() + (1/len(self.batch_files)))
                elif result_type == 'task_failed':
                    task_data, error_msg, exception = data
                    print(f"Task failed.\nData: {task_data}\nError: {error_msg}")
                    in_path = None
                    if isinstance(task_data, tuple) and len(task_data) > 0:
                        if isinstance(task_data[0], np.ndarray): # save task
                            in_path = task_data[4].get('filepath', 'Unknown File')
                        elif isinstance(task_data[0], str): # batch task
                            in_path = task_data[0]
                    if in_path:
                        self._update_batch_item_status(in_path, "❌ Error", "red")

                    self.stop_processing_indicator("An error occurred.", "red")
                    messagebox.showerror("Processing Error", f"Task failed:\n{exception}")
        finally: 
            self.after(100, self.process_results)

    def convert_to_enhanced_grayscale(self, image, mode, target_bit_depth):
        if image.mode == 'L' and target_bit_depth == 8: return np.array(image), None
        
        has_alpha = 'A' in image.getbands()
        alpha_image_out = image.getchannel('A') if has_alpha else None
        
        if image.mode not in ('RGB', 'RGBA'):
             rgb_image = image.convert('RGBA' if has_alpha else 'RGB')
        else:
             rgb_image = image

        source_is_16bit_depth = np.array(image).dtype == np.uint16
        source_dtype = np.float32

        if source_is_16bit_depth:
            rgb_array = np.array(rgb_image.convert('RGB'), dtype=source_dtype) / 65535.0
        else:
            rgb_array = np.array(rgb_image.convert('RGB'), dtype=source_dtype) / 255.0

        if mode == "L*a*b* (L*)":
            lab_image = rgb_image.convert('RGB').convert('LAB')
            l_channel, _, _ = lab_image.split()
            gray_float = np.array(l_channel, dtype=source_dtype) / 255.0
        elif mode == "HSV (Value)":
            hsv_image = rgb_image.convert('RGB').convert('HSV')
            _, _, v_channel = hsv_image.split()
            gray_float = np.array(v_channel, dtype=source_dtype) / 255.0
        elif mode == "HSL (Lightness)":
            def rgb_to_hsl_lightness(rgb):
                cmax = np.maximum.reduce(rgb, axis=-1)
                cmin = np.minimum.reduce(rgb, axis=-1)
                return (cmax + cmin) / 2.0
            gray_float = rgb_to_hsl_lightness(rgb_array)
        elif mode == "Gamma":
            linear_rgb = np.power(rgb_array, 2.2)
            coeffs = [0.2126, 0.7152, 0.0722] # BT.709
            linear_gray = np.dot(linear_rgb, coeffs)
            gray_float = np.power(linear_gray, 1/2.2)
        else: # BT.XXX standards
            coeffs = {"BT.601": [0.299, 0.587, 0.114], "BT.709": [0.2126, 0.7152, 0.0722], "BT.2100": [0.2627, 0.6780, 0.0593]}.get(mode, [0.2126, 0.7152, 0.0722])
            gray_float = np.dot(rgb_array, coeffs)
        
        gray_float = np.clip(gray_float, 0, 1)
        
        dtype = np.uint16 if target_bit_depth == 16 else np.uint8
        multiplier = 65535 if target_bit_depth == 16 else 255
        final_array = (gray_float * multiplier).astype(dtype)

        return final_array, alpha_image_out
    
    def _perform_save(self, gray_array, alpha_image, filepath, settings, original_info):
        file_ext = Path(filepath).suffix.lower()
        is_16bit = settings['bit_depth'] == 16

        # 1. Create Pillow image object from grayscale numpy array
        target_mode = 'I;16' if is_16bit else 'L'
        final_image = Image.fromarray(gray_array, mode=target_mode)

        # 2. Handle Alpha Channel, ensuring modes match before merging
        has_alpha = alpha_image and settings['preserve_alpha'] and file_ext in ['.png', '.tiff', '.webp', '.heic']

        # --- 16-bit TIFF with alpha: use tifffile and stacked array ---
        if is_16bit and has_alpha and file_ext == '.tiff':
            alpha_pil = alpha_image
            # Convert alpha to 16-bit numpy array
            alpha_arr = np.array(alpha_pil.convert('L'), dtype=np.uint8)
            alpha_arr = (alpha_arr.astype(np.uint32) * 65535 // 255).astype(np.uint16)
            stacked = np.stack([gray_array, alpha_arr], axis=-1)
            # Prepare tifffile kwargs
            tiff_kwargs = {'photometric': 'minisblack', 'extrasamples': ['unassalpha']}
            icc_profile = settings.get('icc_profile_path') or original_info.get('icc_profile')
            dpi_val = settings.get('dpi')
            if dpi_val:
                tiff_kwargs['resolution'] = (dpi_val, dpi_val, 'INCH')
            # Handle ICC profile as extratag (34675)
            extratags = []
            if icc_profile:
                with open(icc_profile, "rb") if isinstance(icc_profile, str) else None as f:
                    icc_bytes = f.read() if f else icc_profile
                extratags.append((34675, 'B', len(icc_bytes), icc_bytes, False))
            if extratags:
                tiff_kwargs['extratags'] = extratags
            tifffile.imwrite(filepath, stacked, **tiff_kwargs)
            return

        # --- 16-bit PNG with alpha: warn and drop alpha ---
        if is_16bit and has_alpha and file_ext == '.png':
            messagebox.showwarning("Alpha Channel Not Supported",
                "16-bit PNG with alpha is not supported. The alpha channel will be discarded.")
            has_alpha = False

        # --- 8-bit with alpha: use Image.merge('LA', ...) ---
        if not is_16bit and has_alpha:
            alpha_pil = alpha_image
            if alpha_pil.mode != 'L':
                alpha_pil = alpha_pil.convert('L')
            final_image = Image.merge('LA', (final_image, alpha_pil))

        # 3. Handle Resizing
        if settings['size'] != final_image.size:
            final_image = final_image.resize(settings['size'], Image.Resampling.LANCZOS)

        # 4. Prepare Metadata
        exif_obj = Image.Exif()
        icc_profile = None
        dpi = None

        if not settings['strip_metadata']:
            exif_bytes = original_info.get('exif')
            if exif_bytes:
                try: exif_obj.load(exif_bytes)
                except Exception: pass

            icc_profile = settings.get('icc_profile_path') or original_info.get('icc_profile')
            dpi_val = settings.get('dpi')
            if dpi_val:
                dpi = (dpi_val, dpi_val)

        # 5. Save the image using the most appropriate library/method
        if file_ext == '.tiff':
            # If we get here, it's not 16-bit+alpha, so just save grayscale
            tiff_kwargs = {'photometric': 'minisblack'}
            icc_profile = settings.get('icc_profile_path') or original_info.get('icc_profile')
            if dpi:
                tiff_kwargs['resolution'] = (dpi[0], dpi[1], 'INCH')
            extratags = []
            if icc_profile:
                with open(icc_profile, "rb") if isinstance(icc_profile, str) else None as f:
                    icc_bytes = f.read() if f else icc_profile
                extratags.append((34675, 'B', len(icc_bytes), icc_bytes, False))
            if extratags:
                tiff_kwargs['extratags'] = extratags
            tifffile.imwrite(filepath, np.array(final_image), **tiff_kwargs)
        else:
            save_kwargs = {}
            # Always ensure icc_profile is bytes, not a file path
            if icc_profile:
                if isinstance(icc_profile, str):
                    with open(icc_profile, "rb") as f:
                        save_kwargs['icc_profile'] = f.read()
                else:
                    save_kwargs['icc_profile'] = icc_profile
            if dpi: save_kwargs['dpi'] = dpi
            if len(exif_obj) > 0:
                if file_ext == '.jpeg':
                    subsampling_val = settings.get('subsampling_val')
                    if subsampling_val is not None:
                        subsampling_map = {0: (1, 1), 1: (2, 1), 2: (2, 2)}
                        exif_obj[531] = subsampling_map.get(subsampling_val, (2, 2))
                save_kwargs['exif'] = exif_obj.tobytes()

            if file_ext == '.jpeg':
                if final_image.mode in ['LA', 'RGBA', 'P', 'PA']:
                    final_image = final_image.convert('L')
                final_image.save(filepath, "JPEG", 
                                 quality=settings.get('quality', 95), 
                                 subsampling=settings.get('subsampling_val', 0), 
                                 **save_kwargs)
            else:
                final_image.save(filepath, **save_kwargs)

    def on_drop(self, event):
        paths = self.master.tk.splitlist(event.data)
        if self.tab_view.get() == "Single Conversion" and paths: self.load_image(paths[0])
        else: self.add_batch_files(paths)

    def load_image(self, filepath=None):
        if not filepath:
            filetypes = [("All Images", "*.*")]
            if HEIF_SUPPORT: filetypes.insert(0, ("HEIF/HEIC", "*.heic *.heif"))
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
        if self.original_image is None:
            messagebox.showwarning("Warning", "No image to export.")
            return

        dialog = AdvancedExportDialog(self, self.original_info)
        settings = dialog.result
        if not settings:
            return
        
        default_name = Path(self.original_info.get('filepath', 'image.png')).stem + "_grayscale"
        filepath = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=settings['format'], filetypes=[(f"{settings['format'].upper()[1:]} files", f"*{settings['format']}")])
        if not filepath: return
        self.start_processing_indicator("Exporting image...")
        self.task_queue.put(('export_final', (self.original_image, filepath, settings, self.original_info)))

    def _perform_load(self, source):
        pil_image = None
        if isinstance(source, str):
            pil_image = Image.open(source)
        elif isinstance(source, Image.Image):
            pil_image = source
        
        pil_image.load()
        info = self.analyze_image_properties(pil_image)
        
        return pil_image, info
        
    def add_batch_files(self, filepaths=None):
        if not filepaths: filepaths = filedialog.askopenfilenames(title="Select files for batch processing")
        for path in filepaths:
            if path not in [f['path'] for f in self.batch_files]:
                frame = ctk.CTkFrame(self.batch_list_frame, fg_color=("gray85", "gray28"))
                frame.pack(fill="x", pady=2, padx=2)
                ctk.CTkLabel(frame, text=os.path.basename(path), anchor="w").pack(side="left", expand=True, fill="x", padx=5)
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
        if not (output_folder and os.path.isdir(output_folder)): return messagebox.showerror("Error", "Please select a valid output folder.")
        
        base_info = self.original_info or {'bit_depth': 8, 'size': (0,0), 'dpi': (72,72)}
        dialog = AdvancedExportDialog(self, base_info)
        export_settings = dialog.result
        if not export_settings: return
        export_settings['conversion_mode'] = self.conversion_mode_var.get()
        
        self.batch_progress.grid(row=4, column=0, padx=10, pady=(0,10), sticky="ew")
        self.batch_progress.set(0)
        
        for item in self.batch_files:
            self._update_batch_item_status(item['path'], "Queued", "#cccccc")
            out_name = Path(item['path']).stem + "_grayscale" + export_settings['format']
            self.task_queue.put(('batch_process', (item['path'], os.path.join(output_folder, out_name), export_settings)))
    
    def _update_batch_item_status(self, path, text, color):
        for item in self.batch_files:
            if item['path'] == path: item['status_label'].configure(text=text, text_color=color); return

    def _perform_resize_for_display(self, canvas, image: Image.Image):
        # CORRECTED FUNCTION
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w <= 1 or h <= 1: return None
        
        display_image = image
        # Standardize for display by converting to 8-bit L
        if image.mode != 'L':
            try:
                display_image = image.convert('L')
            except Exception: # Failsafe for unusual modes
                return None
        
        img_w, img_h = display_image.size
        scale = min(w / img_w, h / img_h, 1.0) if img_w > 0 and img_h > 0 else 1.0
        new_size = (int(img_w * scale), int(img_h * scale))
        
        if new_size[0] < 1 or new_size[1] < 1: return None
        return ImageTk.PhotoImage(display_image.resize(new_size, Image.Resampling.LANCZOS))
        
    def update_preview(self, _=None):
        if self.original_image:
            self.start_processing_indicator("Applying grayscale effect...")
            # Always use original bit depth for preview for consistency and speed
            self.task_queue.put(('convert', (self.original_image, self.conversion_mode_var.get(), self.original_info.get('bit_depth', 8))))
            
    def _handle_load_success(self, data):
        self.original_image, self.original_info = data
        self.stop_processing_indicator(f"Loaded: {os.path.basename(self.original_info.get('filepath', 'clipboard'))}")
        self.info_var.set(self.original_info['display_text'])
        self.export_button.configure(state="disabled")
        # Always display the original image as loaded (not grayscale)
        self.request_display_update('original')
        self.update_preview()

    def _handle_convert_success(self, data):
        self.preview_data, _ = data
        self.stop_processing_indicator("Enhanced preview ready.")
        self.export_button.configure(state="normal")
        # Only update the preview canvas, not the original
        self.request_display_update('preview')

    def request_display_update(self, canvas_name):
        image_to_display = None
        if canvas_name == 'original':
            # Always show the original loaded image, not the grayscale
            image_to_display = self.original_image
        elif canvas_name == 'preview' and self.preview_data is not None:
            mode = 'I;16' if self.preview_data.dtype == np.uint16 else 'L'
            image_to_display = Image.fromarray(self.preview_data, mode=mode)

        canvas = self.original_canvas if canvas_name == 'original' else self.preview_canvas
        image_to_display = self.original_image if canvas_name == 'original' else (Image.fromarray(self.preview_data, mode='I;16' if self.preview_data.dtype == np.uint16 else 'L') if self.preview_data is not None else None)
        if image_to_display: self.task_queue.put(('resize_display', (canvas, image_to_display)))
            
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
            if '16' in dtype_str:
                info['bit_depth'] = 16
            elif '32' in dtype_str:
                 info['bit_depth'] = 32
            else:
                 info['bit_depth'] = 8
        except Exception:
            if image.mode in ('I;16', 'I;16B', 'I;16L', 'I;16N', 'I;16LA') or (image.mode == 'LA' and image.getextrema()[0][1] > 255):
                info['bit_depth'] = 16
            elif image.mode in ('I', 'F'):
                info['bit_depth'] = 32
            else:
                info['bit_depth'] = 8

        info['display_text'] = f"Size: {image.size[0]}×{image.size[1]} | Mode: {image.mode} | Bit Depth: {info['bit_depth']}-bit | {'ICC' if info['icc_profile'] else 'No ICC'}"
        return info

    def _update_canvas_image(self, canvas, photo_image):
        if not photo_image: return
        canvas.delete("all")
        canvas.photo = photo_image  # Prevent garbage collection
        x = (canvas.winfo_width() - photo_image.width()) // 2
        y = (canvas.winfo_height() - photo_image.height()) // 2
        canvas.create_image(x, y, anchor='nw', image=photo_image)

if __name__ == "__main__":
    if not HEIF_SUPPORT: print("WARNING: HEIF/HEIC support is not available. Please run 'pip install pillow-heif'.")
    root = TkinterDnD.Tk()
    root.withdraw()
    app_window = ctk.CTkToplevel()
    app_window.title("Enhanced Precision Grayscale Converter v6.5.7")
    app_window.geometry("1400x900")
    app_window.minsize(1200, 800)
    app = MainApplication(master=app_window)
    app_window.protocol("WM_DELETE_WINDOW", root.destroy)
    app_window.mainloop()