# Enhanced Precision Grayscale Converter

A modern desktop application for advanced grayscale image conversion, batch processing, and export, supporting high bit-depth, color profiles, and multiple formats.

## Features
- **Single Image & Batch Processing**: Convert one or many images at once.
- **Multiple Grayscale Modes**: BT.709, L\*a\*b\* (L\*), HSL (Lightness), HSV (Value), BT.601, BT.2100, Gamma.
- **High Bit-Depth Support**: 8-bit and 16-bit grayscale output.
- **Alpha Channel Handling**: Preserve transparency for supported formats.
- **Advanced Export Options**: Choose format, bit depth, color profile, DPI, and metadata handling.
- **Drag & Drop**: Quickly add files for batch processing.
- **Clipboard Support**: Load images directly from the clipboard.
- **Modern UI**: Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter).

## Supported Formats
- PNG (8/16-bit, with/without alpha)
- TIFF (8/16-bit, with/without alpha)
- JPEG
- WEBP
- BMP
- HEIC/HEIF (if `pillow-heif` is installed)

## Requirements
- Python 3.8+
- Windows OS (tested)
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- [Pillow](https://python-pillow.org/)
- [numpy](https://numpy.org/)
- [tifffile](https://pypi.org/project/tifffile/)
- [imageio](https://pypi.org/project/imageio/)
- [pillow-heif](https://pypi.org/project/pillow-heif/) (optional, for HEIC/HEIF)
- [tkinterdnd2](https://pypi.org/project/tkinterdnd2/)

Install all dependencies:
```sh
pip install -r requirements.txt
```

## Usage
1. Run the application:
   ```sh
   python main.py
   ```
2. Use the UI to load images, select conversion mode, and export.
3. For batch processing, add files and select an output folder.

## Building Executable

You can generate a standalone Windows executable using PyInstaller. A build script and a spec file are provided.

### Quick Build (Recommended)

Run the provided batch script from the `build_scripts` folder:

```sh
cd build_scripts
./build.bat
```

This will install dependencies, build the executable, and place it in the main directory.

### Manual Build

Alternatively, you can build manually with PyInstaller:

```sh
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name "PrecisionGrayscaleConverter" --icon=icon.ico main.py
```

The executable will be found in the `dist` folder.

### Creating an Installer

An NSIS installer script is provided in `build_scripts/installer.nsi`. After building the executable, use [NSIS](https://nsis.sourceforge.io/) to generate an installer:

1. Download and install NSIS.
2. Open `installer.nsi` in the NSIS compiler.
3. Compile to create a Windows installer `.exe`.

## Notes
- 16-bit PNG with alpha is not supported (alpha will be discarded).
- For HEIC/HEIF support, install `pillow-heif`.
- ICC profiles are loaded from the Windows system color folder if available.

## License
MIT License

---
**Enhanced Precision Grayscale Converter** — by [Jayesh]
