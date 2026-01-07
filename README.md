# Prefix-Suffix renamer with OCR recognition for the context menu in Windows 11/10

This simple Python script allows you to add custom prefixes and suffixes to file names via the Windows 11/10 context menu. It provides an easy way to rename files directly from the right-click menu in Windows Explorer. 

## Motivation

Although there are sophisticated tools to organize documents such as Mendeley or Zotero, some users still prefer the classic Windows Explorer folder/document organization structure. This tool helps you to quickly rename new academic PDF files based on their content type such as Book, Paper, Thesis, Report, etc., while automatically adding the publication year as a prefix and the author’s name can be added as suffix, making manual organization faster and more consistent.

## Features

- Add predefined prefixes or suffixes to file names according to a list of options.
- Install or uninstall the context menu entries for quick access.
- Automatically handles file name conflicts by appending a counter to the new file name.
  
## Requirements

- Python 3.x
- Administrator privileges (required for installing/uninstalling)
- [pillow](https://pypi.org/project/Pillow/)
- [mss](https://pypi.org/project/mss/)
- [pytesseract](https://pypi.org/project/pytesseract/)
- [Tesseract OCR Engine](https://github.com/UB-Mannheim/tesseract/wiki)

## Installation

To install the context menu entries, run the following command in the `cmd` as `admnistrator`:

```bash
python presuffix.py install
```

## Uninstallation

To uninstall the context menu entries, run the following command:

```bash
python presuffix.py uninstall
```
Optional: kill and restart `explorer.exe` by command to see the changes.

```bash
taskkill /f /im explorer.exe && start explorer.exe
```

## Usage

Once installed, you can right-click on a file in Windows Explorer and select "Add Prefix-Suffix". From there, you can choose a prefix or suffix to apply to the selected file.

## Modify options list

Just modify the following list in the code and `uninstall/install` script.

```python
self.prefix_options = ["+Book+year+", "+Paper+year+", "+Thesis+year+", "+Report+year+", 
                        "+Slides+year+", "+Presentation+year+", "+Draft+year+"]
self.suffix_options = ["+authors"] 
```

## Notes
- The script automatically elevates privileges when needed for installation or uninstallation.
- File name conflicts are resolved by appending a counter to the new file name (e.g., example.pdf → +Book+year+example (1).pdf).

## License

MIT License. 

## Author

paulomarconi
