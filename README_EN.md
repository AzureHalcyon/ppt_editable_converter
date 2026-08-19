# PowerPoint Editable Converter

[简体中文](README.md) | English

A portable Windows script for converting read-only PowerPoint presentations.

Copy `convert_editable.py` into the folder containing the presentations you want to process, then run it.

The script lists the PowerPoint files in that folder. After you select one or more files, the converted presentations are saved alongside their source files.

The script uses Microsoft PowerPoint's COM automation interface to open each source presentation in read-only mode. It then uses PowerPoint's own slide import feature to insert the slides into a newly created presentation and saves the result as a separate file. The source presentation is not modified.

Because the presentation is rebuilt, slide masters, themes, animations, external links, and other advanced features may not be preserved perfectly. Check the layout and playback behavior after conversion. The script cannot bypass open passwords, Information Rights Management (IRM), or other access controls; Microsoft PowerPoint must be able to open the source file normally.

## Supported Formats

Supported input formats:

`.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`, `.ppsm`, `.pot`, `.potx`, and `.potm`.

By default, the script creates a new `.pptx` file and appends `-editable` to its name. For example:

```text
presentation.ppt  ->  presentation-editable.pptx
```

## Requirements

- Windows
- Microsoft PowerPoint desktop application
- Python 3.10 or later
- `pywin32`

Install the dependency with:

```powershell
python -m pip install pywin32
```

## Usage

Copy the script into the folder containing your presentations, then run:

```powershell
python convert_editable.py
```

The script displays a selection menu similar to this:

```text
PowerPoint files found:
  [1] presentation.ppt
  [2] report.pptx
  [a] Convert all
  [q] Quit
Enter a number (separate multiple selections with commas):
```

Enter one number, multiple comma-separated numbers, `a` to convert all files, or `q` to quit.

If an output file already exists, the script asks whether it should be overwritten. Converted files are saved in the same directory as their source files. Files whose names already end in `-editable` are excluded from the default selection list.

## Optional Commands

```powershell
# Scan subdirectories recursively
python convert_editable.py --recursive

# Skip the selection menu and convert every file found
python convert_editable.py --all

# Automatically overwrite existing output files
python convert_editable.py --all --overwrite

# Show the PowerPoint window during conversion
python convert_editable.py --visible

# Preview the source and output paths without converting anything
python convert_editable.py --dry-run
```

You can also pass a file or directory explicitly:

```powershell
python convert_editable.py "path\to\presentation.pptx"
```

## Limitations

- Content that is already stored as an image remains an image; the script cannot reconstruct editable text or shapes from it.
- Slide masters, themes, animations, external links, and other advanced features may not be preserved perfectly.
- When the output format is `.pptm`, only the slides are copied; VBA projects are not copied.
- Password-protected, IRM-protected, or otherwise restricted files must already be accessible to PowerPoint.
- The script requires the desktop version of Microsoft PowerPoint and does not work with Office for the web alone.

## Project Structure

```text
convert_editable.py       Self-contained portable script
tests/                    Automated tests
pyproject.toml            Project metadata and dependencies
```
