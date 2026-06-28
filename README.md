# SolidWorks Drawing Automation Pipeline

Automate the extraction of dimensions from 3D parts (`.SLDPRT`) and insert them cleanly into 2D drawings (`.SLDDRW`), with automated view creation, layout arranging, and PDF exporting.

---

## Features

- **Automatic Drawing Generation**: Generates a standard drawing from a 3D part using your default SolidWorks sheet templates.
- **Auto-arranged Views**: Places Front, Top, Right, and Isometric views in standard configurations.
- **Annotation Import**: Automatically imports marked model dimensions directly into the drawing views.
- **Smart Formatting**: Runs SolidWorks' Auto-Arrange layout tool to ensure dimensions are legible and non-overlapping.
- **PDF Exporting**: Automatically exports the completed drawing sheet to PDF for easy sharing and review.
- **Granular Control**: Supports processing a single part, a whole directory, or performing a connectivity dry-run.

---

## Prerequisites

1. **Windows OS** (required for SolidWorks & pywin32 COM communication).
2. **SolidWorks** installed and running on your machine.
3. **Python 3.8+** installed.
4. **pywin32** library installed for COM integration.

---

## Installation

Install the required Python COM bindings using `pip`:

```bash
pip install pywin32
```

Ensure SolidWorks is open and running on your system before launching the automation script.

---

## Usage

Run the pipeline using the command line with various flags:

### 1. Process All Parts in Current Folder
Process all `.SLDPRT` models in the current directory:
```bash
python solidworks_dimension_pipeline.py
```

### 2. Connect and Test Only (Dry Run)
Verify Python's COM connectivity to SolidWorks without performing any file modifications:
```bash
python solidworks_dimension_pipeline.py --dry-run
```

### 3. Process a Single Part File
Specify a single part file to generate a drawing and PDF for:
```bash
python solidworks_dimension_pipeline.py --part Part1.SLDPRT
```

### 4. Process a Specific Directory
Process all parts in a specific folder path:
```bash
python solidworks_dimension_pipeline.py --dir "C:\Path\To\Your\CAD\Models"
```

---

## Project Structure

For a full explanation of the project layout and guidelines, see [project_structure.md](project_structure.md).

```
SOLIDWORKS/
├── .gitignore                          # Exclude temporary CAD lockfiles and logs
├── README.md                           # Main usage documentation (this file)
├── project_structure.md                # Project layout and descriptions
├── solidworks_dimension_pipeline.py    # Automation Python script
├── Part1.SLDPRT                        # CAD 3D Part
├── Part1.SLDDRW                        # CAD 2D Drawing
└── Part1.pdf                           # Final exported drawing sheet
```

## Logging & Troubleshooting

- **Logs**: Execution logs are stored in `solidworks_pipeline.log`.
- **Active Document Errors**: Make sure that the target files are not currently open in edit mode or locked by another SolidWorks instance.
- **Default Drawing Template**: If the script warns about missing drawing templates, configure it inside SolidWorks:
  `Tools > Options > System Options > Default Templates`
