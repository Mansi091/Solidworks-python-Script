# Project Structure

This document outlines the organization and directory structure of the **SolidWorks Drawing Automation Pipeline** project.

## Directory Layout

```
SOLIDWORKS/
├── .gitignore                          # Specifies intentionally untracked files to ignore
├── README.md                           # Main project documentation and usage guide
├── project_structure.md                # This file, explaining the project structure
├── solidworks_dimension_pipeline.py    # Main automation Python script
├── Part1.SLDPRT                        # Example SolidWorks 3D Part file
├── Part1.SLDDRW                        # Example SolidWorks 2D Drawing file
├── Part1.pdf                           # Example exported 2D Drawing PDF
└── solidworks_pipeline.log             # Local execution log (generated at runtime, git-ignored)
```

## Component Details

### 1. Source Code & Scripting
* **[solidworks_dimension_pipeline.py](file:///c:/Users/Mansi/OneDrive/Desktop/SOLIDWORKS/solidworks_dimension_pipeline.py)**: The core executable script of the project. It uses COM connection (`pywin32`) to automate:
  * Connecting to an active instance of SolidWorks.
  * Loading a `.SLDPRT` 3D Part file.
  * Extracting model dimensions.
  * Generating a corresponding `.SLDDRW` 2D Drawing using the default sheet template.
  * Automatically inserting views and dimensions, followed by auto-arranging annotations.
  * Exporting the finalized drawing to PDF format.

### 2. SolidWorks Design Files (CAD)
* **`Part1.SLDPRT`**: The 3D model that acts as the input source for the pipeline.
* **`Part1.SLDDRW`**: The 2D drawing created/updated by the script. It holds views (e.g. Front, Top, Right, Isometric) and dimensions linked to the 3D part.

### 3. Output Files
* **`Part1.pdf`**: The production-ready 2D drawing exported as a PDF file, suitable for manufacturing or engineering review.

### 4. Configuration & Metadata
* **`.gitignore`**: Defines rules to ignore lock files, debug files, and local logs from version control.
* **`README.md`**: Explains prerequisites, installations, commands, and options to run the automation pipeline.

---

## Workspace Guidelines

1. **CAD File Pairing**: Always ensure that `.SLDPRT` and `.SLDDRW` files have matching basenames in the same folder. The script depends on this convention to match drawings to parts.
2. **CAD Temporary Files**: SolidWorks creates hidden lock files (e.g., `~$Part1.SLDPRT`). These are automatically ignored by Git and should not be manually committed.
3. **Execution Logs**: Execution output is saved to `solidworks_pipeline.log`. Keep this locally for debugging; it is excluded from Git to avoid unnecessary code churn.
