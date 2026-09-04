# Dark Slides

Batch convert PowerPoint files (`.ppt`, `.pptx`) to PDF using Microsoft PowerPoint COM automation, then generate dark-mode versions of each PDF.

## Requirements

- Windows with Microsoft PowerPoint installed
- PowerShell 5.1+ or PowerShell 7
- Node.js (for dark mode conversion)

## Setup

```bash
npm install
```

## Usage

Place all `.ppt` and `.pptx` files in the project root, then run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\ppt-to-pdf.ps1"
```

Or with PowerShell 7:

```powershell
pwsh -ExecutionPolicy Bypass -File ".\ppt-to-pdf.ps1"
```

## Output

| Directory  | Contents                    |
|------------|-----------------------------|
| `pdf/`     | Converted PDFs              |
| `pdf-dark/`| Dark-mode PDFs              |
| `ppt/`     | Original PowerPoint files (moved) |

## How It Works

1. `ppt-to-pdf.ps1` uses PowerPoint's COM interface to export each file to PDF.
2. Original `.ppt`/`.pptx` files are moved to `ppt/`.
3. `convert-dark.js` inverts each PDF's colors against a dark theme (`rgb(42, 37, 34)`) by rendering pages to canvas and remapping pixel values.
