# PBIX to PBIT Converter

A Python script to batch convert Power BI (.pbix) files to templates (.pbit) using [pbi-tools](https://pbi.tools/). The script recursively searches for .pbix files in a specified folder (and subfolders), extracts their contents, compiles them into .pbit files, and saves them to an output folder while mirroring the original directory structure. Temporary extraction folders can be automatically cleaned up.

**Author:** Fernando Mangussi  
**License:** MIT License. See the [LICENSE](LICENSE) file for details.  
**Disclaimer:** This script is provided "AS IS", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.

## Features
- Recursively finds all .pbix files in the input folder and subdirectories.
- Extracts contents using `pbi-tools extract`.
- Compiles to .pbit using `pbi-tools.core compile`.
- Mirrors the original folder structure in the output and temp directories.
- Optional cleanup of temporary extraction folders.
- Progress tracking with ETA and timing logs.
- Handles errors gracefully (e.g., unsupported file models).

## Prerequisites
Before running the script, ensure you have the following installed and configured:

- **Power BI Desktop**: Download from [Microsoft](https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe).
- **.NET Runtime**: Version 8.0 or later from [Microsoft](https://dotnet.microsoft.com/en-us/download/dotnet/8.0/runtime).
- **Python 3**: Version 3.x from [python.org](https://www.python.org/downloads/). Ensure it's added to your PATH.
- **Windows Long Paths Enabled**: For Windows 10 (version 1607) and later, enable long path support via [Microsoft Docs](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry). This prevents issues with deeply nested files.
- **pbi-tools & pbi-tools.core**: Download from [pbi-tools Getting Started](https://pbi.tools/tutorials/getting-started-cli.html).
  - Place the binaries in folders specified via `--cli-path` and `--core-path` (e.g., `pbi-tools.exe` and `pbi-tools.core.exe`).
  - **Note on Licensing**: pbi-tools is licensed under the GNU Affero General Public License version 3 (AGPLv3) or later. If you modify or distribute it (e.g., custom builds), comply with AGPLv3 (e.g., offer source code). For unmodified use, no additional obligations apply. Full license: [pbi-tools License](https://pbi.tools/license).

**Verify Setup:**
- Run `python --version` to confirm Python is installed.
- Ensure `pbi-tools.exe` and `pbi-tools.core.exe` are executable and accessible.

## Installation
1. Clone or download this script (`converter.py`) to your working directory.
2. No additional Python packages are required—the script uses only standard libraries (`argparse`, `logging`, `re`, `shutil`, `subprocess`, `sys`, `time`, `pathlib`, `typing`).
3. Download and extract pbi-tools binaries to separate folders (e.g., `pbi-tools.1.2.0` and `pbi-tools.core.1.2.0_win-x64`).

## Usage
Run the script from the command line (PowerShell or Command Prompt recommended for Windows):

```bash
python converter.py [OPTIONS]
```

### Command-Line Arguments
| Argument | Required | Description | Default |
|----------|----------|-------------|---------|
| `-r, --report-folder` | Yes | Root folder containing .pbix files (searches recursively). | N/A |
| `-o, --pbit-output` | Yes | Root folder to save .pbit files (mirrors input structure). | N/A |
| `-t, --temp-folder` | Yes | Root folder for temporary extraction (mirrors input structure). | N/A |
| `-cli, --cli-path` | Yes | Folder containing `pbi-tools.exe`. | N/A |
| `-core, --core-path` | Yes | Folder containing `pbi-tools.core.exe`. | N/A |
| `--clean` | No | Automatically delete temp extraction folders after successful conversion. | False |

For full help:
```bash
python converter.py --help
```

### Logging
- **INFO and DEBUG**: Output to stdout (progress, timings).
- **ERROR and above**: Output to stderr.
- Redirect for logging: `python converter.py [args] 2> "Error.log" | Tee-Object -FilePath "Output.log"` (PowerShell).

## Examples

### Basic PowerShell Example
Convert .pbix files from a reports folder, output to `pbit`, use temp folder, and clean up:

```powershell
python converter.py `
  --report-folder "C:\Users\myuser\reports" `
  --pbit-output "C:\Users\myuser\pbit" `
  --temp-folder "C:\Users\myuser\temp" `
  --cli-path "C:\Users\myuser\pbi-tools.1.2.0" `
  --core-path "C:\Users\myuser\pbi-tools.core.1.2.0_win-x64" `
  --clean 2> "ConverterError.log" | Tee-Object -FilePath "ConverterOutput.log"
```

### Tutorial: Quick Start with Sample Reports
1. Ensure prerequisites are met (see above).
2. Open PowerShell and navigate to your home directory:
   ```powershell
   cd $HOME
   $cwd = (pwd)
   ```
3. Clone sample Power BI reports:
   ```powershell
   git clone https://github.com/microsoft/powerbi-desktop-samples.git reports
   ```
4. Run the converter (adjust paths as needed):
   ```powershell
   python converter.py `
     --report-folder "$cwd\reports" `
     --pbit-output "$cwd\pbit" `
     --temp-folder "$cwd\temp" `
     --cli-path "$cwd\pbi-tools.1.2.0" `
     --core-path "$cwd\pbi-tools.core.1.2.0_win-x64" `
     --clean 2> "ConverterError.log" | Tee-Object -FilePath "ConverterOutput.log"
   ```

This will process all .pbix files in `reports/`, save .pbit files to `pbit/`, and clean up `temp/`.

## Troubleshooting
- **"pbi-tools executable not found"**: Verify `--cli-path` and `--core-path` point to folders containing the .exe files.
- **"File model is not supported. Model required: V3"**: The .pbix uses an older model; skip or update it in Power BI Desktop.
- **Path length errors**: Enable long paths in Windows (see Prerequisites).
- **No .pbix files found**: Check `--report-folder` path and ensure files exist.
- **Subprocess errors**: Check logs for stdout/stderr details. Ensure .NET Runtime is installed.
- **Large files slow**: Processing can take time; monitor ETA in logs.
- **Antivirus false positives**: pbi-tools binaries may trigger scans—add exceptions if needed.

## Notes
- Only V3 models are supported for compilation.
- The script skips files that don't generate a .pbit (e.g., invalid models) and continues processing others.
- Output .pbit paths are logged for verification.
- Tested on Windows; may require tweaks for other OS (e.g., path separators).

## License
- **Script**: Script: MIT License. See the LICENSE file for the full text.
- **pbi-tools**: AGPLv3 (see [pbi-tools License](https://pbi.tools/license) for details).

If you encounter issues or have suggestions, feel free to fork and improve!
