#!/usr/bin/env python3
#
# Author: Fernando Mangussi
#
# MIT License
#
# Copyright (c) 2025 Fernando Mangussi
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# Dependency Notice:
# This script invokes pbi-tools[](https://pbi.tools/) as an external executable,
# which is licensed under the GNU Affero General Public License version 3
# (AGPLv3) or later. The script itself is not licensed under AGPLv3.
#
# If you modify or distribute pbi-tools (e.g., custom builds or bundles),
# you must comply with AGPLv3 requirements, such as offering modified source
# code to recipients. For unmodified use, no additional obligations apply.
# Full license details: https://pbi.tools/license
#
# The pbi-tools binaries must be obtained separately and placed in the paths
# specified via command-line arguments.

import argparse
import logging
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional


def get_logger(name: str = __name__, level: int = logging.DEBUG) -> logging.Logger:
    """
    Create and return a configured logger.
    - INFO and lower → stdout
    - ERROR and higher → stderr
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Avoid adding duplicate handlers if called multiple times
    if logger.hasHandlers():
        return logger

    # Define log format
    log_format = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # --- DEBUG → stderr ---
    debug_handler = logging.StreamHandler(sys.stderr)
    debug_handler.setLevel(logging.DEBUG)
    debug_handler.addFilter(lambda record: record.levelno == logging.DEBUG)
    debug_handler.setFormatter(log_format)

    # --- INFO and above → stdout ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(log_format)

    # Attach handlers
    logger.addHandler(debug_handler)
    logger.addHandler(console_handler)

    return logger

logger = get_logger("pbi_converter")

def parse_arguments() -> argparse.Namespace:
    """
    Parse command-line arguments for the script.
    """
    parser = argparse.ArgumentParser(
        description="Convert .pbix files to .pbit using pbi-tools.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
PowerShell Examples:

  python converter.py `
    --report-folder "C:\\Users\\myuser\\reports" `
    --pbit-output "C:\\Users\\myuser\\pbit" `
    --temp-folder "C:\\Users\\myuser\\temp" `
    --cli-path "C:\\Users\\myuser\\pbi-tools.1.2.0" `
    --core-path "C:\\Users\\myuser\\pbi-tools.core.1.2.0_win-x64" `
    --clean 2> "ConverterError.log" | Tee-Object -FilePath "ConverterOutput.log"


Notes:
- The script recursively searches for .pbix files in the report folder and subdirectories.
- Converted .pbit files are saved in the output folder, mirroring the original folder structure.
- Extracted contents are temporarily stored in the temp folder, also mirroring the structure.
- Use --clean to automatically remove the extraction folders after successful conversion.
- Run with -h or --help for detailed argument information.

Pre-Requirements:
    - Power BI Desktop ("https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe"):
    - .Net Runtime (https://dotnet.microsoft.com/en-us/download/dotnet/8.0/runtime)
    - Python 3 (https://www.python.org/downloads/)
    - Enable long paths in Windows 10, version 1607, and later (https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry)
    - pbi-tools & pbi-tools.core (https://pbi.tools/tutorials/getting-started-cli.html)


Tutorial:

    # Make sure you have met all the pre-reqs
    # Open PowerShell
    cd $HOME
    $cwd = (pwd)
    git clone https://github.com/microsoft/powerbi-desktop-samples.git reports


    python converter.py `
    --report-folder "$cwd\\reports" `
    --pbit-output "$cwd\\pbit" `
    --temp-folder "$cwd\\temp" `
    --cli-path "$cwd\\pbi-tools.1.2.0" `
    --core-path "$cwd\\pbi-tools.core.1.2.0_win-x64" `
    --clean 2> "ConverterError.log" | Tee-Object -FilePath "ConverterOutput.log"

        """
    )
    parser.add_argument(
        "-r", "--report-folder",
        type=str,
        required=True,
        help="The root folder where the .pbix files are located."
    )
    parser.add_argument(
        "-o", "--pbit-output",
        type=str,
        required=True,
        help="The root folder where .pbit files will be saved."
    )
    parser.add_argument(
        "-t", "--temp-folder",
        type=str,
        required=True,
        help="The root folder where .pbix files will be extracted."
    )
    parser.add_argument(
        "-cli", "--cli-path",
        type=str,
        required=True,
        help="The pbi-tools binary folder."
    )
    parser.add_argument(
        "-core", "--core-path",
        type=str,
        required=True,
        help="The pbi-tools.core binary folder."
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Clean the temp folder after successful conversion."
    )
    return parser.parse_args()


def validate_paths(cli_path: Path, core_path: Path) -> None:
    """
    Validate that the required executable paths exist.
    """
    if not cli_path.exists():
        logger.error("pbi-tools executable not found: %s", cli_path)
        sys.exit(1)
    if not core_path.exists():
        logger.error("pbi-tools.core executable not found: %s", core_path)
        sys.exit(1)


def find_pbix_files(report_folder: Path) -> List[Path]:
    """
    Recursively find all .pbix files in the report folder.
    """
    logger.info("Searching for .pbix files in %s...", report_folder)
    pbix_files = list(report_folder.rglob("*.pbix"))
    logger.info("Found %d files.", len(pbix_files))
    return pbix_files


def ensure_directories(output_folder: Path, target_dir: Path, extract_folder: Path, temp_folder: Path) -> None:
    """
    Ensure the necessary directories exist.
    """
    output_folder.mkdir(parents=True, exist_ok=True)
    target_dir.mkdir(parents=True, exist_ok=True)
    extract_folder.mkdir(parents=True, exist_ok=True)
    temp_folder.mkdir(parents=True, exist_ok=True)


def run_subprocess(command: List[str], description: str) -> Optional[str]:
    """
    Run a subprocess command and return stdout if successful, otherwise log error.
    """
    try:
        result = subprocess.run(    
            command,
            capture_output=True,
            text=True,
            check=True  # Raises CalledProcessError on non-zero exit code
        )
        return result.stdout
    except subprocess.CalledProcessError as e:
        if "could not be deserialized" in e.stdout:
            logger.error("File model is not supported. Model required: V3")    
        else:
            logger.error("Error while processing the file. Check the logs for more information.")
        logger.debug(
            "%s failed with return code %d: %s\nCommand: %s\nStdout: %s\nStderr: %s",
            description, e.returncode, e,
            " ".join(command), e.stdout, e.stderr
        )
        return None
    except Exception as e:
        logger.error("Unexpected error during %s: %s", description, e)
        return None


def extract_pbix(cli_path: Path, pbix_file: Path, extract_folder: Path) -> Optional[str]:
    """
    Extract the contents of a .pbix file using pbi-tools.
    """
    command = [
        str(cli_path),
        "extract",
        str(pbix_file),
        "-extractFolder",
        str(extract_folder)
    ]
    return run_subprocess(command, f"extraction of {pbix_file.name}")


def compile_to_pbit(core_path: Path, extract_folder: Path, target_dir: Path) -> Optional[str]:
    """
    Compile the extracted contents to a .pbit file.
    """
    # Use platform-appropriate path separators
    extract_path = str(extract_folder)
    target_path = str(target_dir)
    command = [
        str(core_path),
        "compile",
        extract_path,
        target_path,
        "PBIT",
        "True"
    ]
    return run_subprocess(command, f"compilation to PBIT in {target_dir}")


def parse_pbit_output(stdout: str) -> Optional[str]:
    """
    Parse the compilation output to extract the .pbit file path.
    """
    pbit_pattern = r"PBIT file written to: (.*)"
    match = re.search(pbit_pattern, stdout, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    if "does not contain a V3 model" in stdout:
        logger.warning("Skipping: PBIX requires V3 model")
    else:
        logger.error("Failed to parse PBIT output path")
    return None


def process_pbix_file(
    pbix_file: Path,
    output_folder: Path,
    cli_path: Path,
    core_path: Path,
    report_folder: Path,
    temp_folder: Path,
    clean_extract: bool = False
) -> None:
    """
    Process a single .pbix file: extract, compile, and log the result.
    """
    try:
        relative_path = pbix_file.relative_to(report_folder)
        target_dir = output_folder / relative_path.parent
        extract_folder = temp_folder / relative_path.parent
        logger.info("-" * 70)
        logger.info("Processing: %s", pbix_file)
        ensure_directories(output_folder, target_dir, extract_folder, temp_folder)
        # Extract
        extract_stdout = extract_pbix(cli_path, pbix_file, extract_folder)
        if not extract_stdout:
            return
        # Compile
        compile_stdout = compile_to_pbit(core_path, extract_folder, target_dir)
        if not compile_stdout:
            return
        # Parse output
        pbit_file_path = parse_pbit_output(compile_stdout)
        if pbit_file_path:
            logger.info("Output: %s", pbit_file_path)
            if clean_extract:
                shutil.rmtree(extract_folder, ignore_errors=True)
                logger.info("Extract folder cleaned: %s", extract_folder)
        else:
            logger.warning("No PBIT file generated for %s", pbix_file.name)
    except Exception as e:
        logger.error("Unexpected error processing %s: %s", pbix_file, e)


def format_duration(seconds: float) -> str:
    """
    Format seconds into a human-readable duration (e.g., '1m 23s').
    """
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    if minutes > 0:
        return f"{minutes}m {secs}s"
    return f"{secs}s"


def main() -> None:
    """
    Main entry point: Converts all .pbix files in the report folder and subdirectories
    to .pbit files in the output folder, mirroring the original folder structure.
    """
    args = parse_arguments()
    report_folder = Path(args.report_folder)
    output_folder = Path(args.pbit_output)
    temp_folder = Path(args.temp_folder)
    cli_path = Path(args.cli_path) / "pbi-tools.exe"
    core_path = Path(args.core_path) / "pbi-tools.core.exe"
    validate_paths(cli_path, core_path)
    pbix_files = find_pbix_files(report_folder)
    if not pbix_files:
        logger.warning("No .pbix files found. Exiting.")
        return

    total_files = len(pbix_files)
    processed_count = 0
    total_time = 0.0

    for pbix_file in pbix_files:
        start_time = time.time()
        process_pbix_file(pbix_file, output_folder, cli_path, core_path,
                          report_folder, temp_folder, args.clean)
        duration = time.time() - start_time
        processed_count += 1
        total_time += duration

        logger.info("Completed %s in %s", pbix_file.name, format_duration(duration))

        if processed_count < total_files:
            avg_time = total_time / processed_count
            remaining_files = total_files - processed_count
            eta_seconds = avg_time * remaining_files
            logger.info("Progress: %d/%d files processed (%.1f%%). ETA: %s",
                        processed_count, total_files,
                        (processed_count / total_files) * 100,
                        format_duration(eta_seconds))

    logger.info("Conversion process completed. Total time: %s", format_duration(total_time))


if __name__ == "__main__":
    main()
