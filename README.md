# PC Status Report Generator

This Python script generates a detailed PDF report of your PC's status, including system information, hardware details, and performance metrics.

## Features

- System Information (OS, version, etc.)
- Display Information (monitors, resolution, manufacturer)
- Input Devices (keyboard, touchpad, mouse)
- I/O Ports (USB, serial ports)
- Power and Charging Information
- GPU Information (name, memory, resolution, drivers)
- CPU Information (cores, frequency, usage)
- Memory Information (total, available, used)
- Disk Information (partitions, usage, file systems)
- Battery Information (if available)
- Network Information (interfaces, addresses)

## Requirements

- Python 3.6 or higher
- Required Python packages (install using `pip install -r requirements.txt`):
  - psutil
  - reportlab
  - wmi
  - pywin32
  - pypiwin32

## Installation

1. Clone this repository or download the files
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the script:
   ```bash
   python pc_report.py
   ```
2. The script will generate a PDF file named `pc_status_report.pdf` in the same directory
3. Open the PDF to view your PC's status report

## Output

The generated PDF report includes:
- A title page with generation timestamp
- System information section
- CPU information section
- Memory information section
- Disk information section
- Battery information section (if available)
- Network information section

Each section is clearly formatted with tables and proper headings for easy reading. 