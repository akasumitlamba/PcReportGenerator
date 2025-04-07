import psutil
import wmi
import platform
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

def get_system_info():
    """Get basic system information"""
    info = {
        'System': platform.system(),
        'Node Name': platform.node(),
        'Release': platform.release(),
        'Version': platform.version(),
        'Machine': platform.machine(),
        'Processor': platform.processor()
    }
    return info

def get_cpu_info():
    """Get CPU information"""
    cpu_info = {}
    try:
        w = wmi.WMI()
        cpu = w.Win32_Processor()[0]  # Get the first CPU
        cpu_info = {
            'Name': cpu.Name,
            'Manufacturer': cpu.Manufacturer,
            'Description': cpu.Description,
            'Physical Cores': psutil.cpu_count(logical=False),
            'Total Cores': psutil.cpu_count(logical=True),
            'Max Clock Speed': f"{cpu.MaxClockSpeed} MHz",
            'Current Clock Speed': f"{psutil.cpu_freq().current:.2f} MHz",
            'CPU Usage': f"{psutil.cpu_percent()}%",
            'Architecture': cpu.Architecture,
            'Family': cpu.Family,
            'ProcessorId': cpu.ProcessorId,
            'Socket': cpu.SocketDesignation
        }
    except Exception as e:
        cpu_info = {'Error': f"Could not retrieve detailed CPU information: {str(e)}"}
    return cpu_info

def get_memory_info():
    """Get memory information"""
    memory = psutil.virtual_memory()
    memory_info = {
        'Total Memory': f"{memory.total / (1024**3):.2f} GB",
        'Available Memory': f"{memory.available / (1024**3):.2f} GB",
        'Used Memory': f"{memory.used / (1024**3):.2f} GB",
        'Memory Usage': f"{memory.percent}%"
    }
    return memory_info

def get_disk_info():
    """Get disk information"""
    disk_info = {}
    partitions = psutil.disk_partitions()
    for partition in partitions:
        try:
            usage = psutil.disk_usage(partition.mountpoint)
            disk_info[partition.device] = {
                'Mount Point': partition.mountpoint,
                'File System': partition.fstype,
                'Total Size': f"{usage.total / (1024**3):.2f} GB",
                'Used': f"{usage.used / (1024**3):.2f} GB",
                'Free': f"{usage.free / (1024**3):.2f} GB",
                'Usage': f"{usage.percent}%"
            }
        except:
            continue
    return disk_info

def get_battery_info():
    """Get battery information"""
    battery_info = {}
    if hasattr(psutil, "sensors_battery"):
        battery = psutil.sensors_battery()
        if battery:
            battery_info = {
                'Percent': f"{battery.percent}%",
                'Power Plugged': battery.power_plugged,
                'Time Left': f"{battery.secsleft / 3600:.2f} hours" if battery.secsleft != -1 else "Unknown"
            }
    return battery_info

def get_network_info():
    """Get network information"""
    network_info = {}
    for interface, addrs in psutil.net_if_addrs().items():
        network_info[interface] = []
        for addr in addrs:
            network_info[interface].append({
                'Family': addr.family.name,
                'Address': addr.address,
                'Netmask': addr.netmask if addr.netmask else 'N/A',
                'Broadcast': addr.broadcast if addr.broadcast else 'N/A'
            })
    return network_info

def get_gpu_info():
    """Get GPU information"""
    gpu_info = []
    try:
        w = wmi.WMI()
        for gpu in w.Win32_VideoController():
            # Convert bytes to GB for video memory
            video_memory_gb = (int(gpu.AdapterRAM) / (1024**3)) if gpu.AdapterRAM else None
            
            gpu_data = {
                'Name': gpu.Name,
                'Manufacturer': gpu.AdapterCompatibility or 'N/A',
                'Video Processor': gpu.VideoProcessor or 'N/A',
                'Video Memory': f"{video_memory_gb:.2f} GB" if video_memory_gb else 'N/A',
                'Driver Version': gpu.DriverVersion or 'N/A',
                'Driver Date': gpu.DriverDate or 'N/A',
                'Video Memory Type': gpu.VideoMemoryType or 'N/A',
                'Current Resolution': f"{gpu.CurrentHorizontalResolution}x{gpu.CurrentVerticalResolution}" if gpu.CurrentHorizontalResolution else 'N/A',
                'Refresh Rate': f"{gpu.CurrentRefreshRate} Hz" if gpu.CurrentRefreshRate else 'N/A',
                'DAC Type': gpu.AdapterDACType or 'N/A',
                'Bits Per Pixel': gpu.CurrentBitsPerPixel or 'N/A'
            }
            gpu_info.append(gpu_data)
    except Exception as e:
        gpu_info.append({'Error': f"Could not retrieve GPU information: {str(e)}"})
    return gpu_info

def get_display_info():
    """Get display information"""
    display_info = []
    try:
        w = wmi.WMI()
        # Get monitor information
        for monitor in w.Win32_DesktopMonitor():
            display_data = {
                'Name': monitor.Name or 'N/A',
                'Manufacturer': monitor.MonitorManufacturer or 'N/A',
                'Model': monitor.MonitorType or 'N/A',
                'Screen Height': f"{monitor.ScreenHeight} pixels" if monitor.ScreenHeight else 'N/A',
                'Screen Width': f"{monitor.ScreenWidth} pixels" if monitor.ScreenWidth else 'N/A'
            }
            display_info.append(display_data)
            
        # Add display configuration information
        for display in w.Win32_VideoController():
            display_data = {
                'Display Adapter': display.Name or 'N/A',
                'Current Resolution': f"{display.CurrentHorizontalResolution}x{display.CurrentVerticalResolution}" if display.CurrentHorizontalResolution else 'N/A',
                'Refresh Rate': f"{display.CurrentRefreshRate} Hz" if display.CurrentRefreshRate else 'N/A',
                'Color Depth': f"{display.CurrentBitsPerPixel} bits" if display.CurrentBitsPerPixel else 'N/A'
            }
            display_info.append(display_data)
    except Exception as e:
        display_info.append({'Error': f"Could not retrieve display information: {str(e)}"})
    return display_info

def get_input_devices_info():
    """Get information about input devices"""
    input_devices = []
    try:
        w = wmi.WMI()
        # Get keyboard information
        for keyboard in w.Win32_Keyboard():
            input_devices.append({
                'Device Type': 'Keyboard',
                'Name': keyboard.Name or 'N/A',
                'Description': keyboard.Description or 'N/A',
                'PNP Device ID': keyboard.PNPDeviceID or 'N/A'
            })
        
        # Get mouse/touchpad information
        for mouse in w.Win32_PointingDevice():
            input_devices.append({
                'Device Type': 'Pointing Device',
                'Name': mouse.Name or 'N/A',
                'Description': mouse.Description or 'N/A',
                'PNP Device ID': mouse.PNPDeviceID or 'N/A'
            })
    except Exception as e:
        input_devices.append({'Error': f"Could not retrieve input devices information: {str(e)}"})
    return input_devices

def get_io_ports_info():
    """Get information about I/O ports"""
    io_ports = []
    try:
        w = wmi.WMI()
        # Get USB ports
        for usb in w.Win32_USBController():
            io_ports.append({
                'Port Type': 'USB',
                'Name': usb.Name or 'N/A',
                'Description': usb.Description or 'N/A',
                'Manufacturer': usb.Manufacturer or 'N/A'
            })
        
        # Get serial ports
        for serial in w.Win32_SerialPort():
            io_ports.append({
                'Port Type': 'Serial',
                'Name': serial.Name or 'N/A',
                'Description': serial.Description or 'N/A'
            })
    except Exception as e:
        io_ports.append({'Error': f"Could not retrieve I/O ports information: {str(e)}"})
    return io_ports

def get_power_info():
    """Get power and charging information"""
    power_info = {}
    try:
        w = wmi.WMI()
        # Get power supply information
        for power in w.Win32_Battery():
            power_info = {
                'Name': power.Name or 'N/A',
                'Description': power.Description or 'N/A',
                'Status': power.Status or 'N/A',
                'Battery Type': power.Chemistry or 'N/A',
                'Design Capacity': power.DesignCapacity or 'N/A',
                'Full Charge Capacity': power.FullChargeCapacity or 'N/A',
                'Health Status': f"{(power.FullChargeCapacity / power.DesignCapacity * 100):.2f}%" if (power.FullChargeCapacity and power.DesignCapacity) else 'N/A',
                'Voltage': f"{power.DesignVoltage/1000:.2f}V" if power.DesignVoltage else 'N/A',
                'Expected Life': power.ExpectedLife or 'N/A',
                'Smart Battery Version': power.SmartBatteryVersion or 'N/A'
            }
            
            # Get current battery status
            battery = psutil.sensors_battery()
            if battery:
                power_info.update({
                    'Current Charge': f"{battery.percent}%",
                    'Power Plugged': "Yes" if battery.power_plugged else "No",
                    'Time Left': f"{battery.secsleft / 3600:.2f} hours" if battery.secsleft != -1 else "Unknown"
                })
    except Exception as e:
        power_info = {'Error': f"Could not retrieve power information: {str(e)}"}
    return power_info

def create_pdf_report():
    """Create PDF report with all system information"""
    # Create PDF document
    doc = SimpleDocTemplate("pc_status_report.pdf", pagesize=letter, rightMargin=30, leftMargin=30)
    styles = getSampleStyleSheet()
    
    # Create custom styles
    styles.add(ParagraphStyle(
        name='TableCell',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        wordWrap='CJK'
    ))
    
    story = []

    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=1
    )
    story.append(Paragraph("PC Status Report", title_style))
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 20))

    def create_table_with_data(title, data, is_nested=False):
        """Helper function to create formatted tables"""
        story.append(Paragraph(title, styles['Heading2']))
        
        if isinstance(data, dict):
            if 'Error' in data:
                story.append(Paragraph(data['Error'], styles['Normal']))
                story.append(Spacer(1, 10))
                return
                
            table_data = []
            for k, v in data.items():
                if isinstance(v, str):
                    # Wrap both key and value in Paragraph for proper word wrapping
                    table_data.append([
                        Paragraph(str(k), styles['TableCell']),
                        Paragraph(str(v), styles['TableCell'])
                    ])
            
            if table_data:
                # Adjust column widths based on content
                table = Table(table_data, colWidths=[2.5*inch, 4.5*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(table)
                story.append(Spacer(1, 10))
        
        elif isinstance(data, list):
            for idx, item in enumerate(data):
                if isinstance(item, dict):
                    if 'Error' in item:
                        story.append(Paragraph(item['Error'], styles['Normal']))
                        story.append(Spacer(1, 10))
                        continue
                        
                    if not is_nested:
                        story.append(Paragraph(f"{title} {idx + 1}", styles['Heading3']))
                    
                    table_data = []
                    for k, v in item.items():
                        if isinstance(v, str):
                            table_data.append([
                                Paragraph(str(k), styles['TableCell']),
                                Paragraph(str(v), styles['TableCell'])
                            ])
                    
                    if table_data:
                        table = Table(table_data, colWidths=[2.5*inch, 4.5*inch])
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
                            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 8),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                            ('TOPPADDING', (0, 0), (-1, -1), 6),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ]))
                        story.append(table)
                        story.append(Spacer(1, 10))

    # System Information
    create_table_with_data("System Information", get_system_info())
    
    # CPU Information
    create_table_with_data("CPU Information", get_cpu_info())
    
    # GPU Information
    create_table_with_data("GPU Information", get_gpu_info())
    
    # Memory Information
    create_table_with_data("Memory Information", get_memory_info())
    
    # Display Information
    create_table_with_data("Display Information", get_display_info())
    
    # Power and Battery Information
    create_table_with_data("Power and Battery Information", get_power_info())
    
    # Disk Information
    disk_info = get_disk_info()
    for disk, info in disk_info.items():
        create_table_with_data(f"Disk: {disk}", info)
    
    # Network Information
    network_info = get_network_info()
    for interface, addrs in network_info.items():
        create_table_with_data(f"Network Interface: {interface}", addrs, is_nested=True)
    
    # Input Devices
    create_table_with_data("Input Devices", get_input_devices_info())
    
    # I/O Ports
    create_table_with_data("I/O Ports", get_io_ports_info())

    # Build PDF
    doc.build(story)
    print("PDF report has been generated successfully as 'pc_status_report.pdf'")

if __name__ == "__main__":
    create_pdf_report() 