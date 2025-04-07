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

# Initialize WMI
wmi_instance = wmi.WMI()

def get_system_info():
    """Get basic system information"""
    return {
        'OS': platform.system(),
        'Node Name': platform.node(),
        'Release': platform.release(),
        'Version': platform.version(),
        'Architecture': platform.machine(),
        'Processor': platform.processor() or 'N/A'
    }

def get_cpu_info():
    """Get CPU information with better error handling"""
    try:
        cpu = wmi_instance.Win32_Processor()[0]
        freq = psutil.cpu_freq()
        return {
            'Name': cpu.Name.strip(),
            'Manufacturer': cpu.Manufacturer,
            'Cores': psutil.cpu_count(logical=False),
            'Threads': psutil.cpu_count(logical=True),
            'Max Speed': f"{cpu.MaxClockSpeed} MHz",
            'Current Speed': f"{freq.current:.2f} MHz" if freq else 'N/A',
            'Usage': f"{psutil.cpu_percent(interval=1)}%",
            'L2 Cache': f"{cpu.L2CacheSize} KB" if cpu.L2CacheSize else 'N/A',
            'L3 Cache': f"{cpu.L3CacheSize} KB" if cpu.L3CacheSize else 'N/A'
        }
    except Exception as e:
        return {'Error': f"CPU Info Error: {str(e)}"}

def get_memory_info():
    """Get memory information with cleaner formatting"""
    mem = psutil.virtual_memory()
    return {
        'Total': f"{mem.total / (1024**3):.2f} GB",
        'Available': f"{mem.available / (1024**3):.2f} GB",
        'Used': f"{mem.used / (1024**3):.2f} GB",
        'Percentage': f"{mem.percent}%",
        'Swap Total': f"{psutil.swap_memory().total / (1024**3):.2f} GB"
    }

def get_disk_info():
    """Get disk information with improved structure"""
    disk_info = {}
    for part in psutil.disk_partitions():
        try:
            usage = psutil.disk_usage(part.mountpoint)
            disk_info[part.device] = {
                'Mount': part.mountpoint,
                'Type': part.fstype,
                'Total': f"{usage.total / (1024**3):.2f} GB",
                'Used': f"{usage.used / (1024**3):.2f} GB",
                'Free': f"{usage.free / (1024**3):.2f} GB",
                'Percent': f"{usage.percent}%"
            }
        except Exception:
            continue
    return disk_info

def get_gpu_info():
    """Get GPU information with better detail"""
    gpu_info = []
    try:
        for gpu in wmi_instance.Win32_VideoController():
            vram = gpu.AdapterRAM / (1024**3) if gpu.AdapterRAM else 0
            gpu_info.append({
                'Name': gpu.Name or 'N/A',
                'Manufacturer': gpu.AdapterCompatibility or 'N/A',
                'VRAM': f"{vram:.2f} GB",
                'Resolution': f"{gpu.CurrentHorizontalResolution}x{gpu.CurrentVerticalResolution}" 
                             if gpu.CurrentHorizontalResolution else 'N/A',
                'Driver': gpu.DriverVersion or 'N/A'
            })
    except Exception as e:
        gpu_info.append({'Error': f"GPU Info Error: {str(e)}"})
    return gpu_info

def get_battery_health_info():
    """Get detailed battery health and charging information with type conversion"""
    try:
        battery = psutil.sensors_battery()
        power = wmi_instance.Win32_Battery()[0] if wmi_instance.Win32_Battery() else None
        
        # Convert WMI values to float/int, handling None or string cases
        design_capacity = float(power.DesignCapacity) if power and power.DesignCapacity else None
        full_capacity = float(power.FullChargeCapacity) if power and power.FullChargeCapacity else None
        design_voltage = float(power.DesignVoltage) if power and power.DesignVoltage else None
        
        # Calculate health percentage safely
        health = None
        if design_capacity and full_capacity and design_capacity > 0:
            health = f"{(full_capacity / design_capacity * 100):.1f}%"
        
        health_info = {
            'Charge': f"{battery.percent}%" if battery else 'N/A',
            'Plugged In': 'Yes' if battery and battery.power_plugged else 'No',
            'Time Left': f"{battery.secsleft / 3600:.2f} hrs" if battery and battery.secsleft != -1 else 'N/A',
            'Battery Type': power.Chemistry if power and power.Chemistry else 'N/A',
            'Design Capacity': f"{design_capacity} mWh" if design_capacity else 'N/A',
            'Full Capacity': f"{full_capacity} mWh" if full_capacity else 'N/A',
            'Health': health if health else 'N/A',
            'Voltage': f"{design_voltage/1000:.2f}V" if design_voltage else 'N/A'
        }
    except Exception as e:
        health_info = {'Error': f"Battery Info Error: {str(e)}"}
    return health_info

def create_pdf_report():
    """Create an enhanced PDF report"""
    doc = SimpleDocTemplate("pc_status_report.pdf", pagesize=letter, 
                          topMargin=0.5*inch, bottomMargin=0.5*inch,
                          leftMargin=0.5*inch, rightMargin=0.5*inch)
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(name='TableCell', fontSize=8, leading=10, wordWrap='CJK'))
    styles.add(ParagraphStyle(name='Footer', fontSize=8, alignment=2, textColor=colors.grey))
    
    story = []
    
    # Header
    story.append(Paragraph("PC Status Report", styles['Heading1']))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 
                          styles['Normal']))
    story.append(Spacer(1, 0.25*inch))

    def add_section(title, data, nested=False):
        """Helper to create consistent tables"""
        story.append(Paragraph(title, styles['Heading2']))
        if not data:
            story.append(Paragraph("No data available", styles['Normal']))
            return
            
        if isinstance(data, dict) and 'Error' in data:
            story.append(Paragraph(data['Error'], styles['Normal']))
            return
            
        table_data = []
        if isinstance(data, dict):
            for key, val in data.items():
                table_data.append([Paragraph(str(key), styles['TableCell']),
                                 Paragraph(str(val), styles['TableCell'])])
        elif isinstance(data, list):
            for i, item in enumerate(data):
                if nested:
                    story.append(Paragraph(f"{title} {i+1}", styles['Heading3']))
                for key, val in item.items():
                    table_data.append([Paragraph(str(key), styles['TableCell']),
                                     Paragraph(str(val), styles['TableCell'])])
                if nested and i < len(data) - 1:
                    story.append(Spacer(1, 0.1*inch))
        
        if table_data:
            table = Table(table_data, colWidths=[2*inch, 4*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BOX', (0,0), (-1,-1), 1, colors.black),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
            ]))
            story.append(table)
        story.append(Spacer(1, 0.2*inch))

    # Add all sections
    add_section("System Information", get_system_info())
    add_section("CPU Information", get_cpu_info())
    add_section("GPU Information", get_gpu_info(), nested=True)
    add_section("Memory Information", get_memory_info())
    add_section("Battery Health & Charging", get_battery_health_info())
    for disk, info in get_disk_info().items():
        add_section(f"Disk: {disk}", info)

    # Custom footer with GitHub link
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(doc.pagesize[0] - 0.5*inch, 0.25*inch, 
                             "github.com/akasumitlamba")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    print("PDF report generated as 'pc_status_report.pdf'")

if __name__ == "__main__":
    create_pdf_report()
