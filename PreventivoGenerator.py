import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from tkinter import Tk, filedialog
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import Image
from reportlab.lib import utils
import os
import sys



def get_logo_path(logo_filename):
    script_dir = os.path.dirname(sys.argv[0])
    logo_path = os.path.join(script_dir, logo_filename)
    return logo_path

def get_transparent_path(transparent_filename):
    script_dir = os.path.dirname(sys.argv[0])
    transparent_path = os.path.join(script_dir, transparent_filename)
    return transparent_path

def scale_logo(logo_path, max_width, max_height):
    img = utils.ImageReader(logo_path)
    img_width, img_height = img.getSize()

    # Calculate the scaling factor for the logo
    width_ratio = max_width / img_width
    height_ratio = max_height / img_height
    scaling_factor = min(width_ratio, height_ratio)

    # Scale the logo dimensions
    scaled_width = img_width * scaling_factor
    scaled_height = img_height * scaling_factor

    return scaled_width, scaled_height

# Prompt the user to select an Excel file
Tk().withdraw()  # Hide the main window
excel_file = filedialog.askopenfilename()

if not excel_file:
    print("No file selected. Exiting...")
    exit()

# Read the Excel file
df = pd.read_excel(excel_file)

# Get the directory of the selected Excel file
excel_dir = os.path.dirname(excel_file)

# Create a PDF file
pdf_file = os.path.join(excel_dir, 'PREVENTIVO.pdf')


# Generate a plot of the DataFrame
plt.figure(figsize=(10, 6))
plt.axis('off')
plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center', colWidths=[0.15] * len(df.columns))

# Save the plot as an image
temp_image_file = excel_dir
plt.savefig(temp_image_file)
plt.close()

# Add logos and Excel table to the PDF
doc = SimpleDocTemplate(pdf_file, pagesize=letter)

elements = []

# Add logo 1 to the top left
logo1_path = get_logo_path('image.png')
logo1_width, logo1_height = scale_logo(logo1_path, max_width=100, max_height=100)
logo1 = Image(logo1_path, width=logo1_width, height=logo1_height)

# Add logo 2 to the top right
logo2_path = get_logo_path('image2.jpeg')
logo2_width, logo2_height = scale_logo(logo2_path, max_width=100, max_height=100)
logo2 = Image(logo2_path, width=logo2_width, height=logo2_height)

# Adjust the top margin of the PDF document
doc.topMargin = 20  # Set the desired top margin value

# Create a table with two cells
logo_table_data = [[logo1, logo2]]
logo_table = Table(logo_table_data, colWidths=[doc.width / 2, doc.width / 2])
logo_table.setStyle(TableStyle([
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
    ('ALIGN', (1, 0), (1, 0), 'RIGHT')
]))

# Add the logo table to the elements list
elements.append(logo_table)

# Create a spacer table between the logos and the spreadsheet table
spacer_height = 30  # Set the desired height of the spacer
spacer_image_path = get_transparent_path('transparent.png')  # Path to a transparent image or a 1x1 pixel transparent image
spacer_table_data = [[Image(spacer_image_path, width=1, height=spacer_height)]]
spacer_table = Table(spacer_table_data)
spacer_table.setStyle(TableStyle([('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
                                  ('BACKGROUND', (0, 0), (-1, -1), colors.white)]))

# Add the spacer table to the elements list
elements.append(spacer_table)


# Add the Excel table
table_data = [df.columns.tolist()] + df.values.tolist()
table = Table(table_data)
table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
                           ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                           ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                           ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                           ('FONTSIZE', (0, 0), (-1, 0), 12),
                           ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                           ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                           ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
elements.append(table)

# Build the PDF document
doc.build(elements)

print(f"PDF file '{pdf_file}' generated successfully.")