################################################
import os

from pptx import Presentation
os.chdir('C:/Users/priya/OneDrive/Desktop/pptx/')
prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "PS was here!"

prs.save('test.pptx')

##############################################################
prs = Presentation()
bullet_slide_layout = prs.slide_layouts[1]

slide = prs.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes

title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = 'Adding a Bullet Slide'

tf = body_shape.text_frame
tf.text = 'Find the bullet slide layout'

p = tf.add_paragraph()
p.text = 'Use _TextFrame.text for first bullet'
p.level = 1

p = tf.add_paragraph()
p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
p.level = 2

prs.save('test1.pptx')
################################################
from pptx import Presentation
from pptx.util import Inches
os.chdir('C:/Users/priya/OneDrive/Desktop/pptx/')
img_path = 'monty-truth.jpg'

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = Inches(1)
pic = slide.shapes.add_picture(img_path, left, top)

left = Inches(5)
height = Inches(5.5)
pic = slide.shapes.add_picture(img_path, left, top)
pic = slide.shapes.add_picture(img_path, left, top, height=height)

prs.save('test3.pptx')
#####################################################
def extract_text_and_position(pptx_file):
    prs = Presentation(pptx_file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                print(f"Text: {text}")
                print(f"Position: Left: {left}, Top: {top}, Width: {width}, Height: {height}")

# Specify the path to your PowerPoint file
pptx_file = "test.pptx"

# Extract text and position information
extract_text_and_position(pptx_file)



##############################
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# Create a sample chart using matplotlib
def create_chart():
    # Sample data
    labels = ['A', 'B', 'C', 'D', 'E']
    values = [23, 45, 56, 78, 90]

    # Create a pie chart
    plt.figure(figsize=(6, 6))
    plt.pie(values, labels=labels, autopct='%1.1f%%')
    plt.title('Sample Chart')

    # Save the chart to a BytesIO object
    chart_bytes = BytesIO()
    plt.savefig(chart_bytes, format='png')
    chart_bytes.seek(0)
    plt.close()

    return chart_bytes

# Insert the chart into a text box in the PowerPoint slide
def insert_chart_into_text_box(pptx_file, chart_bytes):
    prs = Presentation(pptx_file)
    slide = prs.slides[0]  # Assuming you want to insert into the first slide
    left = Inches(1)  # Adjust these values based on your requirements
    top = Inches(2)
    width = Inches(4)
    height = Inches(3)

    # Insert a text box
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = "Chart from Python"

    # Insert the chart image into the text box
    pic = textbox.text_frame.add_paragraph().add_run()
    pic.add_picture = chart_bytes#pic.add_picture(chart_bytes)

    prs.save("output.pptx")

# Path to your PowerPoint file
pptx_file = "test.pptx"

# Create the chart
chart_bytes = create_chart()
chart_bytes
# Insert the chart into a text box in the PowerPoint slide
insert_chart_into_text_box(pptx_file, chart_bytes)

#############################
#Final 

from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import matplotlib.pyplot as plt
import shutil
# Create a sample chart using matplotlib
def create_chart():
    # Sample data
    labels = ['A', 'B', 'C', 'D', 'E']
    values = [23, 45, 56, 78, 90]

    # Create a pie chart
    plt.figure(figsize=(6, 6))
    plt.pie(values, labels=labels, autopct='%1.1f%%')
    plt.title('Sample Chart')

    # Save the chart to a BytesIO object
    chart_bytes = BytesIO()
    plt.savefig(chart_bytes, format='png')
    chart_bytes.seek(0)
    plt.close()

    return chart_bytes

# Replace a text box with the chart in the PowerPoint slide
def replace_text_box_with_chart(Text_Box_Name,chart_bytes,new_ppt):
    prs = Presentation(new_ppt)
    Slide_Number = 0
    for slide_i in prs.slides: # Finding the slide to be modified with text box name
        for shape in slide_i.shapes:
            if shape.has_text_frame and shape.text.strip() == Text_Box_Name:
                # Finalising the slide to be modified
                slide = slide_i  
                
                # Clear the text box content
                shape.text_frame.clear()  

                # Reading the dimesnsions of text box
                left, top, width, heigh = shape.left, shape.top, shape.width, shape.height 
                shape._element.getparent().remove(shape._element)  # Remove the text box

                # Add the chart in the same position
                pic = slide.shapes.add_picture(chart_bytes, left, top, width, height)
                break
        
        prs.save(new_ppt)

# Path to your PowerPoint file
pptx_file = "abc.pptx"
new_ppt = 'final.pptx'

shutil.copy(pptx_file,new_ppt)
print("File copied successfully.")

# Create the chart
chart_bytes = create_chart()

# Replace the text box with the chart in the PowerPoint slide
replace_text_box_with_chart("Abcd", chart_bytes,new_ppt)
chart_bytes = create_chart()
replace_text_box_with_chart("Efgh", chart_bytes,new_ppt)
