import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

# Step 1: Create the bar graph
data = {'Category A': 10, 'Category B': 20, 'Category C': 15}
categories = list(data.keys())
values = list(data.values())

plt.bar(categories, values)
plt.xlabel('Categories')
plt.ylabel('Values')
plt.title('Bar Graph')

# Step 2: Export the graph as an image
plt.savefig('bar_graph.png')
plt.close()

# Step 3: Open the PowerPoint presentation
ppt = Presentation('presentation.pptx')

# Step 4: Find and delete the existing text box (if any)
for slide in ppt.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if 'Replace this text' in paragraph.text:
                    sp = shape
                    slide.shapes._spTree.remove(sp._element)

# Step 5: Insert the created bar graph image
img_path = 'bar_graph.png'
left_inch = Inches(1)
top_inch = Inches(1)
width_inch = Inches(5)
height_inch = Inches(3)

ppt.slides[0].shapes.add_picture(img_path, left_inch, top_inch, width_inch, height_inch)

# Save the modified presentation
ppt.save('modified_presentation.pptx')
