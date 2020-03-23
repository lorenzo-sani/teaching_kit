from pptx import Presentation

prs = Presentation("files/ppt_demo.pptx")

# text_runs will be populated with a list of strings,
# one for each text run in presentation
#text_runs = []


# slide_num = len(prs.slides)  # to calculate slides number
slide_text = []
slide_text = ['---','\n','layout: presentation','\n','author: '+ prs.core_properties.author,'\n','title: ' + prs.core_properties.title]

for slide in prs.slides:
    slide_text.append("\n---\n#")  # new slide, new line, TITLE --- Append is used to add a value at the end of the string
    slide_text.append(slide.shapes.title.text)
    slide_text.append("\n")
    for shape in slide.shapes:
        if shape.has_text_frame and not shape == slide.shapes.title:
             for paragraph in shape.text_frame.paragraphs:
                 slide_text.append("\n\n")
                 for run in paragraph.runs:
                     slide_text.append(run.text)
        elif shape == slide.shapes.title:
            continue
        else:
            slide_text.append("***MISSING FIGURE*** insert manually \n")
            continue


file = open("_presentations/"+prs.core_properties.title+".html","w")

for runs in range(len(slide_text)):
    file.write(slide_text[runs]),
file.close()

# you save files with:
# prs.save('test.pptx')


# transform table.py in one function and use it within the script to create a table
# def read_table(): 
#
