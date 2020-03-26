from pptx import Presentation

prs = Presentation("files/ppt_demo.pptx")


image_type = Presentation("files/ppt_figure.pptx").slides[0].placeholders[1] #needed to
#graphicframe = Presentation("files/ppt_figure.pptx").prs1.slides[2].placeholders[1]
# text_runs will be populated with a list of strings,
# one for each text run in presentation
#text_runs = []


# slide_num = len(prs.slides)  # to calculate slides number
slide_text = []

author = prs.core_properties.author
title = prs.core_properties.title
slide_text = ['---','\n','layout: presentation','\n','author: ', author,'\n','title: ',title]
n=0
for slide in prs.slides:
    slide_text.append("\n---\n#")  # new slide, new line, TITLE --- Append is used to add a value at the end of the string
    slide_text.append(slide.shapes.title.text)
    slide_text.append("\n")
    
    n_plc=len(slide.placeholders)
    n_shp=len(slide.shapes)
    #print(n_shp, n_plc)
    n=n+1
    if n_plc==n_shp:
        #print('oke slide ',n)
        for plc in slide.placeholders:
            if plc.has_text_frame and not plc == slide.shapes.title:
                for paragraph in plc.text_frame.paragraphs:
                    slide_text.append("\n\n")
                    for run in paragraph.runs:
                        slide_text.append(run.text)
            elif plc == slide.shapes.title:
                continue
            elif plc.has_table:
                tab=plc.table
                rows=tab.rows
                n_rows=len(rows)
                col=tab.columns
                n_col=len(col)
                for r in range(n_rows):
                    slide_text.append("\n |")
                    for c in range(n_col):
                        if r == 1 and c == 0: # this step is needed to add the intermediate row 
                            slide_text.append(":---|---:| ---:| ---:| ---:|\n")
                            slide_text.append(" |"+tab.cell(r,c).text+" |")
                        else:
                            slide_text.append(tab.cell(r,c).text+" |")
            elif plc.has_chart:
                slide_text.append("***MISSING CHART*** insert manually \n")
                
            elif type(plc)==type(image_type):
                img=plc.image
                name=img.filename
                with open("_presentations/figures/"+name, 'wb') as presentation_image: 
                    presentation_image.write(img.blob)
                width=str(plc.width.pt)
                height=str(plc.height.pt)
                top=str(plc.top.pt)
                left=str(plc.left.pt)
                slide_text.append("<img src=figures/image.png position=absolute top="+top+"px left="+left+"px width="+width+"px height="+height+"px />")
            else:
                slide_text.append("***MISSING FIGURE*** insert manually \n")
                continue
    else:
        #print('NO slide ',n)
        for shape in slide.shapes:
            if shape.has_text_frame and not shape == slide.shapes.title:
                for paragraph in shape.text_frame.paragraphs:
                    slide_text.append("\n\n")
                    for run in paragraph.runs:
                        slide_text.append(run.text)
            elif shape == slide.shapes.title:
                continue
            elif shape.has_table:
                tab=shape.table
                rows=tab.rows
                n_rows=len(rows)
                col=tab.columns
                n_col=len(col)
                for r in range(n_rows):
                    slide_text.append("\n |")
                    for c in range(n_col):
                        if r == 1 and c == 0: # this step is needed to add the intermediate row 
                            slide_text.append(":---|---:| ---:| ---:| ---:|\n")
                            slide_text.append(" |"+tab.cell(r,c).text+" |")
                        else:
                            slide_text.append(tab.cell(r,c).text+" |")
            elif shape.has_chart:
                 slide_text.append("***MISSING CHART*** insert manually \n")
                
            elif type(shape)==type(image_type):
                img=shape.image
                name=img.filename
                with open("_presentations/figures/"+name, 'wb') as presentation_image: 
                    presentation_image.write(img.blob)
                width=str(shape.width.pt)
                height=str(shape.height.pt)
                top=str(shape.top.pt)
                left=str(shape.left.pt)
                slide_text.append("<img src=figures/image.png position=absolute top="+top+"px left="+left+"px width="+width+"px height="+height+"px />")
            else:
                slide_text.append("***MISSING FIGURE*** insert manually \n")
                continue



#file = open("_presentations/"+prs.core_properties.title+".html","w")

#for runs in range(len(slide_text)):
#    file.write(slide_text[runs]),
#file.close()


with open("_presentations/"+prs.core_properties.title+".html","w") as presentation_file:
    presentation_file.writelines(slide_text)

# you save files with:
# prs.save('test.pptx')


# transform table.py in one function and use it within the script to create a table
# def read_table(): 
#
