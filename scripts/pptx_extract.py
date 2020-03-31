def pptx_extract(path,filename):    
    
    from pptx import Presentation
    from figure_extract import figure_extract
    from table_extract import table_extract

    from pathlib import Path


    prs = Presentation(path+"/"+filename)

    image_type_plc = Presentation("files/ppt_presentations/ppt_figure.pptx").slides[0].placeholders[1] #needed to identify when one placeholder is a picture
    image_type_shp = Presentation("files/ppt_presentations/ppt_figure.pptx").slides[0].shapes[4] #needed to identify when one placeholder is a picture
    
    slide_text = []
    n=0

    author = prs.core_properties.author
    #title = prs.core_properties.title
    title = filename
    slide_text = ['---','\n','layout: presentation','\n','author: ', author,'\n','title: ',title]

    Path("_presentations/figures/"+title).mkdir(parents=True, exist_ok=True) #check if destination folder for pictures exists and/or creates it

    for slide in prs.slides:
        slide_text.append("\n---\n#")  # new slide, new line, TITLE --- Append is used to add a value at the end of the string
        slide_text.append(slide.shapes.title.text)
        slide_text.append("\n")
        
        n_plc=len(slide.placeholders)
        n_shp=len(slide.shapes)
                
        if n_plc==n_shp:
            for plc in slide.placeholders:
                if plc.has_text_frame and not plc == slide.shapes.title:
                    for paragraph in plc.text_frame.paragraphs:
                        slide_text.append("\n\n")
                        for run in paragraph.runs:
                            slide_text.append(run.text)
                elif plc == slide.shapes.title:
                    continue
                elif plc.has_table:
                     table_extract(plc, slide_text)
                elif plc.has_chart:
                    slide_text.append("***MISSING CHART*** insert manually \n")
                    
                elif type(plc)==type(image_type_plc):
                    n=n+1
                    figure_extract(plc, slide_text, n, title)
                else:
                    slide_text.append("***MISSING OBJECT*** insert manually \n")
                    continue
        else:
            for shape in slide.shapes:
                if shape.has_text_frame and not shape == slide.shapes.title:
                    for paragraph in shape.text_frame.paragraphs:
                        slide_text.append("\n\n")
                        for run in paragraph.runs:
                            slide_text.append(run.text)
                elif shape == slide.shapes.title:
                    continue
                elif shape.has_table:
                    table_extract(shape, slide_text)
                elif shape.has_chart:
                    slide_text.append("***MISSING CHART*** insert manually \n")
                    
                elif type(shape)==type(image_type_shp): # <class 'pptx.shapes.placeholder.PlaceholderPicture'>
                    n=n+1
                    figure_extract(shape, slide_text, n, title)
                else:
                    slide_text.append("***MISSING OBJECT*** insert manually \n")
                    continue

    
    with open("_presentations/"+filename+".html","w") as presentation_file:
        presentation_file.writelines(slide_text)

