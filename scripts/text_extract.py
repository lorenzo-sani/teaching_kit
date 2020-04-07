def text_extract(shape, text, n):
    width=str(shape.width.pt)
    height=str(shape.height.pt)
    top=str(shape.top.pt)
    left=str(shape.left.pt)
    style = "box"+str(n)
    text.append("\n<style> div."+style+" {position: absolute; left: "+left+"px; top:"+top+"px; width: "+width+"px; height: "+height+"} </style>\n<div class="+style+">")
    #text.append(shape.text_frame.text)

    NoneType = type(None)

    l=0
    for paragraph in shape.text_frame.paragraphs:
        text.append("\n<br>\n")
        
        l=l+1
        
        skip=0
        try: paragraph.runs[0].font.size
        except: skip=1

        if l==1 and skip==0 and isinstance(paragraph.runs[0].font.size, NoneType):
            size = 18
        elif l==1 and skip==0:
            size = paragraph.runs[0].font.size.pt
        elif l==1 and skip==1:
            size = 18


        text.append("<div style=font-size:"+str(size)+"px>")

        for run in paragraph.runs:
            
            
            if run.font.bold:
                text.append("<b>"+run.text+"</b>")
            elif run.font.italic:
                text.append("<i>"+run.text+"</i>")
            else:
                text.append(run.text)
            
        text.append("</div>")
    text.append("\n</div>\n")
    