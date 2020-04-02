def figure_extract(shape, text, n, title):
    img=shape.image
    name=str(n)+"_"+img.filename
    with open("_presentations/figures/"+title+"/"+name, 'wb') as presentation_image: 
         presentation_image.write(img.blob)
    width=str(shape.width.pt)
    height=str(shape.height.pt)
    top=str(shape.top.pt)
    left=str(shape.left.pt)
    text.append("\n<img src=figures/"+title+"/"+name+" position=absolute top="+top+"px left="+left+"px width="+width+"px height="+height+"px />")