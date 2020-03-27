def figure_extract(shape, text, n):
    img=shape.image
    name=str(n)+"_"+img.filename
    with open("_presentations/figures/"+name, 'wb') as presentation_image: 
         presentation_image.write(img.blob)
    width=str(shape.width.pt)
    height=str(shape.height.pt)
    top=str(shape.top.pt)
    left=str(shape.left.pt)
    text.append("<img src=figures/"+name+" position=absolute top="+top+"px left="+left+"px width="+width+"px height="+height+"px />")