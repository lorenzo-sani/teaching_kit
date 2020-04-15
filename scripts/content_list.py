def content_list():
    
    import os, frontmatter

    # cycle through presentations in folder
    # collcet names and list them¨

    path = '_presentations'

    text = []
    folder = os.fsencode(path)
    
    # LECTURES
    text.append("### Lectures \nHere is the list of the lectures available: \n\n")
    for file in os.listdir(folder):
        
        filename = os.fsdecode(file)
        title, file_extension = os.path.splitext(filename)
        if file_extension==".html":
            try:
                prs = frontmatter.load(path+"/"+filename)
                prs_title = prs['title']
            except:
                prs_title = title
            text.append("\n- ["+prs_title+"](presentations/"+filename+") ") 

    # MODULES
    text.append("\n\n\n### Modules \nHere is the list of the modules available: \n\n")
    path = '_presentations/modules'
    folder = os.fsencode(path)
    for file in os.listdir(folder):
        
        filename = os.fsdecode(file)
        title, file_extension = os.path.splitext(filename)
        if file_extension==".html":
            try: 
                prs = frontmatter.load(path+"/"+filename)
                prs_title = prs['title']
            except:
                prs_title = title    
            text.append("\n- ["+title+"](presentations/modules/"+filename+")")

    with open("_includes/content_list.md","w") as file:
            file.writelines(text)

    print('Updated the content list')



