#from pptx_extract import pptx_extract
#prs = Presentation("files/ppt_demo.pptx")
#n=0
#for slide in prs.slides:
#    plc=len(slide.placeholders)
#    shp=len(slide.shapes)
#    n=n+1
#    if plc==shp:
#        print('oke slide ',n)
#    else:
#        print('no slide ',n)    

import os
import codecs

path = '_presentations'

folder = os.fsencode(path)

tag = []
t = []
for file in os.listdir(folder):
    filename = os.fsdecode(file)
    title, file_extension = os.path.splitext(filename)
    print(file_extension)
    if file_extension==".html":
        {{ file.url | relative_url }}


for presentation in site.presentations 

    <h4>
      <a href="{{ presentation.url | relative_url }}">{{ presentation.title | escape }} <body> - tag: {{ presentation.tag }}</body></a>
    </h4>

{% endfor %}





#import os
#import scripts.pptx_extract_fun
#path = 'files/ppt_presentations'

#folder = os.fsencode(path)

#filenames = []
#slide_text = []

#for file in os.listdir(folder):
 #   filename = os.fsdecode(file)
    #if filename.endswith( ('.jpeg', '.png', '.gif') ): # whatever file types you're using...
    #filenames.append(filename)

 #   pptx_extract(path+"/"+filename)
#

    #prs=Presentation(path+"/"+filename)
    #author = prs.core_properties.author
    #print(author)




#for presentation_file in files/ppt_presentations:
    
    #prs = Presentation("files/ppt_demo.pptx")