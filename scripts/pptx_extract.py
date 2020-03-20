from pptx import Presentation

prs = Presentation("files/ppt_demo.pptx")

# text_runs will be populated with a list of strings,
# one for each text run in presentation
text_runs = []
text_runs = ['---','\n','layout: presentation']
for slide in prs.slides:
    text_runs.append("\n---\n")
    i=1
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            text_runs.append("\n\n")
            for run in paragraph.runs:
                if i==1:
                    text_runs.append("# ")
                    text_runs.append(run.text)
                    i=i + 1
                else: 
                    text_runs.append(run.text)
                    i=i + 1

                

file = open("_presentations/pres.html","w")
# file.write(repr(text_runs))
# file.close()

for text_run in range(len(text_runs)):
    file.write(text_runs[text_run]),
file.close()
# print(text_runs[0:])


# you save files with:
# prs.save('test.pptx')
                 

#    def run_my_code(arguments):
 #       # <do something with arguments>
  #      print(arguments)
   #     return 0

#    if __name__ == '__main__':
 #       args = sys.argv[1:]
  #      run_my_code(args)

    
