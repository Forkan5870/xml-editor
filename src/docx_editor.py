import docx

def CreateDocx(array)
    # create an instance of a word document
    doc = docx.Document()


    # add a heading of level 0 (largest heading)
    doc.add_heading('Heading for the document', 0) # HAPPENS

    # Add a paragraph with "Hello, World" in Verdana font and "World" bolded
    paragraph1 = doc.add_paragraph()
    run1 = paragraph1.add_run('Hello, ')
    run1.font.name = 'Verdana'
    run1 = paragraph1.add_run('World')
    run1.font.name = 'Verdana'
    run1.bold = True

    # Add a paragraph with "whats up" in Assistant font and "up" bolded
    paragraph2 = doc.add_paragraph()
    run2 = paragraph2.add_run('whats ')
    run2.font.name = 'Assistant'
    run2 = paragraph2.add_run('up')
    run2.font.name = 'Assistant'
    run2.bold = True

    run1 = paragraph2.add_run('Im back in verdana')
    run1.font.name = 'Verdana'


    
    # add a page break to start a new page 
    doc.add_page_break() 

    
    # now save the document to a location 
    doc.save('files/test.docx') 