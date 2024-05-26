import fitz
import os

def mergePDFs():
    mylist = os.listdir('output')
    mylist.sort()
    try:
        mylist.remove(".DS_Store")
    except:
        pass
    # print(mylist)

    doc = fitz.open()

    for filename in mylist:
        doc.insert_file('output/' + filename)  # appends it to the end

    doc.save("_combined.pdf")
    doc.close()
    print("PDFs Merged")

    return True
