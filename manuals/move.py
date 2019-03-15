#! pyhton3

import subprocess, glob, shutil, pprint

def listing_pdf(location = '*/**/#*.pdf'):
    return glob.glob(location)

def move_pdf(number_of_pdf):
    if len(number_of_pdf) == 0:
        print('[-]: NO CONTENT... only <#*.pdf> format accepted') 
    else:
        pprint.pprint(number_of_pdf)
        for pdf in number_of_pdf:
            shutil.move(pdf, '.') 
        print(f'task completed: {len(number_of_pdf)} files moved')


pdfs = listing_pdf()

if __name__ == "__main__":
    move_pdf(pdfs)
