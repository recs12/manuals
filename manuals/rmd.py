import os, glob
import click
import subprocess, glob, shutil, pprint
from pathlib import Path

#Create list of .pdf files
list_folder_module = [folder for folder in glob.glob('*/**/*.pdf') if folder.startswith("#")]

@click.group()
def cli():
    pass

@cli.command()
def renaming():
    for i in range(len(list_folder_module)):
        prevName = str(os.path.basename(list_folder_module[i]))
        list_of_elements = list_folder_module[i].split('\\')
        newName = str(list_of_elements[0])+'.pdf'
        dirName = os.path.dirname(list_folder_module[i])
        absPath = os.path.abspath('.')
        linkComplete = os.path.join(absPath, dirName)
        os.chdir(linkComplete)
        if prevName == newName:
            pass
        else:
            os.rename(prevName, newName)
        os.chdir(absPath)
        print("[RENAME]: {0}\t-------------->\t{1}".format(prevName,newName))


@cli.command()
def move_pdf():
    number_of_pdf = glob.glob('*/**/#*.pdf')
    if len(number_of_pdf) == 0:
        print('[-]: NO CONTENT... only <#*.pdf> format accepted')
    else:
        pprint.pprint(number_of_pdf)
        for pdf in number_of_pdf:
            shutil.move(pdf, '.')
        print(f'task completed: {len(number_of_pdf)} files moved')



if __name__ == "__main__":
    cli()
