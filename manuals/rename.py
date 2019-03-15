#! pyhton3

import os, glob
import fire

#Create list of .pdf files
os.chdir(".")
list_folder_module = [folder for folder in glob.glob('*/**/*.pdf') if folder.startswith("#")]

#Loop through the files

def main():
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

if __name__ == "__main__":
    fire.Fire(main)
