from typing import List
import os
from os import mkdir
from os import listdir
from os.path import isfile, isdir, join
from collections import OrderedDict
from shutil import copy2

msExtensions = ['doc', 'docx', 'pptx']

openExtensions = ['txt', 'odf', 'ods']


def RemoveDuplicates(listOfStrings: List[str]):
    listOfStrings = list(OrderedDict.fromkeys(listOfStrings))
    return


def ReplaceGroupNamesWithTheirValues(extensionsList: List[str]):  # extensionsList is mutable, so will be changed as if passed by reference
    while 'ms' in extensionsList:
        extensionsList.remove('ms')
        extensionsList.extend(msExtensions)
    while 'open' in extensionsList:
        extensionsList.remove('open')
        extensionsList.extend(openExtensions)    
    return




print('\nThis script will attempt to create new folders for files with the specified extension.\n\n'
'Type "exit" or "quit" (no quotation marks) or just hit enter without typing anything to close the script and not do anything.\n\n'
'Alternatively, if you wish to proceed, simply type in one of the following:\n'
'1. Names of the extensions you wish to create folders for, separated by spaces. For example:\n'
'txt\n'
'pptx xls doc\n'
'mp3 mp4\n'
'2. Names for groups of extensions (for quicker use). Currently the following groups are supported:\n'
'ms = common Microsoft document file extensions\n'
'open = common "Open Source" (mostly OpenOffice) document file extensions\n'
'These can be used alongside extension names, for example:\n'
'ms exe ods\n\n'
'Type in your command to begin:')


# split the input command (extensions list) into words
chosenExtensions = input().split()

# remove commas and dots from extension names
for i in range(0, len(chosenExtensions)):
    chosenExtensions[i] = chosenExtensions[i].replace(',', '').replace('.', '')

if ('exit' not in chosenExtensions) & ('quit' not in chosenExtensions) & (len(chosenExtensions) > 0):
    RemoveDuplicates(chosenExtensions)

    ReplaceGroupNamesWithTheirValues(chosenExtensions)

    RemoveDuplicates(chosenExtensions)  # in case one of the extensions from the extension groups was explicitly chosen by the user already

    outputDirectoriesNames = chosenExtensions[:]    # cheap hack to copy the list by value, not by reference

    for i in range(0, len(chosenExtensions)):
        chosenExtensions[i] = '.' + chosenExtensions[i] # txt into .txt, doc into .doc, etc. for ease of use with endswith


    pathToCurDir = os.path.dirname(os.path.realpath(__file__))

    # get all paths for files from the current directory with chosen extensions
    chosenFilePaths = [join(pathToCurDir, curFilename) for curFilename in listdir(pathToCurDir) if (isfile(join(pathToCurDir, curFilename))
    and curFilename.endswith(tuple(chosenExtensions)))]


    for filePath in chosenFilePaths:
        fileExtension = filePath[filePath.rfind('.') + 1 : len(filePath)]   # extension name is after the last dot, all the way to the end of the path

        outputDirectory = pathToCurDir + '\\' + fileExtension

        if not isdir(outputDirectory):
            mkdir(outputDirectory)

        copy2(filePath, outputDirectory)    # copy the file (including metadata) to the output directory


    print('Copied files into the following subdirectories: ' + ' '.join(outputDirectoriesNames))
else:
    print('Program exited. No action taken.')