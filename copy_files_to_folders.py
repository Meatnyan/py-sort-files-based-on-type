from typing import List
import os
from os import listdir
from os.path import isfile, join
from collections import OrderedDict


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
'Type "exit" (no quotation marks) or just hit enter without typing anything to close the script and not do anything.\n\n'
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
for extension in chosenExtensions:
    extension = extension.replace(',', '').replace('.', '')

RemoveDuplicates(chosenExtensions)

print('chosenExtensions before:\n' + str(chosenExtensions))


ReplaceGroupNamesWithTheirValues(chosenExtensions)

RemoveDuplicates(chosenExtensions)  # in case one of the extensions from the extension groups was explicitly chosen by the user already


for i in range(0, len(chosenExtensions)):
    chosenExtensions[i] = '.' + chosenExtensions[i] # txt into .txt, doc into .doc, etc.


print('chosenExtensions after:\n' + str(chosenExtensions))


pathToCurDir = os.path.dirname(os.path.realpath(__file__))
chosenFilePaths = [join(pathToCurDir, curFilename) for curFilename in listdir(pathToCurDir) if (isfile(join(pathToCurDir, curFilename))
and curFilename.endswith(tuple(chosenExtensions)))]

