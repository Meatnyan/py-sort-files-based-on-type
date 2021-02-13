from typing import List
import os
from os import mkdir
from os import listdir
from os.path import isfile, isdir, join
from collections import OrderedDict
from shutil import copy2



msExtensions = ['doc', 'dot', 'wbk', 'docx', 'docm', 'dotx', 'dotm', 'docb', 'xls', 'xlt', 'xlm', 'xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xla', 'xlam', 'xll', 'xlw', 'ppt', 'pot', 'pps', 'pptx', 'pptm',
'potx', 'potm', 'ppam', 'ppsx', 'ppsm', 'sldx', 'sldm', 'adn', 'accdb', 'accdr', 'accdt', 'accda', 'mdw', 'accde', 'mam', 'maq', 'mar', 'mat', 'maf', 'laccdb', 'ade', 'adp', 'mdb', 'cdb', 'mda', 'mdn',
'mdt', 'mdf', 'mde', 'ldb', 'one', 'pub', 'xps']

openExtensions = ['odt', 'ott', 'oth', 'odm', 'sxw', 'stw', 'sxg', 'ods', 'ots', 'sxc', 'stc', 'odp', 'otp', 'sxi', 'sti', 'odg', 'otg', 'sxd', 'std', 'odf', 'sxm']

imgExtensions = ['jpeg', 'jpg', 'png', 'gif', 'tiff', 'psd', 'pdf', 'eps', 'ai', 'indd', 'raw', 'svg', 'xcf', 'webp', 'bmp', 'exif']



def RemoveDuplicates(listOfStrings: List[str]):
    newList = list(OrderedDict.fromkeys(listOfStrings))
    listOfStrings.clear()
    listOfStrings.extend(newList)
    return


def ReplaceGroupNamesWithTheirValues(extensionsList: List[str]):  # extensionsList is mutable, so will be changed as if passed by reference
    while 'ms-g' in extensionsList:
        extensionsList.remove('ms-g')
        extensionsList.extend(msExtensions)
    while 'open-g' in extensionsList:
        extensionsList.remove('open-g')
        extensionsList.extend(openExtensions)
    while 'img-g' in extensionsList:
        extensionsList.remove('img-g')
        extensionsList.extend(imgExtensions)    
    return




print('\nThis script will attempt to create new folders for files with the specified extension.\n\n'
'Type "exit" or "quit" (no quotation marks) or just hit enter without typing anything to close the script and not do anything.\n\n'
'If you wish to proceed, simply type in one of the following:\n'
'1. Names of the extensions you wish to create folders for, separated by spaces. For example:\n'
'txt\n'
'pptx xls doc\n'
'mp3 mp4\n'
'2. Names for groups of extensions (for quicker use). Currently the following groups are supported:\n'
'ms-g = common Microsoft document file extensions\n'
'open-g = common OpenOffice document file extensions\n'
'img-g = common image file extensions\n'
'These can be used alongside extension names, for example:\n'
'ms-g exe ods\n'
'You may also type \"all\" (no quotation marks) to copy ALL files in the current directory to corresponding subdirectories. Use with caution.\n\n'
'Type in your command to begin:')


# split the input command (extensions list) into words
chosenExtensions = input().split()

# remove commas and dots from extension names
for i in range(0, len(chosenExtensions)):
    chosenExtensions[i] = chosenExtensions[i].replace(',', '').replace('.', '')

pathToCurDir = os.path.dirname(os.path.realpath(__file__))

outputDirectoriesNames = [] # displayed on-screen after successfully copying to directories

chosenFilePaths = []    # paths of files to copy

if ('exit' not in chosenExtensions) & ('quit' not in chosenExtensions) & (len(chosenExtensions) > 0):
    if 'all' not in chosenExtensions:
        RemoveDuplicates(chosenExtensions)

        writeToMSDirectory = 'ms-g' in chosenExtensions
        writeToOPENDirectory = 'open-g' in chosenExtensions
        writeToIMGDirectory = 'img-g' in chosenExtensions

        ReplaceGroupNamesWithTheirValues(chosenExtensions)

        RemoveDuplicates(chosenExtensions)  # in case one of the extensions from the extension groups was explicitly chosen by the user already


        for i in range(0, len(chosenExtensions)):
            chosenExtensions[i] = '.' + chosenExtensions[i] # txt into .txt, doc into .doc, etc. for ease of use with endswith


        # get all paths for files from the current directory with chosen extensions
        chosenFilePaths = [join(pathToCurDir, curFilename) for curFilename in listdir(pathToCurDir) if (isfile(join(pathToCurDir, curFilename))
        and curFilename.endswith(tuple(chosenExtensions)))]
    else:
        writeToMSDirectory = True
        writeToOPENDirectory = True
        writeToIMGDirectory = True

        chosenFilePaths = [join(pathToCurDir, curFilename) for curFilename in listdir(pathToCurDir) if (isfile(join(pathToCurDir, curFilename)))]

    for filePath in chosenFilePaths:
        fileExtension = filePath[filePath.rfind('.') + 1 : len(filePath)]   # extension name is after the last dot, all the way to the end of the path


        # TODO: clean this up into a function or something
        if writeToMSDirectory & (fileExtension in msExtensions):
            if 'ms-g' not in outputDirectoriesNames:
                outputDirectoriesNames.append('ms-g')
            outputDirectory = pathToCurDir + '\\ms-g'
        elif writeToOPENDirectory & (fileExtension in openExtensions):
            if 'open-g' not in outputDirectoriesNames:
                outputDirectoriesNames.append('open-g')
            outputDirectory = pathToCurDir + '\\open-g'
        elif writeToIMGDirectory & (fileExtension in imgExtensions):
            if 'img-g' not in outputDirectoriesNames:
                outputDirectoriesNames.append('img-g')
            outputDirectory = pathToCurDir + '\\img-g'
        else:
            if fileExtension not in outputDirectoriesNames:
                outputDirectoriesNames.append(fileExtension)
            outputDirectory = pathToCurDir + '\\' + fileExtension



        if not isdir(outputDirectory):
            mkdir(outputDirectory)

        copy2(filePath, outputDirectory)    # copy the file (including metadata) to the output directory


    print('Copied files into the following subdirectories: ' + ' '.join(outputDirectoriesNames))
else:
    print('Script exited. No action taken.')