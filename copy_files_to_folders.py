from collections import OrderedDict
import os
from os import listdir
from os.path import isfile, join


print('\nThis script will attempt to create new folders for files with the specified extension.\n\n'
'Type "exit" (no quotation marks) or just hit enter without typing anything to close the script and not do anything.\n\n'
'Alternatively, if you wish to proceed, simply type in one of the following:\n'
'1. Names of the extensions you wish to create folders for, without the dot, separated by commas and/or spaces. For example:\n'
'txt\n'
'pptx,xls,doc\n'
'mp3 mp4\n'
'2. Names for groups of extensions (for quicker use). Currently the following groups are supported:\n'
'ms = common Microsoft document file extensions\n'
'open = common "Open Source" (mostly OpenOffice) document file extensions\n'
'These can be used alongside extension names, for example:\n'
'ms, exe, ods\n\n'
'Type in your command to begin:')

commandText = input().split()   # split into words

pathToCurDir = os.path.dirname(os.path.realpath(__file__))
textFilenamesInCurDir = [curFilename for curFilename in listdir(pathToCurDir) if (isfile(join(pathToCurDir, curFilename))
and curFilename.endswith('.txt'))]