import os, errno, tempfile, sys
from typing import List
from os import makedirs, listdir, walk
from os.path import isfile, isdir, join
from collections import OrderedDict
from shutil import copy2, move



scriptVersion = 'v1.1'


msExtensions = ['doc', 'dot', 'wbk', 'docx', 'docm', 'dotx', 'dotm', 'docb', 'xls', 'xlt', 'xlm', 'xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xla', 'xlam', 'xll', 'xlw', 'ppt', 'pot', 'pps', 'pptx', 'pptm',
'potx', 'potm', 'ppam', 'ppsx', 'ppsm', 'sldx', 'sldm', 'adn', 'accdb', 'accdr', 'accdt', 'accda', 'mdw', 'accde', 'mam', 'maq', 'mar', 'mat', 'maf', 'laccdb', 'ade', 'adp', 'mdb', 'cdb', 'mda', 'mdn',
'mdt', 'mdf', 'mde', 'ldb', 'one', 'pub', 'xps']

openExtensions = ['odt', 'ott', 'oth', 'odm', 'sxw', 'stw', 'sxg', 'ods', 'ots', 'sxc', 'stc', 'odp', 'otp', 'sxi', 'sti', 'odg', 'otg', 'sxd', 'std', 'odf', 'sxm']

imgExtensions = ['jpeg', 'jpg', 'png', 'gif', 'tiff', 'psd', 'eps', 'ai', 'indd', 'raw', 'svg', 'xcf', 'webp', 'bmp', 'exif']

availableDoubleDashCommands = ['cut', 'all', 'rec'] # preceded by 2 dashes, e.g.    --cut

availableColonQuotesArguments = ['src', 'dst'] # followed by a colon and quotes, e.g.    src:"C:\Folder Name"

charactersBannedInWindowsFilenames = ['\\', '/', ':', '*', '?', '\"', '<', '>', '|']


class InvalidUseOfQuotesError(Exception):

    """
    Doesn't do anything on its own, purely informative.
    """
    pass




# ***** copy-pasted from https://stackoverflow.com/questions/9532499/check-whether-a-path-is-valid-in-python-without-creating-a-file-at-the-paths-ta/
# ***** don't ask me how this works
# Sadly, Python fails to provide the following magic number for us.

ERROR_INVALID_NAME = 123
'''
Windows-specific error code indicating an invalid pathname.

See Also
----------

https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-

    Official listing of all such codes.
'''


def is_pathname_valid(pathname: str) -> bool:
    '''
    `True` if the passed pathname is a valid pathname for the current OS;
    `False` otherwise.
    '''

    # If this pathname is either not a string or is but is empty, this pathname
    # is invalid.

    try:
        if not isinstance(pathname, str) or not pathname:
            return False


        # Strip this pathname's Windows-specific drive specifier (e.g., `C:\`)
        # if any. Since Windows prohibits path components from containing `:`
        # characters, failing to strip this `:`-suffixed prefix would
        # erroneously invalidate all valid absolute Windows pathnames.

        _, pathname = os.path.splitdrive(pathname)

        # Directory guaranteed to exist. If the current OS is Windows, this is
        # the drive to which Windows was installed (e.g., the "%HOMEDRIVE%"
        # environment variable); else, the typical root directory.

        root_dirname = os.environ.get('HOMEDRIVE', 'C:') \
            if sys.platform == 'win32' else os.path.sep

        assert os.path.isdir(root_dirname)   # ...Murphy and her ironclad Law

        # Append a path separator to this directory if needed.
        root_dirname = root_dirname.rstrip(os.path.sep) + os.path.sep

        # Test whether each path component split from this pathname is valid or
        # not, ignoring non-existent and non-readable path components.
        for pathname_part in pathname.split(os.path.sep):
            try:
                os.lstat(root_dirname + pathname_part)

            # If an OS-specific exception is raised, its error code
            # indicates whether this pathname is valid or not. Unless this
            # is the case, this exception implies an ignorable kernel or
            # filesystem complaint (e.g., path not found or inaccessible).
            #
            # Only the following exceptions indicate invalid pathnames:
            #
            # * Instances of the Windows-specific "WindowsError" class
            #   defining the "winerror" attribute whose value is
            #   "ERROR_INVALID_NAME". Under Windows, "winerror" is more
            #   fine-grained and hence useful than the generic "errno"
            #   attribute. When a too-long pathname is passed, for example,
            #   "errno" is "ENOENT" (i.e., no such file or directory) rather
            #   than "ENAMETOOLONG" (i.e., file name too long).
            # * Instances of the cross-platform "OSError" class defining the
            #   generic "errno" attribute whose value is either:
            #   * Under most POSIX-compatible OSes, "ENAMETOOLONG".
            #   * Under some edge-case OSes (e.g., SunOS, *BSD), "ERANGE".

            except OSError as exc:
                if hasattr(exc, 'winerror'):
                    if exc.winerror == ERROR_INVALID_NAME:
                        return False

                elif exc.errno in {errno.ENAMETOOLONG, errno.ERANGE}:
                    return False

    # If a "TypeError" exception was raised, it almost certainly has the
    # error message "embedded NUL character" indicating an invalid pathname.

    except TypeError as exc:
        return False

    # If no exception was raised, all path components and hence this
    # pathname itself are valid. (Praise be to the curmudgeonly python.)

    else:
        return True

    # If any other exception was raised, this is an unrelated fatal issue
    # (e.g., a bug). Permit this exception to unwind the call stack.
    #
    # Did we mention this should be shipped with Python already?


def is_path_creatable(pathname: str) -> bool:
    '''
    `True` if the current user has sufficient permissions to create the passed
    pathname; `False` otherwise.
    '''

    # Parent directory of the passed path. If empty, we substitute the current

    # working directory (CWD) instead.

    dirname = os.path.dirname(pathname) or os.getcwd()

    return os.access(dirname, os.W_OK)


def is_path_exists_or_creatable(pathname: str) -> bool:
    '''
    `True` if the passed pathname is a valid pathname for the current OS _and_
    either currently exists or is hypothetically creatable; `False` otherwise.
    This function is guaranteed to _never_ raise exceptions.
    '''

    try:
        # To prevent "os" module calls from raising undesirable exceptions on
        # invalid pathnames, is_pathname_valid() is explicitly called first.
        return is_pathname_valid(pathname) and (
            os.path.exists(pathname) or is_path_creatable(pathname))

    # Report failure on non-fatal filesystem complaints (e.g., connection
    # timeouts, permissions issues) implying this path to be inaccessible. All
    # other exceptions are unrelated fatal issues and should not be caught here.
    except OSError:
        return False


def is_path_sibling_creatable(pathname: str) -> bool:
    '''
    `True` if the current user has sufficient permissions to create **siblings**
    (i.e., arbitrary files in the parent directory) of the passed pathname;
    `False` otherwise.
    '''

    # Parent directory of the passed path. If empty, we substitute the current
    # working directory (CWD) instead.
    dirname = os.path.dirname(pathname) or os.getcwd()


    try:
        # For safety, explicitly close and hence delete this temporary file
        # immediately after creating it in the passed path's parent directory.
        with tempfile.TemporaryFile(dir=dirname): pass
        return True

    # While the exact type of exception raised by the above function depends on
    # the current version of the Python interpreter, all such types subclass the
    # following exception superclass.
    except EnvironmentError:
        return False


def is_path_exists_or_creatable_portable(pathname: str) -> bool:
    '''
    `True` if the passed pathname is a valid pathname on the current OS _and_
    either currently exists or is hypothetically creatable in a cross-platform
    manner optimized for POSIX-unfriendly filesystems; `False` otherwise.

    This function is guaranteed to _never_ raise exceptions.
    '''

    try:
        # To prevent "os" module calls from raising undesirable exceptions on
        # invalid pathnames, is_pathname_valid() is explicitly called first.
        return is_pathname_valid(pathname) and (
            os.path.exists(pathname) or is_path_sibling_creatable(pathname))

    # Report failure on non-fatal filesystem complaints (e.g., connection
    # timeouts, permissions issues) implying this path to be inaccessible. All
    # other exceptions are unrelated fatal issues and should not be caught here.

    except OSError:
        return False

# ***** copy-pasted section ends here





def RemoveDuplicates(listOfStrings: List[str]):
    '''
    Returns listOfStrings with all duplicate items removed.

    Doesn't raise exceptions.
    '''
    
    return list(OrderedDict.fromkeys(listOfStrings))



def ReplaceGroupNamesWithTheirValues(extensionsList: List[str]):
    '''
    Replaces names of extension groups with their corresponding values (multiple extension names) in the provided extensionsList.

    No return value - extensionsList itself is modified.

    Doesn't raise exceptions.
    '''

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



def ConsoleInput():
    '''
    Returns input() with leading and trailing whiitespaces removed as a lowercase string.

    Doesn't raise exceptions.
    '''

    return input().strip().lower()



def FindDoubleDashCommands(inputStr: str):
    '''
    Returns all double dash (--) commands found in inputStr as a list of strings, even duplicates, alongside their dashes.

    Doesn't raise exceptions.
    '''

    foundDoubleDashCommands = []

    curStartingIndex = 0


    for i in range(0, inputStr.count('--')):
        # find index of first character after --
        curStartingIndex = inputStr.find('--', curStartingIndex) + 2

        # if -- is at the very end of the inputStr, then curArgumentName is empty
        if curStartingIndex > len(inputStr) - 1:
            curArgumentName = ''

        else:   # if -- isn't at the end, then curArgumentName is any characters provided before the first space after --
            indexOfFirstSpaceAfterDoubleDash = inputStr.find(' ', curStartingIndex)
            curArgumentName = inputStr[curStartingIndex : 
            len(inputStr) if (indexOfFirstSpaceAfterDoubleDash == -1) # if there are no spaces following --, then slice curArgumentName all the way to the end of the string
            else indexOfFirstSpaceAfterDoubleDash]


        # append --curArgumentName even if it's a duplicate
        foundDoubleDashCommands.append('--' + curArgumentName)

    return foundDoubleDashCommands



def FindColonQuotesCommands(inputStr: str):
    '''
    Returns all colon-quotes (:"") commands found in inputStr as a list of strings, even duplicates.

    May raise InvalidUseOfQuotesError if a found command doesn't have exactly 2 quotes.
    '''

    foundColonQuotesCommands = []

    curColonQuoteIndex = -2

    for i in range(0, inputStr.count(':\"')):
        curColonQuoteIndex = inputStr.find(':\"', curColonQuoteIndex + 2)

        # if the first character of input is ":", then an argument couldn't have been provided before it
        if curColonQuoteIndex == 0:
            curArgumentName = ''
        else:
            curStartingIndex = inputStr.rfind(' ', 0, curColonQuoteIndex) + 1  # if it doesn't find a space before ":", that means it was provided as the first argument of the input. -1 (on failure) + 1 = 0
            curArgumentName = inputStr[curStartingIndex : curColonQuoteIndex]


        nextColonQuoteIndex = inputStr.find(':\"', curColonQuoteIndex + 2)

        curNumOfQuotes = inputStr.count('\"', curColonQuoteIndex + 1, len(inputStr) if nextColonQuoteIndex == -1 else nextColonQuoteIndex)


        if curNumOfQuotes != 2:
            raise InvalidUseOfQuotesError

        else:
            foundColonQuotesCommands.append(curArgumentName + inputStr[curColonQuoteIndex : inputStr.find('\"', curColonQuoteIndex + 2) + 1])


    return foundColonQuotesCommands


def DisplayHelp():
    '''
    Prints help text to the console.

    Doesn't raise exceptions.
    '''
    
    print('\n\033[0;33msort-files-based-on-type \033[1;33m' + scriptVersion + '\033[0m. Type \033[1;34mhelp\033[0m to view this message again. Examples of full commands are shown in \033[1;34mblue\033[0m, keywords are shown in \033[1;32mgreen\033[0m.\n'
    '\nThis script will attempt to create new folders for files with the specified extensions.\n\n'
    'Type \033[1;34mexit\033[0m or \033[1;34mquit\033[0m to close the script and not do anything.\n\n'
    'If you wish to proceed, simply type in one of the following:\n'
    '1. Names of the extensions you wish to create folders for, separated by spaces:\n'
    '\033[1;34mpptx doc mp3\033[0m\n'
    '2. Names for groups of extensions (for quicker use). Currently the following groups are supported:\n'
    '\033[1;32mms-g\033[0m = common Microsoft document file extensions\n'
    '\033[1;32mopen-g\033[0m = common OpenOffice document file extensions\n'
    '\033[1;32mimg-g\033[0m = common image file extensions\n'
    'These can be used alongside extension names:\n'
    '\033[1;34mms-g exe ods\033[0m\n\n'
    'By default, the source directory is the directory this script is ran from, but you can provide any source directory by specifying it in quotes after \033[1;32msrc:\033[0m:\n'
    '\033[1;34msrc:\"C:\\Users\\username\\Documents\" mp4 img-g\033[0m\n'
    'The default directory for creating new folders and writing to them is the same as the source directory, but you may change that by providing a destination directory in quotes after \033[1;32mdst:\033[0m:\n'
    '\033[1;34mdst:\"C:\\Users\\username\\Documents\\My Games\" ini exe\033[0m\n'
    'Additionally, you may provide the argument \033[1;32m--cut\033[0m to cut (\"move\" - use with caution, backups aren\'t created automatically!) the files into corresponding directories instead of copying them:\n'
    '\033[1;34msrc:\"C:\\Users\\username\\Documents\" dst:\"C:\\Users\\username\\Documents\\My Games\" exe ini txt --cut\033[0m\n'
    'Do note that while the contents of the files will be copied losslessly, some amount of metadata may rarely be lost in the process due to system limitations.\n'
    'You may add the \033[1;32m--all\033[0m argument to copy ALL files in the source directory instead of having to choose file types or groups.\n'
    'Lastly, you can use \033[1;32m--rec\033[0m to recursively affect (copy/cut, depending on your choice) all subfolders\' contents as well as the main source folder\'s contents.\n')
    
    return



# autorun starts here
os.system("")  # enables ansi escape characters in terminal

DisplayHelp()

while True:
    print('\n---------------------\n'
          'Type in your command:')

    commandsString = ConsoleInput()


    programShouldExitWithoutAction = commandsString in ['exit', 'quit']
    if programShouldExitWithoutAction:
        break

    
    if commandsString == '':
        print('\nNo command provided. Try typing \033[1;34mhelp\033[0m if you\'re unsure what to do.')
        continue

    
    if commandsString == 'help':
        DisplayHelp()
        continue
    
    
    if commandsString in ['blue', 'green']:
        print('\nNice try.')
        continue


    foundDoubleDashCommands = FindDoubleDashCommands(commandsString)

    foundDoubleDashCommandsAreValid = True

    for command in foundDoubleDashCommands:
        # check if any double dash (--) command was provided more than once
        if foundDoubleDashCommands.count(command) > 1:
            print('\nError: Command \"' + command + '\" was provided more than once.')
            foundDoubleDashCommandsAreValid = False
            break
        

        # check if any of the arguments (commands with -- removed) are not in availableDoubleDashCommands
        if command[2:] not in availableDoubleDashCommands:
            print('\nError: Command \"' + command + '\" is not a valid double dash command.\n'
            'Available double dash commands:' + str([' --' + argument for argument in availableDoubleDashCommands]))
            foundDoubleDashCommandsAreValid = False
            break


    if not foundDoubleDashCommandsAreValid:
        continue


    operateOnAllFiles = '--all' in foundDoubleDashCommands

    cutTheFiles = '--cut' in foundDoubleDashCommands

    findFilesRecursively = '--rec' in foundDoubleDashCommands


    try:
        foundColonQuotesCommands = FindColonQuotesCommands(commandsString)

    except InvalidUseOfQuotesError:
        print('\nError: Invalid use of quotes.\n'
        'All paths provided to commands need to be enclosed in 2 \" symbols.')
        continue


    foundColonQuotesCommandsAreValid = True


    for command in foundColonQuotesCommands:
        # check if path provided in command is valid (exists or is creatable)
        providedPath = command[command.find('\"') + 1 : command.rfind('\"')]
        if providedPath.endswith('\\'):
            providedPath = providedPath[: -1]


        if not is_path_exists_or_creatable_portable(providedPath):
            print('\nError: Path \"' + providedPath + '\" is not a valid directory.')
            foundColonQuotesCommandsAreValid = False
            break


        # check if any colon quotes (:"") command was provided more than once
        if foundColonQuotesCommands.count(command) > 1:
            print('\nError: Command \"' + command + '\" was provided more than once.')
            foundColonQuotesCommandsAreValid = False
            break


        # check if any of the arguments (commands with everything starting with and after ":" removed) are not in availableColonQuotesArguments
        if command[0 : command.find(':')] not in availableColonQuotesArguments:
            print('\nError: Command \"' + command + '\" is not a valid colon-quotes command.\n'
            'Available colon-quotes commands: ' + str([argument + ':"" ' for argument in availableColonQuotesArguments]))
            foundColonQuotesCommandsAreValid = False
            break



    if not foundColonQuotesCommandsAreValid:
        continue


    # create extensionsString via removing all -- and :"" commands from commandsString
    # (hopefully only leaving extension names and extension group names)
    extensionsString = commandsString


    for curCommand in foundDoubleDashCommands:
        extensionsString = extensionsString.replace(curCommand, '')


    for curCommand in foundColonQuotesCommands:
        extensionsString = extensionsString.replace(curCommand, '')



    # split extensionsString into extensionsList based on whitespaces
    extensionsList = extensionsString.split()


    # determine whether to write to group directories (ms-g, open-g, etc.)
    if operateOnAllFiles:
        writeToMSDirectory = True
        writeToOPENDirectory = True
        writeToIMGDirectory = True

    else:
        writeToMSDirectory = 'ms-g' in extensionsList
        writeToOPENDirectory = 'open-g' in extensionsList
        writeToIMGDirectory = 'img-g' in extensionsList


    # replace group names (ms-g, open-g, etc.) with their corresponding values (groups of extension names)
    ReplaceGroupNamesWithTheirValues(extensionsList)


    extensionsAreValid = True
    eachCharIsValid = True


    for curExtension in extensionsList:
        # check if any extension was provided more than once, including those from extension groups
        if extensionsList.count(curExtension) > 1:
            print('\nError: Extension \"' + curExtension + '\" was provided more than once.')
            extensionsAreValid = False
            break


        # check if any of the characters in provided extensions are banned in windows filenames
        for curChar in curExtension:
            if curChar in charactersBannedInWindowsFilenames:
                print('\nError: Character \"' + curChar + '\" is not allowed in Windows extensions.')
                eachCharIsValid = False
                break
            

        if not eachCharIsValid:
            extensionsAreValid = False
            break



    if not extensionsAreValid:
        break



    pathToSourceDir = ''
    pathToDestinationDir = ''


    for curCommand in foundColonQuotesCommands:
        # if the src:"" command was provided, set the source directory to the provided path, which was previously checked and is valid
        if curCommand.startswith('src:'):
            pathToSourceDir = curCommand[curCommand.find('\"') + 1 : curCommand.rfind('\"')]
            if pathToSourceDir.endswith('\\'):
                pathToSourceDir = pathToSourceDir[: -1]


        # if the dst:"" command was provided, set the destination directory to the provided path, which was previously checked and is valid
        elif curCommand.startswith('dst:'):
            pathToDestinationDir = curCommand[curCommand.find('\"') + 1 : curCommand.rfind('\"')]
            if pathToDestinationDir.endswith('\\'):
                pathToDestinationDir = pathToDestinationDir[: -1]
            


    # if the src:"" command wasn't provided, set the source directory to whatever directory the script file is in
    if pathToSourceDir == '':
        pathToSourceDir = os.path.dirname(os.path.realpath(__file__))


    # if the dst:"" command wasn't provided, set the destination directory to be the same as the source directory
    if pathToDestinationDir == '':
        pathToDestinationDir = pathToSourceDir



    if not ((len(extensionsList) > 0) or operateOnAllFiles):
        print('\nError: No extension or extension group was provided in the input.\n'
        'If you wish to operate on all files in the chosen directory, use the --all command.')
        continue

    
    chosenFilepaths = []

    # get filenames to operate on from the source directory
    
    # without "--rec": only in the original source directory (no subdirectories)
    if not findFilesRecursively:
        chosenFilepaths = [(pathToSourceDir + '\\' + curFilename) for curFilename in listdir(pathToSourceDir) if (isfile(join(pathToSourceDir, curFilename))   # only get items that are files...
        and (operateOnAllFiles or curFilename.endswith(tuple('.' + curExtension for curExtension in extensionsList))))]             # and end with a proper extension.
    else:
        # with "--rec": in the original source directory and all subdirectories
        for (parentDirName, childDirNames, fileNames) in os.walk(pathToSourceDir):
            for curFilename in fileNames:
                curFilepath = parentDirName + '\\' + curFilename
                if curFilepath in chosenFilepaths:
                    continue
                
                if operateOnAllFiles:
                    chosenFilepaths.append(curFilepath)
                elif curFilename.endswith(tuple('.' + curExtension for curExtension in extensionsList)):
                    chosenFilepaths.append(curFilepath)
            


    outputDirectoriesDictionary = {} # contains paths to directories as keys, and copied filenames as values. displayed on-screen after successfully copying to directories

    overwriteFilesYesAll = False
    overwriteFilesNoAll = False

    for curFilepath in chosenFilepaths:
        curFilename = curFilepath[curFilepath.rfind('\\') + 1 :]
        if ('.' in curFilename) and (len(curFilename) > (curFilename.rfind('.') + 1)):   # if curFilename has an extension (something after the dot)...
            fileExtension = curFilename[curFilename.rfind('.') + 1 :]                   # ...find that extension and assign it to fileExtension.
        else:
            fileExtension = 'no-extension'



        # determine output directory for current file
        # TODO: clean this up into a function or something

        if writeToMSDirectory & (fileExtension in msExtensions):
            outputDirectory = pathToDestinationDir + '\\' + 'ms-g'

        elif writeToOPENDirectory & (fileExtension in openExtensions):
            outputDirectory = pathToDestinationDir + '\\' + 'open-g'

        elif writeToIMGDirectory & (fileExtension in imgExtensions):
            outputDirectory = pathToDestinationDir + '\\' + 'img-g'
        else:
            outputDirectory = pathToDestinationDir + '\\' + fileExtension




        if not isdir(outputDirectory):
            makedirs(outputDirectory)


        overwriteFileInput = ''
        filenameAlreadyExists = False
        firstAttemptAtInput = True

        if isfile(join(outputDirectory, curFilename)):
            filenameAlreadyExists = True
            
            if (not overwriteFilesYesAll) and (not overwriteFilesNoAll):
                while overwriteFileInput not in ['y', 'yes', 'yes all', 'n', 'no', 'no all']:
                    if not firstAttemptAtInput:
                        print('\nError: Unrecognized input. Recognized input: y, yes, yes all, n, no, no all.\n')
                    
                    firstAttemptAtInput = False
                    print('File \"' + curFilename + '\" already exists in directory \"' + outputDirectory + '\".\n'
                    'Overwrite it? (yes/yes all/no/no all)')
                    overwriteFileInput = ConsoleInput()


        if overwriteFileInput == 'yes all':
            overwriteFilesYesAll = True
        if overwriteFileInput == 'no all':
            overwriteFilesNoAll = True
        
        
                
        overwriteFile = (overwriteFileInput in ['y', 'yes']) or overwriteFilesYesAll


        if (not filenameAlreadyExists) or overwriteFile:
            if cutTheFiles:
                move(curFilepath, join(outputDirectory, curFilename))   # move (cut) the source file (including metadata) to the output directory
            else:
                copy2(curFilepath, join(outputDirectory, curFilename))  # copy the source file (including metadata) to the output directory


            # add current outputDirectory as a key to outputDirectoriesDictionary, with curFilename as the value...
            if outputDirectory not in outputDirectoriesDictionary:
                outputDirectoriesDictionary[outputDirectory] = [curFilename]

            # ...or if the key already exists, append another filename to its value
            else:
                outputDirectoriesDictionary[outputDirectory].append(curFilename)



    if len(outputDirectoriesDictionary) > 0:
        amountOfFilesAffected = 0
        for directoryName in outputDirectoriesDictionary:
            amountOfFilesAffected += len(outputDirectoriesDictionary[directoryName])
        
        print('\n' + ('Cut (moved)' if cutTheFiles else 'Copied') + ' \033[1;32m' + str(amountOfFilesAffected) + '\033[0m file(s) into the following directories:\n')

        for directoryName in outputDirectoriesDictionary:
            print(directoryName + ' - \033[1;32m' + str(len(outputDirectoriesDictionary[directoryName])) + '\033[0m file(s):\n'
            + ', '.join(outputDirectoriesDictionary[directoryName]) + '\n')
    else:
        print('\nNo suitable files were found in \"' + pathToSourceDir + '\". No action taken.')




if programShouldExitWithoutAction:
    print('\nProgram exited. No action taken.')