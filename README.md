# py-sort-files-based-on-type
Copies files from the source directory (**src:"C:\Some\Directory"**) to the destination directory (**dst:"C:\Other\Folder"**) in folders corresponding to their extensions (for example, .txt files go into a "txt" folder if you provide the "txt" argument. You can provide multiple file types at once, e.g. txt pdf).

## Installation
No installation required. This is a portable script, just download "**sort-files-based-on-type.py**" and run the script in your preferred environment (e.g. Windows command line).

## Extension groups
Provide "ms-g" in the command to sort by common Microsoft document file extensions;

Provide "open-g" in the command to sort by common OpenOffice document file extensions;

Provide "img-g" in the command to sort by common image file extensions.

## Finding and automatically sorting all files
Provide "--all" in the command to copy all files from the source directory and put them in automatically sorted folders in the destination directory.

Using this argument automatically groups the files into the above mentioned extension groups.

## Recursive file search
Provide "--rec" in the command to recursively find files in all subdirectories of the source directory.

## Cutting
Provide "--cut" in the command to cut (move) files from the source directory to the destination directory.

## Overwrite protection
When a command would result in overwriting files, the script will ask you whether you want this specific file to be overwritten, and whether you want future files to be overwritten as well.

## Drag & drop
For maximum ease of use, you can simply copy the script into whichever directory you want it to operate on, and if you don't provide source and destination directory arguments, it will simply sort the files within the same directory.
