# File Cleanup VBScript

This VBScript is designed to count and delete garbage files from a directory and its subdirectories. Garbage files are defined as those with ".tmp" or ".bak" extensions.

---

## Usage

1. **Download the Script**: Download the `cleanup.vbs` script file to your system.

2. **Run the Script**:
   - Open a command prompt (cmd) as an administrator.
   - Navigate to the location where you saved the `index.vbs` file.
   - Run the script using the command `cscript index.vbs`.
  
# Features:

## Count Garbage Files:

```vbs
' Count garbage files
CountGarbageFiles objFSO.GetSpecialFolder(0) ' 0 represents the Desktop folder
```

This function counts the garbage files in the directory and its subdirectories.

## Delete Garbage Files:

```vbs
' Call the function to search and delete garbage files
SearchAndDeleteGarbageFiles objFSO.GetSpecialFolder(0) ' 0 represents the Desktop folder
```

This function searches for and deletes garbage files from the directory and its subdirectories.

## Determine if a File is Garbage:

```vbs
Function IsGarbageFile(objFile)
    Dim extension
    extension = LCase(objFSO.GetExtensionName(objFile.Path))
    IsGarbageFile = (extension = "tmp" Or extension = "bak")
End Function
```

This auxiliary function determines if a file is garbage based on its extension.

# Customization:

You can customize this script by modifying the following parts:

- The file extensions considered as garbage (`tmp` and `bak`) in the IsGarbageFile function.
- The actions performed on the garbage files found in the `SearchAndDeleteGarbageFiles` function, such as deletion or logging.

# Notes
- Make sure to run the script as an administrator to avoid permission issues when deleting files.
- Basic error handling is provided to handle potential permission or access denied issues.
