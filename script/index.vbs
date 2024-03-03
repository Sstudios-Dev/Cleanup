Option Explicit

Dim objFSO, objFolder, objFile, strFolderPath, totalFiles, filesDeleted

' Current user's folder
strFolderPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

' Create FileSystemObject object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Initialize counters
totalFiles = 0
filesDeleted = 0

' Function to count garbage files in a folder and its subfolders
Sub CountGarbageFiles(folderPath)
    Dim objFolder, objFile
    
    ' Check if folder exists
    If objFSO.FolderExists(folderPath) Then
        ' Get reference to the folder
        Set objFolder = objFSO.GetFolder(folderPath)
        
        ' Count files in the current folder
        totalFiles = totalFiles + objFolder.Files.Count
        
        ' Traverse all subfolders and recursively call the function
        For Each objFolder In objFolder.SubFolders
            CountGarbageFiles objFolder.Path
        Next
    End If
End Sub

' Function to search and delete garbage files in a folder and its subfolders
Sub SearchAndDeleteGarbageFiles(folderPath)
    Dim objFolder, objFile, progressBar, i
    
    ' Check if folder exists
    If objFSO.FolderExists(folderPath) Then
        ' Get reference to the folder
        Set objFolder = objFSO.GetFolder(folderPath)
        
        ' Traverse all files in the folder and delete those that meet the conditions
        For Each objFile In objFolder.Files
            ' Conditions to delete garbage files (you can adjust them according to your needs)
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "tmp" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "bak" Then
                ' Show the file being deleted
                WScript.StdOut.WriteLine "Deleting file: " & objFile.Path
                ' Delete the file
                objFile.Delete
                filesDeleted = filesDeleted + 1
            End If
        Next
        
        ' Traverse all subfolders and recursively call the function
        For Each objFolder In objFolder.SubFolders
            SearchAndDeleteGarbageFiles objFolder.Path
        Next
    End If
End Sub

' Count garbage files
CountGarbageFiles strFolderPath

' Call the function to search and delete garbage files
SearchAndDeleteGarbageFiles strFolderPath

' Completion message
WScript.Echo "Garbage files successfully deleted in the user's folder."

' Release objects
Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
