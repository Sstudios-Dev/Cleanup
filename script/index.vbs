Option Explicit

Dim objFSO, totalFiles, filesDeleted

Set objFSO = CreateObject("Scripting.FileSystemObject")
totalFiles = 0
filesDeleted = 0

' Count garbage files
CountGarbageFiles objFSO.GetSpecialFolder(0) ' 0 represents the Desktop folder

' Call the function to search and delete garbage files
SearchAndDeleteGarbageFiles objFSO.GetSpecialFolder(0) ' 0 represents the Desktop folder

' Completion message
WScript.Echo "Garbage files successfully deleted in the user's Desktop folder."

' Release object
Set objFSO = Nothing

Sub CountGarbageFiles(folder)
    Dim objFolder, objFile
    On Error Resume Next ' Handle potential permission denied errors
    
    If Not folder Is Nothing Then
        For Each objFolder In folder.SubFolders
            CountGarbageFiles objFolder
        Next
        
        For Each objFile In folder.Files
            If IsGarbageFile(objFile) Then
                totalFiles = totalFiles + 1
            End If
        Next
    End If
End Sub

Sub SearchAndDeleteGarbageFiles(folder)
    Dim objFolder, objFile
    
    If Not folder Is Nothing Then
        For Each objFolder In folder.SubFolders
            SearchAndDeleteGarbageFiles objFolder
        Next
        
        For Each objFile In folder.Files
            If IsGarbageFile(objFile) Then
                objFile.Delete(True) ' Force delete
                filesDeleted = filesDeleted + 1
            End If
        Next
    End If
End Sub

Function IsGarbageFile(objFile)
    Dim extension
    extension = LCase(objFSO.GetExtensionName(objFile.Path))
    IsGarbageFile = (extension = "tmp" Or extension = "bak")
End Function
