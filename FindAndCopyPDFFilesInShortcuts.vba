Sub FindAndCopyPDFFilesInShortcuts()
    Dim FileSystem As Object
    Dim SourceFolder As Object
    Dim Shortcut As Object
    Dim DestinationFolder As String
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FileSystem.GetFolder(ThisWorkbook.Path)
    
    DestinationFolder = ThisWorkbook.Path & "\FilteredFiles\"
    
    If Not FileSystem.FolderExists(DestinationFolder) Then
        FileSystem.CreateFolder (DestinationFolder)
    End If
    
    ' Recursively search through subfolders and shortcuts for PDF files
    ProcessFolderForPDFs SourceFolder, DestinationFolder, FileSystem
    
    MsgBox "PDF files containing 'EMEG' or 'PRD' in their filenames have been copied to '" & DestinationFolder & "'."
End Sub

Sub ProcessFolderForPDFs(ByVal Folder As Object, ByVal DestinationFolder As String, ByVal FileSystem As Object)
    Dim SubFolder As Object
    Dim File As Object
    Dim Shortcut As Object
    Dim FileName As String
    
    ' Process files in the current folder
    For Each File In Folder.Files
        If FileSystem.GetExtensionName(File.Path) = "pdf" Then
            FileName = File.Name
            ' Check if the filename contains "EMEG" or "PRD"
            If InStr(1, FileName, "EMEG", vbTextCompare) > 0 Or InStr(1, FileName, "PRD", vbTextCompare) > 0 Then
                ' Copy the PDF file to the destination folder
                File.Copy DestinationFolder & FileName
            End If
        End If
    Next File
    
    ' Process shortcuts in the current folder
    For Each Shortcut In Folder.Files
        If FileSystem.GetExtensionName(Shortcut.Path) = "lnk" Then
            ProcessShortcutForPDFs Shortcut, DestinationFolder, FileSystem
        End If
    Next Shortcut
    
    ' Recursively process subfolders
    For Each SubFolder In Folder.SubFolders
        ProcessFolderForPDFs SubFolder, DestinationFolder, FileSystem
    Next SubFolder
End Sub

Sub ProcessShortcutForPDFs(ByVal Shortcut As Object, ByVal DestinationFolder As String, ByVal FileSystem As Object)
    Dim TargetPath As String
    Dim objShell As Object
    
    ' Create a Shell object to work with shortcuts
    Set objShell = CreateObject("WScript.Shell")
    
    ' Get the target path of the shortcut
    TargetPath = objShell.CreateShortcut(Shortcut.Path).TargetPath
    
    ' Check if the target path is a folder
    If FileSystem.FolderExists(TargetPath) Then
        ' Process the target folder for PDF files
        ProcessFolderForPDFs FileSystem.GetFolder(TargetPath), DestinationFolder, FileSystem
    End If
End Sub
