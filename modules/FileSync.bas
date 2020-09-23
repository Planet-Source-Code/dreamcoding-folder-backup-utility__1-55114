Attribute VB_Name = "FileSync"
Enum CompareDirectoryEnum
    cdeSourceDirOnly = -2   ' file is present only in the source directory
    cdeDestDirOnly = -1     ' file is present only in the dest directory
    cdeEqual = 0            ' file is present in both directories,
                            '  with same date, size, and attributes
    cdeSourceIsNewer = 1    ' file in source dir is newer
    cdeSourceIsOlder = 2    ' file in source dir is older
    cdeDateDiffer = 3       ' date of files are different
    cdeSizeDiffer = 4       ' size of files are different
    cdeAttributesDiffer = 8 ' attributes of files are different
End Enum

' Synchronize two directory subtrees
'
' This routine compares source and dest directory trees and copies files
' from source that are newer than (or are missing in) the destination directory

' if TWOWAYSYNC is True, files are synchronized in both ways

' NOTE: requires the CompareDirectories and SynchronizeDirectories routines
'       and a reference to the Microsoft Scripting Runtime type library

Sub SynchronizeDirectoryTrees(ByVal sourceDir As String, _
    ByVal destDir As String, Optional ByVal TwoWaySync As Boolean)
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFld As Scripting.Folder
    Dim destFld As Scripting.Folder
    Dim fld As Scripting.Folder
    Dim col As New Collection
    
    ' we need this in case the dest subdir doesn't exist
    On Error Resume Next
    
    ' get reference to source and dest folder objects
    Set sourceFld = fso.GetFolder(sourceDir)
    Set destFld = fso.GetFolder(destDir)

    ' create the destination directory, if necessary
    If Err Then
        ' if the destination directory doesn't exist,
        '  create it and copy all files there
        ' (this is all we need)
        fso.CopyFolder sourceDir, destDir
        ' nothing else to do
        Exit Sub
    End If
    
    ' synchronize the root directories
    SynchronizeDirectories sourceDir, destDir, TwoWaySync
    
    ' ensure that dir names have a training backslash
    If Right$(sourceDir, 1) <> "\" Then sourceDir = sourceDir & "\"
    If Right$(destDir, 1) <> "\" Then destDir = destDir & "\"
    
    ' repeat for all the subdirectories in the source directory
    For Each fld In sourceFld.SubFolders
        ' remember that we have processed this subdir
        col.Add fld.Name, fld.Name
        
        ' call this routine recursively
        SynchronizeDirectoryTrees fld.Path, destDir & fld.Name, TwoWaySync
        DoEvents
    Next
    
    ' if two-way synchronization was requested, ensure that all subdirs in dest
    ' directories are copied into source directory
    If TwoWaySync Then
        For Each fld In destFld.SubFolders
            If col(fld.Name) = "" Then
                ' we get here only if the folder name isn't in COL,
                '  and therefore
                ' if this subdirectory isn't in the source directory
                fso.CopyFolder fld.Path, sourceDir & fld.Name
            End If
        Next
    End If

End Sub



' Compare files in two directories
'
' returns a two-dimensional array of variants, where arr(0,
'  n) is the name of the N-th file
' and arr(1, n) is one of the CompareDirectoryEnum values
'
' NOTE: requires a reference to the Microsoft Scripting Runtime type library
'
' Usage example:
'   ' compare the directories C:\DOCS and C:\BACKUP\DOCS
'   Dim arr() As Variant, index As Long
'   arr = CompareDirectories("C:\DOCS", "C:\BACKUP\DOCS")
'   ' display files in C:\DOCS that should be copied into the backup directory
'   ' because they are newer or because they aren't there
'   For index = 1 To UBound(arr, 2)
'       If arr(1, index) = cdeSourceDirOnly Or arr(1, index) = cdeSourceIsNewer
' Then
'           Print arr(0, index)
'   Next

Function CompareDirectories(ByVal sourceDir As String, ByVal destDir As String) _
    As Variant()
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFld As Scripting.Folder
    Dim destFld As Scripting.Folder
    Dim sourceFile As Scripting.File
    Dim destFile As Scripting.File
    Dim col As New Collection
    Dim index As Long
    Dim FileName As String
    
    ' get a reference to source and dest folders
    Set sourceFld = fso.GetFolder(sourceDir)
    Set destFld = fso.GetFolder(destDir)
    
    ' ensure that destination path has a trailing backslash
    If Right$(destDir, 1) <> "\" Then destDir = destDir & "\"
    
    ' prepare result array - make it large enough
    ' (we will shrink it later)
    ReDim res(1, sourceFld.Files.Count + destFld.Files.Count) As Variant
    
    ' we need to ignore errors, in case file doesn't exist in destination dir
    On Error Resume Next
    
    ' load files of source directory into result array
    For Each sourceFile In sourceFld.Files
        ' this is the name of the file
        FileName = sourceFile.Name
        
        ' add file name to array
        index = index + 1
        res(0, index) = FileName
        
        ' add file name to collection (to be used later)
        col.Add FileName, FileName
        
        ' try to get a reference to destination file
        Set destFile = fso.GetFile(destDir & FileName)
        
        If Err Then
            Err.Clear
            ' file exists only in source directory
            res(1, index) = cdeSourceDirOnly
            
        Else
            ' if the file exists in both directories,
            '  start assuming it's the same file
            res(1, index) = cdeEqual
            
            ' compare file dates
            Select Case DateDiff("s", sourceFile.DateLastModified, _
                destFile.DateLastModified)
                Case Is < 0
                    ' source file is newer
                    res(1, index) = cdeSourceIsNewer
                Case Is > 0
                    ' source file is newer
                    res(1, index) = cdeSourceIsOlder
            End Select
            
            ' compare attributes
            If sourceFile.Attributes <> destFile.Attributes Then
                res(1, index) = res(1, index) Or cdeAttributesDiffer
            End If
            
            ' compare size
            If sourceFile.Size <> destFile.Size Then
                res(1, index) = res(1, index) Or cdeSizeDiffer
            End If
        End If
    Next
    
    ' now we only need to add all the files in destination directory
    ' that don't appear in the source directory
    For Each destFile In destFld.Files
        ' it's faster to search in the collection
        If col(destFile.Name) = "" Then
            ' we get here only if the filename isn't in the collection
            ' add the file to the result array
            index = index + 1
            res(0, index) = destFile.Name
            ' remember this only appears in the destination directory
            res(1, index) = cdeDestDirOnly
        End If
    Next
    
    ' trim and return the result
    If index > 0 Then
        ReDim Preserve res(1, index) As Variant
        CompareDirectories = res
    End If

End Function



' Synchronize two directories
'
' This routine compares source and dest directories and copies files
' from source that are newer than (or are missing in) the destination directory

' if TWOWAYSYNC is True, files are synchronized in both ways

' NOTE: requires the CompareDirectories routine and a reference to
'       the Microsoft Scripting Runtime type library

Sub SynchronizeDirectories(ByVal sourceDir As String, ByVal destDir As String, _
    Optional ByVal TwoWaySync As Boolean)
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long
    Dim copyDirection As Integer    ' 1=from source dir, 2=from dest dir,
                                    '  0=don't copy
    
    ' retrieve name of files in both directories
    Dim arr() As Variant
    arr = CompareDirectories(sourceDir, destDir)
    
    ' ensure that both dir names have a trailing backslash
    If Right$(sourceDir, 1) <> "\" Then sourceDir = sourceDir & "\"
    If Right$(destDir, 1) <> "\" Then destDir = destDir & "\"
    
    For index = 1 To UBound(arr, 2)
        ' assume this file doesn't need to be copied
        copyDirection = 0
        
        ' see whether files are
        Select Case arr(1, index)
            Case cdeEqual
                ' this file is the same in both directories
            Case cdeSourceDirOnly
                ' this file exists only in source directory
                copyDirection = 1
            Case cdeDestDirOnly
                ' this file exists only in destination directory
                copyDirection = 2
            Case Else
                If arr(1, index) = cdeAttributesDiffer Then
                    ' ignore files that differ only for their attributes
                ElseIf (arr(1, index) And cdeDateDiffer) = cdeSourceIsOlder Then
                    ' file in destination directory is newer
                    copyDirection = 2
                Else
                    ' in all other cases file in source dir should be copied
                    ' into dest dire
                    copyDirection = 1
                End If
        End Select
        
        If copyDirection = 1 Then
            ' copy from source dir to destination dir
            fso.CopyFile sourceDir & arr(0, index), destDir & arr(0, index), _
                True
        ElseIf copyDirection = 2 And TwoWaySync Then
            ' copy from destination dir to source dir
            ' (only if two-way synchronization has been requested)
            fso.CopyFile destDir & arr(0, index), sourceDir & arr(0, index), _
                True
        End If
        DoEvents
    Next
End Sub


