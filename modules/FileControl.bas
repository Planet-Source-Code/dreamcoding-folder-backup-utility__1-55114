Attribute VB_Name = "FileControl"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Find Files
Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long


'Browse for DIR

'API's for selecting a windows directory
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

'Constants
Public OurFiles As String 'Returns list of files found in DIR

Function IfFileExists(ByVal sFilename As String) As Boolean
Dim I As Long
On Error Resume Next

    I = Len(Dir$(sFilename))
    
    If Err Or I = 0 Then
        IfFileExists = False
    Else
        IfFileExists = True
    End If

End Function



  Sub SaveTextAppend(Path As String, StringName As String)
 
 
    On Error Resume Next

    Open Path$ For Append As #1
        
        Print #1, StringName
       
    Close #1


End Sub
  Sub SaveTextOutput(Path As String, StringName As String)
 
 
    On Error Resume Next

    Open Path$ For Output As #1
        
        Print #1, StringName
       
    Close #1


End Sub

Public Function FolderExist(ByVal pName As String) As Boolean
Rem ---------------------------------
Rem Check folder
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  FolderExist = lFso.FolderExists(pName)
  Exit Function
Cerr:
  FolderExist = False
End Function

Public Function FileExists2(sFilename As String) As Boolean

    If Len(sFilename$) = 0 Then
        
FileExists2 = False
        Exit Function
    End If
    If Len(Dir$(sFilename$)) Then
        
FileExists2 = True
    Else
        
FileExists2 = False
    End If
End Function
Public Function FileExists(ByVal pFilename As String) As Boolean
Rem ---------------------------------
Rem Check for file existence
Rem ---------------------------------

On Error GoTo FileExists_Err
  If FileLen(pFilename) > 0 Then
    FileExists = True
  Else
    FileExists = False
  End If
  GoTo FileExists_Out
FileExists_Err:
  FileExists = False
FileExists_Out:
End Function
Public Function DirExists(strDir As String) As Boolean



'change C:\MyDir

strDir = Dir(strDir, vbDirectory)

If (strDir = "") Then
 DirExists = False
 Else
 DirExists = True
 End If
End Function


Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim savelist As Long
    On Error Resume Next
    fe = FreeFile
    Open Directory$ For Output As #fe
    For savelist = 0 To TheList.ListCount - 1
    bufff = TheList.List(savelist)
    bufff = Replace(bufff, Chr(13), "ªä")
        Print #fe, Trim(bufff)
    Next savelist
    Close #fe
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub

Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim MyString As String
    On Error Resume Next
    If IfFileExists(Directory) = False Then Exit Sub
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Combo.AddItem MyString$
    Wend
    Close #1
End Sub

Public Function FileGetAttributes(TheFile As String) As Long
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub

Public Function GetFromINI(Section As String, key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   key$ = LCase$(key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
'  Get and Write To INI work like this
'  like lets say you have a textbox and you
'  want that box to say what you put in it last time you unloaded
'  the program, then you do this
'  in the unload proc, put
'  writetoini("Prefences","countervalue","39243","C:\windows\desktop\crvalue.ini")
'  in the load proc put
'  text1 = GetFromINI("Prefences","countervalue","C:\windows\desktop\crvalue.ini")
'         get it?

Public Sub WriteToINI(Section As String, key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(key$), KeyValue$, Directory$)
End Sub
Function LoadList(lst, File)
On Error GoTo File
lst.Clear
Dim l00AE As String
Dim l00B0 As Long
 A = CurDir
 ChDir A
 Open File For Input As #1
 Do
    Input #1, l00AE$
    l00AE$ = Trim(l00AE$)
    If l00AE$ <> "" Then lst.AddItem l00AE$
    l00B0 = DoEvents()
  Loop Until EOF(1)
    Close #1
File:
End Function

 Function LoadText(File, TextBoxOrLabel As TextBox)
 On Error GoTo Fle
A = CurDir
ChDir A
Open File For Input As #1
Do
Input #1, Text$
TheText = TheText & Text$ & vbNewLine
Loop Until EOF(1)
Close #1
TextBoxOrLabel = TheText
Fle:
 End Function
 

Public Function GetDriveTypes(DriveLetter As String)
    Select Case GetDriveTypes(DriveLetter)
        Case 2
            GetDriveTypes = "Removable"
        Case 3
            GetDriveTypes = "Drive Fixed"
        Case Is = 4
            Debug.Print "Remote"
        Case Is = 5
            Debug.Print "Cd-Rom"
        Case Is = 6
            Debug.Print "Ram disk"
        Case Else
            Debug.Print "Unrecognized"
    End Select
End Function

Private Function FindFile(sFile As String, sRootPath As String) As String
    ' Search for the file specified and retu
    '     rn the full path if found
    Dim sPathBuffer As String
    Dim iEnd As Integer
    
    'Allocate some buffer space (you may nee
    '     d more)
    sPathBuffer = Space(512)
    


    If SearchTreeForFile(sRootPath, sFile, sPathBuffer) Then
        'Strip off the null string that will be
        '     returned following the path name
        iEnd = InStr(1, sPathBuffer, vbNullChar, vbTextCompare)
        sPathBuffer = Left$(sPathBuffer, iEnd - 1)
        FindFile = sPathBuffer
    Else
        FindFile = vbNullString
    End If
End Function

Public Function ExtractAll(ByVal FilePath As String, Optional DefaultExtension As String = "*.*") As String
Dim RetVal As String
Dim MyName As String
Dim SubDir(500) As String

 If Mid$(FilePath, Len(FilePath), 1) <> "\" Then
  FilePath = FilePath + "\"
 End If
 I = 0
On Error GoTo nodir
 MyName = Dir(FilePath, vbDirectory)
 Do While MyName <> ""
  If MyName <> "." And MyName <> ".." And MyName <> "Directories" Then
   If (GetAttr(FilePath & MyName) And vbDirectory) = vbDirectory Then
    SubDir(I) = MyName
    I = I + 1
   End If
  End If
  MyName = Dir
  If MyName <> "." And MyName <> ".." Then
  OurFiles = OurFiles & MyName & "^"
  End If
 Loop
 DoEvents
 MyName = Dir(FilePath + DefaultExtension, vbNormal)
 Do While MyName <> ""
  If Right(LCase(MyName), 3) = Right(DefaultExtension, 3) Then
   RetVal = RetVal + FilePath + MyName + "^"
  End If
  MyName = Dir
  OurFiles = OurFiles & MyName & "^"
 Loop
nodir:
Dim RetVal2 As String
 For T = 0 To I - 1
  RetVal2 = RetVal2 + ExtractAll(FilePath & SubDir(T) + "\")
 Next T
 'ExtractAll = RetVal + RetVal2
 ExtractAll = OurFiles
 Exit Function
End Function
Public Function GetDirectory(frm As Form) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim Path$, Pos%
    
    bi.hOwner = frm.hWnd
    bi.pidlRoot = 0&
    bi.lpszTitle = "Select directory..."
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(bi)
    Path = Space$(256)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
       Pos = InStr(Path, Chr$(0))
       GetDirectory = Left(Path, Pos - 1)
    End If
    Call CoTaskMemFree(pidl)
End Function

Function SetBytes(Bytes) As String

On Error GoTo hell

Bytes = CLng(Bytes)

If Bytes >= 1073741824 Then
    SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") _
         & " GB"
ElseIf Bytes >= 1048576 Then
    SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
ElseIf Bytes >= 1024 Then
    SetBytes = Format(Bytes / 1024, "#0.00") & " KB"
ElseIf Bytes < 1024 Then
    SetBytes = Fix(Bytes) & " Bytes"
End If

Exit Function
hell:
SetBytes = "0 Bytes"
End Function


Public Function CopyFolder(ByVal pSource As String, ByVal pDest As String, Optional ByVal pMove As Boolean, Optional ByVal AutoReplace As Boolean) As Boolean
Rem ---------------------------------
Rem Copy Folder
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
Dim Lok As Boolean
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  If FileExists(pDest) Then
   If AutoReplace = True Then
     Lok = True
   Else
     If MsgBox("Destination already exists. Do you want replace?", vbQuestion + vbOKCancel, Mtitle) = vbOK Then
       Lok = True
     Else
       Lok = False
     End If
   End If
  Else
    Lok = True
  End If
  If Lok Then
    If pMove Then
      lFso.MoveFolder pSource, pDest
    Else
      lFso.CopyFolder pSource, pDest, True
    End If
  End If
  CopyFolder = True
  Exit Function
Cerr:
  CopyFolder = False
End Function
Public Sub DeleteDirectoryOLD(ByVal dir_name As String, DeleteDirectoryFolder As Boolean)
Dim file_name As String
Dim files As Collection
Dim I As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    file_name = Dir$(dir_name & "\*.*", vbReadOnly + _
        vbHidden + vbSystem + vbDirectory)
    Do While Len(file_name) > 0
        If (file_name <> "..") And (file_name <> ".") Then
            files.Add dir_name & "\" & file_name
        End If
        file_name = Dir$()
    Loop

    ' Delete the files.
    For I = 1 To files.Count
        file_name = files(I)
        ' See if it is a directory.
        If GetAttr(file_name) And vbDirectory Then

                    ' It is a directory. Delete it.
                    RmDir file_name
                   
        Else
        'add temp file
        If I = files.Count Then
        Call SaveTextOutput(App.Path & "\temp.txt", "Temp")
          Call FileCopy(App.Path & "\temp.txt", dir_name)
          
        End If
            ' It's a file. Delete it.
            'lblStatus.Caption = file_name
            'lblStatus.Refresh
            SetAttr file_name, vbNormal
            Kill file_name
        End If
    Next I

    ' The directory is now empty. Delete it.
    'lblStatus.Caption = dir_name
    'lblStatus.Refresh
    If DeleteDirectoryFolder = True Then
    RmDir dir_name
    End If
End Sub

Public Sub DeleteDirectory(ByVal dir_name As String, DeleteDirectoryFolder As Boolean)

'Below is a section of code called: Avoid Auto-default update
'This is only if we are deleting all the favorites
'When this happens, the system sometimes will auto-update the
'folder with it's own links. e.g. network admin sets up
'We insert a link, and then later in local code delete it.
'The inserting of a temp link fools the special folder into thinking
'it's not completely empty.
'Avoid Auto-default update


Dim file_name As String
Dim files As Collection
Dim I As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    file_name = Dir$(dir_name & "\*.*", vbReadOnly + _
        vbHidden + vbSystem + vbDirectory)
    Do While Len(file_name) > 0
        If (file_name <> "..") And (file_name <> ".") Then
        'Default1 = InStr(1, file_name, "[DEFAULT.]")
        'If Default1 Then
        'Else
            files.Add dir_name & "\" & file_name
        '    End If
        End If
        file_name = Dir$()
    Loop

    ' Delete the files.
    For I = 1 To files.Count
        file_name = files(I)
        ' See if it is a directory.
       If GetAttr(file_name) And vbDirectory Then
            ' It is a directory. Delete it.
            DeleteDirectory file_name, True
        Else
                              
        
            ' It's a file. Delete it.
           ' lblStatus.Caption = file_name
           ' lblStatus.Refresh
            SetAttr file_name, vbNormal
            
            Kill file_name
        
       If I = files.Count Then
         
        'FavoritesPath = GetFolder(ftFavorites, Form1.hWnd)
        
        If dir_name = FavoritesPath Then
        Call CreateInternetShortCut("http://www.codepiler.com", GetFavoritesPath & "\temp.url")
        End If
         End If
     
        End If
    Next I

    ' The directory is now empty. Delete it.
   ' lblStatus.Caption = dir_name
   ' lblStatus.Refresh
    If DeleteDirectoryFolder = True Then
    RmDir dir_name
    End If
    
End Sub

Public Function CreateInternetShortCut(TargetURL As String, _
  FullPath As String)
  
'PURPOSE: Creates Internet ShortCut Link
'PARAMETERS: TargetURL: The URL to link to
             'FullPath: The File Path, should have .url extension
'EXAMPLE:  CreateInternetShortCut "http://www.freevbcode.com", _
      '"C:\Windows\Desktop\FreeVBCode.url"

'RETURNS: True if Successful, False otherwise

On Error GoTo ErrorHandler:
Dim iFile As Integer

iFile = FreeFile

Open FullPath For Output As #iFile
Print #iFile, "[InternetShortcut]"
Print #iFile, "URL=" & TargetURL
Close #iFile
CreateInternetShortCut = True
ErrorHandler:

End Function

Public Sub RegSave(ControlToRemember As Control, ControlValue As String)
Call SaveSetting(App.EXEName, "Settings", ControlToRemember.Name, ControlValue)
End Sub
Public Function RegLoad(ControlToRemember As Control)
On Error Resume Next
Dim ControlName As String
ControlName = ControlToRemember.Name
RegLoadSet = GetSetting(App.EXEName, "Settings", ControlName)
If RegLoadSet <> "" Then
RegLoad = RegLoadSet
Else
RegLoad = ""
End If



End Function
