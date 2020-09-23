Attribute VB_Name = "FileDetect"
'Module
Option Explicit

Public Const INFINITE = &HFFFF

Public Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Public Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Public Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Public Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                 FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                 FILE_NOTIFY_CHANGE_LAST_WRITE

Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_ABANDONED = &H80
Public Const WAIT_IO_COMPLETION = &HC0
Public Const WAIT_TIMEOUT = &H102
Public Const STATUS_PENDING = &H103

'Form   ,Add three Button controls,one listbox control,two label controls for Form


Dim hChangeHandle As Long
Dim hWatched As Long
Dim terminateFlag As Long



Public Sub WatchDIR_End()

   If hWatched > 0 Then Call WatchDelete(hWatched)
   hWatched = 0
   
   
End Sub


Public Sub WatchDIR_Start(watchPath As String)

   Dim r As Long
   'Dim watchPath As String
   Dim watchStatus As Long
   
   terminateFlag = False
   
   WatchChangeAction watchPath

   MsgBox "Beginning watching of folder " & watchPath & " .. press OK"
   

   hWatched = WatchCreate(watchPath, FILE_NOTIFY_FLAGS)

   watchStatus = WatchDirectory(hWatched, 100)

   If watchStatus = 0 Then
       WatchChangeAction watchPath
       MsgBox "The watched directory has been changed.  Resuming watch..."
       
       Do
            watchStatus = WatchResume(hWatched, 100)
            If watchStatus = -1 Then
                  MsgBox "Watching has been terminated for " & watchPath
            Else: WatchChangeAction watchPath
                  MsgBox "The watched directory has been changed again."
            End If
      DoEvents
       Loop While watchStatus = 0
   Else
     ' MsgBox "Watching has been terminated for " & watchPath
   End If
End Sub



Private Function WatchCreate(lpPathName As String, flags As Long) As Long
   WatchCreate = FindFirstChangeNotification(lpPathName, False, flags)
End Function


Private Sub WatchDelete(hWatched As Long)
   Dim r As Long
   terminateFlag = True
   DoEvents
   r = FindCloseChangeNotification(hWatched)
End Sub


Private Function WatchDirectory(hWatched As Long, interval As Long) As Long
   Dim r As Long
   Do
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   Loop While r <> 0 And terminateFlag = False
   WatchDirectory = r
End Function

Private Function WatchResume(hWatched As Long, interval) As Boolean
   Dim r As Long
   r = FindNextChangeNotification(hWatched)
   Do
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   Loop While r <> 0 And terminateFlag = False
   WatchResume = r
End Function
Private Sub WatchChangeAction(fPath As String)
   Dim fName As String
   'List1.Clear
   fName = Dir(fPath & "\" & "*.txt")
   If fName > "" Then
   '   List1.AddItem "path: " & vbTab & fPath
   '   List1.AddItem "file: " & vbTab & fName
   '   List1.AddItem "size: " & vbTab & FileLen(fPath & "\" & fName)
   '   List1.AddItem "attr: " & vbTab & GetAttr(fPath & "\" & fName)
   End If
End Sub

