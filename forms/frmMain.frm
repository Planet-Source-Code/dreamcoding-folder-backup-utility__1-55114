VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Back-up Utility"
   ClientHeight    =   2640
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5715
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHide 
      Caption         =   "h"
      Height          =   255
      Left            =   5445
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox chkStartwithWindows 
      Caption         =   "Start With Windows?"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   2370
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4895
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:09 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/22/2004"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameFolders 
      Caption         =   "Folder Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   555
      Width           =   5535
      Begin VB.TextBox txtDestination 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSource 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Destination:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Source:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CheckBox chkFolderChange 
      Caption         =   "on Folder Change"
      Height          =   375
      Left            =   3315
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtFileType 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMins 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   15
      Top             =   1260
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1935
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "/ or /"
      Height          =   255
      Left            =   2715
      TabIndex        =   14
      Top             =   195
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "minute(s)"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   195
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "File Type"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "mins"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Back up Every:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   1335
   End
   Begin VB.Menu menuSystray 
      Caption         =   "doesnt matter"
      Visible         =   0   'False
      Begin VB.Menu menuStart 
         Caption         =   "Start Monitor"
      End
      Begin VB.Menu menuStop 
         Caption         =   "Stop Monitor"
      End
      Begin VB.Menu menuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu menuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu menuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu menuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFolderChange_Click()
 Call RegSave(chkFolderChange, chkFolderChange.Value)
 txtMins.Enabled = Not CBool(chkFolderChange.Value)
End Sub

Private Sub chkStartwithWindows_Click()
 SaveSetting App.EXEName, "start", "autoBoot", chkStartwithWindows
 Select Case chkStartwithWindows
  Case 0: DoNotStartUp App.Path & "\" & App.EXEName & ".exe", App.EXEName
  Case 1: DoStartUp App.Path & "\" & App.EXEName & ".exe", App.EXEName
 End Select
End Sub

Private Sub cmdHide_Click()
 Me.Visible = False
 SaveSetting App.EXEName, "start", "hide", 1
End Sub

Private Sub cmdStart_Click()
On Error Resume Next
 If txtMins = "" Then
  msg = MsgBox("You must enter the amount of minutes before backup. Suggested Value is 1.")
  Exit Sub
 End If
 If txtDestination <> "" And txtSource <> "" Then
  status.Panels(1).Text = "Backing up..."
  'Copy files to destination
  'Call CopyFolder(txtSource, txtDestination, fasle, True)
  Call SynchronizeDirectoryTrees(txtSource, txtDestination, False)
  status.Panels(1).Text = "Complete."
  If cmdStart.Caption = "Start" Then
   If chkFolderChange.Value = 1 Then
    Call WatchDIR_Start(txtSource)
    Exit Sub
   End If
   If txtMins <> "" Then
    ourval = (60 * CInt(txtMins.Text))
    ourval = ourval * 1000
    Timer1.interval = ourval
    Timer1.Enabled = True
   End If
   cmdStart.Caption = "Stop"
  Else
   Call WatchDIR_End
   Timer1.Enabled = False
   cmdStart.Caption = "Start"
  End If
 Else
  msg = MsgBox("You must enter a Source Directory and Destination Directory!")
 End If
End Sub

Private Sub Command1_Click()
 txtSource = GetDirectory(Me)
 Call RegSave(txtSource, txtSource.Text)
End Sub

Private Sub Command2_Click()
 txtDestination = GetDirectory(Me)
 Call RegSave(txtDestination, txtDestination.Text)
End Sub

Private Sub Form_Load()
 With nid 'with system tray
  .cbSize = Len(nid)
  .hwnd = Me.hwnd
  .uId = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallBackMessage = WM_MOUSEMOVE
  .hIcon = Me.Icon 'use form's icon in tray
  .szTip = "File Monitor by Phishbowlerz"
 End With
 Shell_NotifyIcon NIM_ADD, nid 'add to tray
 txtSource = RegLoad(txtSource)
 ChkVal = RegLoad(chkFolderChange)
 If ChkVal <> "" Then
  chkFolderChange.Value = ChkVal
 End If
 txtDestination = RegLoad(txtDestination)
 Me.Visible = Not CBool(GetSetting(App.EXEName, "start", "hide", 0))
 chkStartwithWindows = GetSetting(App.EXEName, "start", "autoboot", 0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result, Action As Long
 If Me.ScaleMode = vbPixels Then
  Action = X
 Else
  Action = X / Screen.TwipsPerPixelX
 End If
 Select Case Action
  Case WM_LBUTTONDBLCLK 'Left Button Double Click
   Me.WindowState = vbNormal 'put into taskbar
   Result = SetForegroundWindow(Me.hwnd)
   Me.Show 'show form
  Case WM_RBUTTONUP 'Right Button Up
   Result = SetForegroundWindow(Me.hwnd)
   PopupMenu menuSystray 'popup menu, cool eh?
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Call WatchDIR_End
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Shell_NotifyIcon NIM_DELETE, nid
 Call WatchDIR_End
End Sub

Private Sub mnuEvery_Click()
 chkFolderChange.Enabled = True
 txtMins.Enabled = False
End Sub

Private Sub menuExit_Click()
 Unload Me: End
End Sub

Private Sub menuHide_Click()
 Me.Visible = False
End Sub

Private Sub menuShow_Click()
 Me.Visible = True
End Sub

Private Sub menuStart_Click()
 If chkFolderChange.Value = 1 Then
  Call WatchDIR_Start(txtSource)
  Exit Sub
 End If
 If txtMins <> "" Then
  ourval = (60 * CInt(txtMins.Text))
  ourval = ourval * 1000
  Timer1.interval = ourval
  Timer1.Enabled = True
 End If
 cmdStart.Caption = "Stop"
End Sub

Private Sub menuStop_Click()
  Call WatchDIR_End
  Timer1.Enabled = False
  cmdStart.Caption = "Start"
End Sub

Private Sub Timer1_Timer()
 status.Panels(1).Text = "Backing up..."
 'Copy files to destination
 'Call CopyFolder(txtSource, txtDestination, fasle, True)
 Call SynchronizeDirectoryTrees(txtSource, txtDestination, False)
 status.Panels(1).Text = "Complete."
End Sub

Private Sub txtMins_Change()
 Call RegSave(txtMins, txtMins.Name)
End Sub
