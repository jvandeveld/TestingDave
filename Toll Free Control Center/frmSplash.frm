VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   2985
   ClientLeft      =   225
   ClientTop       =   1560
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   197.918
   ScaleMode       =   0  'User
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2430
      ScaleHeight     =   165
      ScaleWidth      =   3090
      TabIndex        =   1
      Top             =   2070
      Width           =   3090
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1395
      Top             =   2235
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   750
      Top             =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing Toll Free Control Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   2400
      TabIndex        =   2
      Top             =   2370
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing Toll Free Control Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   2385
      TabIndex        =   0
      Top             =   2355
      Width           =   2910
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   4
      Height          =   2940
      Left            =   30
      Top             =   30
      Width           =   5625
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   60
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   165
      Width           =   5520
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsProgress As Object
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As typRECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal bAlpha As Byte, ByVal alpha As Long) As Boolean
Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&
Const WS_EX_LAYERED = &H80000
Const LWA_ALPHA = &H2

Const SWP_HIDEWINDOW = &H80
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const SWP_SHOWWINDOW = &H40

Const WM_MOVE = &H3

Const SW_HIDE = 0
Const SW_MAXIMIZE = 3
Const SW_MINIMIZE = 6
Const SW_NORMAL = 1
Const SW_SHOWDEFAULT = 10
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOWNORMAL = 1
Const SE_ERR_DLLNOTFOUND = 32&
Private intParentForm As Integer
Private lngParentHandle As Long
'Private WithEvents mdwHook As dwSubClass
Private bytOpaqueValue As Byte

Property Get OpaqueValue() As Byte
   OpaqueValue = bytOpaqueValue
End Property

Private Sub SetOpaqueValue(OpaqueValue As Byte)
   bytOpaqueValue = OpaqueValue
   Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, _
     GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
   Call SetLayeredWindowAttributes(Me.hWnd, 0, bytOpaqueValue, _
     LWA_ALPHA)
   SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
End Sub

Public Sub ShowSplash()
   CheckForNewRevision
   'SetOnTop Me.hWnd
   Timer1.Enabled = True
   
   'Me.Show
   
'   Set clsProgress = CreateObject("OOPProgress.ProgressBar")
'
'   'If Not LenTrim(m_strMessage) Then
'   '   m_strMessage = "Loading Towed Vehicle Application"
'   'End If
'
'   Label2.Caption = m_strMessage
'   Label3.Caption = m_strMessage
'
'   'Me.Show
'   DoEvents
'   If Err.Number = 0 Then
'      clsProgress.ParentHandle = Me.hwnd
'      Set clsProgress.ContainerControl = Picture1
'      clsProgress.ShowProgress
'   End If
End Sub

Public Sub HideSplash()
   On Error Resume Next
   Err.Clear
   clsProgress.EndProgress
   Timer2.Enabled = True
End Sub

Private Sub Form_Load()

   bytOpaqueValue = 1
   SetOpaqueValue bytOpaqueValue
   SetOnTop Me.hWnd
   Me.Top = Screen.Height / 2 - Me.Height / 2
   Me.Left = Screen.Width / 2 - Me.Width / 2
         
   Set clsProgress = CreateObject("OOPProgress.ProgressBar")
   
   If Err.Number = 0 Then
      clsProgress.ParentHandle = Me.hWnd
      Set clsProgress.ContainerControl = Picture1
      clsProgress.ShowProgress
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Timer1.Enabled = False
   Set clsProgress = Nothing
End Sub


Private Sub Timer1_Timer()
   If bytOpaqueValue >= 251 Then
      bytOpaqueValue = 255
   Else
      bytOpaqueValue = bytOpaqueValue + 5
   End If
   SetOpaqueValue bytOpaqueValue
   If bytOpaqueValue = 255 Then
      Timer1.Enabled = False
      'Timer2.Enabled = True
   End If
End Sub

Private Sub Timer2_Timer()
   If bytOpaqueValue <= 5 Then
      bytOpaqueValue = 0
   Else
      bytOpaqueValue = bytOpaqueValue - 5
   End If
   SetOpaqueValue bytOpaqueValue
   If bytOpaqueValue = 0 Then
      Timer2.Enabled = False
      Unload Me
   End If
End Sub

Private Sub CheckForNewRevision()
   Dim strPath As String
   Dim strSource As String
   strSource = "\\192.168.1.105\apps\Toll Free Control Center\Toll Free Control Center.exe"
   If Dir(strSource) <> "" Then
      strPath = App.Path
      If Right$(strPath, 1) <> "\" Then
         strPath = strPath & "\"
      End If
      strPath = strPath & App.EXEName & ".exe"
      'Dim fso As FileSystemObject
      Dim fso As New FileSystemObject
      Dim strVersion As String
      strVersion = fso.GetFileVersion(strSource)
      Dim strV() As String
      strV() = Split(strVersion, ".")
      Dim intControl As Integer
      Dim intRunning As Integer
      intControl = Val(strV(0)) * 100 + Val(strV(1)) * 10 + Val(strV(3))
      intRunning = Val(App.Major) * 100 + Val(App.Minor) * 10 + Val(App.Revision)
      If intControl > intRunning Or ShiftKey Then
         'There is a newer version - get it
         Label3.Caption = "Downloading a new version."
         Label1.Caption = "Downloading a new version."
         
         On Error Resume Next
         Err.Clear
         If Dir(strPath & ".bak") <> "" Then
            Kill strPath & ".bak"
         End If

         Err.Clear
         Name strPath As strPath & ".bak"
         If Err.Number <> 0 Then
            Beep
            MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error renaming"
         End If
         
                  
         Err.Clear
         FileCopy strSource, strPath
                           
'         On Error Resume Next
'         Err.Clear
         
'         MsgBox strPath & ".bak"
'         If fso.FileExists(strPath & ".bak") Then
'            fso.DeleteFile strPath & ".bak"
'         End If
''         If Dir(strPath & ".bak") <> "" Then
''            Kill strPath & ".bak"
''         End If
'
'         MsgBox "10"
         
         
         clsProgress.EndProgress
         If Err.Number <> 0 Then
            Beep
            MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error copying from server"
            End
         Else
            Beep
            MsgBox "A new version has been downloaded." & vbCrLf & "Re-start the application.", vbInformation, " "
            End
         End If
      End If
   End If
End Sub

Public Function SetOnTop(hWindow As Long)
   On Error Resume Next
   Err.Clear
   Dim tmpRect As typRECT
   Dim lngTemp As Long
   lngTemp = GetWindowRect(hWindow, tmpRect)
   lngTemp = SetWindowPos(hWindow, -1, tmpRect.Left, tmpRect.Top, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW)
End Function
