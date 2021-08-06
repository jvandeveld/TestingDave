VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Toll Free Control Center"
   ClientHeight    =   2445
   ClientLeft      =   705
   ClientTop       =   900
   ClientWidth     =   5940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1687.583
   ScaleMode       =   0  'User
   ScaleWidth      =   5577.968
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   435
      TabIndex        =   5
      Top             =   1770
      Width           =   1980
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   165
      Picture         =   "FRMABOUT.frx":0000
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   210
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   4530
      TabIndex        =   0
      Top             =   1875
      Width           =   1260
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright © 2021"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   3540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5520.686
      Y1              =   1190.626
      Y2              =   1190.626
   End
   Begin VB.Label lblTitle 
      Caption         =   "Toll Free Phone Number Control Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1155
      TabIndex        =   2
      Top             =   300
      Width           =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5521.625
      Y1              =   1200.979
      Y2              =   1200.979
   End
   Begin VB.Label lblVersion 
      Caption         =   "1.0.101"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1170
      TabIndex        =   3
      Top             =   690
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
'Private Const SW_SHOW = 5
'
'
'' Reg Key Security Options...
'Const READ_CONTROL = &H20000
'Const KEY_QUERY_VALUE = &H1
'Const KEY_SET_VALUE = &H2
'Const KEY_CREATE_SUB_KEY = &H4
'Const KEY_ENUMERATE_SUB_KEYS = &H8
'Const KEY_NOTIFY = &H10
'Const KEY_CREATE_LINK = &H20
'Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'
'' Reg Key ROOT Types...
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const ERROR_SUCCESS = 0
'Const REG_SZ = 1                         ' Unicode nul terminated string
'Const REG_DWORD = 4                      ' 32-bit number
'
'Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'Const gREGVALSYSINFOLOC = "MSINFO"
'Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'Const gREGVALSYSINFO = "PATH"
'
'Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    'lblSubComponent.Caption = lblSubComponent.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub


