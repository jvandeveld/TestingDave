VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUnAssign 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Un-assign number"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3510
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2505
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Remove Assignment"
      Default         =   -1  'True
      Height          =   990
      Left            =   3510
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmUnAssign.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   960
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   1785
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   635
      _Version        =   393216
      Format          =   59768833
      CurrentDate     =   42662
   End
   Begin VB.Label Label3 
      Caption         =   "Date Last Used"
      Height          =   165
      Left            =   240
      TabIndex        =   7
      Top             =   1515
      Width           =   1410
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   840
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number:"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   285
      Width           =   1110
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1515
      TabIndex        =   1
      Top             =   195
      Width           =   2040
   End
   Begin VB.Label lblCustomer 
      AutoSize        =   -1  'True
      Caption         =   "lblCustomer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1605
      TabIndex        =   0
      Top             =   810
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUnAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_boolLoaded As Boolean
Private m_lngID As Long
Private m_strPhoneNumber As String
Private m_strLastUsed As String
Private m_strAccount As String
Private m_strName As String
Private m_sngRate As Single
Event DoUpdate(ID As Long, dtLastUsed As Date)

Property Let ID(value As Long)
   m_lngID = value
End Property

Property Let LastUsed(value As String)
   m_strLastUsed = value
End Property

Property Let Account(value As String)
   m_strAccount = value
End Property

Property Let CustomerName(value As String)
   m_strName = value
End Property

Property Let Rate(value As Single)
   m_sngRate = value
End Property

Property Let PhoneNumber(value As String)
   m_strPhoneNumber = value
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Beep
   If MsgBox("Are you sure you want to release this number?", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
      RaiseEvent DoUpdate(m_lngID, DTPicker1.value)
   End If
End Sub

Private Sub Form_Load()
   On Error Resume Next
   Err.Clear
   Dim hwnd As Long
   hwnd = Me.hwnd
   Call SetClassLong(hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW)
   lblPhone.Caption = m_strPhoneNumber
   lblAccount.Caption = m_strAccount
   lblCustomer.Caption = Replace(m_strName, "&", "&&")
   
   DTPicker1.value = Now
End Sub
