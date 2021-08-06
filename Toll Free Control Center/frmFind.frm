VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4365
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   3135
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox chkSubString 
      Caption         =   "&Substring Search"
      Height          =   360
      Left            =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   810
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next "
      Default         =   -1  'True
      Height          =   315
      Left            =   3135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel "
      Height          =   330
      Left            =   3135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   945
      Width           =   1065
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find What: "
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event FindText(IsSubstring As Boolean, Text As String)
Event FindNext(IsSubstring As Boolean, Text As String)
Event NewButtonClicked()

Private m_strColName As String

Property Let ColName(value As String)
   m_strColName = value
End Property

Private Sub cboSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
   Dim strText As String
   strText = Trim$(cboSearch.Text)
   AddSearchText
   RaiseEvent FindText((chkSubString.value = vbChecked), strText)
End Sub

Private Sub cmdFindNext_Click()
   Dim strText As String
   strText = Trim$(cboSearch.Text)
   AddSearchText
   RaiseEvent FindNext((chkSubString.value = vbChecked), strText)
End Sub

Private Sub Form_Load()
   On Error Resume Next
   Err.Clear
'   Dim hwnd As Long
'   hwnd = Me.hwnd
'   Call SetClassLong(hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW)
   
   'SetDropDownWidth cboSearch.hwnd, 150
   'SendMessage_LONG cboSearch.hwnd, CB_SETEXTENDEDUI, 1, 0
   LoadHistoryList
End Sub

Private Sub LoadHistoryList()
   cboSearch.Clear
   Dim strHistory As String
   strHistory = GetSetting("CoachHouse", "TicketSakes\MainGrid", "SearchHistory", "")
   Dim strHist() As String
   strHist = Split(strHistory, "*&*")
   Dim i%
   For i% = UBound(strHist) To 0 Step -1
      cboSearch.AddItem (strHist(i%))
   Next i%
End Sub

Private Sub AddSearchText()
   Dim strText As String
   strText = Trim$(cboSearch.Text)
   Dim strHistory As String
   strHistory = GetSetting("CoachHouse", "TicketSakes\MainGrid", "SearchHistory", "")
   Dim strHist() As String
   strHist = Split(strHistory, "*&*")
   Dim i%
   For i% = 0 To UBound(strHist)
      If strHist(i%) = strText Then
         Exit For
      End If
   Next i%
   If i% > UBound(strHist) Then
      If UBound(strHist) >= 9 Then
         'save only the top 10
         strHistory = ""
         For i% = 0 To 8
            strHistory = strHistory & strHist(i%) & "*&*"
         Next i%
         strHistory = strHistory & strText
         SaveSetting "CoachHouse", "TicketSakes\MainGrid", "SearchHistory", strHistory
         LoadHistoryList
         cboSearch.Text = strText
      Else
         If LenTrim(strHistory) Then
            strHistory = strHistory & "*&*"
         End If
         strHistory = strHistory & strText
         SaveSetting "CoachHouse", "TicketSakes\MainGrid", "SearchHistory", strHistory
         cboSearch.AddItem strText, 0
      End If
   End If
End Sub

