VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChangeUnused 
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1335
      TabIndex        =   4
      Top             =   1185
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   503
      _Version        =   393216
      Format          =   125239297
      CurrentDate     =   44396
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   720
      Left            =   3255
      TabIndex        =   3
      Top             =   1875
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   720
      Left            =   4560
      TabIndex        =   2
      Top             =   1890
      Width           =   990
   End
   Begin VB.Label lblNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1335
      TabIndex        =   1
      Top             =   210
      Width           =   2130
   End
   Begin VB.Label lblAccNum 
      Caption         =   "Last Used:"
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Top             =   810
      Width           =   1800
   End
End
Attribute VB_Name = "frmChangeUnused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngRecordID As Long
Event RecordUpdated(LastUsed As String)

Property Let RecordID(value As Long)
   m_lngRecordID = value
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim dbData As DAO.Database
   Set dbData = OpenLocalDB
   Dim rsDAO As DAO.Recordset
   Dim strSQL As String
   strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
   Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
   rsDAO.Edit
   With rsDAO
      .Fields("AcctNo") = vbNull
      .Fields("Customer") = vbNull
      .Fields("LastUsed") = Format(DTPicker1.value, "mm/dd/yyyy")
      .Fields("Rate") = vbNull
      .Fields("Assigned") = vbNull
   End With
   rsDAO.Update
   dbData.Close
   Set dbData = Nothing
   Beep
   MsgBox "Number information updated.", vbInformation, " "
   RaiseEvent RecordUpdated(Format(DTPicker1.value, "mm/dd/yyyy"))
   Unload Me
End Sub

Private Sub Form_Load()
   Dim dbData As DAO.Database
   Set dbData = OpenLocalDB
   Dim strSQL As String
   Dim rsDAO As DAO.Recordset
   strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
   Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
   lblNumber.Caption = FormatPhone(FieldValue(rsDAO, "Phone"))
   DTPicker1.value = CDate(FieldValue(rsDAO, "LastUsed"))
   DTPicker1.MaxDate = Now
   dbData.Close
   Set dbData = Nothing
End Sub
