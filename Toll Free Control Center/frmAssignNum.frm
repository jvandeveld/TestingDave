VERSION 5.00
Begin VB.Form frmAssignNum 
   Caption         =   "Assign Number to New Client"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   2010
      TabIndex        =   6
      Top             =   2340
      Width           =   720
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3210
      TabIndex        =   5
      Top             =   2340
      Width           =   720
   End
   Begin VB.CommandButton cmdDifferentNumber 
      Caption         =   "Next Available Number"
      Height          =   735
      Left            =   420
      TabIndex        =   4
      Top             =   2085
      Width           =   1050
   End
   Begin VB.TextBox txtClientName 
      Height          =   660
      Left            =   615
      TabIndex        =   3
      Top             =   1275
      Width           =   3270
   End
   Begin VB.TextBox txtNewNumber 
      Height          =   435
      Left            =   705
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   3210
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Client Name Here:"
      Height          =   390
      Left            =   1290
      TabIndex        =   2
      Top             =   915
      Width           =   1890
   End
   Begin VB.Label Lbl 
      Caption         =   "Number To Assign:"
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   90
      Width           =   1560
   End
End
Attribute VB_Name = "frmAssignNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngRecordID As Long


Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim dbData As DAO.Database
   'Set dbData = OpenDatabase("PhoneNumbers.mdb")
   Set dbData = OpenLocalDB

   Dim rsDAO As DAO.Recordset
   Dim sqlDates As String
   sqlDates = "Select * from datLotus800 where id = " & m_lngRecordID
   
   Set rsDAO = dbData.OpenRecordset(sqlDates, dbOpenDynaset)
   
   If rsDAO.EOF And rsDAO.BOF Then
      'the record must have been deleted - very unlikely
   Else
      rsDAO.MoveFirst
      rsDAO.Edit
      rsDAO.Fields("Customer") = Trim$(txtClientName.Text)
      rsDAO.Update
   End If
   
   dbData.Close
   Set dbData = Nothing
   
End Sub

Private Sub Form_Load()
   loadNum
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Err.Clear
   'm_database.Close
End Sub


Private Sub loadNum()
   Dim dbData As DAO.Database
   Set dbData = OpenDatabase("PhoneNumbers.mdb")

   Dim rsDAO As DAO.Recordset
   Dim sqlDates As String
   sqlDates = "Select top 1 id,Phone, LastUsed from datLotus800 where lastUsed is not null order by lastUsed"
   
   Set rsDAO = dbData.OpenRecordset(sqlDates, dbOpenForwardOnly)
   If Not rsDAO.EOF Then
      m_lngRecordID = Val(FieldValue(rsDAO, "id"))
      txtNewNumber.Text = FieldValue(rsDAO, "Phone") & "  last used:  " & FieldValue(rsDAO, "LastUsed")
   Else
      Stop
   End If
   
   dbData.Close
   Set dbData = Nothing
   
End Sub

