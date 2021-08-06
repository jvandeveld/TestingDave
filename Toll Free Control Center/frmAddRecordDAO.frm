VERSION 5.00
Begin VB.Form frmAddRecordDAO 
   Caption         =   "Add Record DAO"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   480
      Left            =   3285
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2025
      Width           =   990
   End
   Begin VB.TextBox txtPhone 
      Height          =   330
      Left            =   1125
      TabIndex        =   0
      Top             =   270
      Width           =   1620
   End
   Begin VB.TextBox txtCarrier 
      Height          =   330
      Left            =   1140
      TabIndex        =   1
      Top             =   720
      Width           =   1620
   End
   Begin VB.TextBox txtCustomer 
      Height          =   330
      Left            =   1140
      TabIndex        =   2
      Top             =   1215
      Width           =   1620
   End
   Begin VB.CommandButton frmApply 
      Caption         =   "&OK"
      Height          =   480
      Left            =   2070
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2025
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Phone #:"
      Height          =   210
      Left            =   210
      TabIndex        =   6
      Top             =   300
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Carrier:"
      Height          =   210
      Left            =   210
      TabIndex        =   5
      Top             =   780
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Customer"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1215
      Width           =   960
   End
End
Attribute VB_Name = "frmAddRecordDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngRecordID As Long
Private m_database As DAO.Database
Event RecordUpdated(Phone As String, Carrier As String, Customer As String)

Property Let RecordID(value As Long)
   m_lngRecordID = value
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Form_Initialize()
   m_lngRecordID = 0
End Sub

Private Sub Form_Load()
   Set m_database = OpenDatabase("PhoneNumbers.mdb")
   If m_lngRecordID > 0 Then
      LoadCurrentRecord
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_database.Close
   Set m_database = Nothing
End Sub

Private Sub frmApply_Click()
   Dim rsDAO As DAO.Recordset
   Dim strSQL As String
   
   If m_lngRecordID = 0 Then
      strSQL = "Select * from datLotus800 where 1 = 0"
      Set rsDAO = m_database.OpenRecordset(strSQL, dbOpenDynaset)
      rsDAO.AddNew
   Else
      strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
      Set rsDAO = m_database.OpenRecordset(strSQL, dbOpenDynaset)
      rsDAO.Edit
   End If
      
   With rsDAO
      .Fields("Phone") = Trim$(txtPhone.Text)
      .Fields("Carrier") = Trim$(txtCarrier.Text)
      .Fields("Customer") = Trim$(txtCustomer.Text)
   End With
   
   rsDAO.Update
   Beep
   If m_lngRecordID > 0 Then
      RaiseEvent RecordUpdated(txtPhone.Text, txtCarrier.Text, txtCustomer.Text)
      'MsgBox "Record Updated", vbInformation, " "
   Else
      MsgBox "Record Added", vbInformation, " "
   End If
   Unload Me
End Sub

Private Sub LoadCurrentRecord()
   Dim strSQL As String
   strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
   Dim rsDAO As DAO.Recordset
   Set rsDAO = m_database.OpenRecordset(strSQL, dbOpenForwardOnly)
   If Not rsDAO.EOF Then
      txtPhone.Text = FieldValue(rsDAO, "Phone")
      txtCarrier.Text = FieldValue(rsDAO, "Carrier")
      txtCustomer.Text = FieldValue(rsDAO, "Customer")
   End If
End Sub

