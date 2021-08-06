VERSION 5.00
Begin VB.Form frmAddRecordSQL 
   Caption         =   "Add Record SQL"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton frmApply 
      Caption         =   "Add"
      Height          =   480
      Left            =   3030
      TabIndex        =   6
      Top             =   2235
      Width           =   990
   End
   Begin VB.TextBox txtCustomer 
      Height          =   330
      Left            =   1290
      TabIndex        =   5
      Top             =   1515
      Width           =   1620
   End
   Begin VB.TextBox txtCarrier 
      Height          =   330
      Left            =   1290
      TabIndex        =   4
      Top             =   1020
      Width           =   1620
   End
   Begin VB.TextBox txtPhone 
      Height          =   330
      Left            =   1275
      TabIndex        =   3
      Top             =   555
      Width           =   1620
   End
   Begin VB.Label Label3 
      Caption         =   "Customer"
      Height          =   210
      Left            =   405
      TabIndex        =   2
      Top             =   1515
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Carrier:"
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Phone #:"
      Height          =   210
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   960
   End
End
Attribute VB_Name = "frmAddRecordSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_database As DAO.Database

Private Sub Form_Load()
   Set m_database = OpenDatabase("PhoneNumbers.mdb") 'open database
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_database.Close
   Set m_database = Nothing
End Sub

Private Sub frmApply_Click()
   Dim strSQL As String
   strSQL = "Insert into datLotus800 " & vbCrLf
   strSQL = strSQL & "(Phone, Carrier,Customer)" & vbCrLf
   strSQL = strSQL & "Values ('" & txtPhone.Text & "','"
   strSQL = strSQL & txtCarrier.Text & "','"
   strSQL = strSQL & txtCustomer.Text & "')"
   On Error Resume Next
   Err.Clear
   m_database.Execute strSQL, dbFailOnError
   If Err.Number <> 0 Then
      Stop
   End If
End Sub

