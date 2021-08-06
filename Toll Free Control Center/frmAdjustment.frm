VERSION 5.00
Begin VB.Form frmAdjustment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rate"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4110
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   975
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   4110
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1530
      Width           =   960
   End
   Begin VB.TextBox txtNotes 
      Height          =   855
      Left            =   420
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1215
      Width           =   3045
   End
   Begin VB.TextBox txtAdjustedBillable 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   420
      TabIndex        =   1
      Top             =   420
      Width           =   930
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes:"
      Height          =   210
      Left            =   405
      TabIndex        =   2
      Top             =   975
      Width           =   645
   End
   Begin VB.Label lblRate 
      Caption         =   "Adjusted Billable:"
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   135
      Width           =   1605
   End
End
Attribute VB_Name = "frmAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngAdjustmentID As Long
Private m_sngAdjustedBillable As Single
Private m_strNotes As String
Event ApplyAdjustment(AdjustmentID As Long, NewBillable As String, Notes As String)

Property Let AdjustmentID(value As Long)
   m_lngAdjustmentID = value
End Property

Property Let AdjustedBillable(value As Single)
   m_sngAdjustedBillable = value
End Property

Property Let Notes(value As String)
   m_strNotes = value
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   RaiseEvent ApplyAdjustment(m_lngAdjustmentID, Trim$(txtAdjustedBillable.Text), Trim$(txtNotes.Text))
End Sub

Private Sub Form_Load()
   txtAdjustedBillable.Text = m_sngAdjustedBillable
   txtNotes.Text = m_strNotes
End Sub

Private Sub txtAdjustedBillable_GotFocus()
   With txtAdjustedBillable
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtAdjustedBillable_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 32 Then
      If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
