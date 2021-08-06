VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLotus800 
      Caption         =   "Lotus 800 Grid"
      Height          =   480
      Left            =   2865
      TabIndex        =   5
      Top             =   1950
      Width           =   1650
   End
   Begin VB.CommandButton cmdAssignNum 
      Caption         =   "Assign Available Phone Number to Client"
      Height          =   705
      Left            =   2595
      TabIndex        =   4
      Top             =   495
      Width           =   2100
   End
   Begin VB.CommandButton cmdAddRecordDAO 
      Caption         =   "Add record DAO"
      Height          =   525
      Left            =   585
      TabIndex        =   3
      Top             =   2580
      Width           =   1155
   End
   Begin VB.CommandButton cmdGridUnbound 
      Caption         =   "SQL Grid (UnBound)"
      Height          =   585
      Left            =   600
      TabIndex        =   2
      Top             =   1185
      Width           =   1230
   End
   Begin VB.CommandButton cmdAddRecord 
      Caption         =   "Add record SQL"
      Height          =   525
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdSQLGrid 
      Caption         =   "SQL Grid (Bound)"
      Height          =   585
      Left            =   570
      TabIndex        =   0
      Top             =   465
      Width           =   1230
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddRecord_Click()
   Dim frmForm As New frmAddRecordSQL
   frmForm.Show vbModal, Me
End Sub

Private Sub cmdAddRecordDAO_Click()
   Dim frmForm As New frmAddRecordDAO
   frmForm.Show vbModal, Me
End Sub

Private Sub cmdAssignNum_Click()
   Dim frmForm As New frmAssignNum
   frmForm.Show vbModal, Me
End Sub

Private Sub cmdGridUnbound_Click()
   Dim frmForm As New frmSQLGridUnBound
   frmForm.Show , Me
End Sub

Private Sub cmdLotus800_Click()
   Dim frmForm As New frmLotus800
   frmForm.Show vbModeless, Me
End Sub

Private Sub cmdSQLGrid_Click()
   Dim frmForm As New frmSQLGridBound
   frmForm.Show vbModal, Me
End Sub

