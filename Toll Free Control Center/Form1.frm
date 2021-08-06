VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " Rate "
      Enabled         =   0   'False
      Height          =   1470
      Left            =   7275
      TabIndex        =   20
      Top             =   2385
      Width           =   1860
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         TabIndex        =   21
         Top             =   1035
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".00"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".10"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   585
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".20"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   855
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Other"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   22
         Top             =   1110
         Width           =   1140
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   8100
      TabIndex        =   19
      Top             =   600
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1950
      Left            =   6540
      TabIndex        =   15
      Top             =   195
      Width           =   1770
      Begin VB.OptionButton Option7 
         Caption         =   "Option7"
         Height          =   195
         Left            =   510
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   915
         Width           =   885
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   195
         Left            =   345
         TabIndex        =   16
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   225
      Left            =   5700
      TabIndex        =   14
      Top             =   3855
      Width           =   1080
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   225
      Left            =   5535
      TabIndex        =   13
      Top             =   3240
      Width           =   1065
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   300
      Left            =   5295
      TabIndex        =   12
      Top             =   2580
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   330
      Left            =   5295
      TabIndex        =   11
      Top             =   2070
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   4
      Left            =   2535
      TabIndex        =   5
      Top             =   3660
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   3
      Left            =   2205
      TabIndex        =   4
      Top             =   2775
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      Top             =   1830
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   1245
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   570
      Width           =   2460
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   5160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      _Version        =   393216
      Format          =   115736577
      CurrentDate     =   44372
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   4
      Left            =   930
      TabIndex        =   10
      Top             =   3810
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   3
      Left            =   795
      TabIndex        =   9
      Top             =   2865
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Top             =   2055
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   1
      Left            =   615
      TabIndex        =   7
      Top             =   1230
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   6
      Top             =   675
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Stop
End Sub

'Private Sub Text1_GotFocus()
'   Text1.SelStart = 0
'   Text1.SelLength = Len(Text1.Text)
'End Sub
'
'
'Private Sub Text2_GotFocus()
'   Text2.SelStart = 0
'   Text2.SelLength = Len(Text2.Text)
'End Sub

Private Sub Text1_Change(Index As Integer)

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index).Text)
   Label1(Index).FontBold = True
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'   If Not IsNumeric(Chr$(KeyAscii)) Then
'      KeyAscii = 0
'   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
 Label1(Index).FontBold = False
End Sub

