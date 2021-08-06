VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ResetFilters 
      Caption         =   "Reset Filters"
      Height          =   720
      Left            =   3735
      TabIndex        =   15
      Top             =   435
      Width           =   1200
   End
   Begin VB.ComboBox cboPhone 
      Height          =   315
      ItemData        =   "frmFilter.frx":0000
      Left            =   945
      List            =   "frmFilter.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2700
      Width           =   915
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   405
      Left            =   2055
      TabIndex        =   9
      Top             =   2625
      Width           =   2640
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   405
      Left            =   2055
      TabIndex        =   12
      Top             =   3075
      Width           =   2640
   End
   Begin VB.ComboBox cboOperator2 
      Height          =   315
      ItemData        =   "frmFilter.frx":0030
      Left            =   930
      List            =   "frmFilter.frx":003D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3150
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtDNIS 
      Height          =   405
      Left            =   2055
      TabIndex        =   7
      Top             =   2130
      Width           =   1275
   End
   Begin VB.ComboBox cboOperator 
      Height          =   315
      ItemData        =   "frmFilter.frx":0060
      Left            =   945
      List            =   "frmFilter.frx":006D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1710
      Width           =   915
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   1650
      Width           =   2640
   End
   Begin VB.ListBox lstTypes 
      Height          =   1005
      IntegralHeight  =   0   'False
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   405
      Width           =   2115
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   690
      Left            =   3645
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmFilter.frx":0090
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4125
      UseMaskColor    =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   690
      Left            =   4980
      TabIndex        =   0
      Top             =   4125
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Phone #:"
      Height          =   285
      Left            =   195
      TabIndex        =   14
      Top             =   2715
      Width           =   885
   End
   Begin VB.Label Label4 
      Caption         =   "AcctNo:"
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   3165
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "DNIS:"
      Height          =   270
      Left            =   1260
      TabIndex        =   10
      Top             =   2265
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   285
      Left            =   195
      TabIndex        =   5
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   285
      Left            =   195
      TabIndex        =   3
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_boolFormLoaded As Boolean
Event ApplyChanges(Types As String, CustomerWhere As String, DNIS As String, AcctNo As String, Phone As String)

Public Sub ResetParams()
   'reset controls to default values
   Dim i%
   For i% = 0 To lstTypes.ListCount - 1
      lstTypes.Selected(i%) = True
   Next i%
   txtAcctNo.Text = ""
   txtName.Text = ""
   txtDNIS.Text = ""
   txtPhoneNum.Text = ""
End Sub

Private Sub cboPhone_Click()
   If cboPhone.ListIndex = 2 Then
      txtPhoneNum.Text = UnformatPhone(txtPhoneNum.Text)
      'UnformatPhone (txtPhoneNum.Text)
   Else
      txtPhoneNum.Text = UnformatPhone(txtPhoneNum.Text)
      txtPhoneNum.Text = FormatPhone(txtPhoneNum.Text)
   End If
   'Debug.Print cboPhone.Text
End Sub

Private Sub cmdApply_Click()
   ApplyParams
   Me.Hide
End Sub

Private Sub Command1_Click()
   'Unload Me
   Me.Hide
End Sub

Private Sub Form_Load()
   If GetIniSetting(dz_IgnoreEscape) = "Y" Then
      Command1.Cancel = False
   Else
      Command1.Cancel = True
   End If
   cboOperator.ListIndex = 2
   cboOperator2.ListIndex = 1
   cboPhone.ListIndex = 1
   loadtypes
   m_boolFormLoaded = True
End Sub

Private Sub ApplyParams()
   Dim strTypes As String
   Dim strCustomerWhere As String
   Dim strDNIS As String
   Dim strAcctNo As String
   Dim strPhoneNo As String
   strTypes = ""
   Dim i%
   If lstTypes.SelCount <> 4 Then
      For i% = 0 To lstTypes.ListCount - 1
         If lstTypes.Selected(i%) Then
            If LenTrim(strTypes) Then
               strTypes = strTypes & ","
            End If
            strTypes = strTypes & " '" & lstTypes.List(i%) & "'"
         End If
      Next i%
   End If
   
   If LenTrim(txtName.Text) Then
      Select Case cboOperator.ListIndex
         Case 0
            strCustomerWhere = "Customer = '" & Trim$(txtName.Text) & "'"
         Case 1
            strCustomerWhere = "Customer like '" & Trim$(txtName.Text) & "*'"
         Case 2
            strCustomerWhere = "Customer like '*" & Trim$(txtName.Text) & "*'"
      End Select
   End If
   If LenTrim(txtAcctNo.Text) Then
      Select Case cboOperator.ListIndex
         Case 0
            strAcctNo = "AcctNo = '" & Trim$(txtAcctNo.Text) & "'"
         Case 1
            strAcctNo = "AcctNo like '" & Trim$(txtAcctNo.Text) & "*'"
         Case 2
            strAcctNo = "AcctNo like '*" & Trim$(txtAcctNo.Text) & "*'"
      End Select
   End If
   If LenTrim(txtPhoneNum.Text) Then
      Select Case cboPhone.ListIndex
         Case 0
            strPhoneNo = "Phone = '" & Trim$(UnformatPhone(txtPhoneNum.Text)) & "'"
         Case 1
            strPhoneNo = "Phone like '" & Trim$(UnformatPhone(txtPhoneNum.Text)) & "*'"
         Case 2
            strPhoneNo = "Phone like '*" & Trim$(UnformatPhone(txtPhoneNum.Text)) & "*'"
      End Select
   End If
   If LenTrim(txtDNIS.Text) Then
      strDNIS = "DNIS like '*" & txtDNIS.Text & "*'"
   End If
   RaiseEvent ApplyChanges(strTypes, strCustomerWhere, strDNIS, strAcctNo, strPhoneNo)
      
End Sub

Private Sub loadtypes()
   lstTypes.Clear
   lstTypes.AddItem "T"
   lstTypes.AddItem "V"
   lstTypes.AddItem "R"
   lstTypes.AddItem "RV"
   Dim i%
   For i% = 0 To lstTypes.ListCount - 1
      lstTypes.Selected(i%) = True
   Next i%
End Sub

Private Sub ResetFilters_Click()
   RaiseEvent ApplyChanges("", "", "", "", "")
   ResetParams
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtAcctNo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

'Private Sub txtPhoneNum_Change()
'   If Len(txtPhoneNum.Text) = 3 Then
'      txtPhoneNum.Text = txtPhoneNum.Text & "-"
'      txtPhoneNum.SelStart = Len(txtPhoneNum.Text)
'   End If
'   If Len(txtPhoneNum.Text) = 7 Then
'      txtPhoneNum.Text = txtPhoneNum.Text & "-"
'      txtPhoneNum.SelStart = Len(txtPhoneNum.Text)
'   End If
'
'End Sub

Private Sub txtPhoneNum_GotFocus()
   txtPhoneNum.Text = UnformatPhone(txtPhoneNum.Text)
   txtPhoneNum.SelStart = 0
   txtPhoneNum.SelLength = Len(txtPhoneNum.Text)
   txtPhoneNum.MaxLength = 10
End Sub

Private Sub txtPhoneNum_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 32 Then
      If Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtPhoneNum_LostFocus()
   txtPhoneNum.MaxLength = 0
   If cboPhone.ListIndex <> 2 Then
      txtPhoneNum.Text = FormatPhone(txtPhoneNum.Text)
   End If
End Sub


