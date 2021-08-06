VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAssign 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Assign Number"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2415
      TabIndex        =   14
      Top             =   1965
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   635
      _Version        =   393216
      Format          =   126025729
      CurrentDate     =   42662
   End
   Begin VB.Frame Frame1 
      Caption         =   " Rate "
      Enabled         =   0   'False
      Height          =   1470
      Left            =   210
      TabIndex        =   8
      Top             =   1695
      Width           =   1860
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         TabIndex        =   13
         Top             =   1035
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Other"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   1110
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".20"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   855
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".10"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   570
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".00"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2970
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2670
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   4050
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2655
      Width           =   960
   End
   Begin VB.TextBox txtAccount 
      Height          =   330
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1170
      Width           =   1425
   End
   Begin VB.CommandButton cmdNameSearch 
      Height          =   435
      Left            =   1590
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmAssign.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1050
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Assigned Date"
      Height          =   195
      Left            =   2415
      TabIndex        =   15
      Top             =   1695
      Width           =   1410
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
      Left            =   2205
      TabIndex        =   7
      Top             =   1095
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Account:"
      Height          =   165
      Left            =   195
      TabIndex        =   2
      Top             =   915
      Width           =   1215
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
      Left            =   1815
      TabIndex        =   1
      Top             =   210
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number:"
      Height          =   195
      Left            =   465
      TabIndex        =   0
      Top             =   300
      Width           =   1110
   End
End
Attribute VB_Name = "frmAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private m_boolLoaded As Boolean
'Private m_boolAssign As Boolean
'Private m_boolAddNew As Boolean
'Private m_lngID As Long
'Private m_strPhoneNumber As String
'Private m_strLastUsed As String
'Private m_strAccount As String
'Private m_strName As String
'Private m_sngRate As Single
''Private WithEvents CustomerLookup As frmCustomerLookUp
'Event DoAssign(ID As Long, PhoneNumber As String, Account As String, Name As String, AssignedDate As String, Rate As Single)
'Event DoUpdate(ID As Long, PhoneNumber As String, Account As String, Name As String, AssignedDate As String, Rate As Single)
'
'Property Let ID(Value As Long)
'   m_lngID = Value
'End Property
'
'Property Let Assign(Value As Boolean)
'   m_boolAssign = Value
'End Property
'
'Property Let LastUsed(Value As String)
'   m_strLastUsed = Value
'End Property
'
'Property Let Account(Value As String)
'   m_strAccount = Value
'End Property
'
'Property Let CustomerName(Value As String)
'   m_strName = Value
'End Property
'
'Property Let Rate(Value As Single)
'   m_sngRate = Value
'End Property
'
'Property Let PhoneNumber(Value As String)
'   m_strPhoneNumber = Value
'End Property
'
'Private Sub cmdCancel_Click()
'   Unload Me
'End Sub
'
''Private Sub cmdNameSearch_Click()
''   Set CustomerLookup = New frmCustomerLookUp
''   CustomerLookup.Show vbModal
''End Sub
'
'Private Sub cmdOK_Click()
'   Dim sRate As Single
'   Dim strDate As String
'   strDate = ""
'   If optRate(0).Value = True Then
'      sRate = 0
'   ElseIf optRate(1).Value = True Then
'      sRate = 0.1
'   ElseIf optRate(2).Value = True Then
'      sRate = 0.2
'   Else
'      sRate = Val(txtRate.Text)
'   End If
'
'   If DTPicker1.Visible = True Then
'      strDate = Format$(DTPicker1.Value, "MM/DD/YYYY")
'      'strSQL = strSQL & ",DateLastUsed = #" & strLastUsed & "#"
'   End If
'
'   Dim boolT
'   On Error GoTo ET
'   Err.Clear
'   DBEngine.Workspaces(0).BeginTrans
'   Dim strSQL As String
'   Dim rs800 As DAO.Recordset
'   strSQL = "Select * from dat800Numbers where id = " & m_lngID
'   Set rs800 = g_dbPhone.OpenRecordset(strSQL, dbOpenDynaset)
'   If (rs800.EOF And rs800.BOF) Then
'      Beep
'      MsgBox "Unable to read header record!", vbCritical, "ERROR"
'      DBEngine.Workspaces(0).Rollback
'      GoTo ExitSub
'   Else
'      rs800.Edit
'   End If
'
'   If m_boolAddNew = True Then
'      'New assignment
'      rs800("AcctNum") = Trim$(txtAccount.Text)
'      rs800("Name") = Trim$(lblCustomer.Caption)
'      rs800("Rate") = sRate
'      rs800("DateAssigned") = strDate
'   ElseIf LenTrim(m_strAccount) Then
'      'Modifying Account
'      rs800("AcctNum") = Trim$(txtAccount.Text)
'      rs800("Name") = Trim$(lblCustomer.Caption)
'      rs800("Rate") = sRate
'      rs800("DateAssigned") = strDate
'   Else
'      rs800("AcctNum") = ""
'      rs800("Name") = ""
'      rs800("Rate") = 0
'      rs800("DateAssigned") = Null
'      rs800("DateLastUsed") = strDate
'   End If
'
'   rs800.Update
'
'
'   If Not m_boolAddNew Then
'        'Delete the Assigment record if it exists
'      strSQL = "Delete from datAssignments where id in " & vbCrLf
'      strSQL = strSQL & "(" & vbCrLf
'      strSQL = strSQL & "  Select top 1 id from datAssignments where PhoneNumber = '" & Trim$(Replace(m_strPhoneNumber, "-", "")) & "'" & vbCrLf
'      strSQL = strSQL & "   And AcctNum = '" & m_strAccount & "'" & vbCrLf
'      strSQL = strSQL & "   And (DateCanceled is null or DateCanceled = '')" & vbCrLf
'      strSQL = strSQL & ")" & vbCrLf
'      g_dbPhone.Execute strSQL, dbFailOnError
'   End If
'
'   strSQL = "Insert into datAssignments (PhoneNumber,DateAssigned,AcctNum,Rate) " & vbCrLf
'   strSQL = strSQL & "Values ('" & Trim$(Replace(m_strPhoneNumber, "-", "")) & "'," & vbCrLf
'   strSQL = strSQL & "#" & strDate & "#," & vbCrLf
'   strSQL = strSQL & "'" & Trim$(txtAccount.Text) & "'," & vbCrLf
'   strSQL = strSQL & sRate & ")"
'   g_dbPhone.Execute strSQL, dbFailOnError
'
'   DBEngine.Workspaces(0).CommitTrans
'
'   If m_boolAddNew = True Then
'      RaiseEvent DoAssign(m_lngID, m_strPhoneNumber, txtAccount.Text, Replace(lblCustomer.Caption, "&&", "&"), strDate, sRate)
'   Else
'      RaiseEvent DoUpdate(m_lngID, m_strPhoneNumber, txtAccount.Text, Replace(lblCustomer.Caption, "&&", "&"), strDate, sRate)
'   End If
'
'
'ExitSub:
'Exit Sub
'
'ET:
'   Screen.MousePointer = vbNormal
'   'clsProgress.EndProgress
'   Beep
'   MsgBox "Error durring processing." & vbCrLf & "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, ""
'
'   DBEngine.Workspaces(0).Rollback
'   Resume ExitSub
'End Sub
'
'Private Sub CustomerLookup_CustomerSelected(AcctNo As String, Name As String)
'   txtAccount.Text = AcctNo
'   lblCustomer.Caption = Replace(Name, "&", "&&")
'   Unload CustomerLookup
'   Dim i%
'   For i% = 0 To optRate.Count - 1
'      optRate(i%).Enabled = True
'   Next i%
'   Frame1.Enabled = True
'   optRate(1).SetFocus
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      KeyCode = 0
'      cmdNameSearch.Value = True
'   End If
'End Sub
'
'Private Sub Form_Load()
'   m_boolLoaded = False
'   lblCustomer.Caption = ""
'   lblPhone.Caption = m_strPhoneNumber
'   LoadRecord
'   If LenTrim(m_strAccount) Then
'      m_boolAddNew = False
'      'Label3.Visible = False
'      'DTPicker1.Visible = False
'      Frame1.Enabled = True
'      Dim i%
'      For i% = 0 To 3
'         optRate(i%).Enabled = True
'      Next i%
'
'      txtAccount.Text = m_strAccount
'      lblCustomer.Caption = Replace(m_strName, "&", "&&")
'      Select Case m_sngRate
'         Case 0
'            optRate(0).Value = True
'         Case 0.1
'            optRate(1).Value = True
'         Case 0.2
'            optRate(2).Value = True
'         Case Else
'            txtRate.Text = Format$(m_sngRate, "0.##")
'            txtRate.Visible = True
'            optRate(3).Value = True
'      End Select
'      'txtRate.Text = m_sngRate
'      'txtAccount.Enabled = False
'   Else
'      If m_boolAssign = True Then
'         m_boolAddNew = True
'         lblCustomer.Caption = ""
'         'Label3.Visible = False
'         DTPicker1.Visible = True ' False
'         DTPicker1.Value = Now
'      Else
'         Label3.Top = Label2.Top
'         Label3.Left = Label2.Left
'         Label3.Caption = "Date last used:"
'         DTPicker1.Top = txtAccount.Top
'         DTPicker1.Left = txtAccount.Left
'         Label2.Visible = False
'         txtAccount.Visible = False
'         cmdNameSearch.Visible = False
'         lblCustomer.Visible = False
'         If IsDate(m_strLastUsed) Then
'            DTPicker1.Value = CDate(m_strLastUsed)
'         Else
'            DTPicker1.Value = "12:00 AM"
'         End If
'      End If
'   End If
'   m_boolLoaded = True
'End Sub
'
'Private Sub optRate_Click(Index As Integer)
'   If m_boolLoaded = True Then
'      If optRate(3).Value = True Then
'         txtRate.Visible = True
'         txtRate.SetFocus
'      Else
'         txtRate.Visible = False
'      End If
'   End If
'End Sub
'
'Private Sub LoadRecord()
'   Dim rsTemp As DAO.Recordset
'   Dim strSQL As String
'   strSQL = "Select * from dat800Numbers where id = " & m_lngID
'   Set rsTemp = g_dbPhone.OpenRecordset(strSQL, dbOpenForwardOnly)
'   If rsTemp.EOF Then
'      Beep
'      MsgBox "Unable to read header record!", vbCritical, "ERROR"
'   Else
'      txtAccount.Text = FieldValue(rsTemp, "AcctNum")
'      If IsDate(FieldValue(rsTemp, "DateAssigned")) Then
'         DTPicker1.Value = FieldValue(rsTemp, "DateAssigned")
'      End If
'
'   End If
'
'End Sub
Private Sub lblCustomer_Click()

End Sub
