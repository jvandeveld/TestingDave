VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAssignRecord 
   Caption         =   "Assign Client Info to Phone Number"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Rate "
      Height          =   1470
      Left            =   3720
      TabIndex        =   14
      Top             =   1335
      Width           =   1860
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         TabIndex        =   15
         Top             =   1035
         Width           =   645
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".00"
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".10"
         Height          =   270
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   570
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   ".20"
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   855
         Width           =   1035
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Other"
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   480
      Left            =   345
      TabIndex        =   5
      Top             =   3390
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      _Version        =   393216
      Format          =   125042689
      CurrentDate     =   44375
      MaxDate         =   401404
      MinDate         =   18264
   End
   Begin VB.TextBox txtAddCustName 
      Height          =   450
      Left            =   285
      TabIndex        =   3
      Top             =   2325
      Width           =   3105
   End
   Begin VB.TextBox txtAddAccNum 
      Height          =   450
      Left            =   270
      TabIndex        =   1
      Top             =   1425
      Width           =   3090
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   720
      Left            =   4485
      TabIndex        =   11
      Top             =   4035
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   720
      Left            =   3180
      TabIndex        =   10
      Top             =   4005
      Width           =   990
   End
   Begin VB.Label lblDateAssign 
      Caption         =   "Date &Assigned:"
      Height          =   345
      Left            =   315
      TabIndex        =   4
      Top             =   2955
      Width           =   1800
   End
   Begin VB.Label lblCustName 
      Caption         =   "Customer Name:"
      Height          =   345
      Left            =   300
      TabIndex        =   2
      Top             =   1935
      Width           =   1800
   End
   Begin VB.Label lblAccNum 
      Caption         =   "Account Number:"
      Height          =   345
      Left            =   300
      TabIndex        =   0
      Top             =   1020
      Width           =   1800
   End
   Begin VB.Label lblLastUsed 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3030
      TabIndex        =   13
      Top             =   330
      Width           =   1890
   End
   Begin VB.Label lblNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   330
      Width           =   2130
   End
End
Attribute VB_Name = "frmAssignRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngRecordID As Long
Private m_boolEditMode As Boolean
'Private m_database As DAO.Database
Event RecordUpdated(AcctNo As String, Customer As String, DateAssigned As String, OverMinRate As String, TV As String)

Property Let RecordID(value As Long)
   m_lngRecordID = value
End Property

Property Let editMode(value As Boolean)
   m_boolEditMode = value
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim dbData As DAO.Database
   Set dbData = OpenLocalDB
   Dim rsDAO As DAO.Recordset
   Dim rsChk As DAO.Recordset
   Dim strSQL As String

   Dim sRate As Single
   If optRate(0).value = True Then
      sRate = 0
   ElseIf optRate(1).value = True Then
      sRate = 0.1
   ElseIf optRate(2).value = True Then
      sRate = 0.2
   Else
      sRate = Val(txtRate.Text)
   End If
   
   Dim boolContinue As Boolean
   boolContinue = True ' Assume we are updating the database
   
   ' Sanity Checks
   strSQL = "Select count(*) as RecordCount from datLotus800 where AcctNo = '" & Trim$(txtAddAccNum.Text) & "'"
   Set rsChk = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
   If Err.Number <> 0 Then
      Beep
      MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
   Else
      If Val(FieldValue(rsChk, "RecordCount")) > 1 Then
         Dim strMsg As String
         strMsg = "Updating this record will update all " & FieldValue(rsChk, "RecordCount") & " Phone #'s assigned to this account." & vbCrLf & "Do you wish to proceed?"
         Beep
         Select Case MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Proceed")
            Case vbYes
               boolContinue = True
            Case Else
               boolContinue = False
         End Select
      End If
      
      '===============================================
      'TODO: Complete logic above to ensure all phone
      '      numbers are assigned the same rate
      boolContinue = True
      '===============================================
      
      If boolContinue Then
         strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
         Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
         rsDAO.Edit
         
         If Not m_boolEditMode Then
            Select Case UCase$(rsDAO.Fields("T_V"))
               Case "RV"
                  rsDAO.Fields("T_V") = "V"
               Case Else
                  rsDAO.Fields("T_V") = "T"
            End Select
         End If
         With rsDAO
             .Fields("AcctNo") = Trim$(txtAddAccNum.Text)
             .Fields("Customer") = Trim$(txtAddCustName.Text)
             .Fields("Assigned") = Format(DTPicker1.value, "mm/dd/yyyy")
             .Fields("Rate") = sRate
             .Fields("LastUsed") = 0
          End With
          Dim strTV As String
          strTV = FieldValue(rsDAO, "T_V")
          rsDAO.Update
          dbData.Close
          Set dbData = Nothing
          Beep
          If m_boolEditMode = True Then
             MsgBox "Number information updated.", vbInformation, " "
          Else
             MsgBox "Customer and Phone# Linked.", vbInformation, " "
          End If
          RaiseEvent RecordUpdated(txtAddAccNum.Text, txtAddCustName.Text, Format(DTPicker1.value, "mm/dd/yyyy"), CStr(sRate), strTV)
          
          Unload Me
      
      End If
'      Do Until rsChk = EOF
'         'If rsChk.Fields("AcctNo") = Trim$(txtAddAccNum.Text) Then
'            If rsChk.Fields("Rate") <> Str(sRate) Then
'
'            'End If
'         End If
'         rsChk.MoveNext
'      Loop
   End If
   
   Stop
   
 
End Sub


Private Sub Form_Initialize()
   m_boolEditMode = False
End Sub

Private Sub Form_Load()
   Dim dbData As DAO.Database
   Set dbData = OpenLocalDB
   Dim strSQL As String
   Dim rsDAO As DAO.Recordset
'   DTPicker1.MaxDate = DateAdd("m", 1, Now)
'
'   DTPicker1.MinDate = DateAdd("m", -2, Now)
   strSQL = "Select * from datLotus800 where id = " & m_lngRecordID
   Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
   If rsDAO.EOF And rsDAO.BOF Then
      'no record found - should never happen
      
   Else
      If m_boolEditMode = True Then
         'Load Controls from database
         txtAddAccNum.Text = FieldValue(rsDAO, "AcctNo")
         txtAddCustName.Text = FieldValue(rsDAO, "Customer")
         If IsDate(FieldValue(rsDAO, "Assigned")) Then
            DTPicker1.value = CDate(FieldValue(rsDAO, "Assigned"))
         End If
         Dim sRate As Single
         sRate = FieldValue(rsDAO, "Rate")
         If sRate = 0 Then
            optRate(0).value = True
         ElseIf sRate = 0.1 Then
            optRate(1).value = True
         ElseIf sRate = 0.2 Then
            optRate(2).value = True
         Else
            txtRate.Text = sRate
         End If
      End If
   End If
   lblNumber.Caption = FormatPhone(FieldValue(rsDAO, "Phone"))
   lblLastUsed.Caption = FieldValue(rsDAO, "LastUsed")
   dbData.Close
   Set dbData = Nothing
End Sub

Private Sub txtAddAccNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      ' Tab ??
   ElseIf Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub
Private Sub txtAddCustName_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      ' Tab ??
   Else
      KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
   End If
End Sub

Private Sub optRate_GotFocus(Index As Integer)
   Frame1.FontBold = True
End Sub

Private Sub optRate_LostFocus(Index As Integer)
   Frame1.FontBold = False
End Sub

Private Sub txtAddAccNum_GotFocus()
   txtAddAccNum.SelStart = 0
   txtAddAccNum.SelLength = Len(txtAddAccNum.Text)
   lblAccNum.FontBold = True
End Sub

Private Sub txtAddAccNum_LostFocus()
   lblAccNum.FontBold = False
End Sub

Private Sub txtAddCustName_GotFocus()
   txtAddCustName.SelStart = 0
   txtAddCustName.SelLength = Len(txtAddAccNum.Text)
   lblCustName.FontBold = True
End Sub

Private Sub txtAddCustName_LostFocus()
   lblCustName.FontBold = False
End Sub

Private Sub DTPicker1_GotFocus()
   lblDateAssign.FontBold = True
End Sub

Private Sub DTPicker1_LostFocus()
   lblDateAssign.FontBold = False
End Sub
