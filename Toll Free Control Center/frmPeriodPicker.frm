VERSION 5.00
Begin VB.Form frmPeriodPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Billing Period"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRegenerateRecords 
      Caption         =   "Regenerate Records"
      Height          =   435
      Left            =   2565
      TabIndex        =   3
      Top             =   450
      Width           =   1200
   End
   Begin VB.ListBox lstPeriod 
      Height          =   2385
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   2
      Top             =   225
      Width           =   2190
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   2670
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2085
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2670
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1530
      Width           =   960
   End
End
Attribute VB_Name = "frmPeriodPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event PeriodSelected(Period As String, regeneratePeriod As Boolean)

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   SelectPeriod
End Sub

Private Sub Form_Load()
   LoadPeriods
End Sub


Private Sub LoadPeriods()
   lstPeriod.Clear
   Dim dtStart As Date
   dtStart = DateAdd("m", 1, Now)
   Do Until dtStart < #7/1/2021#   ' < 2019
      Dim dtDate As Date
      dtDate = DateSerial(Year(dtStart), Month(dtStart), 1)
      lstPeriod.AddItem Format$(dtDate, "MM-DD-YYYY") & " Invoices"  'Format$(Year(dtStart), "####") & Format$(Month(dtStart), "00")
      dtStart = DateAdd("m", -1, dtStart)
      
   Loop
   lstPeriod.ListIndex = 0
End Sub

Private Sub lstPeriod_Click()
   If lstPeriod.ListIndex <> 0 Then
      chkRegenerateRecords.value = 0
      chkRegenerateRecords.Enabled = False
   Else
      chkRegenerateRecords.Enabled = True
   End If
End Sub

Private Sub lstPeriod_DblClick()
   SelectPeriod
End Sub


Private Sub SelectPeriod()
   If chkRegenerateRecords.value Then
      Beep
      If MsgBox("Are you sure you would like to regenerate this period's grid?", vbQuestion + vbYesNo + vbDefaultButton2, " ") = vbYes Then
         RaiseEvent PeriodSelected(lstPeriod.List(lstPeriod.ListIndex), True)
      End If
   Else
      RaiseEvent PeriodSelected(lstPeriod.List(lstPeriod.ListIndex), False)
   End If
End Sub
