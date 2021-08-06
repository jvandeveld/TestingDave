VERSION 5.00
Object = "{88146B05-ED42-4533-9459-1BD0254F54B4}#29.3#0"; "zatk_reportBuilder.ocx"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   6540
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin zatk_ReportBuilder.Report Report1 
      Height          =   6300
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   11113
      ThumbPaneWidth  =   1.29
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lngReportID As Long

Property Let ReportID(Value As Long)
   m_lngReportID = Value
End Property

Private Sub ResizeControls()
   On Error Resume Next
   Err.Clear
   Report1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Load()
   Report1.Clear
   Select Case m_lngReportID
      Case 1
         BuildVMReport
         
   End Select
   
   'AddHeadersAndFooters
   Report1.Render
End Sub

Private Sub Form_Resize()
   ResizeControls
End Sub

Private Sub BuildVMReport()
   Dim clsp As zatk_ReportBuilder.Page
   
   Stop
End Sub



