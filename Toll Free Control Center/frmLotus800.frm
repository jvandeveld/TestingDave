VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{95559FD0-8A4C-11D3-905E-00A04B0669E7}#1.1#0"; "SmartUI.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLotus800 
   Caption         =   "Lotus 800 Grid"
   ClientHeight    =   8610
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13320
   Icon            =   "frmLotus800.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1575
      Top             =   7050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "CSV"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Usage Billing"
      Height          =   645
      Left            =   8325
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   2025
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4920
      Left            =   600
      TabIndex        =   0
      Top             =   1530
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   8678
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Phone #"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DNIS"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "T/V"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Account #"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Name"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Last Used"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Date Assigned"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Over Rate"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "#.#0"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   979
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2937"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2858"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131601"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2249"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2170"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131601"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1323"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1244"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131601"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2487"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2408"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131601"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=6350"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=6271"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131600"
      Splits(0)._ColumnProps(25)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=131601"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2302"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=131601"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1561"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1482"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=131602"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.alignment=2,.valignment=2"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(9)   =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1,.locked=0"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0,.wraptext=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin BoWSmartUI.SmartUI smartPopUp 
      Height          =   360
      Left            =   9570
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   635
      Template        =   "frmLotus800.frx":3A0A
   End
   Begin BoWSmartUI.SmartUI smartStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8250
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   635
      Template        =   "frmLotus800.frx":73F6
   End
   Begin BoWSmartUI.SmartUI smartToolBar 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   1588
      Template        =   "frmLotus800.frx":7C0E
   End
   Begin VB.Menu mnuGridPopup 
      Caption         =   "Grid Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAssignTASNumber 
         Caption         =   "Assign TAS Number"
      End
      Begin VB.Menu mnuAssignVMNumber 
         Caption         =   "Assign Voicemail Number"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReleaseNumber 
         Caption         =   "Release Number"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu mnuEditLastUsed 
         Caption         =   "Edit Last Used"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmLotus800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_xRows As XArrayDB
Private WithEvents PeriodPicker As frmPeriodPicker
Attribute PeriodPicker.VB_VarHelpID = -1
Private WithEvents FindDialogue As frmFind
Attribute FindDialogue.VB_VarHelpID = -1
Private WithEvents AssignNumber As frmAssignRecord
Attribute AssignNumber.VB_VarHelpID = -1
Private WithEvents ChangeUnused As frmChangeUnused
Attribute ChangeUnused.VB_VarHelpID = -1
Private WithEvents UnAssignNumberForm As frmUnAssign
Attribute UnAssignNumberForm.VB_VarHelpID = -1
Private WithEvents FilterForm As frmFilter
Attribute FilterForm.VB_VarHelpID = -1

Private m_lngRow As Long
Private m_lngCol As Long
Private m_totalRows As Long
Private m_boolAsc As Boolean
Private boolAsc(20) As Boolean


Private Sub ResizeControls()
   On Error Resume Next
   Err.Clear
   If Me.WindowState <> vbMinimized Then
'      If Me.Width < 10380 Then
'         Me.Width = 10380
'      End If
'      If Me.Height < 5000 Then
'         Me.Height = 5000
'      End If
               
      On Error Resume Next
        
      With TDBGrid1
         .Top = smartToolBar.Top + smartToolBar.Height
         .Left = 0
         .Width = Me.ScaleWidth
         .Height = smartStatus.Top - .Top
      End With
   End If
      
End Sub

Private Sub AssignNumber_RecordUpdated(AcctNo As String, Customer As String, DateAssigned As String, OverMinRate As String, TV As String)
   Dim lngRow As Long
   ' Updating Records
   lngRow = TDBGrid1.Bookmark
   m_xRows.value(lngRow, 3) = AcctNo
   m_xRows.value(lngRow, 2) = TV
   m_xRows.value(lngRow, 4) = Customer
   m_xRows.value(lngRow, 5) = ""
   m_xRows.value(lngRow, 6) = DateAssigned
   m_xRows.value(lngRow, 7) = OverMinRate
   TDBGrid1.RefetchRow
End Sub

Private Sub ChangeUnused_RecordUpdated(LastUsed As String)
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   m_xRows.value(lngRow, 5) = LastUsed
   TDBGrid1.RefetchRow
End Sub


Private Sub Command1_Click()
   On Error Resume Next
   Err.Clear
   If True Then
      'clear the grid
      'MsgBox GetIniSetting(dz_MDBPath)
      'LoadGrid
      'ShowProgressWindow "Loading Grid..."
      'Dim frmForm As New frmBillingUsage
      'frmForm.Show vbModal, Me
      Set PeriodPicker = New frmPeriodPicker
      PeriodPicker.Show vbModal, Me
   Else
      If False Then
         Dim intFile As Integer
         intFile = FreeFile
         Dim strFile As String
         strFile = "d:\temp\texxt.txt"
         Open strFile For Input As #intFile
         If Err.Number <> 0 Then
            Beep
            MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
         Else
            'eveyting is ok
            
         End If
         
         Close #intFile
      Else
         On Error Resume Next
         Err.Clear
         With CommonDialog1
            .CancelError = True
            .DialogTitle = "Save CSV File"
            .DefaultExt = "csv"
            
            
            
            .ShowSave
            If Err.Number = 0 Then
               MsgBox CommonDialog1.FileName
            End If
            
         End With
      End If
   End If
End Sub

Private Sub Form_Initialize()
   If App.PrevInstance Then
      Beep
      MsgBox "Application is already running!", vbInformation, "Already Running"
      End
   Else
      Load frmSplash
      frmSplash.ShowSplash
      
      Do
         DoEvents
      Loop Until frmSplash.OpaqueValue >= 255
      
   End If
End Sub

Private Sub Form_Load()
   On Error Resume Next
   Err.Clear
   Dim hWnd As Long
   hWnd = Me.hWnd
   Call SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW)
         
   Set g_clsProgress = CreateObject("OOPProgress.ProgressBar")
   
   Command1.Visible = DesignMode  'DesignMode
   LoadGrid
   TDBGrid1.Columns(5).FetchStyle = dbgFetchCellStyleColumn
   TDBGrid1.Columns(6).FetchStyle = dbgFetchCellStyleColumn
   With TDBGrid1
      .ColumnFooters = False
      .AllowColSelect = False
      .AlternatingRowStyle = True
      .EvenRowStyle.BackColor = RGB(200, 235, 255)
      .RecordSelectors = False
      .MarqueeStyle = dbgHighlightRowRaiseCell
      '.SetFocus
   End With
   
   m_totalRows = m_xRows.UpperBound(1) + 1
   UpdateStatusBar
   frmSplash.HideSplash
End Sub

Private Sub Form_Resize()
   ResizeControls
End Sub

Private Sub LoadGrid(Optional WhereClause As String = "")
   Dim dbData As DAO.Database
   Set dbData = OpenLocalDB
   If dbData Is Nothing Then
      If IsLoaded("frmSplash") Then
         Unload frmSplash
      End If
      Beep
      MsgBox g_strErrorMessage, vbInformation, "ERROR"
   Else
      Dim iRows As Integer
      iRows = -1
      
      Set m_xRows = New XArrayDB
      
      m_xRows.ReDim 0, 10000, 0, 20
      
      On Error Resume Next
      Err.Clear
      
      Dim strSQL As String
      strSQL = "Select * from datLotus800 " '
      If LenTrim(WhereClause) Then
         strSQL = strSQL & WhereClause
      End If
      strSQL = strSQL & " Order by Phone"
      Dim rsDAO As DAO.Recordset
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
      If Err.Number <> 0 Then
         Beep
         MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
      Else
         Do Until rsDAO.EOF
            iRows = iRows + 1
            'Debug.Print "Row: " & iRows
            m_xRows.value(iRows, 0) = FormatPhone(FieldValue(rsDAO, "Phone"))
            m_xRows.value(iRows, 1) = FieldValue(rsDAO, "DNIS")
            m_xRows.value(iRows, 2) = FieldValue(rsDAO, "T_V")
            m_xRows.value(iRows, 3) = FieldValue(rsDAO, "AcctNo")
            m_xRows.value(iRows, 4) = FieldValue(rsDAO, "Customer")
            If Not IsDate(FieldValue(rsDAO, "LastUsed")) Then
               m_xRows.value(iRows, 5) = CDate("12/31/1899")
            Else
               m_xRows.value(iRows, 5) = FieldValue(rsDAO, "LastUsed")
            End If
            If Not IsDate(FieldValue(rsDAO, "Assigned")) Then
               m_xRows.value(iRows, 6) = CDate("12/31/1899")
            Else
               m_xRows.value(iRows, 6) = FieldValue(rsDAO, "Assigned")
            End If
            m_xRows.value(iRows, 7) = FieldValue(rsDAO, "Rate")
            m_xRows.value(iRows, 20) = FieldValue(rsDAO, "id")
            If Err.Number <> 0 Then
               Stop
            End If
            'Stop
            rsDAO.MoveNext
         Loop
         'Stop
      End If
         
      dbData.Close
      Set dbData = Nothing
         
      m_xRows.ReDim 0, iRows, 0, 20
      Set TDBGrid1.Array = m_xRows
      TDBGrid1.ReBind
      TDBGrid1.Bookmark = 0
      
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Err.Clear
   EndProgress
   Set g_clsProgress = Nothing
End Sub

Private Sub EditLast()
   Set ChangeUnused = New frmChangeUnused
   ChangeUnused.RecordID = CurrentRecordID
   ChangeUnused.Show vbModal, Me
End Sub

Private Sub EditRecord()
   Set AssignNumber = New frmAssignRecord
   AssignNumber.RecordID = CurrentRecordID
   AssignNumber.editMode = True
   AssignNumber.Show vbModal, Me
End Sub

Private Sub mnuFind_Click()
   ShowFindForm
End Sub

Private Sub mnuHelpAbout_Click()
   Dim frmForm As New frmAbout
   frmForm.Show vbModal, Me
End Sub

Private Sub PeriodPicker_PeriodSelected(Period As String, regeneratePeriod As Boolean)
   Unload PeriodPicker
   Dim frmForm As New frmBillingUsage
   frmForm.BillingPeriod = Mid(Period, 7, 4) & Mid(Period, 1, 2)
   frmForm.Regen = regeneratePeriod
   frmForm.Show vbModal, Me
End Sub

Private Sub smartPopUp_Click(Item As BoWSmartUI.SmartItem)
   Select Case Trim$(UCase$(Item.Key))
      Case "ASSIGN"
         Select Case Trim$(UCase$(m_xRows.value(TDBGrid1.Bookmark, 2)))
            Case "R"
               AssignTASNum
            Case "RV"
               AssignVMNum
         End Select
      Case "RELEASE"
         ReleaseNum
      Case "EDITALL"
         EditRecord
      Case "EDITLAST"
         EditLast
      Case "DEACTIVATE"
         DeactivateNum
      Case "ACTIVATE"
         ActivateNum
   End Select
End Sub

Private Sub smartToolBar_Click(Item As BoWSmartUI.SmartItem)
   Select Case UCase$(Trim$(Item.Key))
      Case "PRINT"
         PrintGrid TDBGrid1, "PHONE NUMBER INFO"
      Case "FIND"
         ShowFindForm
      Case "FILTER", "FILTERALL"
         ShowFilter
      Case "EXPORT"
         ExportToCSV
      Case "CLEARFILTER"
         ClearFilter
      Case "REPORTVM"
         BuildReport 1
      Case "USAGE"
         Set PeriodPicker = New frmPeriodPicker
         PeriodPicker.Show vbModal, Me
   End Select
End Sub

Private Sub BuildReport(ID As Long)
   Dim frmForm As New frmReport
   frmForm.ReportID = ID
   frmForm.Show vbModal, Me
End Sub

Private Sub ActivateNum()
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   Dim strTV As String
   Dim strnewTV As String
   strTV = UCase$(m_xRows.value(lngRow, 2))
   If strTV = "O" Then
      strnewTV = "R"
   ElseIf strTV = "OV" Then
      strnewTV = "RV"
   End If
   m_xRows.value(lngRow, 2) = strnewTV
   m_xRows.value(lngRow, 3) = ""
   m_xRows.value(lngRow, 4) = ""
   m_xRows.value(lngRow, 5) = Format(Now, "mm/dd/yyyy")
   m_xRows.value(lngRow, 6) = 0  ' Null
   m_xRows.value(lngRow, 7) = Chr$(0) ' Null
   TDBGrid1.RefetchRow
   
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim strSQL As String
   
   Set dbData = OpenLocalDB
   strSQL = "Select * from datLotus800 where id = " & CurrentRecordID
   Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
   rsDAO.Edit
   rsDAO.Fields("T_V") = strnewTV
   rsDAO.Fields("Term") = ""
   rsDAO.Fields("Customer") = ""
   rsDAO.Fields("AcctNo") = ""
   rsDAO.Fields("LastUsed") = Format(Now, "mm/dd/yyyy")
   rsDAO.Fields("Rate") = 0
   rsDAO.Update
   
End Sub

Private Sub DeactivateNum()
   Beep
   If MsgBox("Are you sure you would like to deactivate this number?", vbQuestion + vbYesNo + vbDefaultButton2, " ") = vbYes Then
      Dim strTV As String
      Dim lngRow As Long
      lngRow = TDBGrid1.Bookmark
      strTV = UCase$(m_xRows.value(lngRow, 2))
      Dim strnewTV As String
      If strTV = "T" Or strTV = "R" Then
         strnewTV = "O"
      Else
         strnewTV = "OV"
      End If
      Dim dbData As DAO.Database
      Dim rsDAO As DAO.Recordset
      Dim strSQL As String
      
      Set dbData = OpenLocalDB
      strSQL = "Select * from datLotus800 where id = " & CurrentRecordID
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
      If rsDAO.EOF And rsDAO.BOF Then
         'very bad - no record
      
      Else
         rsDAO.Edit
         rsDAO.Fields("T_V") = strnewTV
         rsDAO.Fields("Term") = ""
         rsDAO.Fields("Customer") = ""
         rsDAO.Fields("AcctNo") = ""
         rsDAO.Fields("LastUsed") = Format(Now, "mm/dd/yyyy")
         rsDAO.Fields("Rate") = 0
         rsDAO.Update
         m_xRows.value(lngRow, 2) = strnewTV
         m_xRows.value(lngRow, 3) = ""
         m_xRows.value(lngRow, 4) = ""
         m_xRows.value(lngRow, 5) = Format(Now, "mm/dd/yyyy")
         m_xRows.value(lngRow, 6) = 0  ' Null
         m_xRows.value(lngRow, 7) = Chr$(0) ' Null
         TDBGrid1.RefetchRow
      End If
   End If
End Sub

Private Sub ExportToCSV(Optional SelectedRows As Boolean = False)
   On Error Resume Next
   Err.Clear
   With CommonDialog1
      .CancelError = True
      .DialogTitle = "Save CSV File"
      .DefaultExt = "CSV"
      .ShowSave
      If Err.Number = 0 Then
         TDBGrid1.ExportToDelimitedFile .FileName, , ",", Chr$(34), Chr$(34), True
         MsgBox CommonDialog1.FileName
      End If
   End With
End Sub

Private Sub ShowFindForm()
   If Not IsLoaded("frmFind") Then
      Set FindDialogue = New frmFind
      FindDialogue.Show , Me
   Else
      FindDialogue.Show , Me
   End If
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid80.StyleDisp)
   Select Case Col
      Case 5, 6 'Last Used Date
         If IsDate(m_xRows.value(Bookmark, Col)) Then
            If Year(m_xRows.value(Bookmark, Col)) < 1900 Then
               CellStyle.ForegroundPicturePosition = dbgFPPictureOnly
            End If
         Else
            CellStyle.ForegroundPicturePosition = dbgFPPictureOnly
         End If
   End Select
End Sub

Private Sub TDBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_lngRow = -1
   m_lngCol = -1
   On Error Resume Next
   Err.Clear
   m_lngCol = TDBGrid1.ColContaining(X)
   m_lngRow = TDBGrid1.RowContaining(Y)

   If Button = 2 Then
      If m_lngRow >= 0 Then
         If m_lngCol >= 0 Then
            TDBGrid1.Row = m_lngRow
         End If
      End If
   End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If m_lngRow >= 0 Then
         If m_lngCol >= 0 Then
            'Set enabled properties of menu items
'            mnuEditRecord.Visible = True
'            mnuEditLastUsed.Visible = True
'            mnuReleaseNumber.Visible = True
'            mnuAssignVMNumber.Visible = True
'            mnuAssignTASNumber.Visible = True
'            mnuEditRecord.Enabled = True
'            mnuEditLastUsed.Enabled = True
'            mnuReleaseNumber.Enabled = True
'            mnuAssignVMNumber.Enabled = True
'            mnuAssignTASNumber.Enabled = True
            smartPopUp.SmartItems("ASSIGN").Visibility = viVisible
            smartPopUp.SmartItems("RELEASE").Visibility = viVisible
            smartPopUp.SmartItems("editall").Visibility = viVisible
            smartPopUp.SmartItems("editlast").Visibility = viVisible
            smartPopUp.SmartItems("DEACTIVATE").Visibility = viVisible
            smartPopUp.SmartItems("ACTIVATE").Visibility = viVisible
            smartPopUp.SmartItems("ACTIVATE").Enabled = True
            smartPopUp.SmartItems("editlast").Enabled = True
            smartPopUp.SmartItems("release").Enabled = True
            smartPopUp.SmartItems("assign").Enabled = True
            smartPopUp.SmartItems("editall").Enabled = True
            smartPopUp.SmartItems("DEACTIVATE").Enabled = True
            Select Case UCase$(m_xRows.value(TDBGrid1.Bookmark, 2))
               Case "R", "RV"
                  smartPopUp.SmartItems("editall").Visibility = viHide
                  smartPopUp.SmartItems("release").Enabled = False
                  smartPopUp.SmartItems("ACTIVATE").Visibility = False
               Case "T", "V"
                  smartPopUp.SmartItems("editlast").Visibility = viHide
                  smartPopUp.SmartItems("assign").Enabled = False
                  smartPopUp.SmartItems("ACTIVATE").Visibility = False
               Case "O"
                  smartPopUp.SmartItems("assign").Visibility = False
                  smartPopUp.SmartItems("release").Visibility = False
                  smartPopUp.SmartItems("editall").Visibility = False
                  smartPopUp.SmartItems("DEACTIVATE").Visibility = False
                  smartPopUp.SmartItems("editLast").Visibility = True
                  smartPopUp.SmartItems("ACTIVATE").Visibility = True
               Case Else
                  Stop
            End Select
            smartPopUp("POPUP").PopupMenu TDBGrid1.Left + X, TDBGrid1.Top + Y
            'Me.PopupMenu mnuGridPopup  'Show the popup menu
         End If
      End If
   End If
     
   
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
   On Error Resume Next
   Err.Clear

   ShowProgressWindow "Sorting..."
   Dim intColStart As Integer
   intColStart = TDBGrid1.LeftCol
   Screen.MousePointer = vbHourglass
   boolAsc(ColIndex) = Not boolAsc(ColIndex)
   Select Case ColIndex
      Case 0, 1, 2, 3, 4 'string
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_STRING
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_STRING
         End If
      Case 5, 6 'Date
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_DATE
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_DATE
         End If
      Case 7 'fp
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_DOUBLE
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_DOUBLE
         End If
   End Select
   TDBGrid1.ReBind
   TDBGrid1.Bookmark = 0
   TDBGrid1.LeftCol = intColStart
   EndProgress
   Screen.MousePointer = vbNormal

End Sub

Private Sub FindDialogue_FindText(IsSubstring As Boolean, Text As String)
   FindDialogueFindText TDBGrid1, IsSubstring, Text
End Sub

Private Sub FindDialogue_FindNext(IsSubstring As Boolean, Text As String)
   FindDialogueFindNext TDBGrid1, IsSubstring, Text
End Sub

Private Sub AssignTASNum()
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   Dim strTV As String
   strTV = UCase$(m_xRows.value(lngRow, 2))
   If strTV <> "R" Then
      'Invalid number for TAS
      Beep
      MsgBox "Invalid number", vbInformation, " "
   Else
      'Code here to assign tas number
      'Dim iRow As Integer
      'iRow = TDBGrid1.Bookmark
      Dim lngRecordID As Long
      lngRecordID = m_xRows.value(lngRow, 20)
      'Dim frmForm As New frmAddRecordDAO
      Set AssignNumber = New frmAssignRecord
      AssignNumber.RecordID = lngRecordID
      AssignNumber.Show vbModal, Me
   End If
End Sub

Private Sub AssignVMNum()
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   Dim strTV As String
   strTV = UCase$(m_xRows.value(lngRow, 2))
   If strTV <> "RV" Then
      'Invalid number for VM
      Beep
      MsgBox "Invalid number", vbInformation, " "
   Else
      'Code here to assign Voice Mail number
      Dim iRow As Integer
      iRow = TDBGrid1.Bookmark
      Dim lngRecordID As Long
      lngRecordID = m_xRows.value(iRow, 20)
      'Dim frmForm As New frmAddRecordDAO
      Set AssignNumber = New frmAssignRecord
      AssignNumber.RecordID = lngRecordID
      AssignNumber.Show vbModal, Me
   End If
End Sub

Private Sub ReleaseNum()
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark

   If True Then
      Set UnAssignNumberForm = New frmUnAssign
      UnAssignNumberForm.PhoneNumber = m_xRows.value(lngRow, 0)
      UnAssignNumberForm.Account = m_xRows.value(lngRow, 3)
      UnAssignNumberForm.CustomerName = m_xRows.value(lngRow, 4)
      UnAssignNumberForm.Show vbModal, Me
   Else
      Dim strTV As String
      strTV = UCase$(m_xRows.value(lngRow, 2))
      Select Case strTV
         Case "T", "V"
            Beep
            If MsgBox("Are you sure you want to release this number?", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
            
               Dim strnewTV As String
               If strTV = "T" Then
                  strnewTV = "R"
               Else
                  strnewTV = "RV"
               End If
               Dim dbData As DAO.Database
               Dim rsDAO As DAO.Recordset
               Dim strSQL As String
               
               Set dbData = OpenLocalDB
               strSQL = "Select * from datLotus800 where id = " & CurrentRecordID
               Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
               If rsDAO.EOF And rsDAO.BOF Then
                  'very bad - no record
               
               Else
                  rsDAO.Edit
                  rsDAO.Fields("T_V") = strnewTV
                  rsDAO.Fields("Term") = ""
                  rsDAO.Fields("Customer") = ""
                  rsDAO.Fields("AcctNo") = ""
                  rsDAO.Fields("LastUsed") = Format(Now, "mm/dd/yyyy")
                  rsDAO.Fields("Rate") = 0
                  rsDAO.Update
                  m_xRows.value(lngRow, 2) = strnewTV
                  m_xRows.value(lngRow, 3) = ""
                  m_xRows.value(lngRow, 4) = ""
                  m_xRows.value(lngRow, 5) = Format(Now, "mm/dd/yyyy")
                  m_xRows.value(lngRow, 6) = 0  ' Null
                  m_xRows.value(lngRow, 7) = Chr$(0) ' Null
                  TDBGrid1.RefetchRow
               End If
               
               
            End If
         Case Else
            Beep
            MsgBox "Number not currently in use", vbInformation, " "
         
      End Select
   End If

End Sub

Private Function CurrentRecordID() As Long
   Dim iRow As Integer
   iRow = TDBGrid1.Bookmark
   Dim lngRecordID As Long
   CurrentRecordID = m_xRows.value(iRow, 20)
End Function

Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)
   On Error Resume Next
   Select Case MsgId
      Case 300
         EndProgress

   End Select
End Sub

Private Sub UnAssignNumberForm_DoUpdate(ID As Long, dtLastUsed As Date)
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   Dim strTV As String
   strTV = UCase$(m_xRows.value(lngRow, 2))
           
   Dim strnewTV As String
   If strTV = "T" Then
      strnewTV = "R"
   Else
      strnewTV = "RV"
   End If
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim strSQL As String
   
   Set dbData = OpenLocalDB
   strSQL = "Select * from datLotus800 where id = " & CurrentRecordID
   Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
   If rsDAO.EOF And rsDAO.BOF Then
      'very bad - no record
      Stop
   Else
      rsDAO.Edit
      rsDAO.Fields("T_V") = strnewTV
      rsDAO.Fields("Term") = ""
      rsDAO.Fields("Customer") = ""
      rsDAO.Fields("AcctNo") = ""
      rsDAO.Fields("LastUsed") = Format(dtLastUsed, "mm/dd/yyyy")
      rsDAO.Fields("Rate") = 0
      rsDAO.Fields("Assigned") = CDate("12-31-1899")
      rsDAO.Update
      m_xRows.value(lngRow, 2) = strnewTV
      m_xRows.value(lngRow, 3) = ""
      m_xRows.value(lngRow, 4) = ""
      m_xRows.value(lngRow, 5) = Format(dtLastUsed, "mm/dd/yyyy")
      m_xRows.value(lngRow, 6) = 0  ' Null
      m_xRows.value(lngRow, 7) = Chr$(0) ' Null
      TDBGrid1.RefetchRow
   End If
   Unload UnAssignNumberForm
End Sub

Private Sub ShowFilter()
   If FilterForm Is Nothing Then
      Set FilterForm = New frmFilter
   End If
   FilterForm.Show vbModal, Me
End Sub

Private Sub ClearFilter()
   LoadGrid
   FilterForm.ResetParams
   UpdateStatusBar
End Sub

Private Sub FilterForm_ApplyChanges(Types As String, CustomerWhere As String, DNIS As String, AcctNo As String, Phone As String)
   Dim strWhere As String
   If LenTrim(Types) Then
      strWhere = "T_V In(" & Types & ")"
   End If
   If LenTrim(CustomerWhere) Then
      If LenTrim(strWhere) Then
         strWhere = strWhere & vbCrLf & "And" & vbCrLf
      End If
      strWhere = strWhere & CustomerWhere
   End If
   If LenTrim(DNIS) Then
      If LenTrim(strWhere) Then
         strWhere = strWhere & vbCrLf & "And" & vbCrLf
      End If
      strWhere = strWhere & DNIS
   End If
   If LenTrim(AcctNo) Then
      If LenTrim(strWhere) Then
         strWhere = strWhere & vbCrLf & "And" & vbCrLf
      End If
      strWhere = strWhere & AcctNo
   End If
   If LenTrim(Phone) Then
      If LenTrim(strWhere) Then
         strWhere = strWhere & vbCrLf & "And" & vbCrLf
      End If
      strWhere = strWhere & Phone
   End If
   If LenTrim(strWhere) Then
      strWhere = " Where " & strWhere
   End If

   LoadGrid strWhere
   UpdateStatusBar
End Sub

Private Sub UpdateStatusBar()
   Dim intShownRows As Integer
   intShownRows = m_xRows.UpperBound(1) + 1
   smartStatus.SmartItems("MDBPath").Caption = GetIniSetting(dz_MDBPath)
   smartStatus.SmartItems("numShown").Caption = intShownRows & " of " & m_totalRows
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 2 Then
      If KeyCode = vbKeyC Then
         KeyCode = 0
         Clipboard.Clear
         Clipboard.SetText TDBGrid1.Text
      End If
   End If
End Sub
