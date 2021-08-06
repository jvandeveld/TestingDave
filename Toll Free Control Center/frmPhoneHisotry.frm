VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{95559FD0-8A4C-11D3-905E-00A04B0669E7}#1.1#0"; "SmartUI.ocx"
Begin VB.Form frmPhoneHistory 
   Caption         =   "Phone History"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "frmPhoneHisotry.frx":0000
      Height          =   3450
      Left            =   285
      TabIndex        =   0
      Top             =   585
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   6085
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "To_DID"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "From_DID"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Local Time"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Call Length"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   979
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2170"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=532"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2805"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2725"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=532"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3069"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2990"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1667"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1588"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=77,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(9)   =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin BoWSmartUI.SmartUI smartStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5175
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   635
      Template        =   "frmPhoneHisotry.frx":0014
   End
   Begin BoWSmartUI.SmartUI smartToolBar 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   794
      Template        =   "frmPhoneHisotry.frx":082C
   End
End
Attribute VB_Name = "frmPhoneHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_strBillingPeriod As String
Private m_strPhoneNumber As String
Private m_dblSumMins As Double
Private boolAsc(20) As Boolean
Private m_xRows As XArrayDB

Property Let BillingPeriod(value As String)
   m_strBillingPeriod = value
End Property

Property Let PhoneNumber(value As String)
   m_strPhoneNumber = value
End Property

Private Sub Form_Load()
   With TDBGrid1
      .AllowUpdate = False
      .ColumnFooters = False
      .AllowColSelect = False
      .AlternatingRowStyle = True
      .EvenRowStyle.BackColor = RGB(200, 235, 255)
      .RecordSelectors = False
      .MarqueeStyle = dbgHighlightRowRaiseCell
      '.FetchRowStyle = True
      '.SetFocus
   End With
   LoadGrid
End Sub

Private Sub LoadGrid()
   TDBGrid1.Caption = "Call details for : " & m_strPhoneNumber
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim sumMins As Long
   Dim numCalls As Long
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
      strSQL = "Select to_did, from_did, LocalTime, billable / 60 as CallLength from datInboundLive where Month(localTime) = " & Mid$(m_strBillingPeriod, 5) & vbCrLf
      strSQL = strSQL & "And Year(localTime) = " & Left(m_strBillingPeriod, 4) & vbCrLf
      strSQL = strSQL & "And to_did = '1" & m_strPhoneNumber & "'"
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
      If Err.Number <> 0 Then
         Beep
         MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
      Else
         m_dblSumMins = 0
         Do Until rsDAO.EOF
            iRows = iRows + 1
            m_xRows.value(iRows, 0) = FieldValue(rsDAO, "to_did")
            m_xRows.value(iRows, 1) = FieldValue(rsDAO, "from_did")
            m_xRows.value(iRows, 2) = FieldValue(rsDAO, "LocalTime")
            m_xRows.value(iRows, 3) = FieldValue(rsDAO, "CallLength")
            m_dblSumMins = m_dblSumMins + m_xRows.value(iRows, 3)
            rsDAO.MoveNext
         Loop
      End If
      
      dbData.Close
      Set dbData = Nothing
         
      m_xRows.ReDim 0, iRows, 0, 20
      Set TDBGrid1.Array = m_xRows
      TDBGrid1.ReBind
      TDBGrid1.Bookmark = 0
   End If
   UpdateStatusBar
End Sub

Private Sub Form_Resize()
   ResizeControls
End Sub

Private Sub ResizeControls()
   On Error Resume Next
   Err.Clear
   TDBGrid1.Move 0, smartToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - smartToolBar.Height - smartStatus.Height
End Sub

Private Sub UpdateStatusBar()
   Dim intShownRows As Integer
   intShownRows = m_xRows.UpperBound(1) + 1
   smartStatus.SmartItems("numShown").Caption = intShownRows & " Call(s) Totaling " & m_dblSumMins & " Minutes"
End Sub

Private Sub smartToolBar_Click(Item As BoWSmartUI.SmartItem)
   Select Case UCase$(Trim$(Item.Key))
      Case "PRINT"
         PrintGrid TDBGrid1, FormatPhone(m_strPhoneNumber) & " CALL HISTORY"
   End Select
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
      Case 0, 1 'string
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_STRING
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_STRING
         End If
      Case 2 'Date
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_DATE
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_DATE
         End If
      Case 3 ' Call length
         If boolAsc(ColIndex) Then
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_SINGLE
         Else
            m_xRows.QuickSort 0, m_xRows.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE_SINGLE
         End If
   End Select
   TDBGrid1.ReBind
   TDBGrid1.Bookmark = 0
   TDBGrid1.LeftCol = intColStart
   EndProgress
   Screen.MousePointer = vbNormal
End Sub
