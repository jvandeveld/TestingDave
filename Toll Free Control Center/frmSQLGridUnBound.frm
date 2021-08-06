VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmSQLGridUnBound 
   Caption         =   "Unbound Grid"
   ClientHeight    =   5835
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   510
      Left            =   6390
      TabIndex        =   4
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox txtSQL 
      Height          =   960
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   405
      Width           =   4110
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   360
      Left            =   4530
      TabIndex        =   1
      Top             =   375
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3585
      Left            =   615
      TabIndex        =   0
      Top             =   1845
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   6324
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   979
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4207"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4128"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2302"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
      HeadLines       =   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Sentence:"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   75
      Width           =   1260
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuGridPopUp 
      Caption         =   "Grid Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuEditRow 
         Caption         =   "Edit Current Record"
      End
   End
End
Attribute VB_Name = "frmSQLGridUnBound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmEditRecord As frmAddRecordDAO
Attribute frmEditRecord.VB_VarHelpID = -1
Private m_database As DAO.Database
Private m_lngCol As Long
Private m_lngRow As Long
Private m_xRows As New XArrayDB
Private WithEvents FindDialogue As frmFind
Attribute FindDialogue.VB_VarHelpID = -1

Private Sub cmdApply_Click()
   LoadGrid
End Sub

Private Sub cmdFind_Click()
   ShowFindForm
End Sub

Private Sub Form_Load()
   Set m_database = OpenDatabase("PhoneNumbers.mdb")
   With TDBGrid1
      .AllowColSelect = False
      .AlternatingRowStyle = True
      .EvenRowStyle.BackColor = RGB(200, 235, 255)
      .RecordSelectors = False
      .MarqueeStyle = dbgHighlightRowRaiseCell
   End With
   
   txtSQL.Text = "Select * from datLotus800"
End Sub


Private Sub LoadGrid()
   Dim iRows As Integer
   iRows = -1
   m_xRows.ReDim 0, 10000, 0, 20
   
   On Error Resume Next
   Err.Clear
   Dim rsDAO As DAO.Recordset
   Set rsDAO = m_database.OpenRecordset(txtSQL.Text, dbOpenForwardOnly)
   If Err.Number <> 0 Then
      Beep
      MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
   Else
      Do Until rsDAO.EOF
         iRows = iRows + 1
         m_xRows.value(iRows, 0) = FieldValue(rsDAO, "Phone") '  rsDAO.Fields("Phone").Value 'Needs error check for NULL, use wrapper
         m_xRows.value(iRows, 1) = FieldValue(rsDAO, "Customer")
         m_xRows.value(iRows, 2) = FieldValue(rsDAO, "Carrier")
         m_xRows.value(iRows, 20) = FieldValue(rsDAO, "id")
         If Err.Number <> 0 Then
            Stop
         End If
         'Stop
         rsDAO.MoveNext
      Loop
      'Stop
   End If
   
   m_xRows.ReDim 0, iRows, 0, 20
   Set TDBGrid1.Array = m_xRows
   TDBGrid1.Row = 0
   TDBGrid1.ReBind
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Err.Clear
   With TDBGrid1
      .Left = 0
      .Width = Me.ScaleWidth
      .Top = txtSQL.Top + txtSQL.Height + 100
      .Height = Me.ScaleHeight - .Top
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Err.Clear
   m_database.Close
   Set m_database = Nothing
End Sub

Private Sub frmEditRecord_RecordUpdated(Phone As String, Carrier As String, Customer As String)
   'Stop
   Dim lngRow As Long
   lngRow = TDBGrid1.Bookmark
   m_xRows.value(lngRow, 0) = Phone
   m_xRows.value(lngRow, 1) = Customer
   m_xRows.value(lngRow, 2) = Carrier
   TDBGrid1.RefetchRow
End Sub

Private Sub mnuEditRow_Click()
   Dim iRow As Integer
   iRow = TDBGrid1.Bookmark
   Dim lngRecordID As Long
   lngRecordID = m_xRows.value(iRow, 20)
   'Dim frmForm As New frmAddRecordDAO
   Set frmEditRecord = New frmAddRecordDAO
   frmEditRecord.RecordID = lngRecordID
   frmEditRecord.Show vbModal, Me
   
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
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
            'smartPopUp("POPUP").PopupMenu TDBGridHeaders.Left + X, TDBGridHeaders.Top + Y
            Me.PopupMenu mnuGridPopup
         End If
      End If
   End If
End Sub

Private Sub ShowFindForm()
   If Not IsLoaded("frmFind") Then
      Set FindDialogue = New frmFind
      FindDialogue.Show , Me
   Else
      FindDialogue.Show , Me
   End If
End Sub
Private Sub FindDialogue_FindText(IsSubstring As Boolean, Text As String)
   FindDialogueFindText TDBGrid1, IsSubstring, Text
End Sub

Private Sub FindDialogue_FindNext(IsSubstring As Boolean, Text As String)
   FindDialogueFindNext TDBGrid1, IsSubstring, Text
End Sub
