VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{95559FD0-8A4C-11D3-905E-00A04B0669E7}#1.1#0"; "SmartUI.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBillingUsage 
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12930
   Icon            =   "frmBillingUsage.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5415
      Left            =   255
      TabIndex        =   0
      Top             =   1140
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   9551
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Acct #"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Customer"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Phone #"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Minutes Used"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Included Minutes"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Over Minutes"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "#####.#0"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Over Rate"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Billable"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Currency"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Adjusted Billable"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Currency"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Notes"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).RecordSelectorWidth=   979
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=529"
      Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4154"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4075"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=532"
      Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1640"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1561"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(24)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(25)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=1958"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=1879"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=530"
      Splits(0)._ColumnProps(31)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(32)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(5).Width=1402"
      Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=1323"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=530"
      Splits(0)._ColumnProps(38)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(39)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(41)=   "Column(6).Width=1217"
      Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=530"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(48)=   "Column(7).Width=1773"
      Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=1693"
      Splits(0)._ColumnProps(51)=   "Column(7)._ColStyle=530"
      Splits(0)._ColumnProps(52)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(55)=   "Column(8).Width=1614"
      Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=1535"
      Splits(0)._ColumnProps(58)=   "Column(8)._ColStyle=530"
      Splits(0)._ColumnProps(59)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(60)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(62)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(63)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(65)=   "Column(9)._ColStyle=532"
      Splits(0)._ColumnProps(66)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(67)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(68)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(69)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(10)._ColStyle=532"
      Splits(0)._ColumnProps(73)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(74)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(75)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(76)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(77)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(79)=   "Column(11)._ColStyle=532"
      Splits(0)._ColumnProps(80)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(81)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(83)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(12)._ColStyle=532"
      Splits(0)._ColumnProps(87)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(89)=   "Column(12).Order=13"
      Splits(1)._UserFlags=   0
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   979
      Splits(1).DividerColor=   15790320
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=13"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2302"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=532"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(1)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(8)=   "Column(1).Width=2963"
      Splits(1)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(10)=   "Column(1)._WidthInPix=2884"
      Splits(1)._ColumnProps(11)=   "Column(1)._ColStyle=532"
      Splits(1)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(1)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(15)=   "Column(2).Width=2328"
      Splits(1)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(17)=   "Column(2)._WidthInPix=2249"
      Splits(1)._ColumnProps(18)=   "Column(2)._ColStyle=529"
      Splits(1)._ColumnProps(19)=   "Column(2).WrapText=1"
      Splits(1)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(21)=   "Column(3).Width=1640"
      Splits(1)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(23)=   "Column(3)._WidthInPix=1561"
      Splits(1)._ColumnProps(24)=   "Column(3)._ColStyle=530"
      Splits(1)._ColumnProps(25)=   "Column(3).WrapText=1"
      Splits(1)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(27)=   "Column(4).Width=1958"
      Splits(1)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(4)._WidthInPix=1879"
      Splits(1)._ColumnProps(30)=   "Column(4)._ColStyle=530"
      Splits(1)._ColumnProps(31)=   "Column(4).WrapText=1"
      Splits(1)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(33)=   "Column(5).Width=1402"
      Splits(1)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(5)._WidthInPix=1323"
      Splits(1)._ColumnProps(36)=   "Column(5)._ColStyle=530"
      Splits(1)._ColumnProps(37)=   "Column(5).WrapText=1"
      Splits(1)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(39)=   "Column(6).Width=1217"
      Splits(1)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(41)=   "Column(6)._WidthInPix=1138"
      Splits(1)._ColumnProps(42)=   "Column(6)._ColStyle=530"
      Splits(1)._ColumnProps(43)=   "Column(6).WrapText=1"
      Splits(1)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(45)=   "Column(7).Width=1773"
      Splits(1)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(7)._WidthInPix=1693"
      Splits(1)._ColumnProps(48)=   "Column(7)._ColStyle=530"
      Splits(1)._ColumnProps(49)=   "Column(7).WrapText=1"
      Splits(1)._ColumnProps(50)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(51)=   "Column(8).Width=1614"
      Splits(1)._ColumnProps(52)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(8)._WidthInPix=1535"
      Splits(1)._ColumnProps(54)=   "Column(8)._ColStyle=530"
      Splits(1)._ColumnProps(55)=   "Column(8).WrapText=1"
      Splits(1)._ColumnProps(56)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(57)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(58)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(60)=   "Column(9)._ColStyle=532"
      Splits(1)._ColumnProps(61)=   "Column(9).WrapText=1"
      Splits(1)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(63)=   "Column(10).Width=2725"
      Splits(1)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(10)._WidthInPix=2646"
      Splits(1)._ColumnProps(66)=   "Column(10)._ColStyle=532"
      Splits(1)._ColumnProps(67)=   "Column(10).WrapText=1"
      Splits(1)._ColumnProps(68)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(69)=   "Column(11).Width=2725"
      Splits(1)._ColumnProps(70)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(71)=   "Column(11)._WidthInPix=2646"
      Splits(1)._ColumnProps(72)=   "Column(11)._ColStyle=532"
      Splits(1)._ColumnProps(73)=   "Column(11).WrapText=1"
      Splits(1)._ColumnProps(74)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(75)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(76)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(78)=   "Column(12)._ColStyle=532"
      Splits(1)._ColumnProps(79)=   "Column(12).WrapText=1"
      Splits(1)._ColumnProps(80)=   "Column(12).Order=13"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   2
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   0
      DirectionAfterTab=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2,.wraptext=-1"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(82)  =   "Splits(1).Style:id=87,.parent=1"
      _StyleDefs(83)  =   "Splits(1).CaptionStyle:id=96,.parent=4"
      _StyleDefs(84)  =   "Splits(1).HeadingStyle:id=88,.parent=2"
      _StyleDefs(85)  =   "Splits(1).FooterStyle:id=89,.parent=3"
      _StyleDefs(86)  =   "Splits(1).InactiveStyle:id=90,.parent=5"
      _StyleDefs(87)  =   "Splits(1).SelectedStyle:id=92,.parent=6"
      _StyleDefs(88)  =   "Splits(1).EditorStyle:id=91,.parent=7"
      _StyleDefs(89)  =   "Splits(1).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(90)  =   "Splits(1).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(91)  =   "Splits(1).OddRowStyle:id=95,.parent=10"
      _StyleDefs(92)  =   "Splits(1).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(93)  =   "Splits(1).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(94)  =   "Splits(1).Columns(0).Style:id=102,.parent=87"
      _StyleDefs(95)  =   "Splits(1).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(96)  =   "Splits(1).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(97)  =   "Splits(1).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(98)  =   "Splits(1).Columns(1).Style:id=106,.parent=87"
      _StyleDefs(99)  =   "Splits(1).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(100) =   "Splits(1).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(101) =   "Splits(1).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(102) =   "Splits(1).Columns(2).Style:id=110,.parent=87,.alignment=2"
      _StyleDefs(103) =   "Splits(1).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(104) =   "Splits(1).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(105) =   "Splits(1).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(106) =   "Splits(1).Columns(3).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(107) =   "Splits(1).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(108) =   "Splits(1).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(109) =   "Splits(1).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(110) =   "Splits(1).Columns(4).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(111) =   "Splits(1).Columns(4).HeadingStyle:id=115,.parent=88"
      _StyleDefs(112) =   "Splits(1).Columns(4).FooterStyle:id=116,.parent=89"
      _StyleDefs(113) =   "Splits(1).Columns(4).EditorStyle:id=117,.parent=91"
      _StyleDefs(114) =   "Splits(1).Columns(5).Style:id=122,.parent=87,.alignment=1"
      _StyleDefs(115) =   "Splits(1).Columns(5).HeadingStyle:id=119,.parent=88"
      _StyleDefs(116) =   "Splits(1).Columns(5).FooterStyle:id=120,.parent=89"
      _StyleDefs(117) =   "Splits(1).Columns(5).EditorStyle:id=121,.parent=91"
      _StyleDefs(118) =   "Splits(1).Columns(6).Style:id=126,.parent=87,.alignment=1"
      _StyleDefs(119) =   "Splits(1).Columns(6).HeadingStyle:id=123,.parent=88"
      _StyleDefs(120) =   "Splits(1).Columns(6).FooterStyle:id=124,.parent=89"
      _StyleDefs(121) =   "Splits(1).Columns(6).EditorStyle:id=125,.parent=91"
      _StyleDefs(122) =   "Splits(1).Columns(7).Style:id=130,.parent=87,.alignment=1"
      _StyleDefs(123) =   "Splits(1).Columns(7).HeadingStyle:id=127,.parent=88"
      _StyleDefs(124) =   "Splits(1).Columns(7).FooterStyle:id=128,.parent=89"
      _StyleDefs(125) =   "Splits(1).Columns(7).EditorStyle:id=129,.parent=91"
      _StyleDefs(126) =   "Splits(1).Columns(8).Style:id=134,.parent=87,.alignment=1"
      _StyleDefs(127) =   "Splits(1).Columns(8).HeadingStyle:id=131,.parent=88"
      _StyleDefs(128) =   "Splits(1).Columns(8).FooterStyle:id=132,.parent=89"
      _StyleDefs(129) =   "Splits(1).Columns(8).EditorStyle:id=133,.parent=91"
      _StyleDefs(130) =   "Splits(1).Columns(9).Style:id=138,.parent=87"
      _StyleDefs(131) =   "Splits(1).Columns(9).HeadingStyle:id=135,.parent=88"
      _StyleDefs(132) =   "Splits(1).Columns(9).FooterStyle:id=136,.parent=89"
      _StyleDefs(133) =   "Splits(1).Columns(9).EditorStyle:id=137,.parent=91"
      _StyleDefs(134) =   "Splits(1).Columns(10).Style:id=142,.parent=87"
      _StyleDefs(135) =   "Splits(1).Columns(10).HeadingStyle:id=139,.parent=88"
      _StyleDefs(136) =   "Splits(1).Columns(10).FooterStyle:id=140,.parent=89"
      _StyleDefs(137) =   "Splits(1).Columns(10).EditorStyle:id=141,.parent=91"
      _StyleDefs(138) =   "Splits(1).Columns(11).Style:id=146,.parent=87"
      _StyleDefs(139) =   "Splits(1).Columns(11).HeadingStyle:id=143,.parent=88"
      _StyleDefs(140) =   "Splits(1).Columns(11).FooterStyle:id=144,.parent=89"
      _StyleDefs(141) =   "Splits(1).Columns(11).EditorStyle:id=145,.parent=91"
      _StyleDefs(142) =   "Splits(1).Columns(12).Style:id=150,.parent=87"
      _StyleDefs(143) =   "Splits(1).Columns(12).HeadingStyle:id=147,.parent=88"
      _StyleDefs(144) =   "Splits(1).Columns(12).FooterStyle:id=148,.parent=89"
      _StyleDefs(145) =   "Splits(1).Columns(12).EditorStyle:id=149,.parent=91"
      _StyleDefs(146) =   "Named:id=33:Normal"
      _StyleDefs(147) =   ":id=33,.parent=0"
      _StyleDefs(148) =   "Named:id=34:Heading"
      _StyleDefs(149) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(150) =   ":id=34,.wraptext=-1"
      _StyleDefs(151) =   "Named:id=35:Footing"
      _StyleDefs(152) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(153) =   "Named:id=36:Selected"
      _StyleDefs(154) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(155) =   "Named:id=37:Caption"
      _StyleDefs(156) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(157) =   "Named:id=38:HighlightRow"
      _StyleDefs(158) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(159) =   "Named:id=39:EvenRow"
      _StyleDefs(160) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(161) =   "Named:id=40:OddRow"
      _StyleDefs(162) =   ":id=40,.parent=33"
      _StyleDefs(163) =   "Named:id=41:RecordSelector"
      _StyleDefs(164) =   ":id=41,.parent=34"
      _StyleDefs(165) =   "Named:id=42:FilterBar"
      _StyleDefs(166) =   ":id=42,.parent=33"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1530
      Top             =   6540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "CSV"
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5685
      Picture         =   "frmBillingUsage.frx":3A0A
      Top             =   6735
      Width           =   240
   End
   Begin BoWSmartUI.SmartUI smartToolBar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   661
      Template        =   "frmBillingUsage.frx":3D4C
   End
   Begin BoWSmartUI.SmartUI smartPopUp 
      Height          =   360
      Left            =   9555
      TabIndex        =   1
      Top             =   6945
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   635
      Template        =   "frmBillingUsage.frx":4F80
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileCreateCSV 
         Caption         =   "Create CSV Billing File"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileNextBookmark 
         Caption         =   "Next Bookmark"
      End
   End
End
Attribute VB_Name = "frmBillingUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_intFlagCol As Integer
Private m_strBillingPeriod As String
Private m_dtInvoiceDate As Date
Private m_dtBillingPeriod As Date
Private m_lngAdjustmentRecordID As Long
Private m_xRows As XArrayDB
Private WithEvents FindDialogue As frmFind
Attribute FindDialogue.VB_VarHelpID = -1
Private WithEvents BillableAdjustment As frmAdjustment
Attribute BillableAdjustment.VB_VarHelpID = -1

Private m_lngRow As Long
Private m_lngCol As Long
Private m_regen As Boolean
Private m_currBMCount As Integer

Property Let BillingPeriod(value As String)
   m_strBillingPeriod = value
End Property

Property Let Regen(value As Boolean)
   m_regen = value
End Property

Private Sub LoadGrid()
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim rsTemp As DAO.Recordset

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
   
      If m_regen Then
         strSQL = "Delete from datMonthlyUsage where BillingPeriod = '" & m_strBillingPeriod & "'"
         dbData.Execute strSQL, dbFailOnError
         If Err.Number <> 0 Then
            Stop
         End If
      End If
      
      strSQL = "Select count(*) from datMonthlyUsage where BillingPeriod = '" & m_strBillingPeriod & "'"
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
      If rsDAO(0).value = 0 Then
         'First Time this is being loaded
         strSQL = "Select A.AcctNum, A.Name,  B.Phone, B.AcctNo, C.MinutesUsed,D.Quan1Inc, B.Rate From"
         strSQL = strSQL & "(((" & vbCrLf
         strSQL = strSQL & "CustomerData A" & vbCrLf
         strSQL = strSQL & "Left Join datLotus800 B on A.AcctNum = B.AcctNo)" & vbCrLf
         strSQL = strSQL & "Left Join" & vbCrLf
         strSQL = strSQL & "(" & vbCrLf
         strSQL = strSQL & "Select Mid(to_did,2) as Phone, sum(billable)/60 as MinutesUsed" & vbCrLf
         strSQL = strSQL & "From datInboundLive" & vbCrLf
         strSQL = strSQL & "Where Month(LocalTime) = " & Month(m_dtBillingPeriod) & vbCrLf
         strSQL = strSQL & "And" & vbCrLf
         strSQL = strSQL & "Year(LocalTime) = " & Year(m_dtBillingPeriod) & vbCrLf
         strSQL = strSQL & "Group by Mid(to_did,2)" & vbCrLf
         strSQL = strSQL & ") C on B.Phone = C.Phone)" & vbCrLf
         strSQL = strSQL & "Left Join CustomerData D on D.AcctNum = B.AcctNo)" & vbCrLf
         strSQL = strSQL & "Where" & vbCrLf
         strSQL = strSQL & "C.MinutesUsed > 0" & vbCrLf
         strSQL = strSQL & "And" & vbCrLf
         strSQL = strSQL & "B.Phone is not null" & vbCrLf
         strSQL = strSQL & "Order By 1"
         Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
         
         Dim rsAdd As DAO.Recordset
         strSQL = "Select * from datMonthlyUsage where 1 = 0"
         Set rsAdd = dbData.OpenRecordset(strSQL, dbOpenDynaset)
         
         Do Until rsDAO.EOF
            rsAdd.AddNew
            rsAdd("BillingPeriod") = m_strBillingPeriod
            rsAdd("AcctNo") = FieldValue(rsDAO, "AcctNo")
            rsAdd("Customer") = FieldValue(rsDAO, "Name")
            rsAdd("Phone") = FieldValue(rsDAO, "Phone")
            rsAdd("MinutesUsed") = FieldValue(rsDAO, "MinutesUsed")
            rsAdd("MinutesIncluded") = FieldValue(rsDAO, "Quan1Inc")
            rsAdd("OverRate") = Val(FieldValue(rsDAO, "Rate"))
            
            rsAdd.Update
            rsDAO.MoveNext
         Loop
      End If
      
      Dim intGroupCnt As Integer
      intGroupCnt = 0
      Dim strAct As String
      strAct = ""
      Dim dblCurrMinTotal As Double
      dblCurrMinTotal = 0
      Dim intFreeMin As Integer
      intFreeMin = -1
      Dim dblCurrRate As Double
      dblCurrRate = 0
      
      strSQL = "Select A.ID as HeaderID, A.BillingPeriod, A.AcctNo, A.Customer, A.Phone,A.MinutesIncluded, A.MinutesUsed, A.OverRate, "
      strSQL = strSQL & "B.ID as AdjustmentID, B.AdjustedBillable, B.Notes" & vbCrLf
      strSQL = strSQL & "From datMonthlyUsage A" & vbCrLf
      strSQL = strSQL & "Left Join datMonthlyUsageAdjustments B on A.BillingPeriod = B.BillingPeriod And A.AcctNo = B.AcctNo" & vbCrLf
      strSQL = strSQL & "Where A.BillingPeriod = '" & m_strBillingPeriod & "' Order By A.AcctNo"
      'from datMonthlyUsage where BillingPeriod = '" & m_strBillingPeriod & "' Order By AcctNo"
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
      If Err.Number <> 0 Then
         Beep
         MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
      Else
         Do Until rsDAO.EOF
            iRows = iRows + 1
            If strAct <> "" Then
               If strAct <> FieldValue(rsDAO, "AcctNo") Then
                  If intGroupCnt > 1 Then
                     m_xRows.value(iRows, 17) = -1
                     m_xRows.value(iRows, 3) = dblCurrMinTotal
                     m_xRows.value(iRows, 4) = intFreeMin
                     m_xRows.value(iRows, 5) = Round(dblCurrMinTotal - intFreeMin, 1)
                     m_xRows.value(iRows, 6) = Round(dblCurrRate, 2)
                     'm_xRows.value(iRows, 11) = FieldValue(rsDAO, "HeaderID")
                     m_xRows.value(iRows, 12) = m_xRows.value(iRows - 1, 12)
                     m_xRows.value(iRows, 0) = m_xRows.value(iRows - 1, 0)
                     m_xRows.value(iRows, 8) = m_xRows.value(iRows - 1, 8)
                     m_xRows.value(iRows, 9) = m_xRows.value(iRows - 1, 9)
                     m_xRows.value(iRows, 1) = m_xRows.value(iRows - 1, 1)

                     If (m_xRows.value(iRows, 5) < 0) Then
                        m_xRows.value(iRows, 5) = 0
                     End If
                     m_xRows.value(iRows, 7) = Round(dblCurrRate * m_xRows.value(iRows, 5), 1)
                     iRows = iRows + 1
                  End If
                  intGroupCnt = 0
                  dblCurrMinTotal = 0
                  intFreeMin = -1
               Else
                  m_xRows.value(iRows - 1, 4) = ""
                  m_xRows.value(iRows - 1, 18) = -1
                  m_xRows.value(iRows - 1, 5) = ""
                  m_xRows.value(iRows - 1, 6) = ""
                  m_xRows.value(iRows - 1, 7) = ""
                  intFreeMin = FieldValue(rsDAO, "MinutesIncluded")
               End If
            End If
            m_xRows.value(iRows, 8) = Val(FieldValue(rsDAO, "AdjustedBillable"))
            m_xRows.value(iRows, 9) = FieldValue(rsDAO, "Notes")
            m_xRows.value(iRows, 11) = FieldValue(rsDAO, "HeaderID")
            m_xRows.value(iRows, 12) = FieldValue(rsDAO, "AdjustmentID")

            strAct = FieldValue(rsDAO, "AcctNo")
            intGroupCnt = intGroupCnt + 1
            dblCurrMinTotal = dblCurrMinTotal + Val(FieldValue(rsDAO, "MinutesUsed"))
            dblCurrRate = Val(FieldValue(rsDAO, "OverRate"))
            m_xRows.value(iRows, 0) = FieldValue(rsDAO, "AcctNo")
            m_xRows.value(iRows, 1) = FieldValue(rsDAO, "Customer")
            m_xRows.value(iRows, 2) = FormatPhone(FieldValue(rsDAO, "Phone"))
            
            m_xRows.value(iRows, 3) = FieldValue(rsDAO, "MinutesUsed")
            If intFreeMin = -1 Then
               m_xRows.value(iRows, 4) = FieldValue(rsDAO, "MinutesIncluded")
            Else
               m_xRows.value(iRows, 18) = -1 'marker for row apart of larger total
            End If
            m_xRows.value(iRows, 5) = Val(m_xRows.value(iRows, 3)) - Val(FieldValue(rsDAO, "MinutesIncluded"))
            If m_xRows.value(iRows, 5) < 0 Then
               m_xRows.value(iRows, 5) = 0
            End If
            m_xRows.value(iRows, 6) = FieldValue(rsDAO, "OverRate")
            m_xRows.value(iRows, 7) = Round(Val(m_xRows.value(iRows, 6)) * Val(m_xRows.value(iRows, 5)), 1)
            If intFreeMin <> -1 Then
               m_xRows.value(iRows, 5) = ""
               m_xRows.value(iRows, 6) = ""
               m_xRows.value(iRows, 7) = ""
            End If
            If Err.Number <> 0 Then
               Stop
            End If
            'Stop
            rsDAO.MoveNext
         Loop
      End If
      dbData.Close
      Set dbData = Nothing
         
      m_xRows.ReDim 0, iRows, 0, 20
      LoadBookmarks

      Set TDBGrid1.Array = m_xRows
      TDBGrid1.ReBind
      TDBGrid1.Bookmark = 0
   End If
End Sub

Private Sub ShowFindForm()
   Set FindDialogue = New frmFind
   FindDialogue.Show vbModal, Me
End Sub

Private Sub DeleteAdjustment(dbData As DAO.Database, AdjustmentRow As Integer)
   On Error Resume Next
   Err.Clear
   Dim strSQL As String
   strSQL = "Delete from datMonthlyUsageAdjustments where id = " & m_lngAdjustmentRecordID
   dbData.Execute strSQL, dbFailOnError
   If Err.Number <> 0 Then
      Stop
   Else
      m_lngAdjustmentRecordID = 0
      m_xRows.value(AdjustmentRow, 12) = ""
      m_xRows.value(AdjustmentRow, 8) = ""
      m_xRows.value(AdjustmentRow, 9) = ""
      TDBGrid1.RefetchRow (AdjustmentRow)
   End If
End Sub

Private Sub BillableAdjustment_ApplyAdjustment(AdjustmentID As Long, NewBillable As String, Notes As String)
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim strSQL As String
   Set dbData = OpenLocalDB
   Dim intRow As Integer
   intRow = GetTotalRow
   If m_lngAdjustmentRecordID = 0 Then
      'Insert a new adjustment row
      strSQL = "Select * from datMonthlyUsageAdjustments where 1 = 0"
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
      rsDAO.AddNew
      rsDAO("BillingPeriod") = m_strBillingPeriod
      rsDAO("AcctNo") = m_xRows.value(intRow, 0)
      rsDAO("AdjustedBillable") = Val(NewBillable)
      rsDAO("Notes") = Notes
      rsDAO.Update
      rsDAO.Bookmark = rsDAO.LastModified
      Dim lngId As Long
      lngId = Val(FieldValue(rsDAO, "ID"))
      m_xRows.value(intRow, 8) = Val(NewBillable)
      m_xRows.value(intRow, 9) = Notes
      m_xRows.value(intRow, 12) = lngId
      TDBGrid1.RefetchRow (intRow)
   Else
      If Val(NewBillable) = Val(m_xRows(intRow, 7)) Then
         'Delete the adjustment
         DeleteAdjustment dbData, intRow
      Else
         'Update adjustment
         strSQL = "Select * from datMonthlyUsageAdjustments where id = " & m_lngAdjustmentRecordID
         Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenDynaset)
         rsDAO.Edit
         rsDAO("AdjustedBillable") = Val(NewBillable)
         rsDAO("Notes") = Notes
         rsDAO.Update
         m_xRows.value(intRow, 8) = Val(NewBillable)
         m_xRows.value(intRow, 9) = Notes
         TDBGrid1.RefetchRow (intRow)
      End If
   End If
   
   Unload BillableAdjustment
   dbData.Close
   Set dbData = Nothing
End Sub

Private Sub Form_Load()
   m_dtInvoiceDate = CDate(Mid(m_strBillingPeriod, 5) & "-01-" & Left(m_strBillingPeriod, 4))
   m_dtBillingPeriod = DateAdd("m", -2, m_dtInvoiceDate)
   
   Me.Caption = "Invoice Date: " & m_dtInvoiceDate & " (" & MonthName(Month(m_dtBillingPeriod)) & " " & Year(m_dtBillingPeriod) & " Usage)"
   '==========================
   'Add Flag Column
   m_intFlagCol = TDBGrid1.Columns.Count
   Dim colFlag As TrueDBGrid80.Column
   Set colFlag = TDBGrid1.Columns.Add(m_intFlagCol)
   colFlag.Width = 350
   colFlag.Order = 0
   colFlag.Visible = True
   TDBGrid1.Splits(1).Columns(m_intFlagCol).Visible = False
   TDBGrid1.Splits(0).Size = 3
   TDBGrid1.Splits(0).SizeMode = dbgNumberOfColumns
   '==========================
   
   TDBGrid1.Splits(0).SelectedStyle.BackColor = TDBGrid1.HighlightRowStyle.BackColor
   TDBGrid1.Splits(0).SelectedForeColor = TDBGrid1.HighlightRowStyle.ForeColor
   With TDBGrid1
      .AllowUpdate = False
      .ColumnFooters = False
      .AllowColSelect = False
      .AlternatingRowStyle = True
      .EvenRowStyle.BackColor = RGB(200, 235, 255)
      .RecordSelectors = False
      .MarqueeStyle = dbgHighlightRowRaiseCell
      .FetchRowStyle = True
      '.SetFocus
   End With
   
   TDBGrid1.Splits(0).AlternatingRowStyle = False
   
   'TDBGrid1.Columns(0).Alignment = dbgCenter
   TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRow
   TDBGrid1.Columns(m_intFlagCol).FetchStyle = dbgFetchCellStyleColumn
   TDBGrid1.Columns(0).FetchStyle = dbgFetchCellStyleColumn
   'TDBGrid1.Columns(2).FetchStyle = dbgFetchCellStyleColumn
   TDBGrid1.Columns(8).FetchStyle = dbgFetchCellStyleColumn
   TDBGrid1.Columns(9).FetchStyle = dbgFetchCellStyleColumn
   TDBGrid1.Columns(0).Merge = dbgMergeFree
   TDBGrid1.Columns(1).Merge = dbgMergeRestricted
      
   If Not DesignMode Then
      TDBGrid1.Columns(10).Visible = False
      TDBGrid1.Columns(10).AllowFocus = False
      TDBGrid1.Columns(11).Visible = False
      TDBGrid1.Columns(11).AllowFocus = False
      TDBGrid1.Columns(12).Visible = False
      TDBGrid1.Columns(12).AllowFocus = False
   End If
      
      
   LoadGrid
   
   EndProgress
End Sub

Private Sub Form_Resize()
   ResizeControls
End Sub

Private Sub ResizeControls()
   On Error Resume Next
   Err.Clear
   TDBGrid1.Move 0, smartToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - smartToolBar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveBookmarks
End Sub

Private Sub mnuFileCreateCSV_Click()
   ExportTASBillerCSV CommonDialog1, m_strBillingPeriod
End Sub

Private Sub mnuFileNextBookmark_Click()
   NextBookmark
End Sub

Private Sub NextBookmark()
   Dim iRow As Integer
   iRow = TDBGrid1.Bookmark
   Dim i%
   For i% = iRow + 1 To m_xRows.UpperBound(1)
      If m_xRows.value(i%, 16) = -1 Then
         TDBGrid1.Bookmark = i%
         Exit For
      End If
   Next i%
End Sub

Private Sub PreviousBookmark()
   Dim iRow As Integer
   iRow = TDBGrid1.Bookmark
   Dim i%
   'Step backward
   For i% = iRow - 1 To 0 Step -1
      If m_xRows.value(i%, 16) = -1 Then
         TDBGrid1.Bookmark = i%
         Exit For
      End If
   Next i%
End Sub

Private Sub ClearAllBookmarks()
   Beep
   If MsgBox("Clear All Bookmarks?", vbQuestion + vbYesNo + vbDefaultButton2, " ") = vbYes Then
      Dim i%
      For i% = 0 To m_xRows.UpperBound(1)
         If m_xRows.value(i%, 16) <> 0 Then
            m_xRows.value(i%, 16) = 0
            TDBGrid1.RefreshRow i%
         End If
      Next i%
      m_currBMCount = 0
      smartToolBar.SmartItems.Item("nextBM").Enabled = False
      smartToolBar.SmartItems.Item("prevBM").Enabled = False
      smartToolBar.SmartItems.Item("clearBM").Enabled = False
   End If
End Sub

Private Sub mnuFind_Click()
   ShowFindForm
End Sub

Private Sub FindDialogue_FindText(IsSubstring As Boolean, Text As String)
   FindDialogueFindText TDBGrid1, IsSubstring, Text
End Sub

Private Sub FindDialogue_FindNext(IsSubstring As Boolean, Text As String)
   FindDialogueFindNext TDBGrid1, IsSubstring, Text
End Sub

Private Sub smartToolBar_Click(Item As BoWSmartUI.SmartItem)
   Select Case UCase$(Trim$(Item.Key))
      Case "TOGGLEBM"
         ToggleBookmark
      Case "NEXTBM"
         NextBookmark
      Case "PREVBM"
         PreviousBookmark
      Case "CLEARBM"
         ClearAllBookmarks
   End Select
End Sub

Private Sub TDBGrid1_DblClick()
   editCurr
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid80.StyleDisp)
   Select Case Col
      Case m_intFlagCol  'Bookmarks column
         If m_xRows.value(Bookmark, 16) = -1 Then
            Set CellStyle.ForegroundPicture = Image1.Picture
            CellStyle.ForegroundPicturePosition = dbgFPLeft
            CellStyle.TransparentForegroundPicture = True
         End If
      Case 8, 9 'Adjusted Billable
         If m_xRows.value(Bookmark, 18) = -1 Then
            'Single row w/o group
            CellStyle.ForegroundPicturePosition = dbgFPPictureOnly
         Else
            If Val(m_xRows.value(Bookmark, 12)) = 0 Then
               CellStyle.ForegroundPicturePosition = dbgFPPictureOnly
            End If
         End If
   End Select
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid80.StyleDisp)
   If m_xRows.value(Bookmark, 17) = -1 Then
      RowStyle.BackColor = vbYellow
      RowStyle.Font.Bold = True
   End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 2 Then
      If KeyCode = vbKeyC Then
         KeyCode = 0
         Clipboard.Clear
         Clipboard.SetText TDBGrid1.Text
      End If
   ElseIf KeyCode = vbKeyReturn Then
      KeyCode = 0
      editCurr
   End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyReturn Then
'      'KeyAscii = 0
'      editCurr
'   End If
End Sub

Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)
   On Error Resume Next
   Select Case MsgId
      Case 300
         EndProgress

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
            smartPopUp.SmartItems("EDIT").Visibility = viVisible
            smartPopUp("POPUP").PopupMenu TDBGrid1.Left + X, TDBGrid1.Top + Y
            'Me.PopupMenu mnuGridPopup  'Show the popup menu
         End If
      End If
   End If
End Sub

Private Sub smartPopUp_Click(Item As BoWSmartUI.SmartItem)
   Select Case Trim$(UCase$(Item.Key))
      Case "EDIT"
         editCurr
      Case "BOOKMARK"
         ToggleBookmark
      Case "NEXTBM"
         NextBookmark
      Case "COPY"
         Clipboard.Clear
         Clipboard.SetText TDBGrid1.Text
      Case "SHOWDETAIL"
         ShowCallDetail
      Case "RESET"
         DeleteAdjustment OpenLocalDB, GetTotalRow
   End Select
End Sub

Private Sub editCurr()
   Set BillableAdjustment = New frmAdjustment
   BillableAdjustment.AdjustmentID = Val(m_xRows(TDBGrid1.Bookmark, 12))
   BillableAdjustment.AdjustedBillable = Val(m_xRows(TDBGrid1.Bookmark, 8))
   BillableAdjustment.Notes = m_xRows(TDBGrid1.Bookmark, 9)
   BillableAdjustment.Show vbModal, Me
End Sub

Private Function GetTotalRow() As Integer
   m_lngAdjustmentRecordID = 0
   Dim iRow As Integer
   If m_xRows.value(TDBGrid1.Bookmark, 18) <> -1 Then
      'Not part of a group
      iRow = TDBGrid1.Bookmark
      If m_xRows.value(TDBGrid1.Bookmark, 17) = -1 Then
         'This is the total row
         m_lngAdjustmentRecordID = Val(m_xRows.value(TDBGrid1.Bookmark, 12))
      Else
         'Stand alone row
         m_lngAdjustmentRecordID = Val(m_xRows.value(TDBGrid1.Bookmark, 12))
      End If
      
   Else
      '18 is -1 so this is part of a group
      Dim i%
      i% = TDBGrid1.Bookmark
      Do Until m_xRows.value(i%, 17) = -1
         i% = i% + 1
      Loop
      iRow = i%
      m_lngAdjustmentRecordID = Val(m_xRows.value(iRow, 12))
   End If
   GetTotalRow = iRow
      
End Function

Private Sub LoadBookmarks()
   Dim strFile As String
   strFile = App.Path
   If Right$(strFile, 1) <> "\" Then
      strFile = strFile & "\"
   End If
   m_currBMCount = 0
   strFile = strFile & m_strBillingPeriod & "Bookmarks.dat"
   If Dir(strFile) <> "" Then
      'Bookmarks file exists
      Dim strLin As String
      Dim intFile As Integer
      intFile = FreeFile
      Open strFile For Input As #intFile
      Do Until EOF(intFile)
         Line Input #intFile, strLin
         If IsNumeric(strLin) Then
            m_xRows.value(Val(strLin), 16) = -1
            m_currBMCount = m_currBMCount + 1
         End If
      Loop
      Close #intFile
   End If
   If m_currBMCount = 0 Then
      smartToolBar.SmartItems.Item("nextBM").Enabled = False
      smartToolBar.SmartItems.Item("prevBM").Enabled = False
      smartToolBar.SmartItems.Item("clearBM").Enabled = False
   Else
      smartToolBar.SmartItems.Item("nextBM").Enabled = True
      smartToolBar.SmartItems.Item("prevBM").Enabled = True
      smartToolBar.SmartItems.Item("clearBM").Enabled = True
   End If
End Sub

Private Sub ToggleBookmark()
'   Dim i%
'   Dim curr As String
'   curr = m_xRows.value(TDBGrid1.Bookmark, 0)
'   Dim toggleOff As Boolean
'   toggleOff = False
'   For i% = 0 To bml.ListCount
'      If bml.List(i%) = curr Then
'         toggleOff = True
'         bml.RemoveItem (i%)
'         Exit For
'      ElseIf Val(bml.List(i%)) > Val(curr) Then
'         bml.AddItem curr, i%
'         toggleOff = True
'         Exit For
'      End If
'   Next i%
'   If Not toggleOff Then
'      bml.AddItem (curr)
'   End If
'   If m_currBMIndex = -1 Then m_currBMIndex = 0
      
      
   If m_xRows.value(TDBGrid1.Bookmark, 16) = -1 Then
      m_xRows.value(TDBGrid1.Bookmark, 16) = 0
      m_currBMCount = m_currBMCount - 1
   Else
      m_xRows.value(TDBGrid1.Bookmark, 16) = -1
      m_currBMCount = m_currBMCount + 1
   End If
   If m_currBMCount > 0 Then
      smartToolBar.SmartItems.Item("nextBM").Enabled = True
      smartToolBar.SmartItems.Item("prevBM").Enabled = True
      smartToolBar.SmartItems.Item("clearBM").Enabled = True
   Else
      smartToolBar.SmartItems.Item("nextBM").Enabled = False
      smartToolBar.SmartItems.Item("prevBM").Enabled = False
      smartToolBar.SmartItems.Item("clearBM").Enabled = False
   End If
   TDBGrid1.RefreshRow
End Sub

Private Sub SaveBookmarks()
   Dim strFile As String
   strFile = App.Path
   If Right$(strFile, 1) <> "\" Then
      strFile = strFile & "\"
   End If
   strFile = strFile & m_strBillingPeriod & "Bookmarks.dat"
   Dim intFile As Integer
   intFile = FreeFile
   Open strFile For Output As #intFile
   Dim i%
   For i% = 0 To m_xRows.UpperBound(1)
      If m_xRows.value(i%, 16) = -1 Then
         Print #intFile, i%
      End If
   Next i%
   Close #intFile
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   TDBGrid1.SelBookmarks.Clear
   TDBGrid1.SelBookmarks.Add (TDBGrid1.Bookmark)
End Sub

Private Sub ShowCallDetail()
   Dim frmForm As New frmCallDetail
   frmForm.BillingPeriod = Year(m_dtBillingPeriod) & Month(m_dtBillingPeriod)
   frmForm.PhoneNumber = UnformatPhone(m_xRows.value(TDBGrid1.Bookmark, 2))
   frmForm.Show vbModal, Me
End Sub
