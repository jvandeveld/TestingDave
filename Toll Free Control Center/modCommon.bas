Attribute VB_Name = "modCommon"
Option Explicit
Public Declare Function GetAsyncKeyState% Lib "user32" (ByVal nVirtKey As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As typRECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage_LONG Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParm As Long, ByVal lParam As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetTempPath& Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String)
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const CS_DROPSHADOW As Long = &H20000
Public Const GCL_STYLE     As Long = -26
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_ERR = -1
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_FINDSTRING = &H14C

Public Type typRECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public g_clsProgress As Object
Public g_strErrorMessage As String
Public g_lngErrorNumber As Long

Public Enum eSettingName
   dz_MDBPath = 1
   dz_IgnoreEscape
   dz_DebugOn
End Enum

Public Function FieldValue(rsDAO As DAO.Recordset, FldName As String) As String
   If IsNull(rsDAO(FldName).value) Then
      FieldValue = ""
   Else
      FieldValue = rsDAO(FldName).value
   End If
End Function

Public Function OpenLocalDB() As DAO.Database
   g_strErrorMessage = ""
   g_lngErrorNumber = 0
   Dim strFile As String
   strFile = GetIniSetting(dz_MDBPath)
   If Dir(strFile) = "" Then
      'File does not exist
      g_strErrorMessage = "File does not exist: " & strFile & vbCrLf & "Check TollFree.ini file."
   Else
      On Error Resume Next
      Err.Clear
      Dim dbData As DAO.Database
      If Not LenTrim(strFile) Then
         Set OpenLocalDB = OpenDatabase("D:\PCMSI\Development\Data\PhoneNumbers\Good\PhoneNumbers.mdb")
      Else
         Set OpenLocalDB = OpenDatabase(strFile)
      End If
      If Err.Number <> 0 Then
         g_lngErrorNumber = Err.Number
         g_strErrorMessage = "Error #: " & g_lngErrorNumber & vbCrLf & Err.Description
      End If
   End If
End Function

Public Sub PrintGrid(TDBGrid1 As TrueDBGrid80.TDBGrid, Title As String)
    With TDBGrid1.PrintInfo
      .PageHeader = CStr(Now) & "\t" & Title & "\tPage \p"

      .RepeatColumnHeaders = True
      .RepeatColumnFooters = True
      .VariableRowHeight = True
      .RepeatSplitHeaders = False
      
      ' 1/4 inch margins all around
      .SettingsMarginBottom = 360
      .SettingsMarginTop = 360
      .SettingsMarginLeft = 360
      .SettingsMarginRight = 360
      
      .PageSetup
      If Not .PageSetupCancelled Then
         .PrintPreview
      End If
   End With
End Sub

Public Function FormatPhone(value As String) As String
   Dim strRet As String
   strRet = value
   If Len(strRet) > 6 Then
      strRet = Mid(strRet, 1, 3) & "-" & Mid(strRet, 4, 3) & "-" & Mid(strRet, 7)
   ElseIf Len(strRet) > 3 Then
      strRet = Mid(strRet, 1, 3) & "-" & Mid(strRet, 4)
   End If
   FormatPhone = strRet
End Function

Public Function UnformatPhone(value As String) As String
   UnformatPhone = Replace(value, "-", "")
End Function

Public Sub FindDialogueFindText(TDBGrid1 As TrueDBGrid80.TDBGrid, IsSubstring As Boolean, Text As String)
   Dim i%
   Dim iCol As Integer
   iCol = TDBGrid1.Col
   Dim xRows As XArrayDB
   Set xRows = TDBGrid1.Array
   If IsSubstring = True Then
      For i% = 0 To xRows.UpperBound(1)
         If InStr(1, xRows.value(i%, iCol), Text, vbTextCompare) > 0 Then
            TDBGrid1.Bookmark = i%
            Exit For
         End If
      Next i%
   Else
      For i% = 0 To xRows.UpperBound(1)
         If Left$(UCase$(xRows.value(i%, iCol)), Len(Text)) = UCase$(Text) Then
            TDBGrid1.Bookmark = i%
            Exit For
         End If
      Next i%
   End If
   TDBGrid1.PostMsg 300
End Sub

Public Sub FindDialogueFindNext(TDBGrid1 As TrueDBGrid80.TDBGrid, IsSubstring As Boolean, Text As String)
   Dim xRows As XArrayDB
   Set xRows = TDBGrid1.Array

   ShowProgressWindow "Find next match..."
   
   Dim iCol As Integer
   iCol = TDBGrid1.Col
   
   Dim i%
   Dim lRow As Long
   If IsSubstring = True Then
      For i% = TDBGrid1.Bookmark + 1 To xRows.UpperBound(1)
         If InStr(1, xRows.value(i%, iCol), Text, vbTextCompare) > 0 Then
            TDBGrid1.Bookmark = i%
            Exit For
         End If
      Next i%
      If i% > xRows.UpperBound(1) Then
         For i% = 0 To TDBGrid1.Bookmark
            If InStr(1, xRows.value(i%, iCol), Text, vbTextCompare) > 0 Then
               TDBGrid1.Bookmark = i%
               Exit For
            End If
         Next i%
      End If
   Else
      For i% = TDBGrid1.Bookmark + 1 To xRows.UpperBound(1)
         If Left$(UCase$(xRows.value(i%, iCol)), Len(Text)) = UCase$(Text) Then
            TDBGrid1.Bookmark = i%
            Exit For
         End If
      Next i%
      If i% > xRows.UpperBound(1) Then
         For i% = 0 To TDBGrid1.Bookmark
            If Left$(UCase$(xRows.value(i%, iCol)), Len(Text)) = UCase$(Text) Then
               TDBGrid1.Bookmark = i%
               Exit For
            End If
         Next i%
      End If
   End If
   TDBGrid1.PostMsg 300
End Sub

Public Function LenTrim(strData As String) As Boolean
   On Error Resume Next
   Err.Clear
   LenTrim = Len(Trim(strData)) > 0
End Function

Public Function IsLoaded(strFormName As String) As Boolean
   On Error GoTo ET
     Dim i%
     strFormName = Trim$(UCase$(strFormName))
     For i% = 0 To Forms.Count - 1
          If UCase$(Trim$(Forms(i%).Name)) = strFormName Then
               IsLoaded = True
               Exit For
          End If
     Next i%
ExitFunction:
Exit Function
ET:
   Err.Raise Err.Number, "MDIModule" & "." & "IsLoaded" & vbCr & Err.Source
End Function

Public Sub ShowProgressWindow(Caption As String)
   If Not g_clsProgress Is Nothing Then
      g_clsProgress.ShowProgressWindow Caption
   End If
End Sub

Public Sub EndProgress()
   If Not g_clsProgress Is Nothing Then
      g_clsProgress.EndProgress
   End If
End Sub

Public Function DesignMode() As Boolean
   DesignMode = CBool(App.LogMode = 0)
End Function

Public Function ControlKey() As Boolean
   On Error Resume Next
   Err.Clear
   'returns -32767 if the CTRL Key is down
   ControlKey = GetAsyncKeyState(&H11) < 0
End Function

Public Function ShiftKey() As Boolean
   On Error Resume Next
   Err.Clear

   'returns -32767 if the SHIFT Key is down
   ShiftKey = GetAsyncKeyState(&H10) < 0
End Function

Public Function GetIniSetting(intSetting As eSettingName) As String
   Static boolFileRead As Boolean
   Static strPaths(15) As String
   Dim strTmp As String
   strTmp = ""
   
   'LOCAL INI SECTION
   If Not boolFileRead Or True Then
      Dim intFile As Integer
      Dim strFileName As String
      Dim strThumbData As String
      Dim strVar As String
      Dim lin$
      Dim strAppPath As String
      strAppPath = App.Path
      If Right$(strAppPath, 1) <> "\" Then
         strAppPath = strAppPath & "\"
      End If
      strFileName = strAppPath & "TollFree.ini"
      If Dir(strFileName) <> "" Then
         intFile = FreeFile
         Dim i%
         Open strFileName For Input As #intFile
         
         Do While Not EOF(intFile)
            Line Input #intFile, lin$
            'lin$ = UCase$(Trim$(lin$))
            If Left$(Trim$(lin$), 1) <> "#" Then

               i% = InStr(lin$, "=")
               If i% > 0 Then
                  strVar = UCase$(Trim$(Left$(lin$, i% - 1)))
                  Select Case strVar
   
                     Case "MDBPATH"
                        strPaths(dz_MDBPath) = UCase$(Trim$(Mid$(lin$, i% + 1)))
                     
                     Case "IGNOREESCAPE"
                        strPaths(dz_IgnoreEscape) = UCase$(Trim$(Mid$(lin$, i% + 1)))  ' = "Y"
   

                     Case Else
                     
                  End Select
               End If
            End If
         Loop
         boolFileRead = True
         Close #intFile
      End If
         
      strFileName = "\\192.168.1.105\Apps\Toll Free Control Center\TollFree.ini"
      intFile = FreeFile
      
      
      'GLOBAL INI SECTION
      On Error Resume Next
      Err.Clear
      Open strFileName For Input As #intFile
      If Err.Number <> 0 Then
         'set default values if missing ini file
         If Not LenTrim(strPaths(dz_MDBPath)) Then
            strPaths(dz_MDBPath) = "D:\PCMSI\Development\Data\PhoneNumbers\Good\PhoneNumbers.mdb"
         End If
      Else
         Do While Not EOF(intFile)
            Line Input #intFile, lin$
            lin$ = UCase$(Trim$(lin$))
            i% = InStr(lin$, "=")
            If i% > 0 Then
               strVar = Left$(lin$, i% - 1)
               Select Case strVar
                  
                  Case "MDBPATH"
                     If Not LenTrim(strPaths(dz_MDBPath)) Then
                        strPaths(dz_MDBPath) = UCase$(Trim$(Mid$(lin$, i% + 1)))
                     End If
                  
                  Case "IGNOREESCAPE"
                     If Not LenTrim(strPaths(dz_IgnoreEscape)) Then
                        strPaths(dz_IgnoreEscape) = UCase$(Trim$(Mid$(lin$, i% + 1)))  ' = "Y"
                     End If
                     
               End Select
            End If
         Loop
      End If
      Close #intFile
         
   End If
   GetIniSetting = strPaths(intSetting)
End Function

Public Sub ExportTASBillerCSV(CommonDialog1 As CommonDialog, Period As String)
   On Error Resume Next
   Err.Clear
   Dim dbData As DAO.Database
   Dim rsDAO As DAO.Recordset
   Dim currVal As Double
   Dim strCommas As String
   Dim strSQL As String
   strCommas = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
   With CommonDialog1
      .CancelError = True
      .DialogTitle = "Save CSV File"
      .DefaultExt = "CSV"
      .ShowSave
   End With
   
   Set dbData = OpenLocalDB
   If dbData Is Nothing Then
      If IsLoaded("frmSplash") Then
         Unload frmSplash
      End If
      Beep
      MsgBox g_strErrorMessage, vbInformation, "ERROR"
   Else
   
      If True Then
         strSQL = "Select  A.AcctNo, A.MinutesIncluded, A.OverRate, B.TotalMinutesUsed, Round(B.TotalMinutesUsed - A.MinutesIncluded,0) as OverMinutes, C.AdjustedBillable" & vbCrLf
         strSQL = strSQL & "From" & vbCrLf
         strSQL = strSQL & "((" & vbCrLf
         strSQL = strSQL & "(Select Distinct AcctNo,BillingPeriod, OverRate, MinutesIncluded from datMonthlyUsage Where BillingPeriod = '" & Period & "') A" & vbCrLf
         strSQL = strSQL & "Left Join" & vbCrLf
         strSQL = strSQL & "(" & vbCrLf
         strSQL = strSQL & "SELECT AcctNo, Round(Sum(MinutesUsed),1) as TotalMinutesUsed" & vbCrLf
         strSQL = strSQL & "From datMonthlyUsage" & vbCrLf
         strSQL = strSQL & "Where BillingPeriod = '" & Period & "'" & vbCrLf
         strSQL = strSQL & "Group By AcctNo" & vbCrLf
         strSQL = strSQL & ") B on A.AcctNo = B.AcctNo)" & vbCrLf
         strSQL = strSQL & "Left Join datMonthlyUsageAdjustments C on A.AcctNo = C.AcctNo and A.BillingPeriod = C.BillingPeriod)" & vbCrLf
         strSQL = strSQL & "Where B.TotalMinutesUsed > A.MinutesIncluded" & vbCrLf
      Else
         strSQL = "Select A.ID as HeaderID, A.BillingPeriod, A.AcctNo, A.Customer, A.Phone,A.MinutesIncluded, A.MinutesUsed,A.OverRate,"
         strSQL = strSQL & "B.ID as AdjustmentID, B.AdjustedBillable, B.Notes" & vbCrLf
         strSQL = strSQL & "From datMonthlyUsage A" & vbCrLf
         strSQL = strSQL & "Left Join datMonthlyUsageAdjustments B on A.BillingPeriod = B.BillingPeriod And A.AcctNo = B.AcctNo" & vbCrLf
         strSQL = strSQL & "Where A.BillingPeriod = '" & Period & "' Order By A.AcctNo"
      End If
      Set rsDAO = dbData.OpenRecordset(strSQL, dbOpenForwardOnly)
      Dim intFile As Integer
      intFile = FreeFile
      Open CommonDialog1.FileName For Output As #intFile
      If Err.Number <> 0 Then
         Beep
         MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
      Else
         Dim strAct As String
         strAct = ""
         Dim dblCurrRate As Double
         dblCurrRate = 0
         Dim dblOverMins As Double
         dblOverMins = 0
'         Dim dblCurrMinTotal As Double
'         dblCurrMinTotal = 0
'         Dim intGroupCnt As Integer
'         intGroupCnt = 0
'         Dim dblIncludedMins As Double
'         dblIncludedMins = 0
'         Dim partOfTotal As Boolean
'         partOfTotal = False
'         Dim adjustedNullCheck As Boolean
'         adjustedNullCheck = False
'         Dim printLast As Boolean
'         printLast = False
'         Dim printStatement As String
'         printStatement = ""
         Do Until rsDAO.EOF
            dblCurrRate = Val(FieldValue(rsDAO, "OverRate"))
            strAct = FieldValue(rsDAO, "AcctNo")
            dblOverMins = Round(Val(FieldValue(rsDAO, "OverMinutes")), 0)
            If strAct <> "" Then
               If dblCurrRate > 0 Then
                  If Not IsNull(rsDAO("AdjustedBillable").value) Then
                     currVal = Val(FieldValue(rsDAO, "AdjustedBillable"))
                     currVal = Round(currVal / dblCurrRate, 0)
                     If currVal > 0 Then
                        Print #intFile, strAct & strCommas & Round(currVal, 0) & ","
                     End If
                  ElseIf dblOverMins > 0 Then
                     Print #intFile, strAct & strCommas & Round(dblOverMins, 0) & ","
                  End If
               End If
            End If
            rsDAO.MoveNext
         Loop
         Close #intFile
      End If
   End If
'   Dim strCommas As String
'   strCommas = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
'   With CommonDialog1
'      .CancelError = True
'      .DialogTitle = "Save CSV File"
'      .DefaultExt = "CSV"
'      .ShowSave
'      If Err.Number = 0 Then
'         Dim intFile As Integer
'         intFile = FreeFile
'         Open .FileName For Output As #intFile
'         Dim i%
'         For i% = 0 To m_xRows.UpperBound(1)
'            If m_xRows.value(i%, 17) = -1 Then
'               'If LenTrim(m_xRows.value(i% - 1, 0)) Then
'                  If IsNumeric(m_xRows.value(i% - 1, 0)) Then
'                     If Round(m_xRows.value(i%, 5), 1) > 0 Then ' <> "0" Then
'                        If m_xRows.value(i%, 7) <> "0" Then
'                           Print #intFile, m_xRows.value(i% - 1, 0) & strCommas & m_xRows.value(i%, 5) & ","
'                        End If
'                     End If
'                  End If
'               'End If
'            Else
'               If LenTrim(m_xRows.value(i%, 0)) Then
'                  If IsNumeric(m_xRows.value(i%, 0)) Then
'                     If Val(m_xRows.value(i%, 8)) <> 0 Then ' check adjusted billing
'                        If m_xRows.value(i%, 7) <> "0" Then ' check for normal billing value
'                           If m_xRows.value(i%, 18) <> "-1" Then ' check for not being part of a total account
'                              Print #intFile, m_xRows.value(i%, 0) & strCommas & m_xRows.value(i%, 5) & ","
'                           End If
'                        End If
'                     ElseIf Round(m_xRows.value(i%, 5), 1) > 0 Then ' check for rate > 0
'                        If m_xRows.value(i%, 7) <> "0" Then ' check for normal billing value
'                           If m_xRows.value(i%, 18) <> "-1" Then ' check for not being part of a total account
'                              Print #intFile, m_xRows.value(i%, 0) & strCommas & m_xRows.value(i%, 5) & ","
'                           End If
'                        End If
'                     End If
'                  End If
'               End If
'            End If
'         Next i%
'         Close #intFile
'      End If
'   End With
End Sub








'            If strAct <> "" Then
'               If strAct <> FieldValue(rsDAO, "AcctNo") Then
'                  If printLast Then
'                     Print #intFile, printStatement
'                  End If
'                  If intGroupCnt > 1 Then
'                     ' TOTAL ROW
'                     If dblCurrRate > 0 Then 'Check for an over rate
'                        If Not adjustedNullCheck Then 'Check if adj. billable exists
'                           currVal = Round(currVal, 0)
'                           If currVal > 0 Then 'Check if adj. billable is not 0 (to print to CSV)
'                              currVal = currVal / dblCurrRate
'                              Print #intFile, strAct & strCommas & currVal & ","
'                           End If
'                        Else
'                           'normal billable used here
'                           dblOverMins = dblCurrMinTotal - dblIncludedMins
'                           dblOverMins = Round(dblOverMins, 0)
'                           If dblOverMins > 0 Then
'                              Print #intFile, strAct & strCommas & Round(dblOverMins, 0) & ","
'                           End If
'                        End If
'                     End If
'                  End If
'                  partOfTotal = False
'                  intGroupCnt = 0
'                  dblCurrMinTotal = 0
'               Else
'                  ' PART OF TOTAL
'                  partOfTotal = True
'               End If
'            End If
'            printLast = False
'            'SET VALS NEEDED FOR GROUPING
'            adjustedNullCheck = IsNull(rsDAO("AdjustedBillable").value)
'            strAct = FieldValue(rsDAO, "AcctNo")
'
'            intGroupCnt = intGroupCnt + 1
'            dblIncludedMins = FieldValue(rsDAO, "MinutesIncluded")
'            dblCurrMinTotal = dblCurrMinTotal + Val(FieldValue(rsDAO, "MinutesUsed"))
'            dblCurrRate = Val(FieldValue(rsDAO, "OverRate"))
'            currVal = Val(FieldValue(rsDAO, "AdjustedBillable"))
'            ' CHECKING ADJ BILLABLE
'            If dblCurrRate > 0 Then 'Check for an over rate
'               If Not partOfTotal Then 'Check if Phone # is part of bigger acc total
'                  If Not adjustedNullCheck Then 'Check if adj. billable exists
'                     currVal = Round(currVal, 0)
'                     If currVal > 0 Then 'Check if adj. billable is not 0 (to print to CSV)
'                        currVal = currVal / dblCurrRate
'                        printLast = True
'                        printStatement = strAct & strCommas & Round(currVal, 0) & ","
'                     End If
'                  Else
'                     'normal billable used here
'                     dblOverMins = dblCurrMinTotal - dblIncludedMins
'                     dblOverMins = Round(dblOverMins, 0)
'                     If dblOverMins > 0 Then
'                        printLast = True
'                        printStatement = strAct & strCommas & Round(dblOverMins, 0) & ","
'                     End If
'                  End If
'               End If
'            End If
