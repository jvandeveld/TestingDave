VERSION 5.00
Begin VB.Form frmThinqManager 
   Caption         =   "Number Manager"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   2160
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   645
      Left            =   4455
      TabIndex        =   2
      Top             =   1155
      Width           =   750
   End
   Begin VB.CommandButton cmdGetTF 
      Caption         =   "GET TOLL-FREE"
      Height          =   465
      Left            =   1995
      TabIndex        =   1
      Top             =   1200
      Width           =   1320
   End
   Begin VB.CommandButton cmdGetDom 
      Caption         =   "GET DOMESTIC"
      Height          =   465
      Left            =   360
      TabIndex        =   0
      Top             =   1140
      Width           =   1320
   End
End
Attribute VB_Name = "frmThinqManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetDom_Click()
   Dim strToken As String
   Dim strUser As String
   Dim strAuth As String
   
   Dim strURL As String
   Dim http As New MSXML2.XMLHTTP30
         
   Dim strRq As String
   strRq = ""
   Dim q$
   q$ = Chr$(34)
   
   'clean the message
   Dim strMsg As String
   strMsg = Replace(msg, "\", "\\")
   strMsg = Replace(strMsg, q$, "\" & q$)
   strMsg = Replace(strMsg, vbCrLf, "\n")
   
   On Error Resume Next
   Err.Clear
   Dim iLevel As Integer
   Dim strError As String
   Dim intTries As Integer
   intTries = 0
   Do
      http.Open "POST", strURL, False
      If Err.Number <> 0 Then
         intTries = intTries + 1
         If intTries >= 5 Then
            iLevel = 100  'Critical
            strError = "Can't open http" & vbCrLf & "http.status = " & http.Status & vbCrLf & "http.responseText = " & http.responseText & vbCrLf & "Err #:" & Err.Number & Err.Description
            'ShutDownApplication strError
         Else
            Sleep 500  'sleep for 1/2 send then try again
            Err.Clear
         End If
      Else
         Exit Do
      End If
   Loop Until Err.Number = 0
   
   http.setRequestHeader "Content-type", "application/json"
   http.setRequestHeader "Authorization", "Basic " & strAuth
                  
   strRq = strRq & "{" & vbCrLf
   'strRq = strRq & "    " & q$ & "from_did" & q$ & ":" & q$ & TODOOOO & q$ & "," & vbCrLf
   strRq = strRq & "    " & q$ & "to_did" & q$ & ":" & q$ & ToDID & q$ & "," & vbCrLf
   strRq = strRq & "    " & q$ & "message" & q$ & ":" & q$ & strMsg & q$ & vbCrLf
   strRq = strRq & "}" & vbCrLf
   
   On Error Resume Next
   'make 5 attempts to send the message before critical error
   intTries = 0
   Do
      
      Err.Clear
      http.send strRq
      If Err.Number <> 0 Then
         intTries = intTries + 1
         If intTries > 5 Then
            'Critical Error
            iLevel = 100  'Critical
            strError = "Can't send http" & vbCrLf & "http.status = " & http.Status & vbCrLf & "http.statusText: " & http.statusText & vbCrLf & "http.responseText = " & http.responseText & vbCrLf & strRq & vbCrLf & "Err #:" & Err.Number & Err.Description
            'ShutDownApplication strError
         Else
            Sleep 1000   'pause for 1 second and retry
         End If
      Else
         Exit Do
      End If
   Loop
   
   Set PostMessageNew = http
End Sub
