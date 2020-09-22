VERSION 5.00
Begin VB.Form frmRetrieve 
   Caption         =   "Domain to IP Application"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNav 
      Caption         =   "Navigate"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox chkNav 
      Caption         =   "Navigate when found."
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox lstResults 
      Height          =   1425
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "Automatically copy  to clipboard if found."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "www."
      Top             =   360
      Width           =   3375
   End
   Begin VB.ComboBox cboQuery 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "&Copy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopySub 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "frmRetrieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application was written by Daniel M.
'Enjoy!

Option Explicit

'Declarations
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Constants
Private Const DNS_DESCRIPTION   As Long = 1 'Retrieve Description
Private Const DNS_ADDRESS       As Long = 2 'Retrieve Address

Private Const DNS_FOUNDSINGLE   As Long = 7 'Found Single Address
Private Const DNS_FOUNDMULTIPLE As Long = 9 'Found Multiple Addresses

Private Const sServerINI        As String = "servers.ini"       'Server names
Private Const sTempPath         As String = "C:\dns_lookup.tmp" 'Temporary Path

Private Function IsConsoleOpen() As Boolean
'Checks if command prompt is running or not.

Dim sWindow     As String
Dim retVal      As Long

retVal = FindWindow("ConsoleWindowClass", sWindow)

If retVal Then
    IsConsoleOpen = True
Else
    IsConsoleOpen = False
End If
End Function

Private Sub cmdQuery_Click()
lstResults.Clear
Dim sDomain As String
sDomain = Trim(txtDomain.Text)
Shell "cmd /c nslookup " & sDomain & " " & DNSServer(cboQuery, DNS_ADDRESS) & " > " & sTempPath
Do
    Sleep 50
Loop Until IsConsoleOpen = False

Dim FileNum         As Integer

Dim sTemp           As String  'Holds temporary data
Dim sPrevious       As String  'Holds previous temp
Dim sAllAddresses   As String  'Holds addresses if multiple

Dim bolMultiple     As Boolean 'Determines if multiple addresses

FileNum = FreeFile
Open sTempPath For Input As #FileNum
    Do While Not EOF(FileNum)
        Input #FileNum, sTemp
        If InStr(sPrevious, "Server") Then
            'skip this!
        Else
            If Left(sTemp, 9) = "Addresses" Then 'Found multiple addresses!
                bolMultiple = True
                sAllAddresses = sTemp
            ElseIf Left(sTemp, 7) = "Address" Then 'Found single address!
                ProcessResults lstResults, sTemp, DNS_FOUNDSINGLE 'Process single address
            ElseIf bolMultiple = True And InStr(sTemp, "Alias") = 0 Then
                sAllAddresses = sAllAddresses & "-" & sTemp 'Retrieve more addresses
            End If
        End If
        sPrevious = sTemp
    Loop
Close #FileNum

If bolMultiple = True Then 'Process results for multiple addresses
    sAllAddresses = Left(sAllAddresses, Len(sAllAddresses) - 1)
    ProcessResults lstResults, sAllAddresses, DNS_FOUNDMULTIPLE
End If

End Sub
Private Function ProcessResults(ByRef ResultBox As ListBox, ByRef sResult As String, _
                                ByVal dwFlags As Long)

sResult = Trim(sResult)
If dwFlags = DNS_FOUNDSINGLE Then 'If found single then
        ResultBox.AddItem Right(sResult, Len(sResult) - 10) 'Add the result
        
        If chkNav Then 'If checked, navigate to page.
            Shell "explorer http://" & Right(sResult, Len(sResult) - 10)
        End If
        
        If chkCopy Then 'If checked, copy to clipboard.
            Clipboard.Clear
            Clipboard.SetText Right(sResult, Len(sResult) - 10)
        End If
ElseIf dwFlags = DNS_FOUNDMULTIPLE Then 'If found multiple then
    
    Dim rCount      As Long
    Dim sReplace    As String
    Dim i           As Long
    
    rCount = Len(sResult) - Len(Replace(sResult, "-", vbNullString)) 'Get result count
    
    sResult = Right(sResult, Len(sResult) - 11) 'Parse results
    sResult = Trim(sResult)
    
    If chkNav Then 'If checked, navigate to first result.
        Shell "explorer http://" & Split(sResult, "-", -1, 1)(0)
    End If
    
    If chkCopy Then 'If checked, copy to clipboard first result.
        Clipboard.Clear
        Clipboard.SetText Split(sResult, "-", -1, 1)(0)
    End If
    
    For i = 0 To rCount 'Add results to list.
        ResultBox.AddItem Split(sResult, "-", -1, 1)(i)
    Next i
    
End If
End Function

Private Function DNSServer(cboBox As ComboBox, Request As Integer) As String
Dim sServer As String
sServer = cboBox.Text
If Request = DNS_DESCRIPTION Then 'Parse for description
    DNSServer = Left(sServer, InStr(sServer, " ") - 1)
ElseIf Request = DNS_ADDRESS Then 'Parse for address
    sServer = StrReverse(sServer)
    sServer = Left(sServer, InStr(sServer, " ") - 1)
    sServer = StrReverse(sServer)
    DNSServer = sServer
End If

End Function


Private Sub cmdNav_Click()
Shell "explorer " & lstResults.Text
End Sub

Private Sub Form_Load()
If RetrieveServers(cboQuery, App.Path & "\" & sServerINI) = False Then
    MsgBox "DNS Servers could not be retrieved!"
End If

End Sub

Private Function RetrieveServers(cboBox As ComboBox, sPath As String) As Boolean
On Error GoTo Catch

Dim FileNum As Integer
Dim sTemp   As String

FileNum = FreeFile
    Open sPath For Input As #FileNum
        Do While Not EOF(FileNum)
            Input #FileNum, sTemp
            If Left(sTemp, 1) <> "[" Then cboBox.AddItem sTemp
        Loop
    Close #FileNum

RetrieveServers = True

    Exit Function
Catch:
'File not found argggh!!
End Function

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuCopy
End If
End Sub

Private Sub mnuCopySub_Click()
Clipboard.Clear
Clipboard.SetText lstResults.Text
End Sub
