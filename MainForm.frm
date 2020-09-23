VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   Caption         =   "Easy updates"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "File list"
      Height          =   1785
      Left            =   105
      TabIndex        =   10
      Top             =   1110
      Width           =   9960
      Begin MSComctlLib.ListView FileList 
         Height          =   1455
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Files"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Server / Status"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Destination"
            Object.Width           =   7832
         EndProperty
      End
   End
   Begin VB.CommandButton cmdGET 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   135
      Picture         =   "MainForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar AutoUpdST 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   5340
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13124
            MinWidth        =   13124
            Text            =   "Ready."
            TextSave        =   "Ready."
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10380
      Top             =   2115
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download Status"
      Height          =   1875
      Left            =   75
      TabIndex        =   0
      Top             =   3435
      Width           =   9975
      Begin VB.Timer TimerTimeOut 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   9345
         Top             =   315
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "Pourcentage accomplie"
         Top             =   1185
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   1
         Scrolling       =   1
      End
      Begin VB.Label LabelFName 
         Caption         =   "File name:"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   465
         Width           =   9060
      End
      Begin VB.Label LabelDest 
         Caption         =   "Destination:"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   9690
      End
      Begin VB.Label Label1 
         Caption         =   "File Size:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LabelFS 
         Caption         =   "0 Kb"
         Height          =   225
         Left            =   1305
         TabIndex        =   6
         Top             =   270
         Width           =   4065
      End
      Begin VB.Label Label3 
         Caption         =   "Estimated time : "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   930
         Width           =   7815
      End
      Begin VB.Label LabelST 
         Caption         =   "0 Kb of 0Kb"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   1545
         Width           =   6930
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10395
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.harveysolution.t2u.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label LabelStatus1 
      Height          =   525
      Left            =   1650
      TabIndex        =   13
      Top             =   525
      Width           =   8340
   End
   Begin VB.Label LabelTitle 
      Caption         =   "Program updates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1665
      TabIndex        =   12
      Top             =   180
      Width           =   8325
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*  This application program has been created by Carl Harvey
'*  This code is copyrighted and may be utilized for personal purposes ONLY
'*  Created: 2002-18-05
'*  Last Modif: 2002-08-07
'*
'*
'*  http://www.harveysolution.t2u.com
'******************************************************

' TODO :  Error handling, Validation, Logs
Dim MyFileNameDestination As String
Dim ASec As Date
Dim EstimLeft As Long
Dim PingTime As String
Dim LookFastest As Boolean
Dim CurrFastestPing As Long
Dim CurrentlyLooking As Boolean
Dim CurrentFileLook As Long
Dim CurrentServerLook As Long
Dim MTitle As String

Private Sub cmdGET_Click()
StartDownload
LabelST = "Update finnished."
End Sub


Private Sub ShowUpdFiles()
For I = 0 To UBound(FilesToUdp)
  FileList.ListItems.Add , , FilesToUdp(I).FileToGet
  FileList.ListItems(I + 1).SubItems(1) = FilesToUdp(I).FileVer
  FileList.ListItems(I + 1).SubItems(2) = "Available servers (" & UBound(FilesToUdp(I).Servers) + 1 & ")"
  FileList.ListItems(I + 1).SubItems(3) = FilesToUdp(I).Destination & FilesToUdp(I).FileToSave
Next
End Sub

Private Sub Form_Initialize()
strArg = Command()
CheckArg strArg
End Sub


Private Sub Form_Load()
UpdateFileName = App.Path & "\MTUpdates.upd"
LabelTitle.Caption = "Program updates"
ChangeOBJColor ProgressBar1.hwnd, 240, 240, 255, 32, 43, 136, PB_BGCOLOR, PB_FGCOLOR
CurrFastestPing = 0
StoreEnvVarInfo
'Load the files to be updated from the UpdateFileName
LoadUpdateFile
'Show the files to update
ShowUpdFiles
End Sub
Public Function GetFtime(Count) As String
Dim Sec1 As Long, Sec2 As Long, MSec1 As Long, MSec2 As Long

rep1 = InStr(1, PingTime, ":")
rep2 = InStr(1, Count, ":")

Sec1 = Mid(PingTime, 1, rep1 - 1)
Sec2 = Mid(Count, 1, rep2 - 1)
MSec1 = Mid(PingTime, rep1 + 1)
MSec2 = Mid(Count, rep2 + 1)

If Sec1 <> Sec2 Then
  Sec1 = Sec2 - Sec1
Else
  Sec1 = 0
End If

If MSec1 <> MSec2 Then
  MSec1 = MSec2 - MSec1 + (Sec1 * 1000)
Else
  MSec1 = 0
End If
GetFtime = MSec1
End Function
Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
    Case icNone: AutoUpdST.Panels(1).Text = "No state to report. 1"
                   
    Case icResolvingHost:            AutoUpdST.Panels(1).Text = "The control is looking up the IP address of the specified host computer.2 "
                                     FileList.ListItems(CurrentFileLook).SubItems(2) = "Serveur non disponible."
    Case icHostResolved:             AutoUpdST.Panels(1).Text = " The control successfully found the IP address of the specified host computer.3"
    Case icConnecting:               AutoUpdST.Panels(1).Text = " The control is connecting to the host computer.4"
    Case icConnected:                AutoUpdST.Panels(1).Text = " The control successfully connected to the host computer.5"
    Case icRequesting:               AutoUpdST.Panels(1).Text = " The control is sending a request to the host computer.6"
    Case icRequestSent:              AutoUpdST.Panels(1).Text = " The control successfully sent the request.7"
    Case icReceivingResponse:        AutoUpdST.Panels(1).Text = " The control is receiving a response from the host computer.8"
    Case icResponseReceived:         AutoUpdST.Panels(1).Text = " The control successfully received a response from the host computer.9"
                                     If LookFastest Then SetGotResponse
                              
    Case icDisconnecting:            AutoUpdST.Panels(1).Text = " The control is disconnecting from the host computer.10"
    Case icDisconnected:             AutoUpdST.Panels(1).Text = " The control successfully disconnected from the host computer.11"
    Case icError:                    AutoUpdST.Panels(1).Text = " An error occurred in communicating with the host computer.12"
    Case icResponseCompleted
                                     AutoUpdST.Panels(1).Text = "Download..."
                                     Timer1.Enabled = True
                                     RetreiveFile Inet1, MyFileNameDestination, ProgressBar1, LabelFS, LabelST
                                     AutoUpdST.Panels(1).Text = "Download completed successfully. File saved to " & MyFileNameDestination
                                     Timer1.Enabled = False
                             
End Select
End Sub

Private Sub Timer1_Timer()
Label3.Caption = "Estimated time : " & Round(ProgressBar1.Max / BytesRead) & " sec"
EstimLeft = (ProgressBar1.Max - ProgressBar1.Value) / BytesRead
CurrentStatus = "Estimated time left " & EstimLeft & " sec "
BytesRead = 0
End Sub

Private Sub StartDownload()
LookFastest = False


For I = 0 To UBound(FilesToUdp)
  
  'This for/next will check for the fastest server
  For i2 = 0 To UBound(FilesToUdp(I).Servers)
  
    LabelStatus1.Caption = "Looking for the fastest server..." & vbCrLf & FilesToUdp(I).Servers(i2).ServerName
 
    FileList.ListItems(I + 1).Selected = True
  
    CurrentFileLook = I + 1
    CurrentServerLook = i2
  
    PingTime = GetMillime(GetTickCount())
    Inet1.OpenURL FilesToUdp(I).Servers(i2).ServerName
    Inet1.RequestTimeout = 40
    
    Do
       DoEvents
    Loop While Inet1.StillExecuting
  
  Next
    ProgressBar1.Value = 0
    LabelStatus1.Caption = "Downloading file..." & vbCrLf & FilesToUdp(I).Servers(FilesToUdp(I).FastestServer).FullPath
    LabelFName = "File name: " & FilesToUdp(I).FileToGet
    LabelDest = "Destination: " & FilesToUdp(I).Destination
    FileList.ListItems(I + 1).Selected = True
    MyFileNameDestination = FilesToUdp(I).Destination & FilesToUdp(I).FileToSave
    Inet1.Protocol = icHTTP
    TimerTimeOut.Enabled = True
    Inet1.Execute FilesToUdp(I).Servers(FilesToUdp(I).FastestServer).FullPath, "GET"
    Do
      DoEvents
    Loop While Inet1.StillExecuting
    
    TimerTimeOut.Enabled = False
    SetPercent
Next

LookFastest = False

End Sub

Private Sub SetPercent()
 If ProgressBar1.Value = 0 Then
      tmp = 0
 Else
      tmp = Format((ProgressBar1.Value / ProgressBar1.Max) * 100, "#")
 End If
 
 FileList.ListItems(CurrentFileLook).SubItems(2) = tmp & "%"
End Sub

'Here we got a response so check the time that we got this response
Private Sub SetGotResponse()
  pingnow = GetFtime(GetMillime(GetTickCount()))
  If CurrFastestPing > pingnow Or CurrFastestPing = 0 Then
     CurrFastestPing = pingnow
     FilesToUdp(CurrentFileLook - 1).FastestServer = CurrentServerLook
  End If
  FileList.ListItems(CurrentFileLook).SubItems(2) = FilesToUdp(CurrentFileLook - 1).Servers(FilesToUdp(CurrentFileLook - 1).FastestServer).ServerName
End Sub

'Validate the parameter from here
Private Sub CheckArg(ArgStr)
If ArgStr = "" Then GoTo ErrArg
LabelTitle.Caption = "Autoupdate"
ArgStr = Replace(ArgStr, Chr(34), "")
rep = InStr(1, ArgStr, ";;")
rep1 = InStr(rep + 2, ArgStr, ";;")
ErrArg:
MsgBox "TO DO." & vbCrLf & "Invalid parameter !" & vbCrLf & "Parse the parameter from here !", vbCritical + vbOKOnly, "Error"
'end
End Sub

Private Sub TimerTimeOut_Timer()
SetPercent
End Sub
