Attribute VB_Name = "GeneralMod"
Public Const APPPATH = "<APPPATH>"
Public Const SYSPATH = "<SYSPATH>"
Public Const S32PATH = "<S32PATH>"
Public Const WINPATH = "<WINPATH>"
Public Const WOSPATH = "<WOSPATH>"
Private Type SysInfo
    SysTempPath As String
    SysWinPath As String
    CompName As String
    UserName As String
    SysOSName As String
    SysDrive As String
End Type
Private Type ServerType
  FullPath As String
  ServerName As String
End Type
Private Type UpdateInfo
    Servers() As ServerType
    FastestServer As Long
    FileToGet As String
    FileToSave As String
    Destination As String
    FileVer As String
End Type

Public EnvironmentVar As SysInfo
Public FilesToUdp() As UpdateInfo

Public UpdateFileName As String

Public Sub StoreEnvVarInfo()
Dim I, nPos, StrLen As Integer
Dim First, StrLine As String

    For I = 1 To 256
        nPos = InStr(Environ(I), "=")
        If nPos > 0 Then
            StrLen = Len(Environ(I))
            StrLine = Environ(I)
            First = Left(StrLine, nPos - 1)
            
            Select Case First
                Case "TEMP": EnvironmentVar.SysTempPath = Mid(StrLine, nPos + 1, StrLen)
                Case "windir": EnvironmentVar.SysWinPath = Mid(StrLine, nPos + 1, StrLen)
                Case "COMPUTERNAME": EnvironmentVar.CompName = Mid(StrLine, nPos + 1, StrLen)
                Case "USERNAME": EnvironmentVar.UserName = Mid(StrLine, nPos + 1, StrLen)
                Case "OS": EnvironmentVar.SysOSName = Mid(StrLine, nPos + 1, StrLen)
                Case "SystemDrive": EnvironmentVar.SysDrive = Mid(StrLine, nPos + 1, StrLen)
            End Select
        End If
        
        If Len(Environ(I)) <= 0 Then Exit For
    Next
    First = "": StrLine = ""
    StrLen = 0: nPos = 0: I = 0
    
End Sub
Public Sub LoadUpdateFile()
On Error GoTo UdpErr
Dim Filep1 As Long: Filep1 = FreeFile
Dim NBFiles: NBFiles = -1
Dim NBServer: NBSever = -1
Dim Pos As Long

Open UpdateFileName For Input As Filep1

Do
    
    Line Input #Filep1, strtemp
    
    'If the line is not blank
    If Len(strtemp) <> 0 Then
      Pos = 1
      NBFiles = NBFiles + 1
      ReDim Preserve FilesToUdp(NBFiles)
      NBServer = -1
      Do
        'Look for open and close statement
          rep = InStr(Pos, strtemp, "<")
          rep3 = InStr(Pos, strtemp, "<s-")
          rep4 = InStr(Pos, strtemp, "<v-")
          rep2 = InStr(Pos, strtemp, ">")
          If rep = 0 Or rep2 = 0 Then Exit Do
               
           
               fieldtemp = Mid(strtemp, rep + 1, rep2 - rep - 1)
          
               If rep4 <> 0 Then
                  FilesToUdp(NBFiles).FileVer = Mid(fieldtemp, 3)
               ElseIf rep3 <> 0 Then
                 NBServer = NBServer + 1
                 ReDim Preserve FilesToUdp(NBFiles).Servers(NBServer)
                 FilesToUdp(NBFiles).Servers(NBServer).FullPath = Mid(fieldtemp, 3)
                 FilesToUdp(NBFiles).Servers(NBServer).ServerName = GetSNameFromURL(FilesToUdp(NBFiles).Servers(NBServer).FullPath)
                 FilesToUdp(NBFiles).FileToGet = GetFName(FilesToUdp(NBFiles).Servers(NBServer).FullPath, "/")
                 
                 
               Else
                 
                 rep = InStr(1, fieldtemp, "-")
                 If rep <> 0 Then
                     FilesToUdp(NBFiles).FileToSave = GetFNameFromDest(fieldtemp)
                     FilesToUdp(NBFiles).Destination = GetRealDest1(fieldtemp)
                 Else
                    FilesToUdp(NBFiles).FileToSave = GetFName(fieldtemp, "\")
                    FilesToUdp(NBFiles).Destination = GetRealDest2(fieldtemp)
                 End If
               
               End If
          
          Pos = rep2 + 1
       Loop
     End If
Loop Until EOF(Filep1)
Close
Exit Sub
UdpErr:

End Sub
Private Function GetFName(url, Del)
Pos = 1
Do
 rep = InStr(Pos, url, Del)
  If rep = 0 Then Exit Do
 Pos = rep + 1
Loop

GetFName = Mid(url, Pos)
End Function


Private Function GetFNameFromDest(url) As String
pathtemp = Mid(url, 10)
GetFNameFromDest = pathtemp
End Function
Private Function GetRealDest1(url) As String
pathtemp = "<" & Mid(url, 1, 7) & ">"
Select Case pathtemp
  Case APPPATH: pathtemp = App.Path & "\"
  Case SYSPATH: pathtemp = EnvironmentVar.SysWinPath & "\" & "system" & "\"
  Case S32PATH: pathtemp = EnvironmentVar.SysWinPath & "\" & "system32" & "\"
  Case WINPATH: pathtemp = EnvironmentVar.SysWinPath & "\"
  Case WOSPATH: pathtemp = EnvironmentVar.SysDrive & "\"
End Select
GetRealDest1 = pathtemp
End Function
Private Function GetRealDest2(url)
Pos = 1
Do
 rep = InStr(Pos, url, "\")
  If rep = 0 Then Exit Do
 Pos = rep + 1
Loop
GetRealDest2 = Mid(url, 1, Pos - 1)
End Function
Private Function GetSNameFromURL(urlt)
Pos = InStr(1, urlt, "//")
rep = InStr(Pos + 2, urlt, "/")
GetSNameFromURL = Mid(urlt, 1, rep - 1)
End Function
Public Function EnsurePath(pathtemp) As String
EnsurePath = IIf(Right(pathtemp, Len(pathtemp) - 1) <> "\", "\", "")
End Function

