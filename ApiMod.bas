Attribute VB_Name = "ApiMod"
'To send a message to any window
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function GetTickCount Lib "Kernel32" () As Long


'Got this part from PSC
Public Const PB_BGCOLOR = &H400 + 9
Public Const PB_FGCOLOR = &H2000 + 1
'End of PSC copyright

Public Sub ChangeOBJColor(OBJ As Long, BGR As Integer, BGG As Integer, BGB As Integer, FGR As Integer, FGG As Integer, FGB As Integer, FG_Message, BG_Message)
    SendMessage OBJ, BG_Message, 0, ByVal RGB(BGR, BGG, BGB)
    SendMessage OBJ, FG_Message, 0, ByVal RGB(FGR, FGG, FGB)
End Sub

Public Function GetMillime(Count) As String

    Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
    
    Miliseconds = Count Mod 1000
    Count = Count \ 1000
    Minutes = Count \ 60
    Seconds = Count Mod 60

    GetMillime = Seconds & ":" & Miliseconds
End Function
