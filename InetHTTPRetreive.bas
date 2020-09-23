Attribute VB_Name = "InetRetreive"
'*******************************************************
'*  This module has been created by Carl Harvey
'*  This code is copyrighted and may be utilized for personal purposes ONLY
'*  Created: 2002-18-05
'*  Last Modif: 2002-18-05
'*      -----------     TIPS     --------------
'*    'To show the real file name you must remove all before the last '/'
'*    'TextShowRealFileName = ParseName(Inet1.Document)
'******************************************************

' TODO :  Error handling, Validation, Logs
Public CurrentStatus As String
Public Const ChunkSize = 1024
Public BytesRead As Long
Public Sub RetreiveFile(Inet1 As Inet, FileDestination As String, StatusBar1 As ProgressBar, LabelFileSize As Label, ByRef LabelStatus As Label)

    Dim FileSize As Long
    Dim FileSizeKB As Long
    Dim FileData() As Byte
    Dim DComplete As Boolean: DComplete = False
    Dim Filep1 As Long
    
    'Get the file size from the file header
    FileSize = Inet1.GetHeader("Content-Length")
    
    FileSizeMB = Round(FileSize / 1000000, 2)
    
    'Set Labels to show up info
    LabelFileSize.Caption = FileSizeMB & " Mb"
    LabelStatus.Caption = "0 bytes" & " of " & FileSize & " bytes"
      
    'Set the progress bar
    StatusBar1.Max = FileSize
    StatusBar1.Value = 0
    
    BytesRead = 0
    
    ' Get first chunk.
    FileData() = Inet1.GetChunk(ChunkSize, 1)
    BytesRead = 1024
    'Assign freefile number
    Filep1 = FreeFile
    
    ' Open binary file to save
    Open FileDestination For Binary Access Write As #Filep1
      
    ' loop until we have retreive all file bytes , by chunk of 1024 bytes
    Do While Not DComplete
      
        'write to the file
        Put #Filep1, , FileData()
        
        'update progress bar and status label
        'the iif() is in case the file is less than 1024 bytes, oups !
        StatusBar1.Value = IIf(StatusBar1.Value + ChunkSize > FileSize, FileSize, StatusBar1.Value + ChunkSize)
        LabelStatus.Caption = CurrentStatus & "(" & Round(StatusBar1.Value / 1000000, 2) & " MB" & " of " & FileSizeMB & " MB)"
        
        ' Get other chunks
        FileData() = Inet1.GetChunk(ChunkSize, 1)
        BytesRead = BytesRead + 1024
        ' Let window do its things
        DoEvents
         
        'If the FileData is empty, we are done !
        If UBound(FileData()) = -1 Then DComplete = True
         
      Loop
      
      'Close the file
      Close #Filep1
      'cancel any left over proccess (just in case)
      'Inet1.Cancel
End Sub

