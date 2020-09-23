Attribute VB_Name = "MP3Play"
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Function MP3Play(ByVal Mp3File As String)

'ovdje je sva mudrolija... kak je to bilo jednostavno :)))))

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
  
path = GetShortPath(Mp3File)
  
cmdToDo = "open " & path & " type MPEGVideo Alias MP3Play"
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If dwReturn <> 0 Then
   mciGetErrorString dwReturn, ret, 128
   mmOpen = ret
MsgBox ret, vbCritical
Exit Function
End If
mciSendString "play MP3Play", 0, 0, 0
End Function

Sub MP3Stop()
mciSendString "stop MP3Play", 0, 0, 0
mciSendString "close MP3Play", 0, 0, 0
End Sub

Sub PauseMp3()
mciSendString "pause MP3Play", 0, 0, 0
End Sub

Sub StopMp3()
mciSendString "stop MP3Play", 0, 0, 0
mciSendString "close MP3Play", 0, 0, 0
End Sub

Function GetShortPath(strFileName As String) As String
Dim lngRes As Long, strPath As String
strPath = String$(165, 0)
lngRes = GetShortPathName(strFileName, strPath, 164)
GetShortPath = Left$(strPath, lngRes)
End Function

Sub Back()
mciSendString "stop MP3Play", 0, 0, 0
mciSendString "play MP3Play from 0", 0, 0, 0
End Sub

Sub UnPauseMp3()
mciSendString "play MP3Play", 0, 0, 0
End Sub

Function IsPlaying() As Boolean
Static s As String * 30
mciSendString "status MP3Play mode", s, Len(s), 0
IsPlaying = (Mid$(s, 1, 7) = "playing")
End Function

