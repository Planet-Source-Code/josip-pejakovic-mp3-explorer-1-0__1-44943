Attribute VB_Name = "modMP3Header"
Type HeaderInfo
    Layer As String
    Frequency As String
    Bitrate As String
    Mode As String
    MpegVersion As String
    Emphasis As String
    FPlayTime As String
   mFileSize As String
   sekunde As String
End Type

Public MP3HeaderInfo As HeaderInfo

Public Function ReadMP3Header(sPassFileName As String)
On Error Resume Next
Dim z, i
Dim BinaryString As String
Dim byteArray(4) As Byte
Dim bin As String
Dim BinString As String
Dim DecString As Integer

Close #3
Open sPassFileName For Binary As #3
   For z = 1 To 4
   Get #3, z, byteArray(z)
   Next z
 Close #3
 bin = ""
   For z = 1 To 4
     For i = 0 To 7 Step 1
         If byteArray(z) And (2 ^ i) Then
            bin = bin + "1"
            Else
            bin = bin + "0"
         End If
         Next i
Next z
BinaryString = bin

DecString = 0
BinString = Mid(bin, 19, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Frequency = 44100
Case 1
MP3HeaderInfo.Frequency = 32000
Case 2
MP3HeaderInfo.Frequency = 48000
Case 3
End Select

DecString = 0
BinString = Mid(bin, 10, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Layer = ""
Case 1
MP3HeaderInfo.Layer = 2
Case 2
MP3HeaderInfo.Layer = 3
Case 3
MP3HeaderInfo.Layer = 1
End Select

DecString = 0
BinString = Mid(bin, 31, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Mode = "Stereo"
Case 1
MP3HeaderInfo.Mode = "Dual Channel"
Case 2
MP3HeaderInfo.Mode = "Joint stereo"
Case 3
MP3HeaderInfo.Mode = "Mono"
End Select

If Mid(bin, 12, 1) = 0 Then
MP3HeaderInfo.MpegVersion = 2
Else
MP3HeaderInfo.MpegVersion = 1
End If

DecString = 0
BinString = Mid(bin, 21, 4)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Bitrate = 0
Case 1
MP3HeaderInfo.Bitrate = 112
Case 2
MP3HeaderInfo.Bitrate = 56
Case 3
MP3HeaderInfo.Bitrate = 224
Case 4
MP3HeaderInfo.Bitrate = 40
Case 5
MP3HeaderInfo.Bitrate = 160
Case 6
MP3HeaderInfo.Bitrate = 80
Case 7
MP3HeaderInfo.Bitrate = 320
Case 8
MP3HeaderInfo.Bitrate = 32
Case 9
MP3HeaderInfo.Bitrate = 128
Case 10
MP3HeaderInfo.Bitrate = 64
Case 11
MP3HeaderInfo.Bitrate = 256
Case 12
MP3HeaderInfo.Bitrate = 48
Case 13
MP3HeaderInfo.Bitrate = 192
Case 14
MP3HeaderInfo.Bitrate = 96
Case 15
MP3HeaderInfo.Bitrate = 0
If MP3HeaderInfo.Layer = 1 Then
    Select Case DecString
    Case 0
MP3HeaderInfo.Bitrate = 0
    Case 1
  MP3HeaderInfo.Bitrate = 128
    Case 2
   MP3HeaderInfo.Bitrate = 64
    Case 3
MP3HeaderInfo.Bitrate = 256
    Case 4
MP3HeaderInfo.Bitrate = 48
    Case 5
MP3HeaderInfo.Bitrate = 192
    Case 6
MP3HeaderInfo.Bitrate = 96
    Case 7
    MP3HeaderInfo.Bitrate = 384
    Case 8
MP3HeaderInfo.Bitrate = 32
    Case 9
MP3HeaderInfo.Bitrate = 160
    Case 10
    MP3HeaderInfo.Bitrate = 80
    Case 11
MP3HeaderInfo.Bitrate = 320
    Case 12
MP3HeaderInfo.Bitrate = 56
    Case 13
MP3HeaderInfo.Bitrate = 224
    Case 14
  MP3HeaderInfo.Bitrate = 112
    Case 15
MP3HeaderInfo.Bitrate = 0
End Select
End If
End Select

DecString = 0
BinString = Mid(bin, 25, 2)
For i = 1 To Len(BinString)
If Mid(BinString, i, 1) = 1 Then
DecString = DecString + 2 ^ (Len(BinString) - i)
End If
Next i
Select Case DecString
Case 0
MP3HeaderInfo.Emphasis = "No"
Case 1
MP3HeaderInfo.Emphasis = "-?-"
Case 2
MP3HeaderInfo.Emphasis = "50/15"
Case 3
MP3HeaderInfo.Emphasis = "CITT j. 17"
End Select

With MP3HeaderInfo
    Dim min, sec
    .Bitrate = Int(.Bitrate)
    .mFileSize = FileSizeMP3(sPassFileName)
    .FPlayTime = ((.mFileSize * 8) / (.Bitrate * 1000))
    min = .FPlayTime \ 60
    sec = .FPlayTime - (min * 60)
    .sekunde = .FPlayTime
    .FPlayTime = Format(min, "#0#") & ":" & Format(sec, "0#")
End With

End Function

Public Function FileSizeMP3(file As String) As String
    Dim LSize As String
    If file = "" Then
    FileSizeMP3 = ""
    Exit Function
    End If
    LSize = FileLen(file)
    FileSizeMP3 = LSize
End Function
