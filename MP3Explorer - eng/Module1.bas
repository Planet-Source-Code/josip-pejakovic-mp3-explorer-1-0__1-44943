Attribute VB_Name = "Module1"
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1


Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const PROGRESS_CANCEL = 1
Public Const PROGRESS_CONTINUE = 0
Public Const PROGRESS_QUIET = 3
Public Const PROGRESS_STOP = 2
Public Const COPY_FILE_FAIL_IF_EXISTS = &H1
Public Const COPY_FILE_RESTARTABLE = &H2

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Const GENERIC_WRITE = &H40000000
Const GENERIC_READ = &H80000000
Const OPEN_EXISTING = 3
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FO_DELETE = &H3
Const FOF_NOCONFIRMATION = &H10

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Type id3
izvodac As String
Album As String
godina As String
komentar As String
velicinafajla As Long
End Type

Public artistv2 As String
Public titlev2 As String
Public albumv2 As String
Public godinav2 As Integer
Public komentarv2 As String

Dim mp3i As id3
Public bleh As ListItem
Public b
Public putanja As String
Public trenutni_dir As String
Public fajlovi()

Public SHFileOp As SHFILEOPSTRUCT

Public prekini_proces As Boolean
Public sFile As String 'ime filea
Public sPath As String 'ime filea+PATH
Public sFilter As String

Public bojaslova
Public bojapozadine

Public odabrana_datoteka As String
Public odabrana_datoteka_putanja As String

Public editirano As Boolean
Public otvorena_lista As Boolean

Public velicina As Double

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Function FindFilesAPI(path As String, SearchStr As String, FileCount As Long, DirCount As Long, subdir As Boolean)

If prekini_proces = True Then
Exit Function
End If


Dim filename As String                          ' Walking filename variable...
Dim DirName As String                           ' SubDirectory Name
Dim dirNames() As String                        ' Buffer for directory name entries
Dim nDir As Long                                    ' Number of directories in this path
Dim i As Long                                         ' For-loop counter...
Dim hSearch As Long                              ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Long

If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DoEvents
DirName = StripNulls(WFD.cFileName)
'Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
'Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)                   'Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
filename = StripNulls(WFD.cFileName)
Form1.Label5 = "Searching..... " & path

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
art = ""
album1 = ""
naziv_pjesme = ""
titlev2 = ""
albumv2 = ""

naziv_pjesme = GetTag(path & filename)     'Za ispisivanje TAGOVA iz MP3 fajla
naziv_pjesmev2 = GetTagID3v2(path & filename)

art = nas
album1 = alb
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If (filename <> ".") And (filename <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1

'/////// Uèitavanje u ListView /////////////////////////////////////////
Set bleh = Form1.lv.ListItems.Add(, , filename)
bleh.SubItems(1) = path & filename

vel (path & filename)

If naziv_pjesme = "" Then
bleh.SubItems(2) = titlev2
ElseIf titlev2 = "" Then
bleh.SubItems(2) = naziv_pjesme
ElseIf naziv_pjesme <> "" And titlev2 <> "" Then
bleh.SubItems(2) = titlev2
End If

If naziv_pjesme = "" And titlev2 = "" Then
bleh.SubItems(2) = Left(filename, Len(filename) - 4)
End If

naziv_pjesme = "": titlev2 = ""

If art = "" Then
bleh.SubItems(3) = artistv2
ElseIf artistv2 = "" Then
bleh.SubItems(3) = art
ElseIf art <> "" And artistv2 <> "" Then
bleh.SubItems(3) = artistv2
End If
art = "": artistv2 = ""

If album1 = "" Then
bleh.SubItems(4) = albumv2
ElseIf albumv2 = "" Then
bleh.SubItems(4) = album1
ElseIf album1 <> "" And albumv2 <> "" Then
bleh.SubItems(4) = albumv2
End If
album1 = "": albumv2 = ""


modMP3Header.ReadMP3Header (path & filename)
bleh.SubItems(5) = modMP3Header.MP3HeaderInfo.FPlayTime
bleh.SubItems(6) = Format(velicina / 1048576, "#0.000")
filename = ""

'bleh.SubItems(4) = bleh.SubItems(4) & album1
'////////////////////////////////////////////////////////////////////////////////

End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
'If there are sub-directories...
If subdir = False Then
Exit Function
Else
If nDir > 0 Then
'Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", SearchStr, FileCount, DirCount, subdir)
Next i
End If
End If
End Function

Function GetTag(file)
On Error Resume Next
Dim Buf As String * 128
Dim tmpStr As String
Dim i As Byte

Mp3File = file
mp3size = FileLen(Mp3File)
Buf = ""
file = ""
Open Mp3File For Binary As #1
Get #1, mp3size - 127, Buf
If Format(Left(Buf, 3), "<") <> "tag" Then
mp3i.izvodac = ""
mp3i.Album = ""
Else
naslov = Trim(Mid(Buf, 4, 30))
mp3i.izvodac = Trim(Mid(Buf, 34, 30))
mp3i.Album = Trim(Mid(Buf, 64, 30))
mp3i.godina = Trim(Mid(Buf, 94, 4))
mp3i.komentar = Trim(Mid(Buf, 98, 30))
End If
Close #1
If naslov = "" Then
naslov = ""
End If
Close #1
GetTag = naslov
file = ""
End Function
Function nas()
nas = ""
nas = mp3i.izvodac
End Function
Function vel(nazivfajla)
Dim horgfile As Long
horgfile = 0
horgfile = CreateFile(nazivfajla, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
velicina = GetFileSize(horgfile, 0)
CloseHandle horgfile
horgfile = 0
End Function
Function alb()
alb = ""
alb = mp3i.Album
End Function
Function god()
god = ""
god = mp3i.godina
End Function

Function komentari()
komentari = ""
komentari = mp3i.komentar
End Function

Public Function Export(ByVal fajl As String)
Open "c:\test.txt" For Output As #1
For x = 1 To Form1.lv.ListItems.Count
If Form1.lv.ListItems(x).SubItems(2) = "" Then
Print #1, Form1.lv.ListItems(x) & " - " & Form1.lv.ListItems(x).SubItems(1)
Else
Print #1, Form1.lv.ListItems(x).SubItems(2) & " - " & Form1.lv.ListItems(x).SubItems(3)
End If
Next x
Close #1
End Function

Function Export_Path_Filename(ByVal fajl As String, lvc As Control)
Open fajl For Output As #1
For epf = 1 To lvc.ListItems.Count
Print #1, lvc.ListItems(epf).SubItems(1) & " - " & lvc.ListItems(epf)
Next epf
Close #1

End Function

Function Export_Song_Title(ByVal fajl As String, lvc As Control)
Open fajl For Output As #1
For st = 1 To lvc.ListItems.Count
Print #1, lvc.ListItems(st).SubItems(2)
Next st
Close #1
End Function

Function Export_SongTitle_Artist(ByVal fajl As String, lvc As Control)


Open fajl For Output As #1
For sta = 1 To lvc.ListItems.Count

If lvc.ListItems(sta).SubItems(3) <> "" And lvc.ListItems(sta).SubItems(2) <> "" Then
Print #1, lvc.ListItems(sta).SubItems(3) & " - " & lvc.ListItems(sta).SubItems(2)
End If

If lvc.ListItems(sta).SubItems(3) = "" Or lvc.ListItems(sta).SubItems(2) = "" Then
Print #1, lvc.ListItems(sta)
End If
Next sta
Close #1

End Function

Function Export_SongTitle_Artist_Album(fajl As String, lvc As Control)
Open fajl + ".txt" For Output As #1
For staa = 1 To lvc.ListItems.Count
Print #1, lvc.ListItems(staa).SubItems(3); Tab(30); " - "; lvc.ListItems(staa).SubItems(4); Tab(30); " - "; lvc.ListItems(staa).SubItems(5)
Next staa
Close #1
End Function

Function savefavorite(fajl As String, lvc As Control)
Open fajl For Output As #1
For sve = 1 To lvc.ListItems.Count
Write #1, lvc.ListItems(sve), lvc.ListItems(sve).SubItems(1), lvc.ListItems(sve).SubItems(2), lvc.ListItems(sve).SubItems(3), lvc.ListItems(sve).SubItems(4)
Next sve
Close #1
End Function

Function openfavorite(fajl As String, lvc As ListView)
lvc.ListItems.Clear
Open fajl For Input As #1
Do
Input #1, kol0, kol1, kol2, kol3, kol4
Set bleh = lvc.ListItems.Add(, , kol0)
bleh.SubItems(1) = kol1
bleh.SubItems(2) = kol2
bleh.SubItems(3) = kol3
bleh.SubItems(4) = kol4
Loop Until EOF(1)
Close #1
End Function

Function GetTagID3v2(tmpfile As String)
  
  Dim StrGanzerTag As String * 10
  Dim bytGanzerTag() As Byte
  Dim strIDdv2Tag As String
  Dim strTagLength As Long

On Error Resume Next

Close #2

Open tmpfile For Binary As #2
Get #2, , StrGanzerTag

If Left$(StrGanzerTag, 3) = "ID3" Then
strTagLength = Asc(Mid$(StrGanzerTag, 7, 1)) * 2 ^ 21 + Asc(Mid$(StrGanzerTag, 8, 1)) * CLng(2 ^ 14) + Asc(Mid$(StrGanzerTag, 9, 1)) * 2 ^ 7 + Asc(Mid$(StrGanzerTag, 10, 1)) * 2 ^ 0
        
ReDim bytGanzerTag(strTagLength) As Byte
Seek #2, 1
Get #2, , bytGanzerTag
Close #2
        
strIDdv2Tag = StrConv(bytGanzerTag, vbUnicode)

        tmp1 = InStr(11, strIDdv2Tag, "TPE1" & String$(3, 0))
        tmp2 = Asc(Mid$(strIDdv2Tag, tmp1 + 7, 1)) - 1
        
        tmp3 = InStr(11, strIDdv2Tag, "TIT2" & String$(3, 0))
        tmp4 = Asc(Mid$(strIDdv2Tag, tmp3 + 7, 1)) - 1
        
        tmp5 = InStr(11, strIDdv2Tag, "TALB" & String$(3, 0))
        tmp6 = Asc(Mid$(strIDdv2Tag, tmp5 + 7, 1)) - 1
        
        tmp7 = InStr(11, strIDdv2Tag, "TRCK" & String$(3, 0))
        tmp8 = Asc(Mid$(strIDdv2Tag, tmp7 + 7, 1)) - 1

        tmp9 = InStr(11, strIDdv2Tag, "TYER" & String$(3, 0))
        tmp10 = Asc(Mid$(strIDdv2Tag, tmp9 + 7, 1)) - 1
        
        tmp11 = InStr(11, strIDdv2Tag, "COMM" & String$(3, 0))
        tmp12 = Asc(Mid$(strIDdv2Tag, tmp11 + 7, 1)) - 5

   
        StrArtist = Mid$(strIDdv2Tag, tmp1 + 11, tmp2)
        StrSongName = Mid$(strIDdv2Tag, tmp3 + 11, tmp4)
        StrAlbum = Mid$(strIDdv2Tag, tmp5 + 11, tmp6)
        StrGodina = Mid$(strIDdv2Tag, tmp9 + 11, tmp10)
        StrKomentar = Mid$(strIDdv2Tag, tmp11 + 15, tmp12)
         

        artistv2 = Trim$(StrArtist)
        titlev2 = Trim$(StrSongName)
        albumv2 = Trim$(StrAlbum)
        godinav2 = Trim$(StrGodina)
        komentarv2 = Trim$(StrKomentar)
    
    End If
Close #2
End Function
