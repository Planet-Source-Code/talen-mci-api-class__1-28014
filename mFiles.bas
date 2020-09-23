Attribute VB_Name = "mFiles"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * 260
  cAlternate As String * 14
End Type

Public Type listEntryInfo
    filename As String
    fileLength As Long
    fileCategory As Integer
    fileInstallTo As String
End Type

Public Type BlockInfo
    BlockSize As Long
    blockCount As Long
    blockRest As Long
    isEOF As Long
    currentBlock As Long
End Type

Public Type mp3tag_dummy
    SongTitle As String * 30
    artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As String * 1
End Type

Public Const INVALID_HANDLE_VALUE = -1
Public findbuffer() As String, findbuffercount%

Public Function getBlockInfo(fileLength As Long, desiredBlockSize As Long) As BlockInfo
    getBlockInfo.currentBlock = 0
    If fileLength < desiredBlockSize Then
        getBlockInfo.blockCount = 1
        getBlockInfo.BlockSize = fileLength
        getBlockInfo.blockRest = 0
        getBlockInfo.isEOF = True
    Else
        getBlockInfo.BlockSize = desiredBlockSize
        getBlockInfo.blockCount = fileLength \ getBlockInfo.BlockSize
        getBlockInfo.blockRest = fileLength Mod getBlockInfo.BlockSize
        If getBlockInfo.blockRest > 0 Then getBlockInfo.blockCount = getBlockInfo.blockCount
        getBlockInfo.isEOF = False
    End If
End Function

Public Function getListEntryInfo(ByVal listentry As String) As listEntryInfo

getListEntryInfo.filename = left(listentry, InStr(listentry, ",") - 1)
listentry = right(listentry, Len(listentry) - InStr(listentry, ","))
getListEntryInfo.fileLength = CLng(left(listentry, InStr(listentry, ",") - 1))
listentry = right(listentry, Len(listentry) - InStr(listentry, ","))
getListEntryInfo.fileCategory = CInt(left(listentry, InStr(listentry, ",") - 1))
listentry = right(listentry, Len(listentry) - InStr(listentry, ","))
getListEntryInfo.fileInstallTo = listentry

End Function

Public Function fileExist(ByVal filename As String) As Boolean
Dim hsearch As Long, findinfo As WIN32_FIND_DATA

hsearch = FindFirstFile(ByVal filename, findinfo)

If hsearch = -1 Then
    fileExist = False
Else
    fileExist = True
End If
End Function

Public Function getID3Tag(filename$) As mp3tag_dummy
Dim ffn% 'Freefilenum
Dim mp3td As mp3tag_dummy
Dim buf$

buf = String(3, vbNullChar)
ffn = FreeFile
Open filename For Binary As ffn
Get ffn, LOF(ffn) - Len(mp3td) - 2, buf
If buf = "TAG" Then
    Get ffn, LOF(ffn) - Len(mp3td) + 1, mp3td
    Close ffn
    getID3Tag = mp3td
Else
    getID3Tag.Genre = vbNullChar
End If
End Function

Public Sub SearchDirsPW(pathToSearch$, useFileSpec As Boolean, fileSpec$)
Dim hItem&, hHandle&
'Ýòè äâå ïåðåìåííûå íå ìîãóò áûòü ñòàòè÷åñêèìè,
'ò.ê. îíè äîëæíû áûòü reinit êàæäûé ðàç!
'Ìû âåäü ÷åðåç êàæäóþ äèðåêòîðèþ ñêðîëëèìñÿ!
Dim dirs%, dirbuf$(), i%
Dim findinfo As WIN32_FIND_DATA
Dim t As mp3tag_dummy
       
DoEvents

hItem& = FindFirstFile(pathToSearch$ & "*.*", findinfo)
If hItem& <> INVALID_HANDLE_VALUE Then
    Do
        If (findinfo.dwFileAttributes And vbDirectory) Then
            If left(findinfo.cFileName, 1) <> "." Then
                If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                dirs% = dirs% + 1
                dirbuf$(dirs%) = left$(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1)
            End If
        'ElseIf UseFileSpec Then
            '...
        End If
    Loop While FindNextFile(hItem&, findinfo)
End If

If useFileSpec Then
    hHandle& = FindFirstFile(pathToSearch$ & fileSpec$, findinfo)
    If hHandle& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If (findbuffercount Mod 10) = 0 Then ReDim Preserve mp3(findbuffercount + 10)
            findbuffercount = findbuffercount + 1
            mp3(findbuffercount).filename = pathToSearch + left(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1)
            t = getID3Tag(pathToSearch + left(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1))
            If t.Genre <> vbNullChar Then
                mp3(findbuffercount).Tag = t
            End If
            mp3(findbuffercount).Key = findbuffercount
        Loop While FindNextFile(hHandle&, findinfo)
    End If
End If
    
Call FindClose(hItem&)
Call FindClose(hHandle&)

For i% = 1 To dirs%: SearchDirsPW pathToSearch$ & dirbuf$(i%) & "\", useFileSpec, fileSpec$: Next i%
  
End Sub

Public Sub SearchDirs(pathToSearch$, useFileSpec As Boolean, fileSpec$)
Dim hItem&, hHandle&
'Ýòè äâå ïåðåìåííûå íå ìîãóò áûòü ñòàòè÷åñêèìè,
'ò.ê. îíè äîëæíû áûòü reinit êàæäûé ðàç!
'Ìû âåäü ÷åðåç êàæäóþ äèðåêòîðèþ ñêðîëëèìñÿ!
Dim dirs%, dirbuf$(), i%
Dim findinfo As WIN32_FIND_DATA
       
DoEvents

hItem& = FindFirstFile(pathToSearch$ & "*.*", findinfo)
If hItem& <> INVALID_HANDLE_VALUE Then
    Do
        If (findinfo.dwFileAttributes And vbDirectory) Then
            If left(findinfo.cFileName, 1) <> "." Then
                If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                dirs% = dirs% + 1
                dirbuf$(dirs%) = left$(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1)
            End If
        'ElseIf UseFileSpec Then
            '...
        End If
    Loop While FindNextFile(hItem&, findinfo)
End If

If useFileSpec Then
    hHandle& = FindFirstFile(pathToSearch$ & fileSpec$, findinfo)
    If hHandle& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If (findbuffercount Mod 10) = 0 Then ReDim Preserve findbuffer(findbuffercount + 10)
            findbuffercount = findbuffercount + 1
            findbuffer(findbuffercount) = pathToSearch + left(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1)
        Loop While FindNextFile(hHandle&, findinfo)
    End If
End If
    
Call FindClose(hItem&)
Call FindClose(hHandle&)

For i% = 1 To dirs%: SearchDirs pathToSearch$ & dirbuf$(i%) & "\", useFileSpec, fileSpec$: Next i%
  
End Sub

Public Function searchByTag(ByVal artist$, ByVal title$) As String
Dim i%, j%

If artist$ = "Íåèçâåñòíûé èñïîëíèòåëü" Then artist$ = ""

For i = 1 To UBound(mp3)
    If RTrim(mp3(i).Tag.artist) = artist Then
        For j = 1 To UBound(mp3)
            If RTrim(mp3(j).Tag.SongTitle) = title + ".mp3" Then
                searchByTag = RTrim(mp3(j).filename)
                Exit Function
            End If
        Next
    End If
Next

End Function

Public Function searchByName(ByVal mp3name As String) As String
Dim i%

For i = 1 To UBound(mp3)
    If lr(RTrim(mp3(i).filename), "\", 1) = mp3name + ".mp3" Then
        searchByName = RTrim(mp3(i).filename)
        Exit Function
    End If
Next

End Function
