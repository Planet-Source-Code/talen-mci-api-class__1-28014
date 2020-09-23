VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCIApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class was created by Talen, though
' idea of it wasn't mine - I saw mci-module by Abdullah Al-Ahdal
' and wanted to make some like this, but as class.
' Then, with help of MSDN I made this.
' I surely understand - this is not best solution
' but I like it ... And, besides, it's not the last
' version.
' Usage - Dim mci as clsMCIApi, then Set mci = new clsMCIApi
' Then open device, and use various functions.
' Dont forget to unload class when unloading app
' (Set mci = nothing)
' WARNING! Always close opened mci-devices,
' because they are opened on system-level! If you open
' device and forget to close it, this will result in:
' 1. Noone will be able to open device with same alias, and
' 2. System resources won't be freed.
' All non-string function return 1 in case of error,
' you can read error description by accessing
' 'ErrorDescription' property.
' All string function return "MCIClass_ERROR" in
' case of error.
' It is not important to check for error each
' time you call class' method.
' If you're looking for some missing mci-commands
' or mci-command parameters, you will find all of this
' in MSDN - 'cause I found all of this there :)
' And, one more comment - don't bother voting for me :)

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private ErrorDesc$, ErrorNum&
Private TTime&, TFrames&, FPS&

Public Property Get ErrorDescription$()
' Description of error occured.
ErrorDescription = ErrorDesc$
End Property

Public Property Get ErrorNumber&()
' Number of error occured
ErrorNumber = ErrorNum
End Property

Public Property Get TotalTime&()
' Total time of opened stream
TotalTime = TTime
End Property

Public Property Get TotalFrames()
' Total frames in opened stream
TotalFrames = TFrames
End Property

Public Property Get FramesPerSec()
' Frames per second in opened stream
    FramesPerSec = FPS
End Property
Public Function OpenDevice&(DeviceType$, DeviceAlias$, filename$, ParentHWnd&, WindowStyle$)
' Opens MCI device. If you want to use WS_ constant in
' window style, CStr it first.
Dim rc&

rc = mciSendString("open " & filename$ & " type " & DeviceType$ & " alias " & DeviceAlias$ & " parent " & ParentHWnd & " style " & WindowStyle, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    OpenDevice = 1
Else
    OpenDevice = 0
    rc = cmdTimeFormat(DeviceAlias, "ms")
    TTime = cmdGetLength(DeviceAlias)
    rc = cmdTimeFormat(DeviceAlias, "frames")
    TFrames = cmdGetLength(DeviceAlias)
    FPS = TFrames / (TTime / 1000)
End If
End Function

Public Function CloseDevice&(DeviceAlias$)
' Close MCI device.
' This device must be opened first (obvious, isn't it? :)
Dim rc&
rc = mciSendString("close " & DeviceAlias, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    CloseDevice = 1
Else
    CloseDevice = 0
End If
End Function

Public Function cmdPlay&(DeviceAlias$, FromPosition As Integer, ToPosition As Integer, Fullscreen As Boolean)
' Play device. Set Fullscreen = true to
' play full-screen video.

Dim todo$, rc&

todo = "play " & DeviceAlias
If Not FromPosition Then todo = todo + " from " & FromPosition
If Not ToPosition Then todo = todo + " to " & ToPosition
If Fullscreen Then todo = todo + " fullscreen"

rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdPlay = 1
Else
    cmdPlay = 0
End If
End Function

Public Function cmdStop&(DeviceAlias$)
' Stop device (warning: it doesn't rewind it to start).
Dim todo$, rc&

todo = "stop " & DeviceAlias
rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdStop = 1
Else
    cmdStop = 0
End If
End Function

Public Function cmdPause&(DeviceAlias$)
' Pause device.
Dim rc&

rc = mciSendString("pause " & DeviceAlias, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdPause = 1
Else
    cmdPause = 0
End If
End Function

Public Function cmdChannels&(DeviceAlias$, Channel%, Value As Boolean)
' Turn channels on/off.
' Use 'channel' to specify channel (0=left, 1=right)
' Value may be true or false (on or off).

Dim todo$, rc&

todo$ = "set " + DeviceAlias + " audio "
Select Case Channel
    Case 0
        todo = todo + "left "
    Case 1
        todo = todo + "right "
End Select

If Value Then todo = todo + "on" Else todo = todo + "off"
rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdChannels = 1
Else
    cmdChannels = 0
End If
End Function

Public Function cmdRate&(DeviceAlias$, Rate%)
' Sets speed rate. 1000 is nominal rate,
' set to zero to play as fast as possible without
' dropping frames.
Dim rc&

rc = mciSendString("set " + DeviceAlias + " speed " + CStr(Rate), vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdRate = 1
Else
    cmdRate = 0
End If
End Function

Public Function cmdTimeFormat&(DeviceAlias$, TimeFormat$)
' See MSDN article to acquire valid timeformats.
' Some of them - frames, ms (milliseconds).
Dim rc&

rc = mciSendString("set " + DeviceAlias + " time format " + TimeFormat, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdTimeFormat = 1
Else
    cmdTimeFormat = 0
End If
End Function

Public Function cmdVideoOutput&(DeviceAlias$, VideoOutput As Boolean)
' Sets Video Output (it doesn't seem to work)
Dim rc&

If VideoOutput Then
    rc = mciSendString("set " + DeviceAlias + " video on", vbNullString, 0&, 0&)
Else
    rc = mciSendString("set " + DeviceAlias + " video off", vbNullString, 0&, 0&)
End If
If rc <> 0 Then
    HandleErr rc
    cmdVideoOutput = 1
Else
    cmdVideoOutput = 0
End If
End Function

Public Function cmdConfigure&(DeviceAlias$)
' Calls configuration window (same as above)
Dim rc&

rc = mciSendString("configure " + DeviceAlias, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdConfigure = 1
Else
    cmdConfigure = 0
End If
End Function

Public Function cmdGetCopyRight$(DeviceAlias$)
Dim rc&, t$

t = String$(255, vbNullChar)
rc = mciSendString("info " + DeviceAlias + " copyright", t, Len(t), 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetCopyRight = "MCIClass_ERROR"
Else
    cmdGetCopyRight = left(t, InStr(t, vbNullChar) - 1)
End If
End Function

Public Function cmdGetName$(DeviceAlias$)
Dim rc&, t$

t = String$(255, vbNullChar)
rc = mciSendString("info " + DeviceAlias + " name", t, Len(t), 0&)
If rc <> 0 Then
    cmdGetName = "MCIClass_ERROR"
    HandleErr rc
Else
    cmdGetName = left(t, InStr(t, vbNullChar) - 1)
End If
End Function

Public Function cmdGetProduct$(DeviceAlias$)
Dim rc&, t$

t = String$(32, vbNullChar)
rc = mciSendString("info " + DeviceAlias + " product", t, Len(t), 0&)
If rc <> 0 Then
    cmdGetProduct = "MCIClass_ERROR"
    HandleErr rc
Else
    cmdGetProduct = left(t, InStr(t, vbNullChar) - 1)
End If
End Function

Public Function cmdResize&(DeviceAlias$, x1%, y1%, x2%, y2%)
' Resize device (applies only to video)
' Coordinates are relative to stream parent.
Dim rc&, todo$
todo = "put " + DeviceAlias$ + " window at " _
         & x1 & " " & y1 & " " & x2 & " " & y2
rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdResize = 1
Else
    cmdResize = 0
End If
End Function

Public Function cmdResizeClient(DeviceAlias$, x1%, y1%, x2%, y2%)
' Resize client (?)
Dim rc&, todo$
todo = "put " + DeviceAlias$ + " window client at " _
        & x1 & " " & y1 & " " & x2 & " " & y2
rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdResizeClient = 1
Else
    cmdResizeClient = 0
End If
End Function


Public Function cmdToStart&(DeviceAlias$)
' Rewind stream to start
Dim rc&
rc = mciSendString("seek " & DeviceAlias$ & " to start", vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdToStart = 1
Else
    cmdToStart = 0
End If
End Function

Public Function cmdToEnd&(DeviceAlias$)
' Rewind stream to end
Dim rc&
rc = mciSendString("seek " & DeviceAlias$ & " to end", vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdToEnd = 1
Else
    cmdToEnd = 0
End If
End Function

Public Function cmdSeek&(DeviceAlias$, SeekTo&)
' Seek to specified position
Dim rc&
rc = mciSendString("seek " & DeviceAlias$ & " to " & (SeekTo * 1000), vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdSeek = 1
Else
    cmdSeek = 0
End If

End Function

Public Function cmdVolume&(DeviceAlias$, Channel%, Factor%)
' Sets volume for specified channel
' 0 - left, 1 - right, 2 - both
' Settings both channels volume doesn't equals them
' (but so does my sliders in example :)
Dim rc&, todo$

todo = "setaudio " + DeviceAlias + " "
Select Case Channel
    Case 0
        todo = todo + "left volume to "
    Case 1
        todo = todo + "right volume to "
    Case 2
        todo = todo + "volume to "
End Select
todo = todo + CStr(Factor)
rc = mciSendString(todo, vbNullString, 0&, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdVolume = 1
Else
    cmdVolume = 0
End If
End Function

Public Function cmdGetVolume&(DeviceAlias$, Channel%, ByRef Volume%)
' Get volume for specified channel, same as above
Dim rc&, todo$, t$

todo = "status " + DeviceAlias + " "
Select Case Channel
    Case 0
        todo = todo + "left volume"
    Case 1
        todo = todo + "right volume"
    Case 2
        todo = todo + "volume"
End Select
t = String(255, vbNullChar)
rc = mciSendString(todo, t, Len(t) - 1, 0&)
Volume = Val(t)
If rc <> 0 Then
    HandleErr rc
    cmdGetVolume = 1
Else
    cmdGetVolume = 0
End If

End Function
Public Function cmdGetPosition&(DeviceAlias$)
' Get current position for specified device
Dim t$, rc&

t = String(255, vbNullChar)
rc = mciSendString("status " + DeviceAlias + " position", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetPosition = -1
Else
    cmdGetPosition = CLng(Val(t))
End If
End Function
Public Function cmdGetFormatTag$(DeviceAlias$)
' I dunno what it does, truly.
Dim t$, rc&

t = String(255, vbNullChar)
rc = mciSendString("status " + DeviceAlias + " format tag", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetFormatTag = "MCIClass_ERROR"
Else
    cmdGetFormatTag = left(t, InStr(t, vbNullChar) - 1)
End If
End Function

Public Function cmdGetTimeFormat$(DeviceAlias$)
' Same as above
Dim t$, rc&

t = String(255, vbNullChar)
rc = mciSendString("status " + DeviceAlias + " time format", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetTimeFormat = "MCIClass_ERROR"
Else
    cmdGetTimeFormat = left(t, InStr(t, vbNullChar) - 1)
End If
End Function

Public Function cmdGetLength&(DeviceAlias$)
' Get duration of the specified device
Dim t$, rc&

t = String(255, vbNullChar)
rc = mciSendString("status " + DeviceAlias + " length", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetLength = -1
Else
    cmdGetLength = CLng(Val(t))
End If

End Function

Public Function cmdGetSize$(DeviceAlias$)
' Get size of specified device (video only)
' Warning: it return only width and height!
' If you want to get left and top, modify func
Dim t$, rc&
Dim x1%, y1%, x2%, y2%

t = String(255, vbNullChar)
rc = mciSendString("where " + DeviceAlias + " destination", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetSize = "MCIClass_ERROR"
Else
    t = left(t, InStr(t, vbNullChar))
    x1 = CInt(left(t, InStr(t, " ") - 1)): t = right(t, Len(t) - InStr(t, " "))
    y1 = CInt(left(t, InStr(t, " ") - 1)): t = right(t, Len(t) - InStr(t, " "))
    x2 = CInt(left(t, InStr(t, " ") - 1)): t = right(t, Len(t) - InStr(t, " "))
    y2 = CInt(t)
    x1 = x2 - x1: y1 = y2 - y1: t = x1 & " " & y1
    cmdGetSize = t
End If

End Function

Public Function cmdGetStatus$(DeviceAlias$)
' Get status of the device
Dim t$, rc&

t = String(255, vbNullChar)
rc = mciSendString("status " + DeviceAlias + " mode", t, Len(t) - 1, 0&)
If rc <> 0 Then
    HandleErr rc
    cmdGetStatus = "MCIClass_ERROR"
Else
    cmdGetStatus = left(t, InStr(t, vbNullChar) - 1)
End If
End Function


Private Sub HandleErr(ErrNum&)
'Handles errors (sets ErrorNum and ErrorDesc)
Dim rc&

ErrorNum = ErrNum
ErrorDesc = String$(255, vbNullChar)
rc = mciGetErrorString(ErrNum, ErrorDesc$, Len(ErrorDesc))
End Sub
