<div align="center">

## Record sound CD \(track\) to WAV file\.


</div>

### Description

Record sound from CD (Track1, Track2...) to a WAV file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Damjan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/damjan.md)
**Level**          |Unknown
**User Rating**    |6.0 (610 globes from 102 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/damjan-record-sound-cd-track-to-wav-file__1-2091/archive/master.zip)





### Source Code

```
'This control use MCI to control CD
Public Sub RecordWave(TrackNum As Integer, Filename As String)
' TrackNum: track to record
' Filename: file to save wave as
On Local Error Resume Next
Dim i As Long
Dim RS As String
Dim cb As Long
Dim t
    RS = Space$(128)
    i = mciSendString("stop cdaudio", RS, 128, cb)
    i = mciSendString("close cdaudio", RS, 128, cb)
    Kill Filename
    RS = Space$(128)
    i = mciSendString("status cdaudio position track " & TrackNum, RS, 128, cb)
    i = mciSendString("open cdaudio", RS, 128, cb)
    i = mciSendString("set cdaudio time format milliseconds", RS, 128, cb)
    i = mciSendString("play cdaudio", RS, 128, cb)
    i = mciSendString("open new type waveaudio alias capture", RS, 128, cb)
    i = mciSendString("record capture", RS, 128, cb)
    t# = Timer + 1: Do Until Timer > t#: DoEvents: Loop
    i = mciSendString("save capture " & Filename, RS, 128, cb)
    i = mciSendString("stop cdaudio", RS, 128, cb)
    i = mciSendString("close cdaudio", RS, 128, cb)
End Sub
```

