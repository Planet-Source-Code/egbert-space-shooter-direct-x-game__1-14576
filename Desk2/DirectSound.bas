Attribute VB_Name = "DirectSoundX"
Option Explicit
Private m_dx As New DirectX7
Private m_dxs As DirectSound
Private Type dxBuffers
isLoaded As Boolean
Buffer As DirectSoundBuffer
FileName As String
End Type
Private SoundFolder As String
Private SB() As dxBuffers
Private CurrentBuffer As Integer

Private Sub LoadSound(Buffer As Integer, sfile As String)
Dim FileName As String
Dim bufferDesc As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX
  
bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
  
waveFormat.nFormatTag = WAVE_FORMAT_PCM
waveFormat.nChannels = 2    '2 channels
waveFormat.lSamplesPerSec = 22050
waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign

FileName = SoundFolder & sfile
On Error GoTo Continue
Set SB(Buffer).Buffer = m_dxs.CreateSoundBufferFromFile(FileName, bufferDesc, waveFormat)
SB(Buffer).isLoaded = True
Exit Sub
Continue:
MsgBox "Error can't find file: " & FileName, vbExclamation + vbOKOnly, "Error"
End Sub

Public Function SetupSound(hWnd As Long, NeededBuffers As Long) As Boolean
On Error Resume Next
Err.Clear
Set m_dxs = m_dx.DirectSoundCreate("")
If Err.Number <> 0 Then
MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed" & Chr(13) & "And check you current directx version it must 7 or higher", vbCritical + vbOKOnly, "Error"
Exit Function
End If
m_dxs.SetCooperativeLevel hWnd, DSSCL_PRIORITY
SetupSound = True
ReDim SB(NeededBuffers)
End Function

Public Function PlaySoundGiveBuffer(FileName As String, Volume As Long, PanValue As Long, LoopIt As Long, SearchFrom As Long, SearchTo As Long)
On Error Resume Next
Dim A As Integer

For A = SearchFrom To SearchTo
If SB(A).isLoaded = False Then
Exit For
End If
If SB(A).Buffer.GetStatus <> DSBSTATUS_PLAYING Then
Exit For
End If
Next A

If SB(A).isLoaded = False Or SB(A).FileName <> FileName Then
LoadSound A, FileName
SB(A).FileName = FileName
End If

PlaySoundGiveBuffer = A
PanSound A, PanValue
VolumeLevel A, Volume
If SB(A).isLoaded Then
SB(A).Buffer.Play LoopIt
End If
End Function

Function PlaySoundByBuffer(Buffer As Integer, FileName As String, Volume As Long, PanValue As Long, LoopIt As Long) As Integer
If SB(Buffer).isLoaded = False Or SB(Buffer).FileName <> FileName Then
LoadSound Buffer, FileName
SB(Buffer).FileName = FileName
End If
If SB(Buffer).isLoaded = False Then Exit Function
SB(Buffer).Buffer.SetCurrentPosition 0
PanSound Buffer, PanValue
VolumeLevel Buffer, Volume
If SB(Buffer).isLoaded Then
SB(Buffer).Buffer.Play LoopIt
End If
End Function

Function PanSound(Buffer As Integer, PanValue As Long)
Select Case PanValue
Case 0
SB(Buffer).Buffer.SetPan -10000
Case 100
SB(Buffer).Buffer.SetPan 10000
Case Else
SB(Buffer).Buffer.SetPan (100 * PanValue) - 5000
End Select
End Function

Function VolumeLevel(Buffer As Integer, Volume As Long)
If Volume > 0 Then
SB(Buffer).Buffer.SetVolume (60 * Volume) - 6000
Else
SB(Buffer).Buffer.SetVolume -6000
End If
End Function

Function IsPlaying(Buffer As Integer) As Boolean
If SB(Buffer).Buffer.GetStatus = DSBSTATUS_PLAYING Or SB(Buffer).Buffer.GetStatus = DSBSTATUS_LOOPING Then
IsPlaying = True
End If
End Function


