Attribute VB_Name = "modWaveIn"
Option Explicit

Private Const GMEM_FIXED As Long = &H0

Private Const CALLBACK_WINDOW As Long = &H10000


Private Const WAVE_FORMAT_PCM As Long = 1

Private Const MM_WIM_CLOSE As Long = &H3BF
Private Const MM_WIM_DATA As Long = &H3C0
Private Const MM_WIM_OPEN As Long = &H3BE
Private Const WIM_CLOSE As Long = MM_WIM_CLOSE
Private Const WIM_DATA As Long = MM_WIM_DATA
Private Const WIM_OPEN As Long = MM_WIM_OPEN
Private Const WHDR_DONE As Long = &H1

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type


Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMiliseconds As Long)

Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" ( _
     ByVal sz As String, _
     ByVal uFlags As Long) As Long

Private Type tHeader
    RIFF As Long            ' "RIFF"
    LenR As Long            ' size of following segment
    WAVE As Long            ' "WAVE"
    fmt As Long             ' "fmt
    FormatSize As Long      ' chunksize
    format As WAVEFORMATEX  ' audio format
    data As Long            ' "data"
    DataLength As Long      ' length of datastream
End Type



Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByRef Destination As Any, _
     ByRef Source As Any, _
     ByVal Length As Long)

Public Declare Sub CopyAudioMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByRef Destination As Any, _
     ByVal Source As Long, _
     ByVal Length As Long)


Private Declare Function GlobalAlloc Lib "kernel32.dll" ( _
     ByVal wFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" ( _
     ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" ( _
     ByVal hmem As Long) As Long



Private Const GWL_WNDPROC As Long = -4

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
     ByVal lpPrevWndFunc As Long, _
     ByVal hWnd As Long, _
     ByVal msg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
     ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long




Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" ( _
     ByVal err As Long, _
     ByVal lpText As String, _
     ByVal uSize As Long) As Long


Private Declare Function waveInReset Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInOpen Lib "winmm.dll" ( _
     ByRef lphWaveIn As Long, _
     ByVal uDeviceID As Long, _
     ByRef lpFormat As WAVEFORMATEX, _
     ByVal dwCallback As Long, _
     ByVal dwInstance As Long, _
     ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" ( _
     ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" ( _
     ByVal hWaveIn As Long, _
     ByRef lpWaveInHdr As WAVEHDR, _
     ByVal uSize As Long) As Long

Private Const BUFFERS As Integer = 4
Private BUFFERSIZE As Long
'Private Const BUFFERSIZE As Long = 8192

Dim hWaveIn As Long
Dim ret As Long
Dim format As WAVEFORMATEX
Dim hmem(BUFFERS) As Long
Dim hdr(BUFFERS) As WAVEHDR
Dim lpPrevWndFunc As Long
Dim hWnd As Long
Dim num As Integer
Public msg As String * 255
Dim pos As Long
Dim bHeaderWritten As Boolean
Dim bRecording As Boolean
Dim bPaused As Boolean
Dim bSaveFile As Boolean

Dim curBuffer As Long

Sub Hook(bHook As Boolean)
    hWnd = frmMain.hWnd
    'Exit Sub
    
    If lpPrevWndFunc <> 0 And bHook Then
        MsgBox "Double-hooking not permited!", vbCritical, "ERROR!"
        Exit Sub
    End If
    
    If bHook Then
        lpPrevWndFunc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf CallbackProc)
    Else
        lpPrevWndFunc = 0
        SetWindowLong hWnd, GWL_WNDPROC, lpPrevWndFunc
    End If
End Sub

Sub Pause()
    Dim i As Integer
    
    bPaused = Not bPaused
    
    If bPaused Then
        waveInStop hWaveIn
    Else
        For i = 0 To BUFFERS
            waveInAddBuffer hWaveIn, hdr(i), Len(hdr(i))
        Next
        
        waveInStart hWaveIn
    End If
End Sub

Sub StopRec()
    Dim i As Integer
    
    If Not bRecording Then Exit Sub
    
    ret = waveInStop(hWaveIn)
    'Form1.DebugIt CStr(ret)
    
    ret = waveInReset(hWaveIn)
    
    For i = 0 To BUFFERS
        waveInUnprepareHeader hWaveIn, hdr(i), Len(hdr(i))
    Next i
    'Form1.DebugIt CStr(ret)
    
    ret = waveInClose(hWaveIn)
    'Form1.DebugIt CStr(ret)
    
    If (bHeaderWritten = False) And (bSaveFile) Then WriteWAVHeader
    
    If bSaveFile Then Close #num
    
    bRecording = False
End Sub

Sub PrepareRecording( _
    ByVal iChannels As Integer, _
    ByVal iBits As Integer, _
    ByVal lFrequency As Long, _
    Optional ByVal sFile As String)
    
    Dim i As Integer
    
    bRecording = True
    
    pos = 0
    
    If sFile <> "" Then
        bSaveFile = True
        
        num = FreeFile
        
        If Dir(sFile) <> "" Then Kill sFile
        Open sFile For Binary As #num
        pos = 45
        bHeaderWritten = False
    Else
        bSaveFile = False
    End If
    
    With format
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = iChannels
        .wBitsPerSample = iBits
        .nSamplesPerSec = lFrequency
        .nBlockAlign = .nChannels * .wBitsPerSample / 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
        
        BUFFERSIZE = 8192 '.nSamplesPerSec * .nBlockAlign * .nChannels * 0.1
        BUFFERSIZE = BUFFERSIZE - (BUFFERSIZE Mod .nBlockAlign)
    End With
    
    ret = waveInOpen(hWaveIn, 0, format, hWnd, 0&, CALLBACK_WINDOW)
    If ret <> 0 Then
        waveInGetErrorText ret, msg, Len(msg)
        'Form1.DebugIt Trim(msg)
        MsgBox Trim(msg)
    End If
    
    For i = 0 To BUFFERS
        hmem(i) = GlobalAlloc(GMEM_FIXED, BUFFERSIZE)
        
        With hdr(i)
            .lpData = GlobalLock(hmem(i))
            .dwBufferLength = BUFFERSIZE
            .dwFlags = 0
            .dwLoops = 0
            .dwUser = CLng(i)
        End With
    Next
    
    For i = 0 To BUFFERS
        ret = waveInPrepareHeader(hWaveIn, hdr(i), Len(hdr(i)))
        
        ret = waveInAddBuffer(hWaveIn, hdr(i), Len(hdr(i)))
    Next
    
    ret = waveInStart(hWaveIn)
End Sub

Public Function CallbackProc(ByVal hw As Long, ByVal uMsg As Integer, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long
    Dim temp() As Byte
    Dim i As Integer
    
    On Error Resume Next
    
    If (uMsg = WIM_DATA) And (bPaused = False) Then
        'ret = waveInAddBuffer(hWaveIn, hdr, Len(hdr))
    
        If (wavhdr.dwFlags And WHDR_DONE) Then
            
            If bSaveFile Then
                
                ReDim temp(wavhdr.dwBytesRecorded)
                CopyMemory temp(0), ByVal wavhdr.lpData, wavhdr.dwBytesRecorded
                WriteData StrConv(temp, vbUnicode)
                
            Else
                
                WriteData ""
                
            End If
            
            curBuffer = wavhdr.dwUser
        End If
        
        For i = 0 To (BUFFERS)
            If Not (hdr(i).dwFlags And WHDR_DONE) Then
                ret = waveInAddBuffer(hWaveIn, hdr(i), Len(hdr(i)))
            End If
        Next
        
        If err Then
            msg = err.Description
            err.Clear
        End If
        
        'waveInGetErrorText ret, msg, Len(msg)
    End If
    CallbackProc = CallWindowProc(lpPrevWndFunc, hw, uMsg, wParam, wavhdr)
End Function

Sub WriteData(Optional ByRef data As String)
    Dim temp() As Byte
    
    If data = "" Then
        pos = pos + BUFFERSIZE
        Exit Sub
    End If
    
    If pos = 0 Then pos = 1
    
    temp = StrConv(data, vbFromUnicode)
    
    Put #num, pos, temp
    
    pos = pos + BUFFERSIZE
End Sub

Sub WriteWAVHeader()
    Dim Header As tHeader
    Dim File() As Byte
    
    With Header
        .RIFF = mmioStringToFOURCC("RIFF", 0&)
        .WAVE = mmioStringToFOURCC("WAVE", 0&)
        .fmt = mmioStringToFOURCC("fmt ", 0&)
        .data = mmioStringToFOURCC("data", 0&)
        
        .format = format
        
        .FormatSize = 16
        .DataLength = LOF(1) - 44
        .LenR = Len(Header) + .DataLength - 8
        ReDim File(Len(Header))
        
        CopyMemory File(0), Header, Len(Header)
    End With
    
    Put #num, 1, File
    bHeaderWritten = True
End Sub

Public Property Get BytesWritten() As Long
    BytesWritten = pos
End Property

Public Property Get GetTime() As Long
    GetTime = pos \ format.nAvgBytesPerSec
End Property

Public Property Get CurBuf() As Long
    CurBuf = curBuffer
End Property

Public Property Get SampleFrequency() As Long
    SampleFrequency = format.nSamplesPerSec
End Property

Public Property Get BlockAlign() As Long
    BlockAlign = format.nBlockAlign
End Property

Public Property Get BufferLength() As Long
    BufferLength = BUFFERSIZE
End Property

Public Property Get BufferData(ByVal lBuffer As Long) As Long
    BufferData = hdr(lBuffer).lpData
End Property

Public Property Get PeakMax() As Double
    If format.wBitsPerSample = 16 Then
        PeakMax = 32767
    ElseIf format.wBitsPerSample = 8 Then
        PeakMax = 127
    End If
End Property

Public Function GetCurPeak(ByRef lLeft As Long, ByRef lRight As Long)
    Static buffer As Integer
    Static bFirst As Boolean
    
    Dim maxLeft As Long
    Dim maxRight As Long
    
    Dim bByte() As Byte
    Dim curLeft As Long
    Dim curRight As Long
    Dim i As Integer
    Dim size As Long
    
    size = (format.nSamplesPerSec / 1000) * format.nBlockAlign
    
    If (bFirst = False) And (buffer <> curBuffer) Then
        bFirst = True
    ElseIf (bFirst = False) And (buffer = 0) Then
        Exit Function
    End If
    
    If curBuffer = 0 Then
        buffer = BUFFERS
    Else
        buffer = curBuffer - 1
    End If
    
    ReDim bByte(size - 1)
    
    If format.wBitsPerSample = 8 Then
        
        size = size * 4
        
        CopyAudioMemory bByte(0), hdr(buffer).lpData, size
        
        Do Until i >= (UBound(bByte) - 1)
            
            curLeft = Abs(bByte(i) - 128)
            
            If curLeft > maxLeft Then maxLeft = curLeft
            
            
            If format.nChannels = 2 Then
                curRight = Abs(bByte(i + 1) - 128)
                
                If curRight > maxRight Then maxRight = curRight
            End If
            
            i = i + format.nBlockAlign
        Loop
        
        
    ElseIf format.wBitsPerSample = 16 Then
        
        CopyAudioMemory bByte(0), hdr(buffer).lpData, size
        
        Do Until i >= (UBound(bByte) - 1)
            
            curLeft = Abs(MakeWord(bByte(i), bByte(i + 1)))
            
            If curLeft > maxLeft Then maxLeft = curLeft
            
            
            If format.nChannels = 2 Then
                
                curRight = Abs(MakeWord(bByte(i + 2), bByte(i + 3)))
                
                
                If curRight > maxRight Then maxRight = curRight
                
            End If
            
            i = i + format.nBlockAlign
        Loop
        
    End If
        
    lLeft = maxLeft
    lRight = maxRight
    
    If format.nChannels = 1 Then lRight = lLeft
End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Long
    If HiByte And &H80 Then
        MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
    Else
        MakeWord = (HiByte * &H100) Or LoByte
    End If
End Function

Public Function Unsigned(ByVal iInt As Integer) As Long
    If iInt < 0 Then
        Unsigned = iInt + 65536
    Else
        Unsigned = iInt
    End If
End Function

Function ConvertToByte2(ByVal int_Input As Integer) As Integer
    Dim lInput As Long
    
    lInput = int_Input And &HFFFF&  ' convert to unsigned
    
    If lInput And &H8000& Then ' high bit set, it's negative (test instead of cmp)
        ConvertToByte2 = &H80 Or (lInput \ 256)
    Else
        ConvertToByte2 = (lInput - 32767) \ 256
    End If
End Function
