Attribute VB_Name = "modWaveOut"
' /*******************************************************************************
'   modWaveOut.bas within vbSpec.vbp
'
'   API declarations and support routines for proving beeper emulation using
'   the Windows waveOut* API fucntions.
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'
'   Copyright (C)1999-2000 Grok Developments Ltd.
'   http://www.grok.co.uk/
'
'   This program is free software; you can redistribute it and/or
'   modify it under the terms of the GNU General Public License
'   as published by the Free Software Foundation; either version 2
'   of the License, or (at your option) any later version.
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program; if not, write to the Free Software
'   Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'
' *******************************************************************************/

Option Explicit


Public Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveOutOpen Lib "winmm.dll" (LPHWAVEOUT As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
    
' // Functions for allocations fixed blocks of memory to hold the waveform buffers
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

' // Function for moving waveform data from a VB byte array into a block of memory
' // allocated by GlobalAlloc()
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)

Public Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
End Type
Public Type WAVEHDR
        lpData As Long
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        dwLoops As Long
        lpNext As Long
        Reserved As Long
End Type
Public Const WAVE_ALLOWSYNC = &H2
Public Const WAVE_FORMAT_1M08 = &H1              '  11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1M16 = &H4              '  11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S08 = &H2              '  11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1S16 = &H8              '  11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10             '  22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2M16 = &H40             '  22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S08 = &H20             '  22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2S16 = &H80             '  22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100            '  44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4M16 = &H400            '  44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S08 = &H200            '  44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4S16 = &H800            '  44.1   kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_DIRECT = &H8
Public Const WAVE_FORMAT_PCM = 1  '  Needed in resource files so outside #ifndef RC_INVOKED
Public Const WAVE_FORMAT_QUERY = &H1
Public Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Public Const WAVE_INVALIDFORMAT = &H0              '  invalid format
Public Const WAVE_MAPPED = &H4
Public Const WAVE_MAPPER = -1&
Public Const WAVE_VALID = &H3         '  ;Internal
Public Const WAVECAPS_LRVOLUME = &H8         '  separate left-right volume control
Public Const WAVECAPS_PITCH = &H1         '  supports pitch control
Public Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control
Public Const WAVECAPS_SYNC = &H10
Public Const WAVECAPS_VOLUME = &H4         '  supports volume control
Public Const WAVERR_BASE = 32
Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)    '  unsupported wave format

Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)    '  last error in range
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)    '  device is synchronous
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  header not prepared
Public Const MMSYSERR_NOERROR = 0  '  no error

Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Public Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Public Const CALLBACK_NULL = &H0        '  no callback
Public Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC
Public Const CALL_PENDING = &H2


Public Const WHDR_BEGINLOOP = &H4         '  loop start block
Public Const WHDR_DONE = &H1         '  done bit
Public Const WHDR_INQUEUE = &H10        '  reserved for driver
Public Const WHDR_ENDLOOP = &H8         '  loop end block
Public Const WHDR_PREPARED = &H2         '  set if this header has been prepared
Public Const WHDR_VALID = &H1F        '  valid flags      / ;Internal /

Public Const MM_WOM_CLOSE = &H3BC
Public Const MM_WOM_DONE = &H3BD
Public Const MM_WOM_OPEN = &H3BB  '  waveform output
Public Const WOM_DONE = MM_WOM_DONE
Public Const WOM_OPEN = MM_WOM_OPEN
Public Const WOM_CLOSE = MM_WOM_CLOSE

' // Variables and constants used by the beeper emulation
Public glphWaveOut As Long
Public Const NUM_WAV_BUFFERS = 20
Public Const WAVE_FREQUENCY = 22050
Public Const WAV_BUFFER_SIZE = 441 ' (WAVE_FREQUENCY \ NUM_WAV_BUFFERS)
Public ghMem(1 To NUM_WAV_BUFFERS) As Long
Public gpMem(1 To NUM_WAV_BUFFERS) As Long
Public gtWavFormat As WAVEFORMAT
Public gtWavHdr(1 To NUM_WAV_BUFFERS) As WAVEHDR
Public gcWaveOut(48000) As Byte
Public glWavePtr As Long
Public glWaveAddTStates As Long

Public Sub AddSoundWave(ts As Long)
    Dim lEarVal
    Static WCount As Long
    
    WCount = WCount + 1
    If WCount = 800 Then
        AY8912Update_8
        WCount = 0
    End If
        
    Static lCounter As Long
    lCounter = lCounter + ts
    
    If gbTZXPlaying Then
        If glEarBit = 64 Then lEarVal = 15 Else lEarVal = 0
    End If
    
    Do While lCounter >= glWaveAddTStates
        If gbEmulateAYSound Then
            gcWaveOut(glWavePtr) = glBeeperVal + RenderByte + lEarVal
        Else
            gcWaveOut(glWavePtr) = glBeeperVal + lEarVal
        End If
        glWavePtr = glWavePtr + 1
        lCounter = lCounter - glWaveAddTStates
    Loop
End Sub


