Attribute VB_Name = "modMain"
' /*******************************************************************************
'   modMain.bas within vbSpec.vbp
'
'   Public variable declarations, startup and initialization code,
'   and routines for loading the ROM and snapshots
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'
'   Copyright (C)1999-2002 Grok Developments Ltd.
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

' // Pretty much all variables are declared as long, even those that only hold boolean
' // values or byte values. This is done for performance, because VB handles longs more
' // efficiently than any other type.

Public glTstatesPerInterrupt As Long
Public Parity(256) As Long
Public glInterruptTimer As Long

' // Memory haindling variables for 128K spectrum emulation
Public glMemPagingType As Long ' // 0 = No mem paging (48K speccy)
                               ' // 1 = 128K/+2 paging available
                               ' // 2 = 128K/+2 and +2A/+3 special paging available
Public gRAMPage(11, 16383) As Byte ' // (pages 0 to 7 are RAM, and pages 8 to 11 are ROM)
Public glPageAt(4) As Long
Public glLastOut7FFD As Long
Public glLastOut1FFD As Long
Public glUseScreen As Long
Public glEmulatedModel As Long
Public gbEmulateAYSound As Long
Public glContentionTable(-30 To 70930) As Long
Public glTSToScanLine(70930) As Long

Public gbTextOut As Long
Public glTextLineLen As Long
'Public gbInInput As Long

' // TAP/TZX file parameters
Public gsTAPFileName As String
Public ghTAPFile As Long

Public gbSoundEnabled As Long
Public glSoundRegister As Long  ' // Contains currently indexed AY-3-8912 sound register
Public glBufNum As Long         ' // ID of the last Wave buffer used
Public glKeyPortMask As Long    ' // Mask used for reading keyboard port (&HBF on Speccys, &H1F on TC2048)

' // Main Z80 registers //
Public regA As Long
Public regHL As Long
Public regB As Long
Public regC As Long
Public regDE As Long

' // Z80 Flags
Public fS As Long
Public fZ As Long
Public f5 As Long
Public fH As Long
Public f3 As Long
Public fPV As Long
Public fN As Long
Public fC As Long

' // Flag bit positions
Public Const F_C As Long = 1
Public Const F_N As Long = 2
Public Const F_PV As Long = 4
Public Const F_3 As Long = 8
Public Const F_H As Long = 16
Public Const F_5 As Long = 32
Public Const F_Z As Long = 64
Public Const F_S As Long = 128

' // Alternate registers //
Public regAF_ As Long
Public regHL_ As Long
Public regBC_ As Long
Public regDE_ As Long

' // Index registers  - ID used as temp for ix/iy
Public regIX As Long
Public regIY As Long
Public regID As Long

' // Stack pointer and program counter
Public regSP As Long
Public regPC As Long

' // Interrupt registers and flip-flops, and refresh registers
Public intI As Long
Public intR As Long
Public intRTemp As Long
Public intIFF1 As Long
Public intIFF2 As Long
Public intIM As Long

' //////////////////////////////////////////////////
' // Variables used by the video display rountines
' //////////////////////////////////////////////////
Public ScrnLines(191, 65) As Long    ' // 192 scanlines (0-191) and either 32 bytes per line or 64 in TC2048 hires mode, plus two flag bytes
Public ScrnNeedRepaint As Long       ' // Set to true when an area of the display changes, and back to false by the ScreenPaint function
Public bFlashInverse As Long         ' // Cycles between true/false to indicate the status of 'flashing' attributes
Public glScreenMem(191, 31) As Long  ' // Static lookup table that maps Y,X screen coords to the correct Speccy memory address
Public glScreenMemTC2048HiRes(191, 63) As Long ' // As above, but for the TC2048 512x192 display mode
Public glTC2048HiResColour As Long   ' // The attribute value for the TC2048 512x192, two-colour display mode
Public glTC2048LastFFOut As Long     ' // Contains the last value OUTed to port FF in TC2048 mode (indicates the screen mode in use)
Public glLastFEOut As Long           ' // Contains the last value OUTed to any port with bit 0 reset (saved in snapshots, etc)
Public glTopMost As Long             ' // Top-most row of the screen that has changed since the last ScreenPaint
Public glBottomMost As Long          ' // Bottom-most row of the screen that has changed since the last ScreenPaint
Public glLeftMost As Long            ' // Left-most column of the screen that has changed since the last ScreenPaint
Public glRightMost As Long           ' // Right-most column "  "    "      "   "     "      "    "   "         "

' // Bob Woodring's (RGW) video performance improvements use the following lookup tables
Public glRowIndex(191) As Long
Public glColIndex(191) As Long
Type tBitTable
   dw0 As Long
   dw1 As Long
End Type
Public gtBitTable(255, 255) As tBitTable
Public glMemAddrDiv16384(81919) As Long ' // Lookup table used in pokeb() & elsewhere
Public glMemAddrDiv256(81919) As Long   ' // Lookup table used by pokeb() - faster
Public glMemAddrDiv32(81919) As Long    ' // Lookup table used by pokeb() -  than
Public glMemAddrDiv4(81919) As Long     ' // Lookup table used by pokeb() -   integer division!


' // RGW -- Variables used by scanline video routines
Public glTStatesPerLine As Long  ' // Contains the # of T-states per display line (different on 48K and 128K spectrums)
Public glTStatesAtTop As Long    ' // # of t-states before the start of the first screen line (excluding border)
Public glTStatesAtBottom As Long ' // # of t-states after the end of the last screen line (excluding border)
Public glTStates As Long         ' // Number of T-States for current frame (counts down towards zero, at which time an interrupt occurs)

Public glContendedMemoryDelay As Long  ' // Contains number of tstates added due to memory contention for the current opcode

' // Array of colour values (speeds up screen painting by avoiding
' // multiple calls to RGB() )
Public glBrightColor(0 To 7) As Long
Public glNormalColor(0 To 7) As Long

' // Global picDisplay variable to speed things up
Public gpicDisplay As PictureBox, gpicDC As Long

' // Interrupts/Screen refreshing
Public interruptCounter As Long
Public glInterruptDelay As Long
Public glDelayOverage As Long

' // Keypresses
Public keyB_SPC As Long
Public keyH_ENT As Long
Public keyY_P As Long
Public key6_0 As Long
Public key1_5 As Long
Public keyQ_T As Long
Public keyA_G As Long
Public keyCAPS_V As Long

' // Sadly, I needed to use these high res timer functions to precisely control the
' // speed of emulation. I had hoped to do this without resorting to any API calls :(
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' // If built with the USE_WINAPI compiler directive defined, vbSpec uses Windows API
' // functions to paint the display. This is faster than raw VB code, and provides the
' // option for selecting scaled display sizes (double and triple size, and so on).
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function StretchDIBitsMono Lib "gdi32" Alias "StretchDIBits" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFOMONO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Const SM_CYCAPTION As Long = 4
Public Const SM_CYMENU As Long = 15
Public Const SM_CXFRAME As Long = 32
Public Const SM_CYFRAME As Long = 33

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(16) As RGBQUAD
End Type
Public Type BITMAPINFOMONO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(2) As RGBQUAD
End Type
Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public bmiBuffer As BITMAPINFO
Public glBufferBits(24576) As Long

Public glDisplayHeight As Long
Public glDisplayWidth As Long
Public glDisplayVSource As Long ' // Set to glDisplayHeight - 1 to improve display speed
Public glDisplayVSize As Long   ' // Set to -glDisplayHeight to improve display speed
Public glDisplayXMultiplier As Long
Public glDisplayYMultiplier As Long

' // Used by the ZX Printer emulation
Public bmiZXPrinter As BITMAPINFOMONO
Public gcZXPrinterBits() As Byte ' // 32 cols * 1152 rows
Public glZXPrinterBMPHeight As Long

Public glBeeperVal As Long

' // ShellExecute is used for the clickable web URL in the "About..." dialog, not
' //the actual emulator itself
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' // MouseCapture functions required for emulating the Kempston Mouse Interface
Public Type POINTAPI
        X As Long
        y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_MBUTTON = &H4

Public glMouseType As Long
Public Const MOUSE_NONE = 0&
Public Const MOUSE_KEMPSTON = 1&
Public Const MOUSE_AMIGA = 2&
Public gbMouseGlobal As Long
Public glMouseBtn As Long

' // Most Recently Used (MRU) file class
Public gMRU As MRUList

' // Flag whether SE BASIC ROM is to be used or not
Public gbSEBasicROM As Long

'MM 16.04.2003
Public Type RECT
    iLeft As Integer
    iTop As Integer
    iRight As Integer
    iBottom As Integer
End Type
Public Type POINT
    lX As Long
    lY As Long
End Type
Public Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Public Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
            
Private Function CompressMemoryBlock(lBlock As Long, sData As String) As Long
    Dim i As Long, bLastWasED As Boolean, cRepChar As Byte, lLength As Long
    
    Do While i < 16384
        ' // The last byte, just write it out
        If i = 16383 Then
            sData = sData & Chr$(gRAMPage(lBlock, i))
            Exit Do
        End If
        
        If (gRAMPage(lBlock, i) = gRAMPage(lBlock, i + 1)) And Not (bLastWasED) Then
            ' // It's a run of bytes and we're not immediately following an ED
            cRepChar = gRAMPage(lBlock, i)
            i = i + 2
            lLength = 2
            bLastWasED = False
            
            ' // Find the length of the run (but cap it at 255 bytes)
            Do While (i < 16384) And (gRAMPage(lBlock, i) = cRepChar) And (lLength < 255)
                lLength = lLength + 1
                i = i + 1
            Loop
            
            If (lLength >= 5) Or (cRepChar = &HED) Then
                sData = sData & Chr$(&HED) & Chr$(&HED) & Chr$(lLength) & Chr$(cRepChar)
            Else
                ' // Not compressible, just write out the raw data
                sData = sData & String$(lLength, cRepChar)
            End If
        Else
            ' // Not a run of bytes
            cRepChar = gRAMPage(lBlock, i)
            sData = sData & Chr$(cRepChar)
            If cRepChar = &HED Then bLastWasED = True Else bLastWasED = False
            i = i + 1
        End If
    Loop
    
    If Len(sData) > 16384 Then
        ' // If the compressed block is longer than the original
        ' // just store the uncompressed version
        sData = ""
        For i = 0 To 16383
            sData = sData & Chr$(gRAMPage(lBlock, i))
        Next i
        CompressMemoryBlock = 65535 ' // block is uncompressed
    Else
        CompressMemoryBlock = Len(sData)
    End If
End Function

Private Sub InitAmigaMouseTables()
    glAmigaMouseX(0) = 5  ' 0101
    glAmigaMouseX(1) = 1  ' 0001
    glAmigaMouseX(2) = 0  ' 0000
    glAmigaMouseX(3) = 4  ' 0100
    
    glAmigaMouseY(0) = 10 ' 1010
    glAmigaMouseY(1) = 8  ' 1000
    glAmigaMouseY(2) = 0  ' 0000
    glAmigaMouseY(3) = 2  ' 0010
End Sub


Sub InitScreenIndexs()
    Dim n As Long
   
    For n = 0 To 191
        glRowIndex(n) = 6144 + (n \ 8) * 32
        glColIndex(n) = (n * 256) \ 4
        
    Next n
   
    For n = 0 To 81919
        glMemAddrDiv16384(n) = (n And 65535) \ 16384&
        glMemAddrDiv256(n) = (n And 65535) \ 256&
        glMemAddrDiv32(n) = (n And 65535) \ 32&
        glMemAddrDiv4(n) = (n And 65535) \ 4&
    Next n
End Sub
Public Sub CloseWaveOut()
    Dim lRet As Long
    
    lRet = waveOutReset(glphWaveOut)
    For lRet = 1 To NUM_WAV_BUFFERS
        waveOutUnprepareHeader glphWaveOut, gtWavHdr(lRet), Len(gtWavHdr(lRet))
    Next lRet
    
    For lRet = 1 To NUM_WAV_BUFFERS
        GlobalUnlock ghMem(lRet)
        GlobalFree ghMem(lRet)
    Next lRet
    
    waveOutClose glphWaveOut
    
    gbSoundEnabled = False
End Sub

Public Sub CreateScreenBuffer()
    bmiBuffer.bmiHeader.biSize = Len(bmiBuffer.bmiHeader)
    bmiBuffer.bmiHeader.biWidth = 256
    bmiBuffer.bmiHeader.biHeight = 192
    bmiBuffer.bmiHeader.biPlanes = 1
    bmiBuffer.bmiHeader.biBitCount = 8
    bmiBuffer.bmiHeader.biCompression = BI_RGB
    bmiBuffer.bmiHeader.biSizeImage = 0
    bmiBuffer.bmiHeader.biXPelsPerMeter = 200
    bmiBuffer.bmiHeader.biYPelsPerMeter = 200
    bmiBuffer.bmiHeader.biClrUsed = 16
    bmiBuffer.bmiHeader.biClrImportant = 16
End Sub

Private Function GetBorderIndex(lRGB As Long) As Long
    Dim lCounter As Long
    
    For lCounter = 0 To 7
        If glNormalColor(lCounter) = lRGB Then
            GetBorderIndex = lCounter
        End If
    Next lCounter
End Function

Function GetFilePart(ByVal sFileName As String) As String
    If InStr(sFileName, ":") > 0 Then sFileName = Mid$(sFileName, InStr(sFileName, ":") + 1)
    Do While InStr(sFileName, "\") > 0
        sFileName = Mid$(sFileName, InStr(sFileName, "\") + 1)
    Loop
    GetFilePart = sFileName
End Function

Public Function InitializeWaveOut() As Boolean
    Dim lRet As Long, sErrMsg As String, lCounter As Long, lCounter2 As Long
    
    If val(GetSetting("Grok", "vbSpec", "SoundEnabled", CStr(1))) = 0 Then
        InitializeWaveOut = False
        Exit Function
    End If
    
    glBeeperVal = 128
    
    With gtWavFormat
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = 1
        .nSamplesPerSec = WAVE_FREQUENCY
        .nAvgBytesPerSec = WAVE_FREQUENCY
        .nBlockAlign = 1
        .wBitsPerSample = 8
        .cbSize = 0
    End With
    lRet = waveOutOpen(glphWaveOut, WAVE_MAPPER, gtWavFormat, 0, True, CALLBACK_NULL)
    If lRet <> MMSYSERR_NOERROR Then
        sErrMsg = Space$(255)
        waveOutGetErrorText lRet, sErrMsg, Len(sErrMsg)
        sErrMsg = Left$(sErrMsg, InStr(sErrMsg, Chr$(0)) - 1)
        MsgBox "Error initialising WaveOut device." & vbCrLf & vbCrLf & sErrMsg, vbOKOnly Or vbExclamation, "vbSpec"
        InitializeWaveOut = False
        Exit Function
    End If
    
    For lCounter = 1 To NUM_WAV_BUFFERS
        ghMem(lCounter) = GlobalAlloc(GPTR, WAV_BUFFER_SIZE)
        gpMem(lCounter) = GlobalLock(ghMem(lCounter))
        With gtWavHdr(lCounter)
            .lpData = gpMem(lCounter)
            .dwBufferLength = WAV_BUFFER_SIZE
            .dwUser = 0
            .dwFlags = 0
            .dwLoops = 0
            .lpNext = 0
        End With
        
        lRet = waveOutPrepareHeader(glphWaveOut, gtWavHdr(lCounter), Len(gtWavHdr(lCounter)))
        If lRet <> MMSYSERR_NOERROR Then
            sErrMsg = Space$(255)
            waveOutGetErrorText lRet, sErrMsg, Len(sErrMsg)
            sErrMsg = Left$(sErrMsg, InStr(sErrMsg, Chr$(0)) - 1)
            MsgBox "Error preparing wave header." & vbCrLf & vbCrLf & sErrMsg, vbOKOnly Or vbExclamation, "vbSpec"
            lRet = waveOutClose(glphWaveOut)
            For lCounter2 = 1 To NUM_WAV_BUFFERS
                GlobalUnlock ghMem(lCounter2)
                GlobalFree ghMem(lCounter2)
            Next lCounter2
            InitializeWaveOut = False
            Exit Function
        End If
    Next lCounter
    
    For lCounter = 0 To 48000
        gcWaveOut(lCounter) = glBeeperVal
    Next lCounter
    
    InitializeWaveOut = True
End Function

Sub initParity()
    Dim lCounter As Long, j As Byte, p As Boolean
    
    For lCounter = 0 To 255
        p = True
        For j = 0 To 7
            If (lCounter And (2 ^ j)) <> 0 Then p = Not p
        Next j
        Parity(lCounter) = p
    Next lCounter
End Sub

Sub initscreen()
    Dim i As Long, X As Long
        
    glTopMost = 0
    glBottomMost = 191
    glLeftMost = 0
    glRightMost = 31
    
    For i = 0 To 191
        For X = 0 To 64
            ScrnLines(i, X) = True
        Next X
    Next i
    ScrnNeedRepaint = True
End Sub

Public Sub InitScreenMemTable()
    Dim X As Long, y As Long
    
    For y = 0 To 191
        For X = 0 To 31
            glScreenMem(y, X) = ((((y \ 8) * 32) + (y Mod 8) * 256) + ((y \ 64) * 2048) - (y \ 64) * 256) + X

            glScreenMemTC2048HiRes(y, X * 2) = ((((y \ 8) * 32) + (y Mod 8) * 256) + ((y \ 64) * 2048) - (y \ 64) * 256) + X
            glScreenMemTC2048HiRes(y, X * 2 + 1) = glScreenMemTC2048HiRes(y, X * 2) + 8192
        Next X
    Next y
End Sub


Sub LoadROM(Optional sROMFile As String = "spectrum.rom", Optional lROMPage As Long = 8)
    Dim hFile As Long, sROM As String, lCounter As Long
    
    On Error Resume Next
    
    If Dir$(sROMFile) = "" Then
        MsgBox "The ROM image file '" & sROMFile & "' could not be found.", vbExclamation Or vbOKOnly
        Exit Sub
    End If
    
    err.Clear
    hFile = FreeFile
    Open sROMFile For Binary As hFile
    If err.Number <> 0 Then
        MsgBox "Unable to open ROM image file: " & sROMFile, vbExclamation Or vbOKOnly
        Close hFile
        Exit Sub
    End If
    
    ' // Read the ROM image into sROM
    err.Clear
    sROM = Input(16384, #hFile)
    Close hFile
        
    If err.Number <> 0 Then
        MsgBox "An error ocurred whilst reading the ROM image file: " & sROMFile, vbExclamation Or vbOKOnly
        Exit Sub
    End If
    
    ' // Copy the ROM into the appropriate memory page
    For lCounter = 1 To 16384
        gRAMPage(lROMPage, lCounter - 1) = Asc(Mid$(sROM, lCounter, 1))
    Next lCounter
    resetKeyboard
End Sub

Public Sub LoadScreenSCR(sFileName As String)
    Dim hFile As Long, n As Long, m As Long, sData As String
    
    hFile = FreeFile
    Open sFileName For Binary As hFile
    ' // 6912 - Standard Screen
    ' // 6144+6144 = HiColour TC2048
    ' // 6144+6145 = HiRes TC2048
    
    If LOF(hFile) = 6912 Then
        sData = Input(6912, #hFile)
        If glUseScreen = 1001 Then
            ' // TC2048 Screen1
            For n = 0 To 6911
                gRAMPage(5, n + 8192) = Asc(Mid$(sData, n + 1, 1))
            Next n
        ElseIf glUseScreen = 1002 Then
            ' // TC2048 HiColour
            
            ' // Copy the mono bitmap
            For n = 0 To 6143
                gRAMPage(5, n) = Asc(Mid$(sData, n + 1, 1))
            Next n
            ' // Then expand the normal 768 attributes into the 6144 of the hi-colour mode
            For n = 0 To 255
                For m = 0 To 7
                    gRAMPage(5, n + 8192 + m * 256) = Asc(Mid$(sData, n + 6145, 1))
                    gRAMPage(5, n + 10240 + m * 256) = Asc(Mid$(sData, n + 6401, 1))
                    gRAMPage(5, n + 12288 + m * 256) = Asc(Mid$(sData, n + 6657, 1))
                Next m
            Next n
        ElseIf glUseScreen = 1006 Then
            ' // TC2048 -- We're in hires mode, but we just copy the screen in as usual
            For n = 0 To 6911
                gRAMPage(5, n) = Asc(Mid$(sData, n + 1, 1))
            Next n
        Else
            For n = 0 To 6911
                gRAMPage(glUseScreen, n) = Asc(Mid$(sData, n + 1, 1))
            Next n
        End If
    ElseIf LOF(hFile) = 12288 Then
        sData = Input(12288, #hFile)
        ' // This is a TC2048 HiColour screen
        If glUseScreen <> 1002 Then
            If MsgBox("This file contains a TC2048 high-colour screen, which does not match the current display mode." & vbCrLf & vbCrLf & "Load it anyway?", vbYesNo Or vbDefaultButton1 Or vbQuestion, "vbSpec") = vbNo Then Exit Sub
        End If
        If glUseScreen >= 1000 Then
            For n = 0 To 6143
                gRAMPage(5, n) = Asc(Mid$(sData, n + 1, 1))
                gRAMPage(5, n + 8192) = Asc(Mid$(sData, n + 6145, 1))
            Next n
        Else
            For n = 0 To 6143
                gRAMPage(glUseScreen, n) = Asc(Mid$(sData, n + 1, 1))
                gRAMPage(glUseScreen, n + 8192) = Asc(Mid$(sData, n + 6145, 1))
            Next n
        End If
        If glEmulatedModel = 5 Then outb 255, 2
        
    ElseIf LOF(hFile) = 12289 Then
        sData = Input(12289, #hFile)
        ' // This is a TC2048 HiRes screen
        If glUseScreen <> 1006 Then
            If MsgBox("This file contains a TC2048 high-resolution screen, which does not match the current display mode." & vbCrLf & vbCrLf & "Load it anyway?", vbYesNo Or vbDefaultButton1 Or vbQuestion, "vbSpec") = vbNo Then Exit Sub
        End If
        If glUseScreen >= 1000 Then
            For n = 0 To 6143
                gRAMPage(5, n) = Asc(Mid$(sData, n + 1, 1))
                gRAMPage(5, n + 8192) = Asc(Mid$(sData, n + 6145, 1))
            Next n
        Else
            For n = 0 To 6143
                gRAMPage(glUseScreen, n) = Asc(Mid$(sData, n + 1, 1))
                gRAMPage(glUseScreen, n + 8192) = Asc(Mid$(sData, n + 6145, 1))
            Next n
        End If
        
        If glEmulatedModel = 5 Then outb 255, Asc(Right$(sData, 1))
    Else
        ' // Invalid SCR length
        Close #hFile
        MsgBox "This file does not contain a valid ZX Spectrum or TC2048 screen image.", vbOKOnly Or vbExclamation, "vbSpec"
        Exit Sub
    End If
    
    Close #hFile
    
    initscreen
    screenPaint
    resetKeyboard
End Sub

Private Sub LoadSNA128Snap(hFile As Long)
    Dim sData As String, sTemp As String, lOut7FFD As Long, lBank As Long, lCounter As Long
    
    ' // Read first three banks
    sData = Input(49152, #hFile)
    
    ' // PC
    sTemp = Input(2, #hFile)
    regPC = Asc(Right$(sTemp, 1)) * 256& + Asc(Left$(sTemp, 1))
    
    ' // Last out to 0x7FFD
    lOut7FFD = Asc(Input(1, #hFile))
    
    ' // Is TR-DOS paged? (ignored by vbSpec)
    sTemp = Input(1, #hFile)
    
    ' Setup first three banks
    For lCounter = 0 To 16383
        gRAMPage(5, lCounter) = Asc(Mid$(sData, lCounter + 1, 1))
        gRAMPage(2, lCounter) = Asc(Mid$(sData, lCounter + 16385, 1))
        gRAMPage(lOut7FFD And 7, lCounter) = Asc(Mid$(sData, lCounter + 32769, 1))
    Next lCounter
    
    lBank = 0
    Do While lBank < 8
        sData = Input(16384, #hFile)
        If (lBank = 5) Or (lBank = 2) Or (lBank = (lOut7FFD And 7)) Then
            lBank = lBank + 1
        End If
        If lBank < 8 Then
            For lCounter = 0 To 16383
                gRAMPage(lBank, lCounter) = Asc(Mid$(sData, lCounter + 1, 1))
            Next lCounter
            lBank = lBank + 1
        End If
    Loop
    
    glLastBorder = -1
    outb &H7FFD&, lOut7FFD
End Sub

Public Sub LoadSNASnap(sFileName As String)
    Dim hFile As Long, sData As String, iCounter As Long
    
    hFile = FreeFile
    Open sFileName For Binary As hFile
    
    sData = Input(1, #hFile)
    intI = Asc(sData)
    sData = Input(2, #hFile)
    regHL_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regDE_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regBC_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regAF_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    
    sData = Input(2, #hFile)
    regHL = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regDE = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    setBC Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regIY = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    sData = Input(2, #hFile)
    regIX = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    
    glLastBorder = -1
    
    sData = Input(1, #hFile)
    If Asc(sData) And 4 Then
        intIFF1 = True
        intIFF2 = True
    Else
        intIFF1 = False
        intIFF2 = False
    End If
    
    sData = Input(1, #hFile)
    intR = Asc(sData)
    intRTemp = intR
    
    sData = Input(2, #hFile)
    setAF Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    
    sData = Input(2, #hFile)
    regSP = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))

    sData = Input(1, #hFile)
    intIM = Asc(sData)
    
    sData = Input(1, #hFile)
    ' // Border color
    glNewBorder = (Asc(sData) And &H7&)
    
    If LOF(hFile) > 49180 Then
        LoadSNA128Snap hFile
        Close hFile
        
        initscreen
    
        ' // Set the initial border color
        frmMainWnd.BackColor = glNormalColor(glNewBorder)
    
        If (glEmulatedModel = 0) Or (glEmulatedModel = 5) Then SetEmulatedModel 1, gbSEBasicROM
    
        frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(sFileName)
    
        screenPaint
        resetKeyboard
    Else
        ' // Load a 48K Snapshot file
        err.Clear
        sData = Input(49153, #hFile)
        Close hFile
       
        For iCounter = 0 To 16383
            gRAMPage(5, iCounter) = Asc(Mid$(sData, iCounter + 1, 1))
            gRAMPage(1, iCounter) = Asc(Mid$(sData, iCounter + 16385, 1))
            gRAMPage(2, iCounter) = Asc(Mid$(sData, iCounter + 32769, 1))
        Next iCounter
    
        initscreen
    
        ' // Set the initial border color
        frmMainWnd.BackColor = glNormalColor(glNewBorder)
    
        ' // if not a 48K speccy or a TC2048, then emulate a 48K
        If glEmulatedModel <> 0 And glEmulatedModel <> 5 Then
            SetEmulatedModel 0, gbSEBasicROM
        ElseIf glEmulatedModel = 5 Then
            ' // If we're on a TC2048, ensure we're using the speccy-compatible screen mode
            outb 255, 0
        End If
    
        'MM 16.04.2003
        frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(sFileName)
    
        screenPaint
        resetKeyboard
        poppc
    End If
End Sub


Public Sub LoadZ80Snap(sFileName As String)
    Dim hFile As Long, sData As String, iCounter As Long
    Dim bCompressed As Boolean
    
    hFile = FreeFile
    Open sFileName For Binary As hFile
    
    glPageAt(0) = 8
    glPageAt(1) = 5
    glPageAt(2) = 1
    glPageAt(3) = 2
    glPageAt(4) = 8
    
    glLastBorder = -1
    Z80Reset
    
    If gbSoundEnabled Then AY8912_reset
    
    ' byte 0 - A register
    sData = Input(1, #hFile)
    regA = Asc(sData)
    ' byte 1 - F register
    sData = Input(1, #hFile)
    setF Asc(sData)
    ' bytes 2 + 3 - BC register pair (C first, then B)
    sData = Input(2, #hFile)
    setBC Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 4 + 5 - HL register pair
    sData = Input(2, #hFile)
    regHL = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 6 + 7 - PC (this is zero for v2.x or v3.0 Z80 files)
    sData = Input(2, #hFile)
    regPC = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 8 + 9 - SP
    sData = Input(2, #hFile)
    regSP = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' byte 10 - Interrupt register
    sData = Input(1, #hFile)
    intI = Asc(sData)
    ' byte 11 - Refresh register
    sData = Input(1, #hFile)
    intR = (Asc(sData) And 127)
    
    ' byte 12 - bitfield
    sData = Input(1, #hFile)
    ' if byte 12 = 255 then it must be treated as if it = 1, for compatibility with other emulators
    If Asc(sData) = 255 Then sData = Chr$(1)
    ' bit 0 - bit 7 of R
    If (Asc(sData) And 1) = 1 Then intR = intR Or 128
    intRTemp = intR
    ' bits 1,2 and 3 - border color
    glNewBorder = (Asc(sData) And 14) \ 2
    ' bit 4 - 1 if SamROM switched in (we don't care about this!)
    ' bit 5 - if 1 and PC<>0 then the snapshot is compressed using the
    '         rudimentary Z80 run-length encoding scheme
    If (Asc(sData) And &H20&) Then bCompressed = True
    ' bits 6 + 7 - no meaning

    ' bytes 13 + 14 - DE register pair
    sData = Input(2, #hFile)
    regDE = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 15 + 16 - BC' register pair
    sData = Input(2, #hFile)
    regBC_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 17 + 18 - DE' register pair
    sData = Input(2, #hFile)
    regDE_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 19 + 20 - HL' register pair
    sData = Input(2, #hFile)
    regHL_ = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' bytes 21 + 22 - AF' register pair (A first then F - not Z80 byte order!!)
    sData = Input(2, #hFile)
    regAF_ = Asc(Left$(sData, 1)) * 256& + Asc(Right$(sData, 1))
    ' byte 23 + 24 - IY register pair
    sData = Input(2, #hFile)
    regIY = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' byte 25 + 26 - IX register pair
    sData = Input(2, #hFile)
    regIX = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    ' byte 27 - Interrupt flipflop (0=DI, else EI)
    sData = Input(1, #hFile)
    If Asc(sData) = 0 Then
        intIFF1 = False
        intIFF2 = False
    Else
        intIFF1 = True
        intIFF2 = True
    End If
    ' byte 28 - IFF2 (ignored)
    sData = Input(1, #hFile)
    ' byte 29 - Interrupt mode (bits 2 - 7 contain info about joystick modes etc, which we ignore)
    sData = Input(1, #hFile)
    intIM = Asc(sData) And 3
    
    If regPC = 0 Then
        ' This is a V2 or V3 Z80 file
        ReadZ80V2orV3Snap hFile
        Close hFile
    Else
        ' // V1 .Z80 snapshots are all 48K
        
        ' // if not a 48K speccy or a TC2048, then emulate a 48K
        If glEmulatedModel <> 0 And glEmulatedModel <> 5 Then
            SetEmulatedModel 0, gbSEBasicROM
        ElseIf glEmulatedModel = 5 Then
            ' // If we're on a TC2048, ensure we're using the speccy-compatible screen mode
            outb 255, 0
        End If
        ' // PC<>0, so lets check to see if this is a compressed V1 Z80 file
        If bCompressed Then
            ' Uncompress the RAM data
            ReadZ80V1Snap hFile
            Close hFile
        Else
            ' // Uncompressed Z80 file
            sData = Input(49153, #hFile)
            Close hFile
            
            ' // Copy the RAM data to addressable memory space
            For iCounter = 16384 To 65535
                gRAMPage(glPageAt(iCounter \ 16384), iCounter And 16383) = Asc(Mid$(sData, iCounter - 16383, 1))
            Next iCounter
        End If
    End If
    
    initscreen
    
    ' // Set the initial border color
    Select Case glNewBorder
    Case 0
        frmMainWnd.BackColor = 0&
    Case 1
        frmMainWnd.BackColor = RGB(0, 0, 192)
    Case 2
        frmMainWnd.BackColor = RGB(192, 0, 0)
    Case 3
        frmMainWnd.BackColor = RGB(192, 0, 192)
    Case 4
        frmMainWnd.BackColor = RGB(0, 192, 0)
    Case 5
        frmMainWnd.BackColor = RGB(0, 192, 192)
    Case 6
        frmMainWnd.BackColor = RGB(192, 192, 0)
    Case 7
        frmMainWnd.BackColor = RGB(192, 192, 192)
    End Select

    frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(sFileName)
    
    initscreen
    screenPaint
    resetKeyboard
        
    gpicDisplay.REFRESH
    DoEvents
End Sub


Sub Main()
    InitColorArrays
    InitScreenMemTable
    InitReverseBitValues ' // Used by the ZX Printer emulation - see modSpectrum
    
    ' // RGW's performance improvements
    InitScreenMask
    InitScreenIndexs
      
    frmMainWnd.Show
    
    Set gpicDisplay = frmMainWnd.picDisplay
    gpicDC = gpicDisplay.hdc
    
    CreateScreenBuffer
    glDisplayWidth = val(GetSetting("Grok", "vbSpec", "DisplayWidth", "256"))
    glDisplayHeight = val(GetSetting("Grok", "vbSpec", "DisplayHeight", "192"))
    bFullScreen = CBool(GetSetting("Grok", "vbSpec", "FullScreen", Trim(CStr(CInt(False)))))
    SetDisplaySize glDisplayWidth, glDisplayHeight
    AY8912_init 1773000, WAVE_FREQUENCY, 8
    
    glInterruptDelay = val(GetSetting("Grok", "vbSpec", "InterruptDelay", "20"))
    
    If val(GetSetting("Grok", "vbSpec", "TapeControlsVisible", "0")) <> 0 Then
        frmTapePlayer.Show 0, frmMainWnd
        frmMainWnd.mnuOptions(4).Checked = True
    End If
    
    If val(GetSetting("Grok", "vbSpec", "EmulateZXPrinter", "0")) <> 0 Then
        frmZXPrinter.Show 0, frmMainWnd
        frmMainWnd.mnuOptions(5).Checked = True
    End If
    ' // Make sure the main window has focus
    frmMainWnd.SetFocus
    
    ' // Initialize everything
    initParity
    Z80Reset
    
    If val(GetSetting("Grok", "vbSpec", "SEBasicROM", "0")) <> 0 Then
        gbSEBasicROM = True
    Else
        gbSEBasicROM = False
    End If
    
    SetEmulatedModel val(GetSetting("Grok", "vbSpec", "EmulatedModel", "0")), gbSEBasicROM
    
    ' // Emulated Mouse support (MOUSE_NONE by default)
    glMouseType = val(GetSetting("Grok", "vbSpec", "MouseType", CStr(MOUSE_NONE)))
    gbMouseGlobal = val(GetSetting("Grok", "vbSpec", "MouseGlobal", "0"))
    
    If glMouseType = MOUSE_NONE Then
        frmMainWnd.picDisplay.MousePointer = 0
    Else
        frmMainWnd.picDisplay.MousePointer = 99
    End If
    
    InitAmigaMouseTables
    
    initscreen
    resetKeyboard
    
    timeBeginPeriod (1)
    
    ' // If we have a command line parameter, try to open it as a snapshot/tape/rom
    If Command$ <> "" Then frmMainWnd.FileOpenDialog Command$
    
    glInterruptTimer = timeGetTime()
    gbSoundEnabled = InitializeWaveOut()
        
    ' // Begin the Z80 execution loop, this drives the whole emulation
    execute
End Sub

Sub InitScreenMask()
   ' RGW Prefill the screen color & attribute lookup table
   '     with all possible combinations
   '     When drawing the screen in the bit buffer
   '     a simple lookup produces the required bytes
   
   Dim fC       As Long   ' fore color
   Dim BC       As Long   ' back color
   Dim Bright   As Long
   Dim Flash    As Long
   
   Dim bits     As Long
   Dim Color(1) As Long
   Dim lTemp    As Long
               
   For Flash = 0 To 1
      For Bright = 0 To 1
         For fC = 0 To 7
            For BC = 0 To 7
               For bits = 0 To 255
                  If Flash = 0 Then
                     Color(1) = fC + (Bright * 8)
                     Color(0) = BC + (Bright * 8)
                  Else
                     Color(1) = BC + (Bright * 8)
                     Color(0) = fC + (Bright * 8)
                  End If
                  lTemp = (Flash * 128) + (Bright * 64) + (BC * 8) + fC
                  gtBitTable(bits, lTemp).dw0 = (Color(Abs((bits And 16) = 16)) * 16777216) + _
                                             (Color(Abs((bits And 32) = 32)) * 65536) + _
                                             (Color(Abs((bits And 64) = 64)) * 256) + _
                                              Color(Abs((bits And 128) = 128))
               
                  gtBitTable(bits, lTemp).dw1 = (Color(Abs((bits And 1) = 1)) * 16777216) + _
                                             (Color(Abs((bits And 2) = 2)) * 65536) + _
                                             (Color(Abs((bits And 4) = 4)) * 256) + _
                                              Color(Abs((bits And 8) = 8))
               Next bits
            Next BC
         Next fC
      Next Bright
   Next Flash
End Sub
Public Sub InitColorArrays()
    bmiBuffer.bmiColors(0).rgbRed = 0
    bmiBuffer.bmiColors(0).rgbGreen = 0
    bmiBuffer.bmiColors(0).rgbBlue = 0
    
    bmiBuffer.bmiColors(1).rgbRed = 0
    bmiBuffer.bmiColors(1).rgbGreen = 0
    bmiBuffer.bmiColors(1).rgbBlue = 192
    
    bmiBuffer.bmiColors(2).rgbRed = 192
    bmiBuffer.bmiColors(2).rgbGreen = 0
    bmiBuffer.bmiColors(2).rgbBlue = 0
    
    bmiBuffer.bmiColors(3).rgbRed = 192
    bmiBuffer.bmiColors(3).rgbGreen = 0
    bmiBuffer.bmiColors(3).rgbBlue = 192
    
    bmiBuffer.bmiColors(4).rgbRed = 0
    bmiBuffer.bmiColors(4).rgbGreen = 192
    bmiBuffer.bmiColors(4).rgbBlue = 0
    
    bmiBuffer.bmiColors(5).rgbRed = 0
    bmiBuffer.bmiColors(5).rgbGreen = 192
    bmiBuffer.bmiColors(5).rgbBlue = 192
    
    bmiBuffer.bmiColors(6).rgbRed = 192
    bmiBuffer.bmiColors(6).rgbGreen = 192
    bmiBuffer.bmiColors(6).rgbBlue = 0
    
    bmiBuffer.bmiColors(7).rgbRed = 192
    bmiBuffer.bmiColors(7).rgbGreen = 192
    bmiBuffer.bmiColors(7).rgbBlue = 192
    
    bmiBuffer.bmiColors(8).rgbRed = 0
    bmiBuffer.bmiColors(8).rgbGreen = 0
    bmiBuffer.bmiColors(8).rgbBlue = 0
    
    bmiBuffer.bmiColors(9).rgbRed = 0
    bmiBuffer.bmiColors(9).rgbGreen = 0
    bmiBuffer.bmiColors(9).rgbBlue = 255
    
    bmiBuffer.bmiColors(10).rgbRed = 255
    bmiBuffer.bmiColors(10).rgbGreen = 0
    bmiBuffer.bmiColors(10).rgbBlue = 0
    
    bmiBuffer.bmiColors(11).rgbRed = 255
    bmiBuffer.bmiColors(11).rgbGreen = 0
    bmiBuffer.bmiColors(11).rgbBlue = 255
    
    bmiBuffer.bmiColors(12).rgbRed = 0
    bmiBuffer.bmiColors(12).rgbGreen = 255
    bmiBuffer.bmiColors(12).rgbBlue = 0
    
    bmiBuffer.bmiColors(13).rgbRed = 0
    bmiBuffer.bmiColors(13).rgbGreen = 255
    bmiBuffer.bmiColors(13).rgbBlue = 255
    
    bmiBuffer.bmiColors(14).rgbRed = 255
    bmiBuffer.bmiColors(14).rgbGreen = 255
    bmiBuffer.bmiColors(14).rgbBlue = 0
    
    bmiBuffer.bmiColors(15).rgbRed = 255
    bmiBuffer.bmiColors(15).rgbGreen = 255
    bmiBuffer.bmiColors(15).rgbBlue = 255
    
    glBrightColor(0) = 0
    glBrightColor(1) = RGB(0, 0, 255)
    glBrightColor(2) = RGB(255, 0, 0)
    glBrightColor(3) = RGB(255, 0, 255)
    glBrightColor(4) = RGB(0, 255, 0)
    glBrightColor(5) = RGB(0, 255, 255)
    glBrightColor(6) = RGB(255, 255, 0)
    glBrightColor(7) = RGB(255, 255, 255)
    glNormalColor(0) = 0
    glNormalColor(1) = RGB(0, 0, 192)
    glNormalColor(2) = RGB(192, 0, 0)
    glNormalColor(3) = RGB(192, 0, 192)
    glNormalColor(4) = RGB(0, 192, 0)
    glNormalColor(5) = RGB(0, 192, 192)
    glNormalColor(6) = RGB(192, 192, 0)
    glNormalColor(7) = RGB(192, 192, 192)
End Sub



Private Sub ReadZ80V1Snap(hFile As Long)
    Dim lDataLen As Long, sData As String, lBlockLen As Long
    Dim lCounter As Long, lMemPos As Long, lBlockCounter As Long
 
    lDataLen = LOF(hFile) - Seek(hFile) + 1
    ' // read the compressed data into sData
    sData = Input(lDataLen, #hFile)
        
    ' // Uncompress the block to memory
    lCounter = 1
    lMemPos = 16384
    Do
        If Asc(Mid$(sData, lCounter, 1)) = &HED& Then
            If Asc(Mid$(sData, lCounter + 1, 1)) = &HED& Then
                ' // This is an encoded block
                lCounter = lCounter + 2
                lBlockLen = Asc(Mid$(sData, lCounter, 1))
                lCounter = lCounter + 1
                For lBlockCounter = 1 To lBlockLen
                    gRAMPage(glPageAt(lMemPos \ 16384), lMemPos And 16383) = Asc(Mid$(sData, lCounter, 1))
                    lMemPos = lMemPos + 1
                Next lBlockCounter
            Else
                ' // Just a single ED, write it out
                gRAMPage(glPageAt(lMemPos \ 16384), lMemPos And 16383) = &HED&
                lMemPos = lMemPos + 1
            End If
        Else
            gRAMPage(glPageAt(lMemPos \ 16384), lMemPos And 16383) = Asc(Mid$(sData, lCounter, 1))
            lMemPos = lMemPos + 1
        End If
        lCounter = lCounter + 1
    Loop Until lCounter > Len(sData) - 4
    
    If Mid$(sData, lCounter, 4) <> Chr$(0) & Chr$(&HED) & Chr$(&HED) & Chr$(0) Then
        MsgBox "Error in compressed Z80 file. Block end marker 0x00EDED00 is not present."
    End If

End Sub

Private Sub ReadZ80V2orV3Snap(hFile As Long)
    Dim lHeaderLen As Long, sData As String
    Dim lCounter As Long, bHardwareSupported As Boolean
    Dim lMemPage As Long, lBlockCounter As Long, b128K As Boolean, lMemPos As Long
    Dim lOutFFFD As Long, bTimex As Boolean
    
    sData = Input(2, #hFile)
    lHeaderLen = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
    
    ' // offset 32 - 2 bytes - PC
    If lCounter < lHeaderLen Then
        sData = Input(2, #hFile)
        regPC = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
        lCounter = lCounter + 2
    End If
    
    ' // offset 34 - 1 byte - hardware mode
    If lCounter < lHeaderLen Then
        sData = Input(1, #hFile)
        Select Case Asc(sData)
        Case 0 ' // 48K spectrum
            bHardwareSupported = True
            ' // if not currently emulating a 48K speccy or a TC2048, then emulate a 48K
            If glEmulatedModel <> 0 And glEmulatedModel <> 5 Then
                SetEmulatedModel 0, gbSEBasicROM
            ElseIf glEmulatedModel = 5 Then
                ' // If we're on a TC2048, ensure we're using the speccy-compatible screen mode
                outb 255, 0
            End If
        Case 1 ' // 48K spectrum + Interface 1
            bHardwareSupported = True
            SetEmulatedModel 0, gbSEBasicROM
        Case 2 ' // SamROM
            sData = "SamROM"
            bHardwareSupported = False
        Case 3
            If lHeaderLen = 23 Then
                sData = "128K Spectrum"
                bHardwareSupported = True
                b128K = True
                If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 1, gbSEBasicROM
            Else
                sData = "48K Spectrum + M.G.T."
                bHardwareSupported = True
                SetEmulatedModel 0, gbSEBasicROM
            End If
        Case 4
            If lHeaderLen = 23 Then
                sData = "128K Spectrum + Interface 1"
                bHardwareSupported = True
                b128K = True
                If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 1, gbSEBasicROM
            Else
                sData = "128K Spectrum"
                bHardwareSupported = True
                b128K = True
                If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 1, gbSEBasicROM
            End If
        Case 5
            bHardwareSupported = True
            sData = "128K Spectrum + Interface 1"
            b128K = True
            If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 1, gbSEBasicROM
        Case 6
            bHardwareSupported = False
            sData = "128K Spectrum + M.G.T."
            b128K = True
            If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 1, gbSEBasicROM
        Case 7
            bHardwareSupported = False
            sData = "ZX Spectrum +3"
            b128K = True
            If glEmulatedModel = 0 Or glEmulatedModel = 5 Then SetEmulatedModel 2, gbSEBasicROM
        Case 14
            bHardwareSupported = True
            bTimex = True
            SetEmulatedModel 5, gbSEBasicROM
            ' // If we're on a TC2048, ensure we're using the speccy-compatible screen mode
            outb 255, 0
        Case Else
            bHardwareSupported = False
            sData = "Unknown hardware platform"
        End Select
        
        If bHardwareSupported = False Then
            MsgBox "vbSpec does not support the required hardware platform (" & sData & ") for this snapshot. This snapshot may not function correctly in vbSpec."
        End If
        
        lCounter = lCounter + 1
    End If
    
    ' // offset 35 - 1 byte - last out to 0x7FFD - not required for 48K spectrum
    If lCounter < lHeaderLen Then
        If b128K Then
            outb &H7FFD&, CLng(Asc(Input(1, #hFile)))
        Else
            sData = Input(1, #hFile)
            glPageAt(0) = 8
            glPageAt(1) = 5
            glPageAt(2) = 1
            glPageAt(3) = 2
            glPageAt(4) = 8
        End If
        lCounter = lCounter + 1
    End If
    
    ' // offset 36 - 1 byte - 0xFF if Interface 1 ROM is paged in
    If lCounter < lHeaderLen Then
        sData = Input(1, #hFile)
        lCounter = lCounter + 1
        If bTimex Then
            outb &HFF&, CLng(Asc(sData))
        ElseIf Asc(sData) = 255 Then
            MsgBox "This snapshot was saved with the Interface 1 ROM paged in. It may not run correctly in vbSpec."
        End If
    End If
    
    ' // offset 37 - 1 byte (bit 0: 1=intR emulation on, bit 1: 1=LDIR emulation on
    If lCounter < lHeaderLen Then
        sData = Input(1, #hFile)
        lCounter = lCounter + 1
    End If
    
    ' // offset 38 - Last out to 0xFFFD (+2/+3 sound chip register number)
    If lCounter < lHeaderLen Then
        lOutFFFD = Asc(Input(1, #hFile))
        lCounter = lCounter + 1
    End If
    
    ' // offset 39 - 16 bytes - contents of the sound chip registers
    If lCounter < lHeaderLen Then
        If b128K Then
            AYWriteReg 0, Asc(Input(1, #hFile))
            AYWriteReg 1, Asc(Input(1, #hFile))
            AYWriteReg 2, Asc(Input(1, #hFile))
            AYWriteReg 3, Asc(Input(1, #hFile))
            AYWriteReg 4, Asc(Input(1, #hFile))
            AYWriteReg 5, Asc(Input(1, #hFile))
            AYWriteReg 6, Asc(Input(1, #hFile))
            AYWriteReg 7, Asc(Input(1, #hFile))
            AYWriteReg 8, Asc(Input(1, #hFile))
            AYWriteReg 9, Asc(Input(1, #hFile))
            AYWriteReg 10, Asc(Input(1, #hFile))
            AYWriteReg 11, Asc(Input(1, #hFile))
            AYWriteReg 12, Asc(Input(1, #hFile))
            AYWriteReg 13, Asc(Input(1, #hFile))
            AYWriteReg 14, Asc(Input(1, #hFile))
            AYWriteReg 15, Asc(Input(1, #hFile))
        Else
            sData = Input(16, #hFile)
        End If
        lCounter = lCounter + 16
    End If
    
    If b128K Then
        outb &HFFFD&, lOutFFFD
    End If
    
    ' // read the remaining bytes of the header (we don't care what information they hold)
    If lCounter < lHeaderLen Then
        sData = Input(lHeaderLen - lCounter, #hFile)
    End If
    
    Do
        ' // read a block
        sData = Input(2, #hFile)
        If EOF(hFile) Then Exit Do
        lHeaderLen = Asc(Right$(sData, 1)) * 256& + Asc(Left$(sData, 1))
        sData = Input(1, #hFile)
        Select Case Asc(sData)
        Case 0 ' // Spectrum ROM
            If b128K Then lMemPage = 9 Else lMemPage = 8
        Case 1 ' // Interface 1 ROM, or similar (we discard these blocks)
            lMemPage = -1
        Case 2 ' // 128K ROM (reset)
            If b128K Then lMemPage = 8 Else lMemPage = -1
        Case 3 ' // Page 0 (not used by 48K Spectrum)
            If b128K Then lMemPage = 0 Else lMemPage = -1
        Case 4 ' // Page 1 RAM at 0x8000
            lMemPage = 1
        Case 5 ' // Page 2 RAM at 0xC000
            lMemPage = 2
        Case 6 ' // Page 3 (not used by 48K Spectrum)
            If b128K Then lMemPage = 3 Else lMemPage = -1
        Case 7 ' // Page 4 (not used by 48K Spectrum)
            If b128K Then lMemPage = 4 Else lMemPage = -1
        Case 8 ' // Page 5 RAM at 0x4000
            lMemPage = 5
        Case 9 ' // Page 6 (not used by 48K Spectrum)
            If b128K Then lMemPage = 6 Else lMemPage = -1
        Case 10 ' // Page 7 (not used by 48K Spectrum)
            If b128K Then lMemPage = 7 Else lMemPage = -1
        Case 11 ' // Multiface ROM
            lMemPage = -1
        End Select
        
        
        If lMemPage <> -1 Then
            If lHeaderLen = &HFFFF& Then
                sData = Input(16384, #hFile)
                ' Not a compressed block, just copy it straight into RAM
                For lCounter = 0 To 16383
                    gRAMPage(lMemPage, lCounter) = Asc(Mid$(sData, lCounter + 1, 1))
                Next lCounter
            Else
                sData = Input(lHeaderLen, #hFile)
                ' // Uncompress the block to memory
                lCounter = 1
                lMemPos = 0
                Do
                    If Asc(Mid$(sData, lCounter, 1)) = &HED& Then
                        If Asc(Mid$(sData, lCounter + 1, 1)) = &HED& Then
                            ' // This is an encoded block
                            lCounter = lCounter + 2
                            lHeaderLen = Asc(Mid$(sData, lCounter, 1))
                            lCounter = lCounter + 1
                            If lMemPos + lHeaderLen - 1 > 16383 Then GoTo ErrBlockTooBig
                            For lBlockCounter = 0 To lHeaderLen - 1
                                gRAMPage(lMemPage, lMemPos + lBlockCounter) = Asc(Mid$(sData, lCounter, 1))
                            Next lBlockCounter
                            lMemPos = lMemPos + lBlockCounter
                        Else
                            ' // Just a single ED, write it out
                            gRAMPage(lMemPage, lMemPos) = &HED&
                            lMemPos = lMemPos + 1
                        End If
                    Else
                        gRAMPage(lMemPage, lMemPos) = Asc(Mid$(sData, lCounter, 1))
                        lMemPos = lMemPos + 1
                    End If
                    
                    If lMemPos > 16384 Then GoTo ErrBlockTooBig
                    lCounter = lCounter + 1
                Loop Until lCounter > Len(sData)
            End If
        End If
    Loop Until EOF(hFile)
Exit Sub

ErrBlockTooBig:
    MsgBox "Errors were encountered in the z80 file. Compressed memory block [" & CStr(lMemPage + 3) & "] has an uncompressed length of more than 16384 bytes.", vbOKOnly Or vbExclamation, "vbSpec"
End Sub

Sub resetKeyboard()
    keyB_SPC = &HFF&
    keyH_ENT = &HFF&
    keyY_P = &HFF&
    key6_0 = &HFF&
    key1_5 = &HFF&
    keyQ_T = &HFF&
    keyA_G = &HFF&
    keyCAPS_V = &HFF&
End Sub


Public Sub SaveROM(sFileName As String)
    Dim hFile As Long, lCounter As Long
    
    On Error GoTo SaveROM_Err
    
    hFile = FreeFile
    Open sFileName For Output As hFile
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(0), lCounter));
    Next lCounter
    
SaveROM_Err:
    Close hFile
End Sub

Public Sub SaveScreenBMP(sFileName As String)
    On Error Resume Next
    
    SavePicture frmMainWnd.picDisplay.Image, sFileName
End Sub

Public Sub SaveScreenSCR(sFileName As String)
    Dim hFile As Long, n As Long
    
    hFile = FreeFile
    Open sFileName For Output As hFile
    
    If glUseScreen = 1001 Then
        ' // TC2048 second screen
        For n = 8192 To 15103
            Print #hFile, Chr$(gRAMPage(5, n));
        Next n
    ElseIf glUseScreen = 1002 Then
        ' // TC2048 hicolour
        
        ' // Save the mono bitmap data first
        For n = 0 To 6143
            Print #hFile, Chr$(gRAMPage(5, n));
        Next n
        ' // Immediately followed by the colour data
        For n = 8192 To 14335
            Print #hFile, Chr$(gRAMPage(5, n));
        Next n
    ElseIf glUseScreen = 1006 Then
        ' // TC2048 hires
        ' // Save columns 0,2,4,6... first
        For n = 0 To 6143
            Print #hFile, Chr$(gRAMPage(5, n));
        Next n
        ' // Immediately followed by columns 1,3,5,7...
        For n = 8192 To 14335
            Print #hFile, Chr$(gRAMPage(5, n));
        Next n
        ' // And finally an extra byte to indicate the screen colour
        Print #hFile, Chr$(glTC2048LastFFOut);
    ElseIf glUseScreen < 1000 Then
        For n = 0 To 6911
            Print #hFile, Chr$(gRAMPage(glUseScreen, n));
        Next n
    End If
    
    Close #hFile
End Sub


Private Sub SaveSNA128Snap(hFile As Long)
    Dim lBank As Long, lCounter As Long
    
    Print #hFile, Chr$(intI);
    Print #hFile, Chr$(regHL_ And &HFF&); Chr$(regHL_ \ 256&);
    Print #hFile, Chr$(regDE_ And &HFF&); Chr$(regDE_ \ 256&);
    Print #hFile, Chr$(regBC_ And &HFF&); Chr$(regBC_ \ 256&);
    Print #hFile, Chr$(regAF_ And &HFF&); Chr$(regAF_ \ 256&);
    
    Print #hFile, Chr$(regHL And &HFF&); Chr$(regHL \ 256&);
    Print #hFile, Chr$(regDE And &HFF&); Chr$(regDE \ 256&);
    Print #hFile, Chr$(regC); Chr$(regB);
    Print #hFile, Chr$(regIY And &HFF&); Chr$(regIY \ 256&);
    Print #hFile, Chr$(regIX And &HFF&); Chr$(regIX \ 256&);
    
    ' Interrupt flipflops
    If intIFF1 = True Then
        Print #hFile, Chr$(4);
    Else
        Print #hFile, Chr$(0);
    End If

    ' R
    intRTemp = intRTemp And 127
    Print #hFile, Chr$((intR And &H80&) Or intRTemp);

    ' // AF
    Print #hFile, Chr$(getAF And &HFF&); Chr$(getAF \ 256&);
    
    ' // SP
    Print #hFile, Chr$(regSP And &HFF&); Chr$(regSP \ 256&);
    
    ' // Interrupt Mode
    Print #hFile, Chr$(intIM);
    
    Print #hFile, Chr$(GetBorderIndex(frmMainWnd.BackColor));

    ' // Save the three currently-paged RAM banks
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(1), lCounter));
    Next lCounter
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(2), lCounter));
    Next lCounter
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(3), lCounter));
    Next lCounter

    ' // PC
    Print #hFile, Chr$(regPC And &HFF&); Chr$(regPC \ 256&);
    
    ' // Last out to 0x7FFD
    Print #hFile, Chr$(glLastOut7FFD);
    
    ' // Is TR-DOS paged? (0=not paged, 1=paged)
    Print #hFile, Chr$(0);
    
    ' // Save the remaining RAM banks
    lBank = 0
    Do While lBank < 8
        If lBank <> glPageAt(1) And lBank <> glPageAt(2) And lBank <> glPageAt(3) Then
            For lCounter = 0 To 16383
                Print #hFile, Chr$(gRAMPage(lBank, lCounter));
            Next lCounter
        End If
        lBank = lBank + 1
    Loop
End Sub

Public Sub SaveSNASnap(sFileName As String)
    Dim hFile As Long, sData As String, lCounter As Long
    
    hFile = FreeFile
    Open sFileName For Output As hFile
    
    If (glEmulatedModel <> 0) And (glEmulatedModel <> 5) Then
        ' // We're running in 128 mode
        SaveSNA128Snap hFile
        Close hFile
        Exit Sub
    End If
    pushpc
    
    Print #hFile, Chr$(intI);
    Print #hFile, Chr$(regHL_ And &HFF&); Chr$(regHL_ \ 256&);
    Print #hFile, Chr$(regDE_ And &HFF&); Chr$(regDE_ \ 256&);
    Print #hFile, Chr$(regBC_ And &HFF&); Chr$(regBC_ \ 256&);
    Print #hFile, Chr$(regAF_ And &HFF&); Chr$(regAF_ \ 256&);
    
    Print #hFile, Chr$(regHL And &HFF&); Chr$(regHL \ 256&);
    Print #hFile, Chr$(regDE And &HFF&); Chr$(regDE \ 256&);
    Print #hFile, Chr$(regC); Chr$(regB);
    Print #hFile, Chr$(regIY And &HFF&); Chr$(regIY \ 256&);
    Print #hFile, Chr$(regIX And &HFF&); Chr$(regIX \ 256&);
    
    ' Interrupt flipflops
    If intIFF1 = True Then
        Print #hFile, Chr$(4);
    Else
        Print #hFile, Chr$(0);
    End If

    ' R
    intRTemp = intRTemp And 127
    Print #hFile, Chr$((intR And &H80&) Or intRTemp);

    ' // AF
    Print #hFile, Chr$(getAF And &HFF&); Chr$(getAF \ 256&);
    
    ' // SP
    Print #hFile, Chr$(regSP And &HFF&); Chr$(regSP \ 256&);
    
    ' // Interrupt Mode
    Print #hFile, Chr$(intIM);
    
    Print #hFile, Chr$(GetBorderIndex(frmMainWnd.BackColor));

    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(1), lCounter));
    Next lCounter
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(2), lCounter));
    Next lCounter
    For lCounter = 0 To 16383
        Print #hFile, Chr$(gRAMPage(glPageAt(3), lCounter));
    Next lCounter
    
    Close hFile
    poppc
End Sub

Public Sub SaveZ80Snap(sFileName As String)
    Dim hFile As Long, sData As String, lCounter As Long, lBufSize As Long
    
    hFile = FreeFile
    Open sFileName For Output As hFile
    
    ' A,F
    Print #hFile, Chr$(regA); Chr$(getF);
    ' BC
    Print #hFile, Chr$(regC); Chr$(regB);
    ' HL
    Print #hFile, Chr$(regHL And &HFF&); Chr$(regHL \ 256&);
    ' Set PC to zero to indicate a v2.01 Z80 file
    Print #hFile, Chr$(0); Chr$(0);
    ' SP
    Print #hFile, Chr$(regSP And &HFF&); Chr$(regSP \ 256&);
    ' I
    Print #hFile, Chr$(intI);
    ' R (7 bits)
    intRTemp = intRTemp And &H7F&
    Print #hFile, Chr$(intRTemp);
    ' bitfield
    Print #hFile, Chr$((IIf(intR And &H80& = &H80, 1, 0)) Or (GetBorderIndex(frmMainWnd.BackColor) * 2));
    ' DE
    Print #hFile, Chr$(regDE And &HFF&); Chr$(regDE \ 256&);
    ' BC'
    Print #hFile, Chr$(regBC_ And &HFF&); Chr$(regBC_ \ 256&);
    ' DE'
    Print #hFile, Chr$(regDE_ And &HFF&); Chr$(regDE_ \ 256&);
    ' HL'
    Print #hFile, Chr$(regHL_ And &HFF&); Chr$(regHL_ \ 256&);
    ' AF'
    Print #hFile, Chr$(regAF_ \ 256&); Chr$(regAF_ And &HFF&);
    ' IY
    Print #hFile, Chr$(regIY And &HFF&); Chr$(regIY \ 256&);
    ' IX
    Print #hFile, Chr$(regIX And &HFF&); Chr$(regIX \ 256&);
    ' Interrupt flipflops
    If intIFF1 = True Then
        Print #hFile, Chr$(255);
        Print #hFile, Chr$(255);
    Else
        Print #hFile, Chr$(0);
        Print #hFile, Chr$(0);
    End If
    ' // Interrupt Mode
    Print #hFile, Chr$(intIM);
    
    ' // V2.01 info
    Print #hFile, Chr$(23) & Chr$(0);
    ' PC
    Print #hFile, Chr$(regPC And &HFF&); Chr$(regPC \ 256&);
    ' Hardware mode
    If (glEmulatedModel = 0) Then
        Print #hFile, Chr$(0); ' // 48K Spectrum
        Print #hFile, Chr$(0);
    ElseIf (glEmulatedModel = 5) Then
        Print #hFile, Chr$(14);
        Print #hFile, Chr$(0);
    Else
        Print #hFile, Chr$(3); ' // 128K Spectrum
        ' Last out to 7FFD
        Print #hFile, Chr$(glLastOut7FFD);
    End If
    
    If glEmulatedModel = 5 Then
        '  Last out to 00FF
        Print #hFile, Chr$(glTC2048LastFFOut And &HFF&);
    Else
        ' IF.1 Paged in
        Print #hFile, Chr$(0);
    End If
    ' 1=R emulation on,2=LDIR emulation on
    Print #hFile, Chr$(3);
    ' Last out to FFFD
    Print #hFile, Chr$(glSoundRegister);
    ' AY-3-8912 register contents
    For lCounter = 0 To 15
        Print #hFile, Chr$(AYPSG.Regs(lCounter));
    Next lCounter
    
    If glEmulatedModel = 0 Or glEmulatedModel = 5 Then
        ' // Block 1
        sData = ""
        lBufSize = CompressMemoryBlock(glPageAt(1), sData)
        ' Buffer length
        Print #hFile, Chr$(lBufSize And &HFF&); Chr$(lBufSize \ 256&);
        ' Block number
        Print #hFile, Chr$(glPageAt(1) + 3);
        Print #hFile, sData;
        ' // Block 2
        sData = ""
        lBufSize = CompressMemoryBlock(glPageAt(2), sData)
        ' Buffer length
        Print #hFile, Chr$(lBufSize And &HFF&); Chr$(lBufSize \ 256&);
        ' Block number
        Print #hFile, Chr$(glPageAt(2) + 3);
        Print #hFile, sData;
        ' // Block 3
        sData = ""
        lBufSize = CompressMemoryBlock(glPageAt(3), sData)
        ' Buffer length
        Print #hFile, Chr$(lBufSize And &HFF&); Chr$(lBufSize \ 256&);
        ' Block number
        Print #hFile, Chr$(glPageAt(3) + 3);
        Print #hFile, sData;
    Else
        For lCounter = 0 To 7
            sData = ""
            lBufSize = CompressMemoryBlock(lCounter, sData)
            ' Buffer length
            Print #hFile, Chr$(lBufSize And &HFF&); Chr$(lBufSize \ 256&);
            ' Block number
            Print #hFile, Chr$(lCounter + 3);
            Print #hFile, sData;
        Next lCounter
    End If
        
    Close hFile
    resetKeyboard
End Sub


Private Sub SEPatch128ROM()
    gRAMPage(8, 576) = 0
    gRAMPage(8, 577) = 0
    gRAMPage(8, 578) = 0
    gRAMPage(8, &H37F) = 0
    gRAMPage(8, &H380) = 0
    gRAMPage(8, &H381) = &H15
    gRAMPage(8, &H382) = 0
    gRAMPage(8, &H383) = 0
    gRAMPage(8, &H384) = 0
    gRAMPage(8, &H3A3) = &H3E
    gRAMPage(8, &H3A4) = &H20
    gRAMPage(8, &H3A5) = &HD7
    gRAMPage(8, &H3A6) = &H0
    gRAMPage(8, &H3A7) = &H0
    gRAMPage(8, &H3A8) = &H0
    gRAMPage(8, &H33B4) = &HCB
End Sub

Private Sub SEPatchPlus2ROM()
    gRAMPage(8, 576) = 0
    gRAMPage(8, 577) = 0
    gRAMPage(8, 578) = 0
    gRAMPage(8, &H37F) = 0
    gRAMPage(8, &H380) = 0
    gRAMPage(8, &H381) = &H15
    gRAMPage(8, &H382) = 0
    gRAMPage(8, &H383) = 0
    gRAMPage(8, &H384) = 0
    gRAMPage(8, &H3A3) = &H3E
    gRAMPage(8, &H3A4) = &H20
    gRAMPage(8, &H3A5) = &HD7
    gRAMPage(8, &H3A6) = &H0
    gRAMPage(8, &H3A7) = &H0
    gRAMPage(8, &H3A8) = &H0
    gRAMPage(8, &H33DA) = &HCB
End Sub


Public Sub SetDisplaySize(lWidth As Long, lHeight As Long)
    Dim i As Integer, y As Long, X As Long
    
    'MM 16.04.2003
    Dim rectWindow As RECT
    Dim pointWindow As POINT
    Dim lNewTop As Long, lNewLeft As Long, lRightMargin As Long, lBottomMargin As Long
    Dim lTopMargin As Long
    
    'No full screen modus
    If Not bFullScreen Then
        'Prepear window
        frmMainWnd.FullScreenOff
    End If
    
    y = (GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU) + GetSystemMetrics(SM_CYFRAME) * 2 + frmMainWnd.picStatus.Height) * Screen.TwipsPerPixelY
    X = (GetSystemMetrics(SM_CXFRAME) * 2) * Screen.TwipsPerPixelX
    glDisplayWidth = lWidth
    glDisplayHeight = lHeight
    glDisplayVSource = lHeight - 1
    glDisplayVSize = -lHeight
    
    SaveSetting "Grok", "vbSpec", "DisplayWidth", CStr(glDisplayWidth)
    SaveSetting "Grok", "vbSpec", "DisplayHeight", CStr(glDisplayHeight)
    'MM 16.04.2003
    SaveSetting "Grok", "vbSpec", "FullScreen", CInt(bFullScreen)
    
    Select Case lWidth
    Case 256 ' // Standard
        'MM 16.04.2003
        glDisplayXMultiplier = 1
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, 4200 + X, frmMainWnd.Height
        End If
    Case 512
        glDisplayXMultiplier = 2
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, 8040 + X, frmMainWnd.Height
        End If
    Case 768
        glDisplayXMultiplier = 3
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, 11880 + X, frmMainWnd.Height
        End If
    End Select
        
    Select Case lHeight
    Case 192
        glDisplayYMultiplier = 1
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, frmMainWnd.Width, 3300 + y
        End If
    Case 384
        glDisplayYMultiplier = 2
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, frmMainWnd.Width, 6180 + y
        End If
    Case 576
        glDisplayYMultiplier = 3
        If frmMainWnd.WindowState = vbNormal Then
            frmMainWnd.Move frmMainWnd.Left, frmMainWnd.Top, frmMainWnd.Width, 9060 + y
        End If
    End Select
    
    'MM 16.04.2003
    'If this speccy is in full screen modus
    If bFullScreen Then
        'Prepear window
        frmMainWnd.FullScreenOn
        'Get parameter
        GetClientRect frmMainWnd.hWnd, rectWindow
        pointWindow.lX = rectWindow.iLeft
        pointWindow.lY = rectWindow.iTop
        ClientToScreen frmMainWnd.hWnd, pointWindow
        'Margins
        lNewLeft = frmMainWnd.Left - (pointWindow.lX * Screen.TwipsPerPixelX)
        lNewTop = frmMainWnd.Top - (pointWindow.lY * Screen.TwipsPerPixelY)
        lRightMargin = Abs((frmMainWnd.Width - frmMainWnd.Left) - (rectWindow.iRight * Screen.TwipsPerPixelX))
        lBottomMargin = Abs((frmMainWnd.Height - frmMainWnd.Top) - (rectWindow.iBottom * Screen.TwipsPerPixelY))
        'Set modus
        frmMainWnd.Move 0, _
                        0, _
                        2 * Screen.Width, _
                        2 * Screen.Height
        SetForegroundWindow frmMainWnd.hWnd
        glDisplayXMultiplier = lWidth / 265
        glDisplayYMultiplier = glDisplayXMultiplier
    End If
        
    gpicDisplay.Width = lWidth
    gpicDisplay.Height = lHeight
    
    gpicDC = gpicDisplay.hdc
    initscreen
End Sub


Public Sub SetEmulatedModel(lModel As Long, Optional bSEBasicROM As Long = 0)
    Dim sModel As String
    
    glEmulatedModel = lModel
    
    Select Case lModel
    Case 0
        ' // A 48K Spectrum has 69888 tstates per interrupt (3.50000 MHz)
        sModel = "ZX Spectrum 48K"
        glTstatesPerInterrupt = 69888
        glWaveAddTStates = 158 ' 58
        
        glMemPagingType = 0
        glUseScreen = 5
        gbEmulateAYSound = False
        glKeyPortMask = &HBF&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 1
        glPageAt(3) = 2
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 224
        glTStatesAtTop = -glTstatesPerInterrupt + 14336
        glTStatesAtBottom = -glTstatesPerInterrupt + 14336 + 43007
        
        ' // load the ROM image into memory
        If bSEBasicROM Then
            LoadROM App.Path & "\sebasic.rom", 8
        Else
            LoadROM App.Path & "\spectrum.rom", 8
        End If
        SetupContentionTable
    Case 1
        ' // A 128K Spectrum has 70908 tstates per interrupt (3.54690 MHz)
        sModel = "ZX Spectrum 128"
        glTstatesPerInterrupt = 70908
        
        glWaveAddTStates = 160
        
        
        glMemPagingType = 1
        glUseScreen = 5
        gbEmulateAYSound = True
        glKeyPortMask = &HBF&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 2
        glPageAt(3) = 0
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 228
        glTStatesAtTop = -glTstatesPerInterrupt + 14364
        glTStatesAtBottom = -glTstatesPerInterrupt + 14364 + 43775
        
        If bSEBasicROM Then
            LoadROM App.Path & "\sebasic.rom", 9
            LoadROM App.Path & "\zx128_0.rom", 8
            SEPatch128ROM
        Else
            LoadROM App.Path & "\zx128_1.rom", 9
            LoadROM App.Path & "\zx128_0.rom", 8
        End If
        SetupContentionTable
    Case 2
        ' // A Spectrum +2 has 70908 tstates per interrupt (3.54690 MHz)
        sModel = "ZX Spectrum +2"
        glTstatesPerInterrupt = 70908
        
        glWaveAddTStates = 160
        
        glMemPagingType = 1
        glUseScreen = 5
        gbEmulateAYSound = True
        glKeyPortMask = &HBF&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 2
        glPageAt(3) = 0
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 228
        glTStatesAtTop = -glTstatesPerInterrupt + 14364
        glTStatesAtBottom = -glTstatesPerInterrupt + 14364 + 43775
        
        If bSEBasicROM Then
            LoadROM App.Path & "\sebasic.rom", 9
            LoadROM App.Path & "\plus2_0.rom", 8
            SEPatchPlus2ROM
        Else
            LoadROM App.Path & "\plus2_1.rom", 9
            LoadROM App.Path & "\plus2_0.rom", 8
        End If
        SetupContentionTable
    Case 3 ' // +2A
        sModel = "ZX Spectrum +2A"
        glTstatesPerInterrupt = 70908
        
        glWaveAddTStates = 160
    
        glMemPagingType = 2
        glUseScreen = 5
        gbEmulateAYSound = True
        glKeyPortMask = &HBF&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 1
        glPageAt(3) = 2
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 228
        glTStatesAtTop = -glTstatesPerInterrupt + 14364
        glTStatesAtBottom = -glTstatesPerInterrupt + 14364 + 43775
        
        LoadROM App.Path & "\plus2a_3.rom", 11
        LoadROM App.Path & "\plus2a_2.rom", 10
        LoadROM App.Path & "\plus2a_1.rom", 9
        LoadROM App.Path & "\plus2a_0.rom", 8
        SetupContentionTable
    Case 4 ' // +3
        sModel = "ZX Spectrum +3"
        glTstatesPerInterrupt = 70908
        
        glWaveAddTStates = 160
            
        glMemPagingType = 2
        glUseScreen = 5
        gbEmulateAYSound = True
        glKeyPortMask = &HBF&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 1
        glPageAt(3) = 2
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 228
        glTStatesAtTop = -glTstatesPerInterrupt + 14364
        glTStatesAtBottom = -glTstatesPerInterrupt + 14364 + 43775
        
        LoadROM App.Path & "\plus3_3.rom", 11
        LoadROM App.Path & "\plus3_2.rom", 10
        LoadROM App.Path & "\plus3_1.rom", 9
        LoadROM App.Path & "\plus3_0.rom", 8
        SetupContentionTable
    Case 5 ' // TC2048
        sModel = "Timex TC2048"
        glTstatesPerInterrupt = 69888
        
        glWaveAddTStates = 158
        
        glMemPagingType = 0
        glUseScreen = 5
        gbEmulateAYSound = False
        glKeyPortMask = &H1F&
        
        glPageAt(0) = 8
        glPageAt(1) = 5
        glPageAt(2) = 1
        glPageAt(3) = 2
        glPageAt(4) = 8
        
        ' // T-state information
        glTStatesPerLine = 224
        glTStatesAtTop = -glTstatesPerInterrupt + 14336
        glTStatesAtBottom = -glTstatesPerInterrupt + 14336 + 43007
        
        If bSEBasicROM Then
            LoadROM App.Path & "\sebasic.rom", 8
        Else
            LoadROM App.Path & "\tc2048.rom", 8
        End If
        SetupContentionTable
    End Select
    
    SetupTStatesToScanLines
    
    bmiBuffer.bmiHeader.biWidth = 256
    SetStatus sModel
End Sub


Public Sub SetStatus(sMsg As String)
    frmMainWnd.lblStatusMsg.Caption = sMsg
End Sub




Public Sub SetupContentionTable()
    Dim l As Long, z As Long, X(8) As Long, y As Long
    
    X(0) = 6 '6
    X(1) = 5 '5
    X(2) = 4 '4
    X(3) = 3 '3
    X(4) = 2 '2
    X(5) = 1 '1
    X(6) = 0 '0
    X(7) = 0 '0
    
    l = -glTstatesPerInterrupt
    Do While l <= 0
        If (l >= (glTStatesAtTop)) And (l <= glTStatesAtBottom) Then
            For y = 0 To glTStatesPerLine
                If y < 128 Then
                    glContentionTable(-l - y) = X(z)
                    z = z + 1
                    If z > 7 Then Let z = 0
                Else
                   glContentionTable(-l - y) = 0
                End If
            Next y
            z = 0
            l = l + glTStatesPerLine - 1
        Else
            glContentionTable(-l) = 0
        End If
        l = l + 1
    Loop
End Sub

Private Sub SetupTStatesToScanLines()
    Dim n As Long
    
    For n = -glTstatesPerInterrupt To 0
        If (n >= glTStatesAtTop) And (n <= glTStatesAtBottom) Then
            glTSToScanLine(-n) = (n - glTStatesAtTop) \ glTStatesPerLine
        Else
            glTSToScanLine(-n) = -1 ' // In the border area or vertical retrace
        End If
    Next n
End Sub





