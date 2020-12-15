Attribute VB_Name = "modTAP"
' /*******************************************************************************
'   modTAP.bas within vbSpec.vbp
'
'   Handles loading of ".TAP" files (Spectrum tape images)
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'
'   Copyright (C)2001-2002 Grok Developments Ltd.
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

Private m_lChecksum As Long

Public Sub CloseTAPFile()
    If ghTAPFile > 0 Then
        Close ghTAPFile
        ghTAPFile = 0
    End If
End Sub

Public Function SaveTAP(lID As Long, lStart As Long, lLength As Long) As Boolean
    Dim n As Long, lChecksum As Long
    
    On Error Resume Next
    ' // Move to end of existing TAP data, if there is any
    If LOF(ghTAPFile) > 0 Then Seek #ghTAPFile, LOF(ghTAPFile) + 1
    Debug.Print lID; "  "; lStart; "  "; lLength
    
    Put #ghTAPFile, , CInt(lLength + 2)
    Put #ghTAPFile, , CByte(lID)
    lChecksum = lID
    For n = lStart To lStart + lLength - 1
        Put #ghTAPFile, , gRAMPage(glPageAt(glMemAddrDiv16384(n)), n And 16383&)
        lChecksum = lChecksum Xor gRAMPage(glPageAt(glMemAddrDiv16384(n)), n And 16383&)
    Next n
    Put #ghTAPFile, , CByte(lChecksum)
    
    SaveTAP = True
End Function

Public Function LoadTAP(lID As Long, lStart As Long, lLength As Long) As Boolean
    Dim lBlockLen As Long, s(0 To 1) As Byte, lBlockID As Long, lBlockChecksum As Long
    
    On Error Resume Next
    
    If gsTAPFileName = "" Then
        LoadTAP = False
    Else
        If Seek(ghTAPFile) > LOF(ghTAPFile) Then Seek #ghTAPFile, 1
        If EOF(ghTAPFile) Then Seek #ghTAPFile, 1
        Get #ghTAPFile, , s
        lBlockLen = s(1) * 256& + s(0) - 2
        
        lBlockID = Asc(Input(1, #ghTAPFile))
        m_lChecksum = lBlockID ' // Initialize the checksum
        
        If lBlockID = lID Then
            ' // This block type is the same as the requested block type
            If lLength <= lBlockLen Then
                ' // There are enough bytes in the block to cover this request
                ReadTAPBlock lStart, lLength
                If lLength < lBlockLen Then
                    ' // Skip the rest of the bytes up to the end of the block
                    SkipTAPBytes lBlockLen - lLength
                End If
                lBlockChecksum = Asc(Input(1, #ghTAPFile))
                regIX = (regIX + lLength) And &HFFFF&
                regDE = 0
                If m_lChecksum = lBlockChecksum Then
                    LoadTAP = True
                Else
                    LoadTAP = False
                End If
            Else
                ' // More bytes requested than there are in the block
                ReadTAPBlock lStart, lBlockLen
                lBlockChecksum = Asc(Input(1, #ghTAPFile))
                regIX = (regIX + lBlockLen) And &HFFFF&
                regDE = regDE - lBlockLen
                LoadTAP = False
            End If
        Else
            ' // Wrong block type -- skip this block
            SkipTAPBytes lBlockLen
            lBlockChecksum = Asc(Input(1, #ghTAPFile))
            LoadTAP = False
        End If
    End If
    initscreen
    screenPaint
End Function
Public Sub OpenTAPFile(sName As String)
    If Dir$(sName) = "" Then Exit Sub
    If ghTAPFile > 0 Then Close #ghTAPFile
    
    If Dir$(sName) = "" Then Exit Sub
    
    StopTape ' // Stop the TZX tape player
    
    ghTAPFile = FreeFile
    Open sName For Binary As ghTAPFile
    
    If LOF(ghTAPFile) = 0 Then
        Close #ghTAPFile
        Exit Sub
    End If
    
    gsTAPFileName = sName
    
    frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(sName)
End Sub
Private Sub ReadTAPBlock(lStart As Long, lLen As Long)
    Dim s() As Byte, lCounter As Long, a As Long
    
    On Error Resume Next
    
    ReDim s(0 To (lLen - 1))
    Get #ghTAPFile, , s

    For lCounter = 0 To lLen - 1
        a = lStart + lCounter
        gRAMPage(glPageAt(glMemAddrDiv16384(a)), a And 16383&) = s(lCounter)
        m_lChecksum = m_lChecksum Xor s(lCounter)
    Next lCounter
End Sub

Public Sub SaveTAPFileDlg()
    On Error Resume Next
    
    If ghTAPFile < 1 Then
        err.Clear
        With frmMainWnd.dlgCommon
            .DialogTitle = "Select TAP file for saving"
            .DefaultExt = ".tap"
            .FileName = "*.tap"
            .Filter = "Tape files (*.tap)|*.tap|All Files (*.*)|*.*"
            .Flags = cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
            .CancelError = True
            .ShowSave
            resetKeyboard
            If err.Number = cdlCancel Then
                Exit Sub
            End If
        End With
        If frmMainWnd.dlgCommon.FileName <> "" Then
            ghTAPFile = FreeFile
            Open frmMainWnd.dlgCommon.FileName For Binary As ghTAPFile
            
            gsTAPFileName = frmMainWnd.dlgCommon.FileName
            gMRU.AddMRUFile gsTAPFileName
            frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(gsTAPFileName)
        End If
    End If
    StopTape ' // Stop the TZX tape player
    
    If SaveTAP(glMemAddrDiv256(regAF_), regIX, regDE) Then
        regIX = regIX + regDE
        regDE = 0
    End If
    
    resetKeyboard
    regPC = 1342 ' RET
End Sub

Private Sub SkipTAPBytes(lLen As Long)
    Dim s As String, lCounter As Long
    
    s = Input(lLen, #ghTAPFile)
    For lCounter = 1 To Len(s)
        m_lChecksum = m_lChecksum Xor Asc(Mid$(s, lCounter, 1))
    Next lCounter
End Sub


