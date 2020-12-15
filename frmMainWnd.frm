VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMainWnd 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "vbSpec"
   ClientHeight    =   3630
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4215
   ControlBox      =   0   'False
   Icon            =   "frmMainWnd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picStatus 
      Align           =   2  'Unten ausrichten
      BorderStyle     =   0  'Kein
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   3300
      Width           =   4215
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   1.78260e5
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   1.78260e5
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label lblStatusMsg 
         Caption         =   "vbSpec"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   4035
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   60
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   180
      MouseIcon       =   "frmMainWnd.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   240
      Width           =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -4
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save As..."
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Create blank tape file for saving..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Load &Binary..."
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save B&inary..."
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Reset Spectrum"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&1"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&2"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&3"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&4"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&5"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   17
      End
   End
   Begin VB.Menu mnuOptionsMain 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&General Settings..."
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Display..."
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Tape Controls..."
         Index           =   4
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "ZX &Printer..."
         Index           =   5
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&YZ Switch"
         Index           =   7
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Poke memory"
         Index           =   8
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Spectrum &Keyboard..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About vbSpec..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuFullScreenMenu 
      Caption         =   "Full screen mode menu"
      Visible         =   0   'False
      Begin VB.Menu mnuFullScreenFile 
         Caption         =   "File..."
      End
      Begin VB.Menu mnuFullScreenOptions 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuFullScreenHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuFullNormalView 
         Caption         =   "Normal view!"
      End
   End
End
Attribute VB_Name = "frmMainWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmMainWnd.frm within vbSpec.vbp
'
'   Main application window. Contains the Spectrum display output, processes
'   Keypresses, and contains a timer object used for maintaining the emulator
'   at 50 frames/sec on a fast enough machine, and a common dialog control for
'   use by file open/save operations.
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'
'   Copyright (C)1999-2001  Grok Developments Ltd.
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

'MM 16.04.2003
Private lPopUp As Long
Private sNewCaption As String

'That is a new caption property. If the Speccy is in full screen modus, the normal
'caption of the form should not be set. The Speccy saves the data an restores it,
'if the full screen modus is off
Public Property Let NewCaption(ByVal sData As String)
    If modSpectrum.bFullScreen Then
        sNewCaption = sData
    Else
        Me.Caption = sData
    End If
End Property

'This method sets the caption to an empty string and saves the original caption
Public Sub FullScreenOn()
    sNewCaption = Me.Caption
    Me.Caption = vbNullString
    Me.mnuFileMain.Visible = False
    Me.mnuOptionsMain.Visible = False
    Me.mnuHelpMain.Visible = False
    Me.picStatus.Visible = False
    frmMainWnd.Line1.Item(1).Visible = False
End Sub

'This method restores the original caption
Public Sub FullScreenOff()
    bFullScreen = False
    Me.NewCaption = sNewCaption
    Me.mnuFileMain.Visible = True
    Me.mnuOptionsMain.Visible = True
    Me.mnuHelpMain.Visible = True
    Me.picStatus.Visible = True
    frmMainWnd.Line1.Item(1).Visible = True
    Me.Move 0, 0
End Sub

Public Sub FileCreateNewTap()
    Dim sName As String
    
    On Error Resume Next
    
    err.Clear
    dlgCommon.DialogTitle = "Create New Tap File"
    dlgCommon.DefaultExt = ".tap"
    dlgCommon.FileName = "*.tap"
    dlgCommon.Filter = "Tape files (*.tap)|*.tap|All Files|*.*"
    dlgCommon.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNLongNames
    dlgCommon.CancelError = True
    dlgCommon.ShowSave
    If err.Number = cdlCancel Then
        Exit Sub
    End If
    
    sName = dlgCommon.FileName
    
    If (sName <> "") Then
        If ghTAPFile > 0 Then Close #ghTAPFile
        
        StopTape ' // Stop the TZX tape player
        
        Kill sName
        
        ghTAPFile = FreeFile
        Open sName For Binary As ghTAPFile
        
        gsTAPFileName = sName
        
        frmMainWnd.NewCaption = App.ProductName & " - " & GetFilePart(sName)

        gMRU.AddMRUFile sName
        SetMRUMenu
    End If
End Sub

Public Sub FileOpenDialog(Optional sName As String = "")
    On Error Resume Next
    
    If sName = "" Then
        err.Clear
        dlgCommon.DialogTitle = "Open Snapshot or ROM image"
        dlgCommon.DefaultExt = ".sna"
        dlgCommon.FileName = ""
        dlgCommon.Filter = "All Spectrum Files (*.sna;*.z80;*.tap;*.tzx;*.rom;*.scr)|*.sna;*.z80;*.tap;*.tzx;*.rom;*.scr|SNA snapshots (*.sna)|*.sna|Z80 snapshots (*.z80)|*.z80|ROM images (*.rom)|*.rom|Tape files (*.tap;*.tzx)|*.tap;*.tzx|Screen images (*.scr)|*.scr|All Files|*.*"
        dlgCommon.Flags = cdlOFNFileMustExist Or cdlOFNExplorer Or cdlOFNLongNames
        dlgCommon.CancelError = True
        dlgCommon.ShowOpen
        If err.Number = cdlCancel Then
            Exit Sub
        End If
        
        sName = dlgCommon.FileName
    End If
    
    If (sName <> "") And (Dir$(sName) <> "") Then
        gMRU.AddMRUFile sName
        
        Select Case LCase$(Right$(sName, 4))
        Case ".z80"
            LoadZ80Snap sName
        Case ".rom"
            Z80Reset
            If glEmulatedModel <> 0 And glEmulatedModel <> 5 Then
                SetEmulatedModel 0, False
            End If
            LoadROM sName
        Case ".sna"
            LoadSNASnap sName
        Case ".tap"
            OpenTAPFile sName
        Case ".tzx"
            OpenTZXFile sName
        Case ".scr"
            LoadScreenSCR sName
        Case Else
            ' try opening it as a SNA file
            LoadSNASnap sName
        End Select
        SetMRUMenu
    End If
End Sub




Private Sub FileSaveAsDialog()
    On Error Resume Next
    
    err.Clear
    dlgCommon.DialogTitle = "Save As"
    dlgCommon.DefaultExt = ".z80"
    dlgCommon.FileName = ""
    dlgCommon.Filter = "Z80 snapshot (*.z80)|*.z80|SNA snapshot (*.sna)|*.sna|ROM image (*.rom)|*.rom|Screen Bitmap (*.bmp)|*.bmp|Screen Image (*.scr)|*.scr"
    dlgCommon.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    dlgCommon.CancelError = True
    dlgCommon.ShowSave
    If err.Number = cdlCancel Then
        Exit Sub
    End If
    
    If dlgCommon.FileName <> "" Then
        gMRU.AddMRUFile dlgCommon.FileName
        
        Select Case LCase$(Right$(dlgCommon.FileName, 4))
        Case ".z80"
            SaveZ80Snap dlgCommon.FileName
        Case ".rom"
            SaveROM dlgCommon.FileName
        Case ".sna"
            SaveSNASnap dlgCommon.FileName
        Case ".bmp"
            SaveScreenBMP dlgCommon.FileName
        Case ".scr"
            SaveScreenSCR dlgCommon.FileName
        Case Else
            ' save it as a Z80 file
            SaveZ80Snap dlgCommon.FileName
        End Select
        SetMRUMenu
    End If

End Sub



Private Function SetMRUMenu()
    Dim l As Long
    
    For l = 10 To 15
        mnuFile(l).Visible = False
    Next l
      
    For l = 1 To gMRU.GetMRUCount
        If Len(gMRU.GetMRUFile(l)) <= 40 Then
            mnuFile(10 + l).Caption = "&" & CStr(l) & " " & gMRU.GetMRUFile(l)
        Else
            mnuFile(10 + l).Caption = "&" & CStr(l) & " " & GetFilePart(gMRU.GetMRUFile(l))
        End If
        mnuFile(10 + l).Visible = True
    Next l
    
    If gMRU.GetMRUCount > 0 Then mnuFile(10).Visible = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And 4) Then Exit Sub
    doKey True, KeyCode, Shift
    'MM 16.04.2003
    If (KeyCode = 27) And (Shift = 1) And bFullScreen Then
        lPopUp = -1
        Me.PopupMenu mnuFullScreenMenu
        If lPopUp = 0 Then
            Me.PopupMenu mnuFileMain
        End If
        If lPopUp = 1 Then
            Me.PopupMenu mnuOptionsMain
        End If
        If lPopUp = 2 Then
            Me.PopupMenu mnuHelpMain
        End If
    End If
End Sub

'MM 16.04.2003
Private Sub mnuFullScreenFile_Click()
    lPopUp = 0
End Sub
Private Sub mnuFullScreenOptions_Click()
    lPopUp = 1
End Sub
Private Sub mnuFullScreenHelp_Click()
    lPopUp = 2
End Sub
Private Sub mnuFullNormalView_Click()
    Me.FullScreenOff
    modMain.SetDisplaySize 256, 192
    Me.Resize
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    doKey False, KeyCode, Shift
End Sub


Private Sub Form_Load()
    Dim X As Long, y As Long
    Dim lCounter As Long
    'MM JD
    Dim sKey As String
    
    Set gMRU = New MRUList
    SetMRUMenu
    
    X = val(GetSetting("Grok", "vbSpec", "MainWndX", "-1"))
    y = val(GetSetting("Grok", "vbSpec", "MainWndY", "-1"))
    
    If X >= 0 And X <= (Screen.Width - Screen.TwipsPerPixelX * 16) Then
        Me.Left = X
    End If
    If y >= 0 And y <= (Screen.Height - Screen.TwipsPerPixelY * 16) Then
        Me.Top = y
    End If
    
    'MM JD
    'Load standard joystick tables
    'Joystick 1
    'Kepmston
    PortValueToKeyStroke KEMPSTON_PUP, KEMPSTON_UP, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_UP).lPort = KEMPSTON_PUP
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_UP).lValue = KEMPSTON_UP
    PortValueToKeyStroke KEMPSTON_PDOWN, KEMPSTON_DOWN, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_DOWN).lPort = KEMPSTON_PDOWN
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_DOWN).lValue = KEMPSTON_DOWN
    PortValueToKeyStroke KEMPSTON_PLEFT, KEMPSTON_LEFT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_LEFT).lPort = KEMPSTON_PLEFT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_LEFT).lValue = KEMPSTON_LEFT
    PortValueToKeyStroke KEMPSTON_PRIGHT, KEMPSTON_RIGHT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_RIGHT).lPort = KEMPSTON_PRIGHT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_RIGHT).lValue = KEMPSTON_RIGHT
    PortValueToKeyStroke KEMPSTON_PFIRE, KEMPSTON_FIRE, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_BUTTON1).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_BUTTON1).lPort = KEMPSTON_PFIRE
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, JDT_BUTTON1).lValue = KEMPSTON_FIRE
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjKempston, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    'Cursor
    PortValueToKeyStroke CURSOR_PUP, CURSOR_UP, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_UP).lPort = CURSOR_PUP
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_UP).lValue = CURSOR_UP
    PortValueToKeyStroke CURSOR_PDOWN, CURSOR_DOWN, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_DOWN).lPort = CURSOR_PDOWN
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_DOWN).lValue = CURSOR_DOWN
    PortValueToKeyStroke CURSOR_PLEFT, CURSOR_LEFT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_LEFT).lPort = CURSOR_PLEFT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_LEFT).lValue = CURSOR_LEFT
    PortValueToKeyStroke CURSOR_PRIGHT, CURSOR_RIGHT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_RIGHT).lPort = CURSOR_PRIGHT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_RIGHT).lValue = CURSOR_RIGHT
    PortValueToKeyStroke CURSOR_PFIRE, CURSOR_FIRE, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_BUTTON1).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_BUTTON1).lPort = CURSOR_PFIRE
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, JDT_BUTTON1).lValue = CURSOR_FIRE
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjCursor, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    'Sinclair 1
    PortValueToKeyStroke SINCLAIR1_PUP, SINCLAIR1_UP, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_UP).lPort = SINCLAIR1_PUP
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_UP).lValue = SINCLAIR1_UP
    PortValueToKeyStroke SINCLAIR1_PDOWN, SINCLAIR1_DOWN, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_DOWN).lPort = SINCLAIR1_PDOWN
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_DOWN).lValue = SINCLAIR1_DOWN
    PortValueToKeyStroke SINCLAIR1_PLEFT, SINCLAIR1_LEFT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_LEFT).lPort = SINCLAIR1_PLEFT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_LEFT).lValue = SINCLAIR1_LEFT
    PortValueToKeyStroke SINCLAIR1_PRIGHT, SINCLAIR1_RIGHT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_RIGHT).lPort = SINCLAIR1_PRIGHT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_RIGHT).lValue = SINCLAIR1_RIGHT
    PortValueToKeyStroke SINCLAIR1_PFIRE, SINCLAIR1_FIRE, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_BUTTON1).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_BUTTON1).lPort = SINCLAIR1_PFIRE
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, JDT_BUTTON1).lValue = SINCLAIR1_FIRE
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair1, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    'Sinclair 2
    PortValueToKeyStroke SINCLAIR2_PUP, SINCLAIR2_UP, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_UP).lPort = SINCLAIR2_PUP
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_UP).lValue = SINCLAIR2_UP
    PortValueToKeyStroke SINCLAIR2_PDOWN, SINCLAIR2_DOWN, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_DOWN).lPort = SINCLAIR2_PDOWN
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_DOWN).lValue = SINCLAIR2_DOWN
    PortValueToKeyStroke SINCLAIR2_PLEFT, SINCLAIR2_LEFT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_LEFT).lPort = SINCLAIR2_PLEFT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_LEFT).lValue = SINCLAIR2_LEFT
    PortValueToKeyStroke SINCLAIR2_PRIGHT, SINCLAIR2_RIGHT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_RIGHT).lPort = SINCLAIR2_PRIGHT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_RIGHT).lValue = SINCLAIR2_RIGHT
    PortValueToKeyStroke SINCLAIR2_PFIRE, SINCLAIR2_FIRE, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_BUTTON1).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_BUTTON1).lPort = SINCLAIR2_PFIRE
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, JDT_BUTTON1).lValue = SINCLAIR2_FIRE
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjSinclair2, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    'Fuller box
    PortValueToKeyStroke FULLER_PUP, FULLER_UP, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_UP).lPort = FULLER_PUP
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_UP).lValue = FULLER_UP
    PortValueToKeyStroke FULLER_PDOWN, FULLER_DOWN, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_DOWN).lPort = FULLER_PDOWN
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_DOWN).lValue = FULLER_DOWN
    PortValueToKeyStroke FULLER_PLEFT, FULLER_LEFT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_LEFT).lPort = FULLER_PLEFT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_LEFT).lValue = FULLER_LEFT
    PortValueToKeyStroke FULLER_PRIGHT, FULLER_RIGHT, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_RIGHT).lPort = FULLER_PRIGHT
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_RIGHT).lValue = FULLER_RIGHT
    PortValueToKeyStroke FULLER_PFIRE, FULLER_FIRE, sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_BUTTON1).sKey = sKey
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_BUTTON1).lPort = FULLER_PFIRE
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, JDT_BUTTON1).lValue = FULLER_FIRE
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjFullerBox, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    'User defined
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_UP).sKey = vbNullString
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_UP).lPort = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_UP).lValue = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_DOWN).sKey = vbNullString
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_DOWN).lPort = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_DOWN).lValue = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_LEFT).sKey = vbNullString
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_LEFT).lPort = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_LEFT).lValue = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_RIGHT).sKey = vbNullString
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_RIGHT).lPort = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_RIGHT).lValue = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_BUTTON1).sKey = vbNullString
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_BUTTON1).lPort = 0
    aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, JDT_BUTTON1).lValue = 0
    For lCounter = 2 To JDT_MAXBUTTONS
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, (JDT_BUTTON_BASE + lCounter)).sKey = vbNullString
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, (JDT_BUTTON_BASE + lCounter)).lPort = 0
        aJoystikDefinitionTable(JDT_JOYSTICK1, zxjUserDefined, (JDT_BUTTON_BASE + lCounter)).lValue = 0
    Next lCounter
    
    'MM 12.03.2003 - Joystick support code
    '=================================================================================
    'API call result
    Dim lResult As Long
    Dim uJoyCap As JOYCAPS
    
    'Init
    Let modSpectrum.lPCJoystick1Is = zxjInvalid
    Let modSpectrum.lPCJoystick2Is = zxjInvalid
    Let modSpectrum.lPCJoystick1Fire = pcjbInvalid
    Let modSpectrum.lPCJoystick2Fire = pcjbInvalid
    Let modSpectrum.bJoystick1Valid = False
    Let modSpectrum.bJoystick2Valid = False
    
    'Errorhandler
    On Error GoTo LOAD_ERROR
    
    'Try to initiate PC-joystick nr. 1
    Let lResult = modSpectrum.joyGetDevCaps(modSpectrum.JOYSTICKID1, uJoyCap, Len(uJoyCap))
    'Init succeded
    If lResult = modSpectrum.JOYERR_OK Then
        'PC-joystick nr. 1 can be set up
        Let modSpectrum.bJoystick1Valid = True
        'Init for Kempston
        Let modSpectrum.lPCJoystick1Is = zxjKempston
        Let modSpectrum.lPCJoystick1Fire = pcjbButton1
        'Number of buttons
        Let modSpectrum.lPCJoystick1Buttons = uJoyCap.wNumButtons
    'Init for PC-joystick nr.1 failed
    Else
        'There is no posibility to set up the PC-joystick nr. 1
        Let modSpectrum.bJoystick1Valid = False
        Let modSpectrum.lPCJoystick1Is = zxjInvalid
        Let modSpectrum.lPCJoystick1Fire = pcjbInvalid
        'Number of buttons
        Let modSpectrum.lPCJoystick1Buttons = -1
    End If
    
    'Try to initiate PC-joystick nr. 2
    Let lResult = modSpectrum.joyGetDevCaps(modSpectrum.JOYSTICKID2, uJoyCap, Len(uJoyCap))
    'Init succeded
    If lResult = modSpectrum.JOYERR_OK Then
        'PC-joystick nr. 2 can be set up
        Let modSpectrum.bJoystick2Valid = True
        'Init for Kempston
        Let modSpectrum.lPCJoystick2Is = zxjKempston
        Let modSpectrum.lPCJoystick2Fire = pcjbButton1
        'Number of buttons
        Let modSpectrum.lPCJoystick2Buttons = uJoyCap.wNumButtons
    'Init for PC-joystick nr.2 failed
    Else
        'There is no posibility to set up the PC-joystick nr. 2
        Let modSpectrum.bJoystick2Valid = False
        Let modSpectrum.lPCJoystick2Is = zxjInvalid
        Let modSpectrum.lPCJoystick2Fire = pcjbInvalid
        'Number of buttons
        Let modSpectrum.lPCJoystick2Buttons = -1
    End If
    
    'All right
    Exit Sub
LOAD_ERROR:

    'MM 03.02.2003 -- END
    '=================================================================================

End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Data.GetFormat(vbCFFiles) Then
        If (LCase$(Right$(Data.Files(1), 4)) = ".sna") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".z80") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".tap") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".rom") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".scr") Then
            If (Effect And vbDropEffectCopy) Then
                Select Case LCase$(Right$(Data.Files(1), 4))
                Case ".z80"
                    DoEvents
                    If Dir$(Data.Files(1)) <> "" Then LoadZ80Snap Data.Files(1)
                Case ".rom"
                    DoEvents
                    Z80Reset
                    If glEmulatedModel <> 0 And glEmulatedModel <> 5 Then
                        SetEmulatedModel 0, False
                    End If
                    If Dir$(Data.Files(1)) <> "" Then LoadROM Data.Files(1)
                Case ".sna"
                    DoEvents
                    If Dir$(Data.Files(1)) <> "" Then LoadSNASnap Data.Files(1)
                Case ".tap"
                    DoEvents
                    If Dir$(Data.Files(1)) <> "" Then OpenTAPFile Data.Files(1)
                Case ".scr"
                    DoEvents
                    If Dir$(Data.Files(1)) <> "" Then LoadScreenSCR Data.Files(1)
                End Select
            End If
        End If
    End If
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        If (LCase$(Right$(Data.Files(1), 4)) = ".sna") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".z80") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".tap") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".rom") Or _
           (LCase$(Right$(Data.Files(1), 4)) = ".scr") Then
            If (Effect And vbDropEffectCopy) Then
                Effect = vbDropEffectCopy
            Else
                Effect = vbDropEffectNone
            End If
        Else
            Effect = vbDropEffectNone
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not (frmZXPrinter.Visible) Then
        SaveSetting "Grok", "vbSpec", "EmulateZXPrinter", "0"
    End If
    
    If Not (frmTapePlayer.Visible) Then
        SaveSetting "Grok", "vbSpec", "TapeControlsVisible", "0"
    End If
End Sub

'MM 16.04.2003
Private Sub Form_Resize()
    Resize
End Sub
Private Sub Form_Activate()
    Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbSoundEnabled Then CloseWaveOut
    
    Set gMRU = Nothing
    
    SaveSetting "Grok", "vbSpec", "MainWndX", CStr(Me.Left)
    SaveSetting "Grok", "vbSpec", "MainWndY", CStr(Me.Top)
    
    timeEndPeriod 1
    End
End Sub


Private Sub mnuFile_Click(Index As Integer)
    Dim lWndProc As Long
    mnuFileMain.Visible = Not bFullScreen
    Select Case Index
    Case 1 ' Open
        If gbSoundEnabled Then waveOutReset glphWaveOut
        FileOpenDialog
    Case 2 ' Save As
        If gbSoundEnabled Then waveOutReset glphWaveOut
        FileSaveAsDialog
    Case 4 ' Create blank TAP file
        If gbSoundEnabled Then waveOutReset glphWaveOut
        FileCreateNewTap
    Case 6 ' Load Binary
        frmLoadBinary.Show 1
    Case 7 ' Save Binary
        frmSaveBinary.Show 1
    Case 9 ' Reset
        Z80Reset
    Case Else
        If mnuFile(Index).Caption = "E&xit" Then
            Unload Me
        Else
            ' // MRU File
            FileOpenDialog gMRU.GetMRUFile(Index - 10)
        End If
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    mnuHelpMain.Visible = Not bFullScreen
    Select Case Index
    Case 1 ' // Show spectrum keyboard layout
        frmKeyboard.Show
    Case 3 ' // About... Dialog
        frmAbout.Show 1
    End Select
End Sub

Public Sub mnuOptions_Click(Index As Integer)
    mnuOptionsMain.Visible = Not bFullScreen
    Select Case Index
    Case 1
        frmOptions.Show 1
    Case 2
        frmDisplayOpt.Show 1
        Form_Resize
    Case 4
        If mnuOptions(Index).Checked Then
            frmTapePlayer.Hide
            mnuOptions(Index).Checked = False
        Else
            ' // GS: Fix for tapeplayer window always-on-top
            frmTapePlayer.Show 0, frmMainWnd
            mnuOptions(Index).Checked = True
        End If
        SaveSetting "Grok", "vbSpec", "TapeControlsVisible", IIf(mnuOptions(Index).Checked, "-1", "0")
    Case 5
        If mnuOptions(Index).Checked Then
            frmZXPrinter.Hide
            mnuOptions(Index).Checked = False
        Else
            frmZXPrinter.Show 0, frmMainWnd
            mnuOptions(Index).Checked = True
        End If
        SaveSetting "Grok", "vbSpec", "EmulateZXPrinter", IIf(mnuOptions(Index).Checked, "-1", "0")
    'MM 03.02.2003 - BEGIN
    Case 7
        mnuOptions.Item(Index).Checked = Not mnuOptions.Item(Index).Checked
    Case 8
        frmPoke.Show vbModal
    'MM 03.02.2003 - END
    End Select
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    glMouseBtn = glMouseBtn Or Button
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    glMouseBtn = glMouseBtn Xor Button
End Sub


Private Sub picDisplay_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub picDisplay_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    Form_OLEDragOver Data, Effect, Button, Shift, X, y, State
End Sub

'Resize operations
Public Sub Resize()
    'Variables
    Dim lDisplayWidth As Long, lDisplayHeight As Long
    Dim lNewLeft As Long, lNewTop As Long
    'Get display width and height
    lDisplayWidth = Me.ScaleWidth
    lDisplayHeight = Me.ScaleHeight - picStatus.Height
    'Center the paper
    lNewLeft = CLng((lDisplayWidth - picDisplay.Width) / 2)
    lNewTop = CLng((lDisplayHeight - picDisplay.Height) / 2)
    'Preserve minimum size
    If lNewLeft < 12 Then
        If glDisplayXMultiplier = 1 Then
            Me.Width = 4320
            Exit Sub
        End If
        If glDisplayXMultiplier = 2 Then
            Me.Width = 8160
            Exit Sub
        End If
        If glDisplayXMultiplier = 3 Then
            Me.Width = 12000
            Exit Sub
        End If
    End If
    If lNewTop < 14 Then
        If glDisplayYMultiplier = 1 Then
            Me.Height = 4320
            Exit Sub
        End If
        If glDisplayYMultiplier = 2 Then
            Me.Height = 7200
            Exit Sub
        End If
        If glDisplayYMultiplier = 3 Then
            Me.Height = 10080
            Exit Sub
        End If
    End If
    'Set size
    picDisplay.Top = lNewTop
    picDisplay.Left = lNewLeft
End Sub

