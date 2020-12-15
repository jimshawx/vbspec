VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "General Settings"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'Kein
      Height          =   2775
      Index           =   3
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   5055
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdDefine 
         Caption         =   "Define..."
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   24
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefine 
         Caption         =   "Define..."
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   29
         Top             =   2020
         Width           =   1095
      End
      Begin VB.ComboBox cmbZXJoystick 
         Height          =   315
         Index           =   1
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   26
         Top             =   1305
         Width           =   2835
      End
      Begin VB.ComboBox cmbFire 
         Height          =   315
         Index           =   1
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   28
         Top             =   1660
         Width           =   2835
      End
      Begin VB.ComboBox cmbZXJoystick 
         Height          =   315
         Index           =   0
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   21
         Top             =   60
         Width           =   2835
      End
      Begin VB.ComboBox cmbFire 
         Height          =   315
         Index           =   0
         Left            =   2100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   420
         Width           =   2835
      End
      Begin VB.Label lblZXJoystick 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick 2 Emulates:"
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   25
         Top             =   1365
         Width           =   2055
      End
      Begin VB.Label lblFire 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick 2 Fire Button:"
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   1725
         Width           =   2055
      End
      Begin VB.Label lblZXJoystick 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick 1 Emulates:"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblFire 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick 1 Fire Button:"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'Kein
      Height          =   2775
      Index           =   2
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   5055
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboMouseType 
         Height          =   315
         ItemData        =   "frmOptions.frx":000C
         Left            =   2100
         List            =   "frmOptions.frx":0019
         Style           =   2  'Dropdown-Liste
         TabIndex        =   13
         Top             =   60
         Width           =   2835
      End
      Begin VB.Frame Frame1 
         Caption         =   "Respond to Mouse Buttons"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4815
         Begin VB.OptionButton optMouseGlobal 
            Caption         =   "&Always"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optMouseVB 
            Caption         =   "&Only when Windows pointer is over vbSpec window"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   300
            Width           =   4155
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         Caption         =   "Windows &Mouse Emulates:"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'Kein
      Height          =   2775
      Index           =   1
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   5055
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   5055
      Begin VB.CheckBox chkSEBasic 
         Caption         =   "Use SE Basic ROM"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame fraStd 
         Caption         =   "Emulation Speed"
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   4875
         Begin VB.OptionButton optSpeed200 
            Caption         =   "&Double - Limit emulation to 200% real Spectrum speed"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            Top             =   900
            Width           =   4335
         End
         Begin VB.OptionButton optSpeed50 
            Caption         =   "&Slow - Limit emulation to 50% of real Spectrum speed"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   300
            Width           =   4395
         End
         Begin VB.OptionButton optSpeed100 
            Caption         =   "&Real - Limit emulation to 100% real Spectrum speed"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   600
            Width           =   4335
         End
         Begin VB.OptionButton optSpeedFastest 
            Caption         =   "&Fastest - Do not limit emulation speed"
            Height          =   255
            Left            =   180
            TabIndex        =   11
            Top             =   1200
            Width           =   4275
         End
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Enable sound output"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   780
         Width           =   2115
      End
      Begin VB.ComboBox cboModel 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lblStatic 
         Alignment       =   1  'Rechts
         Caption         =   "Emulated model:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   0
      Top             =   3420
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   3420
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip ts1 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Emulation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mouse"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Joystick"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmOptions.frm within vbSpec.vbp
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'   Joystick support code: Miklos Muhi <vbspec@muhi.org>
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

                                ' // Identifies the emulated model at the time
                                ' // the dialog is displayed. We need to reboot
Private lOriginalModel As Long  ' // the Spectrum if this changes.
                                
'MM 12.03.2003 - Joystick support code
'=================================================================================
'The first activation flag
Private bIsFirstActivate As Boolean
'Property value -- wich PC-joytick will be set up
Private lPCJoystick As PCJOYSTICKS
'So many Buttons will be supported
Private lSupportedButtons As Long
'Supported PC-Joysticks
Public Enum PCJOYSTICKS
    pcjInvalid = -1
    pcjJoystick1 = 0
    pcjJoystick2 = 1
End Enum
'=================================================================================

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    'Variables
    Dim jdTMP As JOYSTICK_DEFINITION
    
    If optSpeed50.Value Then
        glInterruptDelay = 40
    ElseIf optSpeed100.Value Then
        glInterruptDelay = 20
    ElseIf optSpeed200.Value Then
        glInterruptDelay = 10
    ElseIf optSpeedFastest.Value Then
        glInterruptDelay = 0
    End If
    SaveSetting "Grok", "vbSpec", "InterruptDelay", CStr(glInterruptDelay)
    
    glEmulatedModel = cboModel.ItemData(cboModel.ListIndex)
    If (glEmulatedModel <> lOriginalModel) Or ((chkSEBasic.Value = 1) <> gbSEBasicROM) Then
        ' // Model or ROM has been changed - reboot the Spectrum
        SaveSetting "Grok", "vbSpec", "SEBasicROM", CStr(chkSEBasic.Value)
        SaveSetting "Grok", "vbSpec", "EmulatedModel", CStr(glEmulatedModel)
        Z80Reset
    End If
    
    SaveSetting "Grok", "vbSpec", "MouseType", CStr(cboMouseType.ListIndex)
    SaveSetting "Grok", "vbSpec", "SoundEnabled", CStr(chkSound.Value)
    
    gbSEBasicROM = -chkSEBasic.Value
    glMouseType = cboMouseType.ListIndex
    
    If optMouseGlobal.Value Then
        gbMouseGlobal = True
    Else
        gbMouseGlobal = False
    End If
    SaveSetting "Grok", "vbSpec", "MouseGlobal", CStr(gbMouseGlobal)
        
    If glMouseType = MOUSE_NONE Then
        frmMainWnd.picDisplay.MousePointer = 0
    Else
        frmMainWnd.picDisplay.MousePointer = 99
    End If
    
    If gbSoundEnabled Then
        If chkSound.Value = 0 Then
            CloseWaveOut
        End If
    Else
        If chkSound.Value = 1 Then
            gbSoundEnabled = InitializeWaveOut
        End If
    End If

    'MM 12.03.2003 - Joystick support code
    '=================================================================================
    'Save values
    If modSpectrum.bJoystick1Valid Then
        lPCJoystick1Is = cmbZXJoystick.Item(0).ItemData(cmbZXJoystick.Item(0).ListIndex)
        If (lPCJoystick1Is <> zxjInvalid) And (lPCJoystick1Is <> zxjUserDefined) Then
            jdTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lPCJoystick1Fire + 1)
            aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lPCJoystick1Fire + 1).sKey = vbNullString
            aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lPCJoystick1Fire + 1).lPort = 0
            aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lPCJoystick1Fire + 1).lValue = 0
            lPCJoystick1Fire = cmbFire.Item(0).ItemData(cmbFire.Item(0).ListIndex)
            aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lPCJoystick1Fire + 1) = jdTMP
        End If
    End If
    If modSpectrum.bJoystick2Valid Then
        modSpectrum.lPCJoystick2Is = cmbZXJoystick.Item(1).ItemData(cmbZXJoystick.Item(1).ListIndex)
        If (lPCJoystick2Is <> zxjInvalid) And (lPCJoystick2Is <> zxjUserDefined) Then
            jdTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lPCJoystick2Fire + 1)
            aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lPCJoystick2Fire + 1).sKey = vbNullString
            aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lPCJoystick2Fire + 1).lPort = 0
            aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lPCJoystick2Fire + 1).lValue = 0
            lPCJoystick2Fire = cmbFire.Item(1).ItemData(cmbFire.Item(1).ListIndex)
            aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lPCJoystick2Fire + 1) = jdTMP
        End If
    End If
    '=================================================================================
    
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case glInterruptDelay
    Case 40
        ' // 40ms delay = 50% Spectrum speed
        optSpeed50.Value = True
    Case 20
        ' // 20ms delay = 100% Spectrum speed
        optSpeed100.Value = True
    Case 10
        ' // 10ms delay = 200% Spectrum speed
        optSpeed200.Value = True
    Case 0
        ' // 0 delay = run as fast as possible
        optSpeedFastest.Value = True
    End Select
    
    cboModel.AddItem "ZX Spectrum 48K"
    cboModel.ItemData(cboModel.NewIndex) = 0
    cboModel.AddItem "ZX Spectrum 128"
    cboModel.ItemData(cboModel.NewIndex) = 1
    cboModel.AddItem "ZX Spectrum +2"
    cboModel.ItemData(cboModel.NewIndex) = 2
    cboModel.AddItem "Timex TC2048"
    cboModel.ItemData(cboModel.NewIndex) = 5
    
    lOriginalModel = glEmulatedModel
    Select Case glEmulatedModel
    Case 0 ' // 48K
        cboModel.ListIndex = 0
    Case 1 ' // 128K
        cboModel.ListIndex = 1
    Case 2 ' // +2
        cboModel.ListIndex = 2
    Case 5 ' // TC2048
        cboModel.ListIndex = 3
    End Select
    
    If gbSoundEnabled Then chkSound.Value = 1 Else chkSound.Value = 0
    If gbSEBasicROM Then chkSEBasic.Value = 1 Else chkSEBasic.Value = 0
    
    If Dir$(App.Path & "\sebasic.rom") = "" Then chkSEBasic.Enabled = False
    
    If glMouseType = MOUSE_KEMPSTON Then
        cboMouseType.ListIndex = 1
    ElseIf glMouseType = MOUSE_AMIGA Then
        cboMouseType.ListIndex = 2
    Else
        cboMouseType.ListIndex = 0
    End If
    
    If gbMouseGlobal Then optMouseGlobal.Value = True Else optMouseVB.Value = True
    
    'MM 12.03.2003 - Joystick support code
    '=================================================================================
    'Set up the controles
    With cmbZXJoystick.Item(0)
        .Clear
        .AddItem "(Invalid)"
        Let .ItemData(.NewIndex) = zxjInvalid
        .AddItem "Kempston joystick"
        Let .ItemData(.NewIndex) = zxjKempston
        .AddItem "Cursor joystick"
        Let .ItemData(.NewIndex) = zxjCursor
        .AddItem "Sinclair 1"
        Let .ItemData(.NewIndex) = zxjSinclair1
        .AddItem "Sinclair 2"
        Let .ItemData(.NewIndex) = zxjSinclair2
        .AddItem "Fuller Box"
        Let .ItemData(.NewIndex) = zxjFullerBox
        'MM JD
        .AddItem "User defined"
        Let .ItemData(.NewIndex) = zxjUserDefined
    End With
    With cmbZXJoystick.Item(0)
        .Clear
        .AddItem "(Invalid)"
        Let .ItemData(.NewIndex) = zxjInvalid
        .AddItem "Kempston joystick"
        Let .ItemData(.NewIndex) = zxjKempston
        .AddItem "Cursor joystick"
        Let .ItemData(.NewIndex) = zxjCursor
        .AddItem "Sinclair 1"
        Let .ItemData(.NewIndex) = zxjSinclair1
        .AddItem "Sinclair 2"
        Let .ItemData(.NewIndex) = zxjSinclair2
        .AddItem "Fuller Box"
        Let .ItemData(.NewIndex) = zxjFullerBox
        'MM JD
        .AddItem "User defined"
        Let .ItemData(.NewIndex) = zxjUserDefined
    End With
    'The first activation starts
    Let bIsFirstActivate = True
    '=================================================================================
End Sub

'MM 12.03.2003 - Joystick support code
'=================================================================================
'This code executes every time the window becomes visible
Private Sub Form_Activate()

    'Help-Variable
    Dim iCounter As Integer

    'By the first activate only
    If bIsFirstActivate Then
        'Validity check
        lblZXJoystick.Item(0).Enabled = modSpectrum.bJoystick1Valid
        lblFire.Item(0).Enabled = modSpectrum.bJoystick1Valid
        cmbZXJoystick.Item(0).Enabled = modSpectrum.bJoystick1Valid
        cmbFire.Item(0).Enabled = modSpectrum.bJoystick1Valid
        cmdDefine.Item(0).Enabled = modSpectrum.bJoystick1Valid
        lblZXJoystick.Item(1).Enabled = modSpectrum.bJoystick2Valid
        lblFire.Item(1).Enabled = modSpectrum.bJoystick2Valid
        cmbZXJoystick.Item(1).Enabled = modSpectrum.bJoystick2Valid
        cmbFire.Item(1).Enabled = modSpectrum.bJoystick2Valid
        cmdDefine.Item(1).Enabled = modSpectrum.bJoystick2Valid
        'If joystick1 is valid
        If modSpectrum.bJoystick1Valid Then
            'Get supported buttons
            Let lSupportedButtons = modSpectrum.lPCJoystick1Buttons
            'Initialise Combos
            With cmbFire.Item(0)
                'Add invalid button
                .AddItem "(Invalid)"
                'Initialise ItemData
                Let .ItemData(.NewIndex) = -1
                'Add valid buttons
                For iCounter = 1 To lSupportedButtons
                    'Add button
                    .AddItem "Button " & Trim(CStr(iCounter))
                    'Initialise ItemData
                    Let .ItemData(.NewIndex) = iCounter - 1
                Next iCounter
            End With
            'Refresh the form
            Me.REFRESH
            'Joystick1 values
            If modSpectrum.lPCJoystick1Is = zxjInvalid Then
                Let cmbZXJoystick.Item(0).ListIndex = 0
            Else
                Let cmbZXJoystick.Item(0).ListIndex = modSpectrum.lPCJoystick1Is + 1
            End If
            If modSpectrum.lPCJoystick1Fire = pcjbInvalid Then
                Let cmbFire.Item(0).ListIndex = 0
            Else
                Let cmbFire.Item(0).ListIndex = modSpectrum.lPCJoystick1Fire + 1
            End If
        End If
        'If joystick2 is valid
        If modSpectrum.bJoystick2Valid Then
            'Get supported buttons
            Let lSupportedButtons = modSpectrum.lPCJoystick1Buttons
            'Initialise Combos
            With cmbFire.Item(1)
                'Add invalid button
                .AddItem "(Invalid)"
                'Initialise ItemData
                Let .ItemData(.NewIndex) = -1
                'Add valid buttons
                For iCounter = 1 To lSupportedButtons
                    'Add button
                    .AddItem "Button " & Trim(CStr(iCounter))
                    'Initialise ItemData
                    Let .ItemData(.NewIndex) = iCounter - 1
                Next iCounter
            End With
            'Refresh the form
            Me.REFRESH
            'Joystick2 Values
            If modSpectrum.lPCJoystick2Is = zxjInvalid Then
                Let cmbZXJoystick.Item(1).ListIndex = 0
            Else
                Let cmbZXJoystick.Item(1).ListIndex = modSpectrum.lPCJoystick2Is + 1
            End If
            If modSpectrum.lPCJoystick2Fire = pcjbInvalid Then
                Let cmbFire.Item(1).ListIndex = 0
            Else
                Let cmbFire.Item(1).ListIndex = modSpectrum.lPCJoystick2Fire + 1
            End If
        End If
        'The first activate ends here
        Let bIsFirstActivate = False
    End If
End Sub
'=================================================================================

'MM JD
'User Joystick definition support
Private Sub cmdDefine_Click(Index As Integer)
    'Variables
    Dim frmDefJoystickWnd As frmDefJoystick
    Dim lCounter As Long
    'You cannot redefine an invalid joystick
    If cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) = zxjInvalid Then
        Exit Sub
    End If
    'Error handel
    On Error GoTo CMDDEFINECLICK_ERROR
    'Initialise
    Set frmDefJoystickWnd = New frmDefJoystick
    Load frmDefJoystickWnd
    frmDefJoystickWnd.JoystickNo = Index + 1
    frmDefJoystickWnd.JoystickSetting = cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex)
    frmDefJoystickWnd.FireButton = cmbFire.Item(Index).ListIndex
    'Show
    frmDefJoystickWnd.Show vbModal
    'If the changes where taken over
    If frmDefJoystickWnd.OK Then
        'That is user defined
        For lCounter = 0 To (cmbZXJoystick.Item(Index).ListCount - 1)
            If cmbZXJoystick.Item(Index).ItemData(lCounter) = zxjUserDefined Then
                cmbZXJoystick.Item(Index).ListIndex = lCounter
            End If
        Next lCounter
    End If
    'Set to Nothing
    If Not frmDefJoystickWnd Is Nothing Then
        Set frmDefJoystickWnd = Nothing
    End If
    'The End
    Exit Sub
CMDDEFINECLICK_ERROR:
    'Report error
    MsgBox err.Description, vbInformation + vbOKOnly
    'Set to Nothing
    If Not frmDefJoystickWnd Is Nothing Then
        Set frmDefJoystickWnd = Nothing
    End If
End Sub

'MM JD
'You cannot redefine an invalid joystick
Private Sub cmbZXJoystick_Click(Index As Integer)
    cmdDefine.Item(Index).Enabled = CBool(cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) <> zxjInvalid)
    cmbFire.Item(Index).Enabled = CBool(cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) <> zxjInvalid) And _
                                  CBool(cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) <> zxjUserDefined)
    lblFire.Item(Index).Enabled = CBool(cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) <> zxjInvalid) And _
                                  CBool(cmbZXJoystick.Item(Index).ItemData(cmbZXJoystick.Item(Index).ListIndex) <> zxjUserDefined)
End Sub






Private Sub ts1_Click()
    Dim l As Long
    
    For l = 1 To ts1.Tabs.count
        picFrame(l).Visible = ts1.Tabs(l).Selected
    Next l
End Sub


Private Sub ts1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ts1_Click
End Sub

