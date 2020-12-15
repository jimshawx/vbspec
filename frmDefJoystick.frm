VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefJoystick 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "User defined joystick"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmDefJoystick.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load..."
      Height          =   375
      Left            =   4485
      TabIndex        =   32
      ToolTipText     =   "Loads a joystick configuration from a file"
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   375
      Left            =   4485
      TabIndex        =   31
      ToolTipText     =   "Saves the current joystick configuration into a file"
      Top             =   990
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog comdlgFile 
      Left            =   4650
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4485
      TabIndex        =   30
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4485
      TabIndex        =   29
      Top             =   150
      Width           =   1095
   End
   Begin VB.Frame fraButtonDefintions 
      Caption         =   "Buttons"
      Height          =   2025
      Left            =   45
      TabIndex        =   25
      Top             =   1965
      Width           =   4350
      Begin MSComctlLib.ListView lvButtons 
         Height          =   1620
         Left            =   120
         TabIndex        =   26
         Top             =   255
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   2858
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdUndefine 
         Caption         =   "Undefine"
         Height          =   375
         Left            =   2985
         TabIndex        =   28
         Top             =   675
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefine 
         Caption         =   "Define"
         Height          =   375
         Left            =   2985
         TabIndex        =   27
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.Frame fraDirectionDefintions 
      Caption         =   "Directions"
      Height          =   1875
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   4365
      Begin VB.TextBox txtValueDown 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3870
         TabIndex        =   12
         Text            =   "255"
         Top             =   615
         Width           =   390
      End
      Begin VB.TextBox txtPortDown 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3180
         TabIndex        =   10
         Text            =   "65535"
         Top             =   615
         Width           =   570
      End
      Begin VB.TextBox txtValueLeft 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3870
         TabIndex        =   18
         Text            =   "255"
         Top             =   975
         Width           =   390
      End
      Begin VB.TextBox txtPortLeft 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3180
         TabIndex        =   16
         Text            =   "65535"
         Top             =   975
         Width           =   570
      End
      Begin VB.TextBox txtValueRight 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3870
         TabIndex        =   24
         Text            =   "255"
         Top             =   1335
         Width           =   390
      End
      Begin VB.TextBox txtPortRight 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3180
         TabIndex        =   22
         Text            =   "65535"
         Top             =   1335
         Width           =   570
      End
      Begin VB.TextBox txtValueUp 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3870
         TabIndex        =   6
         Text            =   "255"
         Top             =   255
         Width           =   390
      End
      Begin VB.TextBox txtPortUp 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   3180
         TabIndex        =   4
         Text            =   "65535"
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox txtRight 
         Height          =   315
         Left            =   2460
         TabIndex        =   20
         Text            =   "XXX"
         Top             =   1335
         Width           =   435
      End
      Begin VB.TextBox txtLeft 
         Height          =   315
         Left            =   2460
         TabIndex        =   14
         Text            =   "X"
         Top             =   975
         Width           =   435
      End
      Begin VB.TextBox txtDown 
         Height          =   315
         Left            =   2460
         TabIndex        =   8
         Text            =   "X"
         Top             =   615
         Width           =   435
      End
      Begin VB.TextBox txtUp 
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Text            =   "X"
         Top             =   255
         Width           =   435
      End
      Begin VB.Label lblOr 
         Caption         =   "or"
         Height          =   210
         Index           =   3
         Left            =   2955
         TabIndex        =   21
         Top             =   1395
         Width           =   165
      End
      Begin VB.Label lblOr 
         Caption         =   "or"
         Height          =   210
         Index           =   2
         Left            =   2955
         TabIndex        =   15
         Top             =   1035
         Width           =   165
      End
      Begin VB.Label lblOr 
         Caption         =   "or"
         Height          =   210
         Index           =   1
         Left            =   2955
         TabIndex        =   9
         Top             =   675
         Width           =   165
      End
      Begin VB.Label lblOr 
         Caption         =   "or"
         Height          =   210
         Index           =   0
         Left            =   2955
         TabIndex        =   3
         Top             =   315
         Width           =   165
      End
      Begin VB.Label lblKomma 
         Caption         =   ","
         Height          =   210
         Index           =   3
         Left            =   3780
         TabIndex        =   11
         Top             =   675
         Width           =   75
      End
      Begin VB.Label lblKomma 
         Caption         =   ","
         Height          =   210
         Index           =   2
         Left            =   3780
         TabIndex        =   17
         Top             =   1035
         Width           =   75
      End
      Begin VB.Label lblKomma 
         Caption         =   ","
         Height          =   210
         Index           =   1
         Left            =   3780
         TabIndex        =   23
         Top             =   1395
         Width           =   75
      End
      Begin VB.Label lblKomma 
         Caption         =   ","
         Height          =   210
         Index           =   0
         Left            =   3780
         TabIndex        =   5
         Top             =   315
         Width           =   75
      End
      Begin VB.Label lblRight 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick right translated into:"
         Height          =   210
         Left            =   195
         TabIndex        =   19
         Top             =   1395
         Width           =   2220
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick left translated into:"
         Height          =   210
         Left            =   195
         TabIndex        =   13
         Top             =   1035
         Width           =   2220
      End
      Begin VB.Label lblDown 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick down translated into:"
         Height          =   210
         Left            =   195
         TabIndex        =   7
         Top             =   675
         Width           =   2220
      End
      Begin VB.Label lblUp 
         Alignment       =   1  'Rechts
         Caption         =   "Joystick up translated into:"
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   315
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmDefJoystick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmDefJoystick.frm within vbSpec.vbp
'
'   "Options->General settings->Joystick->Define" dialog for vbSpec
'
'   Author: Miklos Muhi <vbspec@muhi.org>
'   http://www.muhi.org/vbspec/
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
'Joystick definition rules:
'1. You can only then define a joystick, if the joystick is connected.
'2. If you open this configuration window, you will see the current joystick settigns,
'   described with ports.
'3. You can save and load joystick configuration files (myconfig.jcf)
'   The file format of then JCF-files looks like this:
'   [vbSpec joystick configuration file 0.01]
'   UP=<Key>*<Port>*<Value>
'   DOWN=<Key>*<Port>*<Value>
'   LEFT=<Key>*<Port>*<Value>
'   RIGHT=<Key>*<Port>*<Value>
'   BUTTON=<Key>*<Port>*<Value>
'   ...
'   BUTTON=<Key>*<Port>*<Value>
'
'4. You can redefine define each supported joystick.
Option Explicit

'Variables
Private bIsFirstActivate As Boolean
'The joystick you redefine
Private lJoystickNo As Long
'The current settings
Private lJoystickSetting As Long
'Fire button
Private lFireButton As Long
'Changes
Private bEdit As Boolean
'Keys, ports and  values
Private sKey As String, lPort As Long, lValue As Long
'Edit flags
Private bKey As Boolean, bPortValue As Boolean
'OK flag
Private bOK As Boolean

'This property sets the joystick you are going to redefine
Friend Property Let JoystickNo(ByVal lData As Long)
    lJoystickNo = lData
End Property
'This property provides the joysticks current settings
Friend Property Let JoystickSetting(ByVal lData As Long)
    lJoystickSetting = lData
End Property
'This property provides the current fire button setting
Friend Property Let FireButton(ByVal lData As Long)
    lFireButton = lData
End Property
'Get the OK-Flag
Friend Property Get OK() As Boolean
    OK = bOK
End Property

'The form will be loaded
Private Sub Form_Load()
    'Variables
    Dim oColHeader As ColumnHeader
    'Setup the controls
    txtUp.Text = vbNullString
    txtUp.MaxLength = 3
    txtUp.Locked = True
    txtPortUp.Text = vbNullString
    txtPortUp.MaxLength = 5
    txtValueUp.Text = vbNullString
    txtValueUp.MaxLength = 3
    txtDown.Text = vbNullString
    txtDown.MaxLength = 3
    txtDown.Locked = True
    txtPortDown.Text = vbNullString
    txtPortDown.MaxLength = 5
    txtValueDown.Text = vbNullString
    txtValueDown.MaxLength = 3
    txtLeft.Text = vbNullString
    txtLeft.MaxLength = 3
    txtLeft.Locked = True
    txtPortLeft.Text = vbNullString
    txtPortLeft.MaxLength = 5
    txtValueLeft.Text = vbNullString
    txtValueLeft.MaxLength = 3
    txtRight.Text = vbNullString
    txtRight.MaxLength = 3
    txtRight.Locked = True
    txtPortRight.Text = vbNullString
    txtPortRight.MaxLength = 5
    txtValueRight.Text = vbNullString
    txtValueRight.MaxLength = 3
    With lvButtons
        .AllowColumnReorder = False
        .Appearance = cc3D
        .FlatScrollBar = False
        .FullRowSelect = True
        .GridLines = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = False
        .View = lvwReport
        Set oColHeader = .ColumnHeaders.Add(, , "Button")
        oColHeader.Width = CLng((.Width / 100) * 25)
        Set oColHeader = .ColumnHeaders.Add(, , "Key")
        oColHeader.Width = CLng((.Width / 100) * 19)
        Set oColHeader = .ColumnHeaders.Add(, , "Port")
        oColHeader.Width = CLng((.Width / 100) * 23)
        oColHeader.Alignment = lvwColumnRight
        Set oColHeader = .ColumnHeaders.Add(, , "Value")
        oColHeader.Width = CLng((.Width / 100) * 22)
        oColHeader.Alignment = lvwColumnRight
    End With
    'Default values
    lJoystickNo = 1
    lJoystickSetting = zxjKempston
    lFireButton = 1
    bEdit = True
    bKey = False
    bPortValue = False
    bOK = False
    'Now it's time for the first activation
    bIsFirstActivate = True
End Sub

'The form will be activated
Private Sub Form_Activate()
    'Variables
    Dim lCounter As Long
    Dim oItem As ListItem
    Dim sKey As String
    Dim lPort As Long, lValue As Long
    'Only by the first activation
    If bIsFirstActivate Then
        'If there is no valid joystick
        If Not ((modSpectrum.bJoystick1Valid And (lJoystickNo = 1)) Or _
               ((modSpectrum.bJoystick2Valid And (lJoystickNo = 2)))) Then
            'Close window
            Unload Me
            Exit Sub
        End If
        'Set caption
        Me.Caption = "Redefine joystick " & Trim(CStr(lJoystickNo))
        'You cannot set an invalid joystick
        If lJoystickSetting = zxjInvalid Then
            Unload Me
            Exit Sub
        End If
        'Set values
        txtPortUp.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_UP).lPort))
        txtPortDown.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_DOWN).lPort))
        txtPortLeft.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_LEFT).lPort))
        txtPortRight.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_RIGHT).lPort))
        txtValueUp.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_UP).lValue))
        txtValueDown.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_DOWN).lValue))
        txtValueLeft.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_LEFT).lValue))
        txtValueRight.Text = Trim(CStr(aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_RIGHT).lValue))
        If lJoystickNo = 1 Then
            For lCounter = 1 To lPCJoystick1Buttons
                Set oItem = lvButtons.ListItems.Add(, , "Button" & Trim(CStr(lCounter)))
                lPort = aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_BUTTON_BASE + lCounter).lPort
                lValue = aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_BUTTON_BASE + lCounter).lValue
                If (lPort <> 0) And (lValue <> 0) Then
                    PortValueToKeyStroke lPort, lValue, sKey
                    oItem.SubItems(1) = sKey
                    oItem.SubItems(2) = lPort
                    oItem.SubItems(3) = lValue
                Else
                    oItem.SubItems(1) = " "
                    oItem.SubItems(2) = " "
                    oItem.SubItems(3) = " "
                End If
            Next lCounter
        End If
        If lJoystickNo = 2 Then
            For lCounter = 1 To lPCJoystick2Buttons
                Set oItem = lvButtons.ListItems.Add(, , " ")
                lPort = aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_BUTTON_BASE + lCounter).lPort
                lValue = aJoystikDefinitionTable(lJoystickNo - 1, lJoystickSetting, JDT_BUTTON_BASE + lCounter).lValue
                If (lPort <> 0) And (lValue <> 0) Then
                    PortValueToKeyStroke lPort, lValue, sKey
                    oItem.SubItems(1) = sKey
                    oItem.SubItems(2) = lPort
                    oItem.SubItems(3) = lValue
                Else
                    oItem.SubItems(1) = " "
                    oItem.SubItems(2) = " "
                    oItem.SubItems(3) = " "
                End If
            Next lCounter
        End If
        'Select the first button
        Set lvButtons.SelectedItem = lvButtons.ListItems.Item(1)
        lvButtons_ItemClick lvButtons.SelectedItem
        'No edit
        bEdit = False
        'The first activation ends here
        bIsFirstActivate = False
    End If
End Sub

'The form will be unloaded
Private Sub Form_Unload(Cancel As Integer)
    'Variables
    Dim vbmrValue As VbMsgBoxResult
    'If there are unsaved data
    If bEdit Then
        'Ask the user
        vbmrValue = MsgBox("Do you want to take over your changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
        'Analyse
        Select Case vbmrValue
            'Save changes and exit
            Case vbYes
                'Save changes
                OKButton_Click
            'Do not save changes
            Case vbNo
                'No changes
                bOK = False
            'Cancel
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

'Close and save
Private Sub OKButton_Click()

    'Variablen
    Dim sKey As String
    Dim lPort As Long
    Dim lValue As Long
    Dim oItem As ListItem
    
    'Validate
    If Not Validate Then
        Exit Sub
    End If
    
    'Take over changes
    lPort = CLng(txtPortUp.Text)
    lValue = CLng(txtValueUp.Text)
    PortValueToKeyStroke lPort, lValue, sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_UP).sKey = sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_UP).lPort = lPort
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_UP).lValue = lValue
    lPort = CLng(txtPortDown.Text)
    lValue = CLng(txtValueDown.Text)
    PortValueToKeyStroke lPort, lValue, sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_DOWN).sKey = sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_DOWN).lPort = lPort
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_DOWN).lValue = lValue
    lPort = CLng(txtPortLeft.Text)
    lValue = CLng(txtValueLeft.Text)
    PortValueToKeyStroke lPort, lValue, sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_LEFT).sKey = sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_LEFT).lPort = lPort
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_LEFT).lValue = lValue
    lPort = CLng(txtPortRight.Text)
    lValue = CLng(txtValueRight.Text)
    PortValueToKeyStroke lPort, lValue, sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_RIGHT).sKey = sKey
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_RIGHT).lPort = lPort
    aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_RIGHT).lValue = lValue
    For Each oItem In lvButtons.ListItems
        If (Trim(oItem.SubItems(2)) <> vbNullString) And (Trim(oItem.SubItems(3)) <> vbNullString) Then
            lPort = CLng(oItem.SubItems(2))
            lValue = CLng(oItem.SubItems(3))
            PortValueToKeyStroke lPort, lValue, sKey
            aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_BUTTON_BASE + oItem.Index).sKey = sKey
            aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_BUTTON_BASE + oItem.Index).lPort = lPort
            aJoystikDefinitionTable(lJoystickNo - 1, zxjUserDefined, JDT_BUTTON_BASE + oItem.Index).lValue = lValue
        End If
    Next oItem

    'No edit
    bEdit = False
    'Taken over
    bOK = True
    'Close form
    Unload Me
End Sub

'Close without save
Private Sub CancelButton_Click()
    bOK = False
    Unload Me
End Sub

'Save the joystick configuration
Private Sub cmdSave_Click()

    'Variables
    Dim sFileName As String
    Dim hFile As Long
    Dim oItem As ListItem

    'Error handling
    On Error GoTo CMDSAVECLICK_ERROR

    'Initialise
    With comdlgFile
        .CancelError = True
        .DialogTitle = "Save joystick configuration"
        .DefaultExt = ".jcf"
        .FileName = ""
        .Filter = "vbSpec joystick configuration file (*.jcf)|*.jcf|"
        .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
        .ShowSave
    End With
    'Get file name
    sFileName = comdlgFile.FileName
    
    'Start
    hFile = FreeFile
    Open sFileName For Output As hFile
    
    'Header
    Print #hFile, "[vbSpec joystick configuration file 0.01]"
    'Up
    Print #hFile, "UP=" & Trim(txtUp.Text) & "*" & Trim(CStr(txtPortUp.Text)) & "*" & _
                  Trim(CStr(txtValueUp.Text))
    'Down
    Print #hFile, "DOWN=" & Trim(txtDown.Text) & "*" & Trim(CStr(txtPortDown.Text)) & "*" & _
                  Trim(CStr(txtValueDown.Text))
    'Left
    Print #hFile, "LEFT=" & Trim(txtLeft.Text) & "*" & Trim(CStr(txtPortLeft.Text)) & "*" & _
                  Trim(CStr(txtValueLeft.Text))
    'Right
    Print #hFile, "RIGHT=" & Trim(txtRight.Text) & "*" & Trim(CStr(txtPortRight.Text)) & "*" & _
                  Trim(CStr(txtValueRight.Text))
    'Buttons
    For Each oItem In lvButtons.ListItems
        Print #hFile, "BUTTON=" & Trim(oItem.SubItems(1)) & "*" & Trim(oItem.SubItems(2)) & "*" & _
                      Trim(oItem.SubItems(3))
    Next oItem
    
    'Close file
    Close hFile

    Exit Sub
CMDSAVECLICK_ERROR:
End Sub

'Load a joystick configuration
Private Sub cmdLoad_Click()

    'Variables
    Dim sFileName As String
    Dim hFile As Long
    Dim oItem As ListItem
    Dim sRow As String
    Dim vRow As Variant

    'Error handling
    On Error GoTo CMDLOADCLICK_ERROR

    'Initialise
    With comdlgFile
        .CancelError = True
        .DialogTitle = "Load joystick configuration"
        .DefaultExt = ".jcf"
        .FileName = ""
        .Filter = "vbSpec joystick configuration file (*.jcf)|*.jcf|"
        .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
        .ShowOpen
    End With
    'Get file name
    sFileName = comdlgFile.FileName
    
    'Start
    hFile = FreeFile
    Open sFileName For Input As hFile
    
    'Read the first row
    Input #hFile, sRow
    'The file must begin with the first row
    If sRow = "[vbSpec joystick configuration file 0.01]" Then
        'Clear the button list
        lvButtons.ListItems.Clear
        'Scan the whole file
        Do While Not EOF(hFile)
            'Up
            If UCase(Left(sRow, 3)) = "UP=" Then
                vRow = Split(Right(sRow, (Len(sRow) - 3)), "*")
                txtPortUp.Text = vRow(1)
                txtValueUp.Text = vRow(2)
            End If
            'Down
            If UCase(Left(sRow, 5)) = "DOWN=" Then
                vRow = Split(Right(sRow, (Len(sRow) - 5)), "*")
                txtPortDown.Text = vRow(1)
                txtValueDown.Text = vRow(2)
            End If
            'Left
            If UCase(Left(sRow, 5)) = "LEFT=" Then
                vRow = Split(Right(sRow, (Len(sRow) - 5)), "*")
                txtPortLeft.Text = vRow(1)
                txtValueLeft.Text = vRow(2)
            End If
            'Right
            If UCase(Left(sRow, 6)) = "RIGHT=" Then
                vRow = Split(Right(sRow, (Len(sRow) - 6)), "*")
                txtPortRight.Text = vRow(1)
                txtValueRight.Text = vRow(2)
            End If
            'Button
            If UCase(Left(sRow, 7)) = "BUTTON=" Then
                vRow = Split(Right(sRow, (Len(sRow) - 7)), "*")
                If lJoystickNo = 1 Then
                    If lvButtons.ListItems.count < lPCJoystick1Buttons Then
                        Set oItem = lvButtons.ListItems.Add(, , "Button" & Trim(CStr((lvButtons.ListItems.count + 1))))
                        oItem.SubItems(1) = vRow(0)
                        oItem.SubItems(2) = vRow(1)
                        oItem.SubItems(3) = vRow(2)
                    End If
                Else
                    If lvButtons.ListItems.count < lPCJoystick2Buttons Then
                        Set oItem = lvButtons.ListItems.Add(, , "Button" & Trim(CStr((lvButtons.ListItems.count + 1))))
                        oItem.SubItems(1) = vRow(0)
                        oItem.SubItems(2) = vRow(1)
                        oItem.SubItems(3) = vRow(2)
                    End If
                End If
            End If
            'Read the next row
            Input #hFile, sRow
        Loop
    End If
    
    'Close file
    Close hFile

    Exit Sub
CMDLOADCLICK_ERROR:
End Sub

'The selection is changed
Private Sub lvButtons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'You can define a button only if it has no definition yet
    cmdDefine.Enabled = CBool(Trim(Item.SubItems(2)) = vbNullString)
    'You can undefine a button if it has a definition yet
    cmdUndefine.Enabled = Not cmdDefine.Enabled
End Sub

'Define button
Private Sub cmdDefine_Click()
    'Variables
    Dim frmDefButtonWnd As frmDefButton
    'Error handel
    On Error GoTo CMDDEFINECLICK_ERROR
    'Initialise
    Set frmDefButtonWnd = New frmDefButton
    Load frmDefButtonWnd
    frmDefButtonWnd.Button = lvButtons.SelectedItem.Index
    'Show form
    frmDefButtonWnd.Show vbModal
    'If the user wants to save
    If frmDefButtonWnd.OK Then
        'Get changes
        lvButtons.SelectedItem.SubItems(1) = Trim(CStr(frmDefButtonWnd.ButtonChar))
        lvButtons.SelectedItem.SubItems(2) = Trim(CStr(frmDefButtonWnd.ButtonPort))
        lvButtons.SelectedItem.SubItems(3) = Trim(CStr(frmDefButtonWnd.ButtonValue))
        'That's a change
        bEdit = True
    End If
    'Anihilate objects
    If Not frmDefButtonWnd Is Nothing Then
        Set frmDefButtonWnd = Nothing
    End If
    'End procedure
    Exit Sub
CMDDEFINECLICK_ERROR:
    'Report error
    MsgBox err.Description, vbInformation + vbOKOnly
    'Anihilate objects
    If Not frmDefButtonWnd Is Nothing Then
        Set frmDefButtonWnd = Nothing
    End If
End Sub

'Anihilate a button definition
Private Sub cmdUndefine_Click()
    'Anihilate definition
    lvButtons.SelectedItem.SubItems(1) = " "
    lvButtons.SelectedItem.SubItems(2) = " "
    lvButtons.SelectedItem.SubItems(3) = " "
    'Regulate selection
    lvButtons_ItemClick lvButtons.SelectedItem
    'That's a change
    bEdit = True
End Sub

'Default change hanlder
Private Sub DefaultChangeHanlder()
    bEdit = True
End Sub

'Default focus hanlder
Private Sub DefaultFocusHanlder(ByRef oTextBox As TextBox)
    If oTextBox.Text <> vbNullString Then
        oTextBox.SelStart = 0
        oTextBox.SelLength = Len(oTextBox.Text)
    End If
End Sub

'Define per keystroke
Private Sub txtUp_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bPortValue Then
        bKey = True
        KeyStrokeToPortValue KeyCode, Shift, sKey, lPort, lValue
        txtUp.Text = sKey
        txtPortUp.Text = Trim(CStr(lPort))
        txtValueUp.Text = Trim(CStr(lValue))
        bKey = False
    End If
End Sub
Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bPortValue Then
        bKey = True
        KeyStrokeToPortValue KeyCode, Shift, sKey, lPort, lValue
        txtDown.Text = sKey
        txtPortDown.Text = Trim(CStr(lPort))
        txtValueDown.Text = Trim(CStr(lValue))
        bKey = False
    End If
End Sub
Private Sub txtLeft_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bPortValue Then
        bKey = True
        KeyStrokeToPortValue KeyCode, Shift, sKey, lPort, lValue
        txtLeft.Text = sKey
        txtPortLeft.Text = Trim(CStr(lPort))
        txtValueLeft.Text = Trim(CStr(lValue))
        bKey = False
    End If
End Sub
Private Sub txtRight_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bPortValue Then
        bKey = True
        KeyStrokeToPortValue KeyCode, Shift, sKey, lPort, lValue
        txtRight.Text = sKey
        txtPortRight.Text = Trim(CStr(lPort))
        txtValueRight.Text = Trim(CStr(lValue))
        bKey = False
    End If
End Sub

'Define per port, in value
Private Sub txtPortUp_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortUp.Text) And IsNumeric(txtValueUp.Text)) Then
            txtUp.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortUp.Text), CLng(txtValueUp.Text), sKey
        txtUp.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtValueUp_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortUp.Text) And IsNumeric(txtValueUp.Text)) Then
            txtUp.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortUp.Text), CLng(txtValueUp.Text), sKey
        txtUp.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtPortDown_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortDown.Text) And IsNumeric(txtValueDown.Text)) Then
            txtDown.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortDown.Text), CLng(txtValueDown.Text), sKey
        txtDown.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtValueDown_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortDown.Text) And IsNumeric(txtValueDown.Text)) Then
            txtDown.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortDown.Text), CLng(txtValueDown.Text), sKey
        txtDown.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtPortLeft_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortLeft.Text) And IsNumeric(txtValueLeft.Text)) Then
            txtLeft.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortLeft.Text), CLng(txtValueLeft.Text), sKey
        txtLeft.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtValueLeft_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortLeft.Text) And IsNumeric(txtValueLeft.Text)) Then
            txtLeft.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortLeft.Text), CLng(txtValueLeft.Text), sKey
        txtLeft.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtPortRight_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortRight.Text) And IsNumeric(txtValueRight.Text)) Then
            txtRight.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortRight.Text), CLng(txtValueRight.Text), sKey
        txtRight.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub
Private Sub txtValueRight_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortRight.Text) And IsNumeric(txtValueRight.Text)) Then
            txtRight.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortRight.Text), CLng(txtValueRight.Text), sKey
        txtRight.Text = sKey
        bPortValue = False
    End If
    DefaultChangeHanlder
End Sub

'Focus
Private Sub txtPortUp_GotFocus()
    DefaultFocusHanlder txtPortUp
End Sub
Private Sub txtValueUp_GotFocus()
    DefaultFocusHanlder txtValueUp
End Sub
Private Sub txtPortDown_GotFocus()
    DefaultFocusHanlder txtPortDown
End Sub
Private Sub txtValueDown_GotFocus()
    DefaultFocusHanlder txtValueDown
End Sub
Private Sub txtPortLeft_GotFocus()
    DefaultFocusHanlder txtPortLeft
End Sub
Private Sub txtValueLeft_GotFocus()
    DefaultFocusHanlder txtValueLeft
End Sub
Private Sub txtPortRight_GotFocus()
    DefaultFocusHanlder txtPortRight
End Sub
Private Sub txtValueRight_GotFocus()
    DefaultFocusHanlder txtValueRight
End Sub

'Validate joystick settings
Private Function Validate() As Boolean

    'Variables
    Dim oItem As ListItem

    'Preset
    Validate = False
    
    'All values must be numeric
    If Not IsNumeric(txtPortUp.Text) Then
        MsgBox "Please enter a numeric value for the up port.", vbInformation + vbOKOnly
        txtPortUp.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtPortDown.Text) Then
        MsgBox "Please enter a numeric value for the down port.", vbInformation + vbOKOnly
        txtPortDown.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtPortLeft.Text) Then
        MsgBox "Please enter a numeric value for the left port.", vbInformation + vbOKOnly
        txtPortLeft.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtPortRight.Text) Then
        MsgBox "Please enter a numeric value for the right port.", vbInformation + vbOKOnly
        txtPortRight.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtValueUp.Text) Then
        MsgBox "Please enter a numeric value for the up in.", vbInformation + vbOKOnly
        txtValueUp.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtValueDown.Text) Then
        MsgBox "Please enter a numeric value for the down in.", vbInformation + vbOKOnly
        txtValueDown.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtValueLeft.Text) Then
        MsgBox "Please enter a numeric value for the left in.", vbInformation + vbOKOnly
        txtValueLeft.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtValueRight.Text) Then
        MsgBox "Please enter a numeric value for the right in.", vbInformation + vbOKOnly
        txtValueRight.SetFocus
        Exit Function
    End If
    
    'All port values must be between 0 and 65535
    If CLng(txtPortUp.Text) < 0 Or CLng(txtPortUp.Text) > 65535 Then
        MsgBox "Please enter a value between 0 and 65535 for the up port value.", vbInformation + vbOKOnly
        txtPortUp.SetFocus
        Exit Function
    End If
    If CLng(txtPortDown.Text) < 0 Or CLng(txtPortDown.Text) > 65535 Then
        MsgBox "Please enter a value between 0 and 65535 for the down port value.", vbInformation + vbOKOnly
        txtPortDown.SetFocus
        Exit Function
    End If
    If CLng(txtPortLeft.Text) < 0 Or CLng(txtPortLeft.Text) > 65535 Then
        MsgBox "Please enter a value between 0 and 65535 for the left port value.", vbInformation + vbOKOnly
        txtPortLeft.SetFocus
        Exit Function
    End If
    If CLng(txtPortRight.Text) < 0 Or CLng(txtPortRight.Text) > 65535 Then
        MsgBox "Please enter a value between 0 and 65535 for the right port value.", vbInformation + vbOKOnly
        txtPortRight.SetFocus
        Exit Function
    End If

    'All in values must be between 0 and 255
    If CLng(txtValueUp.Text) < 0 Or CLng(txtValueUp.Text) > 255 Then
        MsgBox "Please enter a value between 0 and 255 for the up in.", vbInformation + vbOKOnly
        txtValueUp.SetFocus
        Exit Function
    End If
    If CLng(txtValueDown.Text) < 0 Or CLng(txtValueDown.Text) > 255 Then
        MsgBox "Please enter a value between 0 and 255 for the down in.", vbInformation + vbOKOnly
        txtValueDown.SetFocus
        Exit Function
    End If
    If CLng(txtValueLeft.Text) < 0 Or CLng(txtValueLeft.Text) > 255 Then
        MsgBox "Please enter a value between 0 and 255 for the left in.", vbInformation + vbOKOnly
        txtValueLeft.SetFocus
        Exit Function
    End If
    If CLng(txtValueRight.Text) < 0 Or CLng(txtValueRight.Text) > 255 Then
        MsgBox "Please enter a value between 0 and 255 for the right in.", vbInformation + vbOKOnly
        txtValueRight.SetFocus
        Exit Function
    End If
    
    'Validate list items
    For Each oItem In lvButtons.ListItems
        If (Trim(oItem.SubItems(2)) <> vbNullString) And (Trim(oItem.SubItems(3)) <> vbNullString) Then
            If Not IsNumeric(oItem.SubItems(2)) Then
                MsgBox "Please enter numeric value for button " & Trim(CStr(oItem.Index)) & " port.", vbInformation + vbOKOnly
                Set lvButtons.SelectedItem = oItem
                Exit Function
            End If
            If Not IsNumeric(oItem.SubItems(3)) Then
                MsgBox "Please enter numeric value for button " & Trim(CStr(oItem.Index)) & " in.", vbInformation + vbOKOnly
                Set lvButtons.SelectedItem = oItem
                Exit Function
            End If
            If CLng(oItem.SubItems(2)) < 0 Or CLng(oItem.SubItems(2)) > 65535 Then
                MsgBox "Please enter a value between 0 and 65535 for button " & Trim(CStr(oItem.Index)) & " port.", vbInformation + vbOKOnly
                Set lvButtons.SelectedItem = oItem
                Exit Function
            End If
            If Not IsNumeric(oItem.SubItems(3)) Then
                MsgBox "Please enter numeric value for button " & Trim(CStr(oItem.Index)) & " in.", vbInformation + vbOKOnly
                Set lvButtons.SelectedItem = oItem
                Exit Function
            End If
        Else
            If (Trim(oItem.SubItems(2)) <> vbNullString) Or (Trim(oItem.SubItems(3)) <> vbNullString) Then
                MsgBox "You must enter a port number and an in value to define a button."
                Set lvButtons.SelectedItem = oItem
                Exit Function
            End If
        End If
    Next oItem
    
    'That shoul be OK
    Validate = True
End Function
