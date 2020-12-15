VERSION 5.00
Begin VB.Form frmDefButton 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Joystick button definition"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmDefButton.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   585
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   765
      TabIndex        =   6
      Top             =   585
      Width           =   1095
   End
   Begin VB.TextBox txtValueFire 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Text            =   "255"
      Top             =   90
      Width           =   390
   End
   Begin VB.TextBox txtPortFire 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   2670
      TabIndex        =   2
      Text            =   "65535"
      Top             =   90
      Width           =   570
   End
   Begin VB.TextBox txtFire 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Text            =   "X"
      Top             =   90
      Width           =   435
   End
   Begin VB.Label lblOr 
      Caption         =   "or"
      Height          =   210
      Index           =   0
      Left            =   2430
      TabIndex        =   5
      Top             =   150
      Width           =   165
   End
   Begin VB.Label lblKomma 
      Caption         =   ","
      Height          =   210
      Index           =   0
      Left            =   3270
      TabIndex        =   3
      Top             =   150
      Width           =   75
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Rechts
      Caption         =   "Button X translated into:"
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   1710
   End
End
Attribute VB_Name = "frmDefButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmDefButton.frm within vbSpec.vbp
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
Option Explicit

'Variables
Private bIsFirstActivate As Boolean
'Button
Private lButton As Long
'If the Form was Closed with OK
Private bOK As Boolean
'If there was any change
Private bEdit As Boolean
'Button definitions
Private sButton As String
Private lButtonPort As Long
Private lButtonValue As Long
'Keys, ports and  values
Private sKey As String, lPort As Long, lValue As Long
'Edit flags
Private bKey As Boolean, bPortValue As Boolean

'This property provides the number of the button wich you set up
Friend Property Let Button(ByVal lData As Long)
    lButton = lData
End Property
'Provides, if the form was closed with OK
Friend Property Get OK() As Boolean
    OK = bOK
End Property
'Provider the button data
Friend Property Get ButtonChar() As String
    ButtonChar = sButton
End Property
Friend Property Get ButtonPort() As Long
    ButtonPort = lButtonPort
End Property
Friend Property Get ButtonValue() As Long
    ButtonValue = lButtonValue
End Property

'The form will be loaded
Private Sub Form_Load()
    'Setup the controls
    lblButton.Caption = vbNullString
    txtFire.Text = vbNullString
    txtFire.MaxLength = 3
    txtFire.Locked = True
    txtPortFire.Text = vbNullString
    txtPortFire.MaxLength = 5
    txtValueFire.Text = vbNullString
    txtValueFire.MaxLength = 3
    'Default values
    lButton = 1
    bOK = False
    bEdit = False
    sButton = vbNullString
    lButtonPort = 0
    lButtonValue = 0
    bKey = False
    bPortValue = False
    'Now it's time for the first activation
    bIsFirstActivate = True
End Sub

'The form will be activated
Private Sub Form_Activate()
    'Only by the first activation
    If bIsFirstActivate Then
        'If there is no valid joystick
        If Not (modSpectrum.bJoystick1Valid Or modSpectrum.bJoystick2Valid) Then
            'Close window
            Unload Me
            Exit Sub
        End If
        'Set label caption
        lblButton.Caption = "Button " & Trim(CStr(lButton)) & " translated into:"
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
        vbmrValue = MsgBox("Do you want to save your changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
        'Analyse
        Select Case vbmrValue
            'Save changes and exit
            Case vbYes
                'Save changes
                OKButton_Click
            'Do not save changes
            Case vbNo
                'No saved data
                bOK = False
            'Cancel
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

'Close and save
Private Sub OKButton_Click()
    'Save changes
    sButton = txtFire.Text
    If IsNumeric(txtPortFire.Text) Then
        lButtonPort = CLng(txtPortFire.Text)
    Else
        lButtonPort = 0
    End If
    If IsNumeric(txtValueFire.Text) Then
        lButtonValue = CLng(txtValueFire.Text)
    Else
        lButtonValue = 0
    End If
    'Changes over
    bEdit = False
    'Close window
    bOK = True
    Unload Me
End Sub

'Close without save
Private Sub CancelButton_Click()
    bOK = False
    Unload Me
End Sub

'There was a change
Private Sub DefaultChange()
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
Private Sub txtFire_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bPortValue Then
        bKey = True
        KeyStrokeToPortValue KeyCode, Shift, sKey, lPort, lValue
        txtFire.Text = sKey
        txtPortFire.Text = Trim(CStr(lPort))
        txtValueFire.Text = Trim(CStr(lValue))
        bKey = False
    End If
End Sub
Private Sub txtFire_Change()
    DefaultChange
End Sub
Private Sub txtPortFire_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortFire.Text) And IsNumeric(txtValueFire.Text)) Then
            txtFire.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortFire.Text), CLng(txtValueFire.Text), sKey
        txtFire.Text = sKey
        bPortValue = False
    End If
    DefaultChange
End Sub
Private Sub txtValueFire_Change()
    If Not bKey Then
        bPortValue = True
        If Not (IsNumeric(txtPortFire.Text) And IsNumeric(txtValueFire.Text)) Then
            txtFire.Text = vbNullString
            Exit Sub
        End If
        PortValueToKeyStroke CLng(txtPortFire.Text), CLng(txtValueFire.Text), sKey
        txtFire.Text = sKey
        bPortValue = False
    End If
    DefaultChange
End Sub

'Focus
Private Sub txtPortFire_GotFocus()
    DefaultFocusHanlder txtPortFire
End Sub
Private Sub txtValueFire_GotFocus()
    DefaultFocusHanlder txtValueFire
End Sub


