VERSION 5.00
Begin VB.Form frmPoke 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Poke Memory"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmPoke.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdPeek 
      Caption         =   "Peek"
      Height          =   330
      Left            =   2940
      TabIndex        =   6
      Top             =   90
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   2940
      TabIndex        =   5
      ToolTipText     =   "This button closes the window (the others don't)"
      Top             =   810
      Width           =   1635
   End
   Begin VB.CommandButton cmdPoke 
      Caption         =   "&Poke"
      Default         =   -1  'True
      Height          =   330
      Left            =   2940
      TabIndex        =   4
      ToolTipText     =   "The poke-operation will be completed and the controls reset."
      Top             =   450
      Width           =   1635
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   1485
      TabIndex        =   3
      ToolTipText     =   "Here comes the value that must be poked (0..255)"
      Top             =   435
      Width           =   1305
   End
   Begin VB.TextBox txtAddress 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   1485
      TabIndex        =   1
      ToolTipText     =   "Enter here a number greater than 16384 (Spectrum RAM)"
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Rechts
      Caption         =   "&Value"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   510
      Width           =   1260
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Rechts
      Caption         =   "&Address"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   1260
   End
End
Attribute VB_Name = "frmPoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmPoke.frm within vbSpec.vbp
'
'   Author: Miklos Muhi <miklos.muhi@bakonyi.de>
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

'*******************************************************************************
'Cancel operation
'*******************************************************************************
Private Sub cmdCancel_Click()

    'Unload the form
    Unload Me
End Sub

Private Sub cmdPeek_Click()
    txtValue.Text = Trim(CStr(modZ80.peekb(CLng(txtAddress.Text))))
End Sub

'*******************************************************************************
'Poke memory
'*******************************************************************************
Private Sub cmdPoke_Click()

    'Variables
    Dim lAddress As Long, lValue As Long

    'Validation -- there must be a value for the Address
    If Trim(txtAddress.Text) = vbNullString Then
        'Report the problem
        MsgBox "You must give an address to poke to!", vbInformation + vbOKOnly
        'Set the focus
        txtAddress.SetFocus
        'Leave this function
        Exit Sub
    End If
    
    'Validation -- there must be a value to poke
    If Trim(txtValue.Text) = vbNullString Then
        'Report the problem
        MsgBox "Missing a value to poke!", vbInformation + vbOKOnly
        'Set the focus
        txtValue.SetFocus
        'Leave this function
        Exit Sub
    End If
    
    'Initialisation
    Let lAddress = -1
    Let lValue = -1
    
    'Error handling
    On Error GoTo POKE_ERROR
    
    'Convert the values
    Let lAddress = CLng(txtAddress.Text)
    Let lValue = CLng(txtValue.Text)
    
    'Poke value
    modZ80.pokeb lAddress, lValue
    
    'Reset controls
    Let txtAddress.Text = vbNullString
    Let txtValue.Text = vbNullString
    'Reset focus
    txtAddress.SetFocus
    
    'There wasn't any problem
    Exit Sub
POKE_ERROR:

    'There can be only one problem: the conversions of the strings in long or byte failed
    'If the address-conversion failed
    If lAddress = -1 Then
        'Report the problem
        MsgBox "Invalid address format!", vbInformation + vbOKOnly
    'If the address-conversion was a success
    Else
        'If the address-conversion failed
        If lValue = -1 Then
            'Report the problem
            MsgBox "The value has an invalid format!", vbInformation + vbOKOnly
        'I don't know this...
        Else
            'Report the problem
            MsgBox "Unknown error!", vbInformation + vbOKOnly
        End If
    End If
End Sub
