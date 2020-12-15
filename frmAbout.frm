VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About vbSpec"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3555
      TabIndex        =   4
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label lblWebSite 
      Caption         =   "http://www.muhi.org/vbspec/"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   780
      Width           =   2265
   End
   Begin VB.Label lblStatic 
      Caption         =   $"frmAbout.frx":000C
      Height          =   1095
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   1140
      Width           =   3915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   720
      X2              =   9240
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   720
      X2              =   9300
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label lblVerInfo 
      Caption         =   "Version 0.00.0000"
      Height          =   255
      Left            =   780
      TabIndex        =   1
      Top             =   360
      Width           =   3795
   End
   Begin VB.Label lblStatic 
      Caption         =   "vbSpec - Sinclair ZX Spectrum Emulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":00D9
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmAbout.frm within vbSpec.vbp
'
'   "About..." dialog for vbSpec
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

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVerInfo.Caption = "Version " & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblWebSite.Font.Underline = False
End Sub


Private Sub lblStatic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    lblWebSite.Font.Underline = False
End Sub


Private Sub lblVerInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblWebSite.Font.Underline = False
End Sub


Private Sub lblWebSite_Click()
    ShellExecute 0, "open", "http://www.muhi.org/vbspec/", vbNullString, vbNullString, 0
End Sub



Private Sub lblWebSite_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblWebSite.Font.Underline = True
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblWebSite.Font.Underline = False
End Sub



