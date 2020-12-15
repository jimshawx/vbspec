VERSION 5.00
Begin VB.Form frmTapePlayer 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "vbSpec TZX Tape Controls"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lstBlocks 
      Height          =   2595
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   3795
   End
   Begin VB.CommandButton cmdNext 
      Height          =   375
      Left            =   3060
      Picture         =   "frmTapePlayer.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Next Block"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   375
      Left            =   2280
      Picture         =   "frmTapePlayer.frx":0122
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Previous Block"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdRewind 
      Height          =   375
      Left            =   1440
      Picture         =   "frmTapePlayer.frx":0244
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Rewind to start"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdStop 
      Height          =   375
      Left            =   720
      Picture         =   "frmTapePlayer.frx":0366
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Stop/Pause"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   375
      Left            =   0
      Picture         =   "frmTapePlayer.frx":0488
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   675
   End
End
Attribute VB_Name = "frmTapePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmTapePlayer.frm within vbSpec.vbp
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

Public Sub UpdateCurBlock()
    If lstBlocks.ListCount > TZXCurBlock Then lstBlocks.ListIndex = TZXCurBlock
End Sub

Public Sub UpdateTapeList()
    Dim l As Long
    Dim lType As Long, lLen As Long, sText As String
    
    lstBlocks.Clear
    
    For l = 0 To TZXNumBlocks - 1
        GetTZXBlockInfo l, lType, sText, lLen
        lstBlocks.AddItem Hex$(lType) & " - " & sText
    Next l
    
    If lstBlocks.ListCount > TZXCurBlock Then lstBlocks.ListIndex = TZXCurBlock
End Sub

Private Sub cmdNext_Click()
    If (gbTZXInserted) And (TZXCurBlock < TZXNumBlocks - 1) Then
        SetCurTZXBlock TZXCurBlock + 1
    End If
    frmMainWnd.SetFocus
End Sub

Private Sub cmdPlay_Click()
    If gbTZXInserted Then StartTape
    frmMainWnd.SetFocus
End Sub

Private Sub cmdPrev_Click()
    If (gbTZXInserted) And (TZXCurBlock > 0) Then
        SetCurTZXBlock TZXCurBlock - 1
    End If
    frmMainWnd.SetFocus
End Sub

Private Sub cmdRewind_Click()
    If gbTZXInserted Then SetCurTZXBlock 0
    frmMainWnd.SetFocus
End Sub


Private Sub cmdStop_Click()
    If gbTZXInserted Then StopTape
    frmMainWnd.SetFocus
End Sub


Private Sub Form_Activate()
    UpdateCurBlock
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And 4) Then Exit Sub
    doKey True, KeyCode, Shift
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    doKey False, KeyCode, Shift
End Sub

Private Sub Form_Load()
    Dim X As Long, y As Long
    
    X = val(GetSetting("Grok", "vbSpec", "TapeWndX", "-1"))
    y = val(GetSetting("Grok", "vbSpec", "TapeWndY", "-1"))
    
    If X >= 0 And X <= (Screen.Width - Screen.TwipsPerPixelX * 16) Then
        Me.Left = X
    End If
    If y >= 0 And y <= (Screen.Height - Screen.TwipsPerPixelY * 16) Then
        Me.Top = y
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Grok", "vbSpec", "TapeWndX", CStr(Me.Left)
    SaveSetting "Grok", "vbSpec", "TapeWndY", CStr(Me.Top)
    
    frmMainWnd.mnuOptions(4).Checked = False
End Sub

Private Sub lstBlocks_Click()
    UpdateCurBlock
End Sub


