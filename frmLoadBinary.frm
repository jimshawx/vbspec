VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadBinary 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Load Binary Data"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmLoadBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboBase 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown-Liste
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtAddr 
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   60
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   180
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "Memory &Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "&Filename:"
      Height          =   255
      Left            =   540
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmLoadBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmLoadBinary.frm within vbSpec.vbp
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
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim lAddr As Long, h As Long, s As String
    
    On Error Resume Next
    
    If Left$(txtAddr, 1) = "$" Then
        lAddr = val("&H" & Mid$(txtAddr, 2) & "&")
    ElseIf cboBase.Text = "Hex" Then
        lAddr = val("&H" & txtAddr & "&")
    Else
        lAddr = val(txtAddr)
    End If
    
    If Dir$(txtFile) = "" Then
        MsgBox txtFile & vbCrLf & vbCrLf & "File not found.", vbOKOnly Or vbExclamation, "vbSpec"
        Exit Sub
    Else
        err.Clear
        h = FreeFile
        Open txtFile For Binary As h
        
        If err.Number = 76 Then
            Close #h
            MsgBox txtFile & vbCrLf & vbCrLf & "File not found.", vbOKOnly Or vbExclamation, "vbSpec"
            Exit Sub
        End If
        
        If LOF(h) + lAddr > 65536 Then
            If MsgBox(txtFile & vbCrLf & vbCrLf & "File will overrun the 64K upper boundary." & vbCrLf & "Do you want to load the first " & CStr(65536 - lAddr) & " bytes?", vbYesNo Or vbQuestion, "vbSpec") = vbNo Then
                Close #h
                Exit Sub
            End If
        End If
        
        ' // Load as many bytes as will fit into the Z80 memory space
        If LOF(h) > 65536 Then s = Input(65536, #h) Else s = Input(LOF(h), #h)
        If lAddr + Len(s) > 65536 Then s = Left$(s, 65536 - lAddr)
        Close #h
        
        h = 1
        Do
            gRAMPage(glPageAt(glMemAddrDiv16384(lAddr)), lAddr And 16383) = Asc(Mid$(s, h, 1))
            lAddr = lAddr + 1
            h = h + 1
        Loop Until h > Len(s)
    End If
    
    Me.Hide
    
    'MM 23.04.2003
    'Bugfix - you will see instantly if you binary load a screen
    initscreen
    screenPaint
End Sub

Private Sub cmdOpen_Click()
    On Error Resume Next
    
    err.Clear
    dlgCommon.DialogTitle = "Open Binary File"
    dlgCommon.DefaultExt = ".bin"
    dlgCommon.FileName = ""
    dlgCommon.Filter = "All Files|*.*"
    dlgCommon.Flags = cdlOFNFileMustExist Or cdlOFNExplorer Or cdlOFNLongNames
    dlgCommon.CancelError = True
    If Dir$(txtFile) <> "" Then dlgCommon.InitDir = txtFile
    
    dlgCommon.ShowOpen
    If err.Number = cdlCancel Then
        Exit Sub
    End If
    
    If dlgCommon.FileName <> "" Then
        txtFile.Text = dlgCommon.FileName
    End If
End Sub

Private Sub Form_Load()
    cboBase.AddItem "Decimal"
    cboBase.AddItem "Hex"

    cboBase.ListIndex = 0
End Sub


Private Sub txtAddr_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 102 Then KeyAscii = KeyAscii - 32
    If KeyAscii > 70 Then KeyAscii = 0
    If KeyAscii > 57 And KeyAscii < 65 Then KeyAscii = 0
    
    If KeyAscii > 64 Then cboBase.ListIndex = 1
End Sub

