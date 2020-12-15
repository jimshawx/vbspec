VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSaveBinary 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Save Binary Data"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmSaveBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cboBase 
      Height          =   315
      Index           =   1
      Left            =   5280
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtLength 
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   60
      Width           =   4515
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtAddr 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cboBase 
      Height          =   315
      Index           =   0
      Left            =   2460
      Style           =   2  'Dropdown-Liste
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      Caption         =   "&Length:"
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "&Filename:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "Memory &Address:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "frmSaveBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmSaveBinary.frm within vbSpec.vbp
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'
'   Copyright (C)1999-2003 Grok Developments Ltd.
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
    Dim lAddr As Long, lLen As Long, h As Long, s As String
    
    On Error Resume Next
    
    ' // Parse txtAddress and put the decimal equivalent in lAddr
    If Left$(txtAddr, 1) = "$" Then
        lAddr = val("&H" & Mid$(txtAddr, 2) & "&")
    ElseIf cboBase(0).Text = "Hex" Then
        lAddr = val("&H" & txtAddr & "&")
    Else
        lAddr = val(txtAddr)
    End If
    
    ' // Parse txtLength and put the decimal equivalent in lLen
    If Left$(txtLength, 1) = "$" Then
        lLen = val("&H" & Mid$(txtLength, 2) & "&")
    ElseIf cboBase(1).Text = "Hex" Then
        lLen = val("&H" & txtLength & "&")
    Else
        lLen = val(txtLength)
    End If
    
    If Dir$(txtFile) <> "" Then
        If MsgBox(txtFile & vbCrLf & vbCrLf & "Existing file will be overwritten.", vbOKCancel Or vbExclamation, "vbSpec") = vbCancel Then
            Exit Sub
        Else
            Kill txtFile
        End If
    End If
    
    err.Clear
    h = FreeFile
    Open txtFile For Binary Access Write As h
              
    If lAddr + lLen > 65536 Then
        If MsgBox("Output will overrun the 64K upper boundary." & vbCrLf & "Do you want to save the " & CStr(65536 - lAddr) & " bytes up to address 65535?", vbYesNo Or vbQuestion, "vbSpec") = vbNo Then
            Close #h
            Exit Sub
        Else
            lLen = 65536 - lAddr
        End If
    End If
    
    ' // Save the memory block
    lLen = lAddr + lLen
    Do
        Put #h, , gRAMPage(glPageAt(glMemAddrDiv16384(lAddr)), lAddr And 16383)
        lAddr = lAddr + 1
    Loop Until lAddr >= lLen
    
    Me.Hide
End Sub


Private Sub cmdOpen_Click()
    On Error Resume Next
    
    err.Clear
    dlgCommon.DialogTitle = "Open Binary File"
    dlgCommon.DefaultExt = ".bin"
    dlgCommon.FileName = ""
    dlgCommon.Filter = "All Files|*.*"
    dlgCommon.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNLongNames
    dlgCommon.CancelError = True
    If Dir$(txtFile) <> "" Then dlgCommon.InitDir = txtFile
    
    dlgCommon.ShowSave
    If err.Number = cdlCancel Then
        Exit Sub
    End If
    
    If dlgCommon.FileName <> "" Then
        txtFile.Text = dlgCommon.FileName
    End If
End Sub


Private Sub Form_Load()
    cboBase(0).AddItem "Decimal"
    cboBase(1).AddItem "Decimal"
    cboBase(0).AddItem "Hex"
    cboBase(1).AddItem "Hex"

    cboBase(0).ListIndex = 0
    cboBase(1).ListIndex = 0
End Sub


Private Sub txtAddr_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 102 Then KeyAscii = KeyAscii - 32
    If KeyAscii > 70 Then KeyAscii = 0
    If KeyAscii > 57 And KeyAscii < 65 Then KeyAscii = 0
    
    If KeyAscii = 36 Then cboBase(0).ListIndex = 1
    If KeyAscii > 64 Then cboBase(0).ListIndex = 1
End Sub


Private Sub txtLength_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 102 Then KeyAscii = KeyAscii - 32
    If KeyAscii > 70 Then KeyAscii = 0
    If KeyAscii > 57 And KeyAscii < 65 Then KeyAscii = 0
    
    If KeyAscii = 36 Then cboBase(1).ListIndex = 1
    If KeyAscii > 64 Then cboBase(1).ListIndex = 1
End Sub


