VERSION 5.00
Begin VB.Form frmDisplayOpt 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Display Options"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmDisplayOpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   165
      Width           =   1095
   End
   Begin VB.Frame fraDisplaySize 
      Caption         =   "&Display Size"
      Height          =   2145
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2595
      Begin VB.OptionButton optFullScreen 
         Caption         =   "Full screen (auto size)"
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   1725
         Width           =   2295
      End
      Begin VB.OptionButton optTriple 
         Caption         =   "&Triple size (768 x 576)"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optDouble 
         Caption         =   "&Double size (512 x 384)"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1155
         Width           =   2175
      End
      Begin VB.OptionButton optDoubleHeight 
         Caption         =   "Double &height (256 x 384)"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   870
         Width           =   2295
      End
      Begin VB.OptionButton optDoubleWidth 
         Caption         =   "Double &width (512 x 192)"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   585
         Width           =   2235
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "&Normal size (256 x 192)"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmDisplayOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmDisplayOpt.frm within vbSpec.vbp
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
'   Full screen support: Miklos Muhi <vbspec@muhi.org>
'
'   Copyright (C)1999-2002 Grok Developments Ltd.
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
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'MM 16.04.2003
    Dim lFullHeight As Long, lFullWidth As Long
    Dim lXFactor As Long, lYFactor As Long, lCommonFactor As Long
    
    If optNormal.Value Then
        'MM 16.04.2003
        bFullScreen = False
        SetDisplaySize 256, 192
    ElseIf optDoubleWidth.Value Then
        'MM 16.04.2003
        bFullScreen = False
        SetDisplaySize 512, 192
    ElseIf optDoubleHeight.Value Then
        'MM 16.04.2003
        bFullScreen = False
        SetDisplaySize 256, 384
    ElseIf optDouble.Value Then
        'MM 16.04.2003
        bFullScreen = False
        SetDisplaySize 512, 384
    ElseIf optTriple.Value Then
        'MM 16.04.2003
        bFullScreen = False
        SetDisplaySize 768, 576
    
    'MM 16.04.2003
    ElseIf optFullScreen.Value Then
        lXFactor = Fix(CLng(Screen.Width / Screen.TwipsPerPixelX) / 256)
        lYFactor = Fix(CLng(Screen.Height / Screen.TwipsPerPixelY) / 176)
        lCommonFactor = lXFactor
        If lYFactor < lXFactor Then
            lCommonFactor = lYFactor
        End If
        bFullScreen = True
        SetDisplaySize 256 * lCommonFactor, 192 * lCommonFactor
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    If bFullScreen Then
        optFullScreen.Value = True
    Else
        If glDisplayWidth = 256 Then
            If glDisplayHeight = 192 Then
                optNormal.Value = True
            ElseIf glDisplayHeight = 384 Then
                optDoubleHeight.Value = True
            End If
        ElseIf glDisplayWidth = 512 Then
            If glDisplayHeight = 192 Then
                optDoubleWidth.Value = True
            ElseIf glDisplayHeight = 384 Then
                optDouble.Value = True
            End If
        ElseIf glDisplayWidth = 768 Then
            If glDisplayHeight = 576 Then
                optTriple.Value = True
            End If
        End If
    End If
End Sub


