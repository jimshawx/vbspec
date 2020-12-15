VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmZXPrinter 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ZX Printer"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Height          =   345
      Left            =   900
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmZXPrinter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Clear Output"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optAlphacom 
      Height          =   300
      Left            =   2820
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmZXPrinter.frx":015A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optZX 
      Height          =   300
      Left            =   1500
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmZXPrinter.frx":04AC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "ZX Printer"
      Top             =   60
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Timer tmrFormFeed 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   780
   End
   Begin VB.PictureBox picLogo 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   4095
      TabIndex        =   4
      Top             =   3480
      Width           =   4095
      Begin VB.Image imgAlpha 
         Height          =   390
         Left            =   60
         Picture         =   "frmZXPrinter.frx":076E
         Top             =   120
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Image imgSinclair 
         Height          =   390
         Left            =   180
         Picture         =   "frmZXPrinter.frx":10A4
         Top             =   60
         Width           =   3090
      End
   End
   Begin VB.CommandButton cmdFF 
      Height          =   345
      Left            =   480
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmZXPrinter.frx":2A06
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Form Feed"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Left            =   60
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmZXPrinter.frx":2B60
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Output"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.VScrollBar vs 
      Height          =   2655
      LargeChange     =   8
      Left            =   3840
      SmallChange     =   8
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   480
      Width           =   3840
      Begin MSComDlg.CommonDialog dlgSave 
         Left            =   2640
         Top             =   780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "bmp"
         DialogTitle     =   "Save Output"
         FileName        =   "untitled"
         Filter          =   "Windows Bitmap (*.bmp)|*.bmp|All Files (*.*)|*.*"
      End
   End
End
Attribute VB_Name = "frmZXPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*******************************************************************************
'   frmZXPrinter.frm within vbSpec.vbp
'
'   Author: Chris Cowley <ccowley@grok.co.uk>
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

Private iWidth As Long


Private Sub InitZXPrinterBitmap()
    bmiZXPrinter.bmiHeader.biSize = Len(bmiZXPrinter.bmiHeader)
    bmiZXPrinter.bmiHeader.biWidth = 256
    bmiZXPrinter.bmiHeader.biHeight = 1152
    bmiZXPrinter.bmiHeader.biPlanes = 1
    bmiZXPrinter.bmiHeader.biBitCount = 1
    bmiZXPrinter.bmiHeader.biCompression = BI_RGB
    bmiZXPrinter.bmiHeader.biSizeImage = 0
    bmiZXPrinter.bmiHeader.biXPelsPerMeter = 200
    bmiZXPrinter.bmiHeader.biYPelsPerMeter = 200
    bmiZXPrinter.bmiHeader.biClrUsed = 2
    bmiZXPrinter.bmiHeader.biClrImportant = 2
    
    If optZX.Value = True Then
        bmiZXPrinter.bmiColors(0).rgbRed = 192
        bmiZXPrinter.bmiColors(0).rgbGreen = 192
        bmiZXPrinter.bmiColors(0).rgbBlue = 192
        bmiZXPrinter.bmiColors(1).rgbRed = 0
        bmiZXPrinter.bmiColors(1).rgbGreen = 0
        bmiZXPrinter.bmiColors(1).rgbBlue = 0
    Else
        bmiZXPrinter.bmiColors(0).rgbRed = 255
        bmiZXPrinter.bmiColors(0).rgbGreen = 255
        bmiZXPrinter.bmiColors(0).rgbBlue = 255
        bmiZXPrinter.bmiColors(1).rgbRed = 64
        bmiZXPrinter.bmiColors(1).rgbGreen = 64
        bmiZXPrinter.bmiColors(1).rgbBlue = 192
    End If
    
    ReDim gcZXPrinterBits(36864) ' // 1152 * 32
    glZXPrinterBMPHeight = 1152
End Sub


Private Sub SaveMonoBitmap(sFile As String)
    Dim h As Long, lSize As Long, b() As Byte, l As Long, X As Long, d As Long
    
    h = FreeFile
    Open sFile For Binary Access Write As h
    
    lSize = 14 + Len(bmiZXPrinter.bmiHeader) + 8 + lZXPrinterY * 32
    
    ' // BITMAPFILEHEADER - 14
     
    bmiZXPrinter.bmiHeader.biHeight = lZXPrinterY
    bmiZXPrinter.bmiHeader.biSizeImage = lZXPrinterY * 32
     
    ' // Flip the stored image as RGB bitmaps are stored bottom-left to top-right
    d = lZXPrinterY - 1
    ReDim b(lZXPrinterY * 32)
    For l = 0 To lZXPrinterY - 1
        For X = 0 To 31
            b(X + l * 32) = gcZXPrinterBits(X + d * 32)
        Next X
        d = d - 1
    Next l
     
    Put #h, , "BM" ' // bitmap ident
    Put #h, , lSize ' // size
    Put #h, , CLng(0)
    Put #h, , CLng(62)
    Put #h, , bmiZXPrinter.bmiHeader
    Put #h, , bmiZXPrinter.bmiColors(0).rgbBlue
    Put #h, , bmiZXPrinter.bmiColors(0).rgbGreen
    Put #h, , bmiZXPrinter.bmiColors(0).rgbRed
    Put #h, , bmiZXPrinter.bmiColors(0).rgbReserved
    Put #h, , bmiZXPrinter.bmiColors(1).rgbBlue
    Put #h, , bmiZXPrinter.bmiColors(1).rgbGreen
    Put #h, , bmiZXPrinter.bmiColors(1).rgbRed
    Put #h, , bmiZXPrinter.bmiColors(1).rgbReserved
    Put #h, , b
    Close #h
End Sub

Private Sub cmdClear_Click()
    lZXPrinterY = 0
    vs.Min = 0
    vs.Max = 0
    picView.Cls
    InitZXPrinterBitmap
End Sub

Private Sub cmdFF_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    tmrFormFeed.Interval = 50
    tmrFormFeed.Enabled = True
End Sub


Private Sub cmdFF_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    tmrFormFeed.Enabled = False
End Sub


Private Sub cmdSave_Click()
    err.Clear
    
    On Error Resume Next
    
    dlgSave.DefaultExt = "bmp"
    dlgSave.FileName = ""
    dlgSave.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    dlgSave.CancelError = True
    dlgSave.ShowSave
    If err.Number = cdlCancel Then
        Exit Sub
    End If
    
    If dlgSave.FileName <> "" Then
        SaveMonoBitmap dlgSave.FileName

        cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim X As Long, y As Long, h As Long

    InitZXPrinterBitmap
    iWidth = Me.Width
       
    X = val(GetSetting("Grok", "vbSpec", "ZXPrnWndX", "-1"))
    y = val(GetSetting("Grok", "vbSpec", "ZXPrnWndY", "-1"))
    h = val(GetSetting("Grok", "vbSpec", "ZXPrnWndHeight", "-1"))
    
    If X >= 0 And X <= (Screen.Width - Screen.TwipsPerPixelX * 16) Then
        Me.Left = X
    End If
    If y >= 0 And y <= (Screen.Height - Screen.TwipsPerPixelY * 16) Then
        Me.Top = y
    End If
    If h >= 0 And h <= (Screen.Height - Me.Top - Screen.TwipsPerPixelY * 16) Then
        Me.Height = h
    End If
End Sub

Private Sub Form_Resize()
    Dim l As Long
    
    Me.Width = iWidth
    
    l = Me.ScaleHeight - picView.Top - picLogo.Height
    
    If l > 8 Then
        picView.Height = l
    Else
        picView.Height = 8
    End If
    
    vs.Height = picView.Height
    picView.Cls
    vs_Change
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Grok", "vbSpec", "ZXPrnWndX", CStr(Me.Left)
    SaveSetting "Grok", "vbSpec", "ZXPrnWndY", CStr(Me.Top)
    SaveSetting "Grok", "vbSpec", "ZXPrnWndHeight", CStr(Me.Height)
    
    frmMainWnd.mnuOptions(5).Checked = False
End Sub

Private Sub optAlphacom_Click()
    imgAlpha.Visible = True
    imgSinclair.Visible = False
    
    picView.BackColor = RGB(255, 255, 255)
    bmiZXPrinter.bmiColors(0).rgbRed = 255
    bmiZXPrinter.bmiColors(0).rgbGreen = 255
    bmiZXPrinter.bmiColors(0).rgbBlue = 255
    bmiZXPrinter.bmiColors(0).rgbReserved = 0
    bmiZXPrinter.bmiColors(1).rgbRed = 64
    bmiZXPrinter.bmiColors(1).rgbGreen = 64
    bmiZXPrinter.bmiColors(1).rgbBlue = 192
    bmiZXPrinter.bmiColors(1).rgbReserved = 0
    vs_Change
End Sub

Private Sub optZX_Click()
    imgSinclair.Visible = True
    imgAlpha.Visible = False
    
    picView.BackColor = RGB(192, 192, 192)
    bmiZXPrinter.bmiColors(0).rgbRed = 192
    bmiZXPrinter.bmiColors(0).rgbGreen = 192
    bmiZXPrinter.bmiColors(0).rgbBlue = 192
    bmiZXPrinter.bmiColors(0).rgbReserved = 0
    bmiZXPrinter.bmiColors(1).rgbRed = 0
    bmiZXPrinter.bmiColors(1).rgbGreen = 0
    bmiZXPrinter.bmiColors(1).rgbBlue = 0
    bmiZXPrinter.bmiColors(1).rgbReserved = 0
    vs_Change
End Sub

Private Sub picView_Resize()
    If lZXPrinterY > frmZXPrinter.picView.Height Then
        frmZXPrinter.vs.Min = frmZXPrinter.picView.Height \ 8
        frmZXPrinter.vs.Max = lZXPrinterY \ 8
        frmZXPrinter.vs.Value = lZXPrinterY \ 8
    Else
        frmZXPrinter.vs.Min = 0
        frmZXPrinter.vs.Max = 0
    End If
End Sub


Private Sub tmrFormFeed_Timer()
    lZXPrinterEncoder = 0
    lZXPrinterX = 0
    lZXPrinterY = lZXPrinterY + 1
    If lZXPrinterY >= glZXPrinterBMPHeight Then
        glZXPrinterBMPHeight = lZXPrinterY + 32
        ReDim Preserve gcZXPrinterBits(glZXPrinterBMPHeight * 32)
        bmiZXPrinter.bmiHeader.biHeight = glZXPrinterBMPHeight
    End If
    
    If frmZXPrinter.picView.Height > lZXPrinterY Then
        StretchDIBitsMono picView.hdc, 0, picView.Height, 256, -lZXPrinterY - 1, 0, 0, 256, lZXPrinterY + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
    Else
        StretchDIBitsMono picView.hdc, 0, picView.Height, 256, -picView.Height - 1, 0, lZXPrinterY - picView.Height, 256, picView.Height + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
    End If
    picView.REFRESH
        
    ' // Set up the scroll bar properties for the visible display
    ' // to allow scrolling back over the material printed so far
    If lZXPrinterY > frmZXPrinter.picView.Height Then
        frmZXPrinter.vs.Min = picView.Height \ 8
        frmZXPrinter.vs.Max = lZXPrinterY \ 8
        frmZXPrinter.vs.Value = lZXPrinterY \ 8
    Else
        frmZXPrinter.vs.Min = 0
        frmZXPrinter.vs.Max = 0
    End If
End Sub

Private Sub vs_Change()
    If frmZXPrinter.picView.Height > lZXPrinterY Then
        StretchDIBitsMono frmZXPrinter.picView.hdc, 0, frmZXPrinter.picView.Height, 256, -lZXPrinterY - 1, 0, 0, 256, lZXPrinterY + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
    Else
        StretchDIBitsMono frmZXPrinter.picView.hdc, 0, frmZXPrinter.picView.Height, 256, -frmZXPrinter.picView.Height - 1, 0, vs.Value * 8 - frmZXPrinter.picView.Height, 256, frmZXPrinter.picView.Height + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
    End If
    
'    If frmZXPrinter.picView.Height > lZXPrinterY Then
'        BitBlt frmZXPrinter.picView.hdc, 0, frmZXPrinter.picView.Height - lZXPrinterY, 256, frmZXPrinter.picView.Height, frmZXPrinter.picPrn.hdc, 0, 0, SRCCOPY
'    Else
'        BitBlt frmZXPrinter.picView.hdc, 0, 0, 256, frmZXPrinter.picView.Height, frmZXPrinter.picPrn.hdc, 0, vs.Value * 8, SRCCOPY
'    End If

    frmZXPrinter.picView.REFRESH
End Sub

Private Sub vs_Scroll()
    vs_Change
End Sub


