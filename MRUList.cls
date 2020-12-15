VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MRUList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' /*******************************************************************************
'   MRUList.cls within vbSpec.vbp
'
'   Most Recently Used file list handling
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

Private Const MaxMRUFiles = 5
Private MRUCount As Long
Private MRU(MaxMRUFiles) As String

Public Function AddMRUFile(sFile As String) As Long
    Dim l As Long, s As String
    
    ' // Empty filename, ignore
    If sFile = "" Then Exit Function
    
    ' // If file is already in MRU, bring it to the top of the list
    For l = 1 To MaxMRUFiles
        If LCase$(sFile) = LCase$(MRU(l)) Then
            s = MRU(1)
            MRU(1) = sFile
            MRU(l) = s
            Exit Function
        End If
    Next l
    
    ' // Move all the MRU files down one and insert the new file
    ' // into position 1
    For l = MaxMRUFiles To 2 Step -1
        MRU(l) = MRU(l - 1)
    Next l
    MRU(1) = sFile
    
    ' // Set the MRUCount appropriately
    MRUCount = 0
    For l = 1 To MaxMRUFiles
        If MRU(l) <> "" Then MRUCount = MRUCount + 1
    Next l
    
    AddMRUFile = MRUCount
End Function


Public Function GetMRUCount() As Long
    GetMRUCount = MRUCount
End Function

Public Function GetMRUFile(l As Long) As String
    If l > 0 And l <= MRUCount Then
        GetMRUFile = MRU(l)
    End If
End Function

Public Function GetMRUMax() As Long
    GetMRUMax = MaxMRUFiles
End Function

Private Sub Class_Initialize()
    Dim l As Long, s As String
    
    MRUCount = 0
    ' // Read files in from registry
    For l = 1 To MaxMRUFiles
        s = GetSetting("Grok", "vbSpec", "MRU" & CStr(l), "")
        If s <> "" Then
            MRUCount = MRUCount + 1
            MRU(MRUCount) = s
        End If
    Next l
End Sub


Private Sub Class_Terminate()
    Dim l As Long, c As Long
    
    c = 0
    ' // Read files in from registry
    For l = 1 To MaxMRUFiles
        If MRU(l) <> "" Then
            c = c + 1
            SaveSetting "Grok", "vbSpec", "MRU" & CStr(c), MRU(l)
        End If
    Next l
End Sub


