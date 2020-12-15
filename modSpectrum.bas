Attribute VB_Name = "modSpectrum"
' /*******************************************************************************
'   modSpectrum.bas within vbSpec.vbp
'
'   Routines for emulating the spectrum hardware; displaying the
'   video memory (0x4000 - 0x5AFF), reading the keyboard (port
'   0xFE), and displaying the border colour (out (xxFE),x)
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

'MM JD
Option Base 0

'MM 03.02.2003 - BEGIN
'Remarks:
'1. this code supports analog joysticks only
'2. the user must have a joytick driver under Windows
'3. the Y-axes represents the up and down
'4. the X-axes represents the right and left
'5. only one PC-joystick button is as fire supported
'6. the user can theoretically assign two PC-joysticks to one ZX-joystick, but I
'   don't think that't a good idea...
'JoyStick API
Public Const MAXPNAMELEN = 32
'MM 03.02.2003 - BEGIN
Public Const MAX_JOYSTICKOEMVXDNAME = 260
'MM 03.02.2003 - END
' The JOYINFOEX user-defined type contains extended information about the joystick position,
' point-of-view position, and button state.
Type JOYINFOEX
   dwSize As Long                      ' size of structure
   dwFlags As Long                     ' flags to indicate what to return
   dwXpos As Long                      ' x position
   dwYpos As Long                      ' y position
   dwZpos As Long                      ' z position
   dwRpos As Long                      ' rudder/4th axis position
   dwUpos As Long                      ' 5th axis position
   dwVpos As Long                      ' 6th axis position
   dwButtons As Long                   ' button states
   dwButtonNumber As Long              ' current button number pressed
   dwPOV As Long                       ' point of view state
   dwReserved1 As Long                 ' reserved for communication between winmm driver
   dwReserved2 As Long                 ' reserved for future expansion
End Type
' The JOYCAPS user-defined type contains information about the joystick capabilities
Type JOYCAPS
   wMid As Integer                     ' Manufacturer identifier of the device driver for the MIDI output device
                                       ' For a list of identifiers, see the Manufacturer Indentifier topic in the
                                       ' Multimedia Reference of the Platform SDK.
   wPid As Integer                     ' Product Identifier Product of the MIDI output device. For a list of
                                       ' product identifiers, see the Product Identifiers topic in the Multimedia
                                       ' Reference of the Platform SDK.
   szPname As String * MAXPNAMELEN     ' Null-terminated string containing the joystick product name
   wXmin As Long                       ' Minimum X-coordinate.
   wXmax As Long                       ' Maximum X-coordinate.
   wYmin As Long                       ' Minimum Y-coordinate
   wYmax As Long                       ' Maximum Y-coordinate
   wZmin As Long                       ' Minimum Z-coordinate
   wZmax As Long                       ' Maximum Z-coordinate
   wNumButtons As Long                 ' Number of joystick buttons
   wPeriodMin As Long                  ' Smallest polling frequency supported when captured by the joySetCapture function.
   wPeriodMax As Long                  ' Largest polling frequency supported when captured by the joySetCapture function.
   wRmin As Long                       ' Minimum rudder value. The rudder is a fourth axis of movement.
   wRmax As Long                       ' Maximum rudder value. The rudder is a fourth axis of movement.
   wUmin As Long                       ' Minimum u-coordinate (fifth axis) values.
   wUmax As Long                       ' Maximum u-coordinate (fifth axis) values.
   wVmin As Long                       ' Minimum v-coordinate (sixth axis) values.
   wVmax As Long                       ' Maximum v-coordinate (sixth axis) values.
   wCaps As Long                       ' Joystick capabilities as defined by the following flags
                                       '     JOYCAPS_HASZ-     Joystick has z-coordinate information.
                                       '     JOYCAPS_HASR-     Joystick has rudder (fourth axis) information.
                                       '     JOYCAPS_HASU-     Joystick has u-coordinate (fifth axis) information.
                                       '     JOYCAPS_HASV-     Joystick has v-coordinate (sixth axis) information.
                                       '     JOYCAPS_HASPOV-   Joystick has point-of-view information.
                                       '     JOYCAPS_POV4DIR-  Joystick point-of-view supports discrete values (centered, forward, backward, left, and right).
                                       '     JOYCAPS_POVCTS Joystick point-of-view supports continuous degree bearings.
   wMaxAxes As Long                    ' Maximum number of axes supported by the joystick.
   wNumAxes As Long                    ' Number of axes currently in use by the joystick.
   wMaxButtons As Long                 ' Maximum number of buttons supported by the joystick.
   szRegKey As String * MAXPNAMELEN    ' String containing the registry key for the joystick.
   'MM 03.02.2003 - BEGIN
   szOEMVxD As String * MAX_JOYSTICKOEMVXDNAME    ' Null-terminated string identifying the joystick driver OEM.
   'MM 03.02.2003 - END
End Type
Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
' This function queries a joystick for its position and button status. The function
' requires the following parameters;
'     uJoyID-  integer identifying the joystick to be queried. Use the constants
'              JOYSTICKID1 or JOYSTICKID2 for this value.
'     pji-     user-defined type variable that stores extended position information
'              and button status of the joystick. The information returned from
'              this function depends on the flags you specify in dwFlags member of
'              the user-defined type variable.
'
' The function returns the constant JOYERR_NOERROR if successful or one of the
' following error values:
'     MMSYSERR_NODRIVER-      The joystick driver is not present.
'     MMSYSERR_INVALPARAM-    An invalid parameter was passed.
'     MMSYSERR_BADDEVICEID-   The specified joystick identifier is invalid.
'     JOYERR_UNPLUGGED-       The specified joystick is not connected to the system.
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
' This function queries a joystick to determine its capabilities. The function requires
' the following parameters:
'     uJoyID-  integer identifying the joystick to be queried. Use the contstants
'              JOYSTICKID1 or JOYSTICKID2 for this value.
'     pjc-     user-defined type variable that stores the capabilities of the joystick.
'     cbjc-    Size, in bytes, of the pjc variable. Use the Len function for this value.
' The function returns the constant JOYERR_NOERROR if a joystick is present or one of
' the following error values:
'     MMSYSERR_NODRIVER-   The joystick driver is not present.
'     MMSYSERR_INVALPARAM- An invalid parameter was passed.
Public Const JOYERR_OK = 0
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10
Public Const JOY_RETURNV = &H20
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const JOYCAPS_HASZ = &H1&
Public Const JOYCAPS_HASR = &H2&
Public Const JOYCAPS_HASU = &H4&
Public Const JOYCAPS_HASV = &H8&
Public Const JOYCAPS_HASPOV = &H10&
Public Const JOYCAPS_POV4DIR = &H20&
Public Const JOYCAPS_POVCTS = &H40&
Public Const JOYERR_BASE = 160
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)
'Supported ZX-Joysticks
Public Enum ZXJOYSTICKS
    zxjInvalid = -1
    zxjKempston = 0
    zxjCursor = 1
    zxjSinclair1 = 2
    zxjSinclair2 = 3
    zxjFullerBox = 4
    'MM JD
    zxjUserDefined = 5
End Enum
'Supported PC-joystick buttons
Public Enum PCJOYBUTTONS
    pcjbInvalid = -1
    pcjbButton1 = 0
    pcjbButton2 = 1
    pcjbButton3 = 2
    pcjbButton4 = 3
    pcjbButton5 = 4
    pcjbButton6 = 5
    pcjbButton7 = 6
    pcjbButton8 = 7
End Enum
'Settings for PC-joystick nr. 1
'The ZX-joystick type to be emulated by PC-joystick nr. 1
Public lPCJoystick1Is As ZXJOYSTICKS
'The Fire button of the ZX-joystick played by the button nr. ... of the
'PC-joystick  nr. 1
Public lPCJoystick1Fire As PCJOYBUTTONS
'Number of the buttons by PC-joystick nr. 1
'Remarks: up to 8 PC-joystick buttons supported
Public lPCJoystick1Buttons As Long
'Settings for PC-joystick nr. 2
'The ZX-joystick type to be emulated by PC-joystick nr. 2
Public lPCJoystick2Is As ZXJOYSTICKS
'The Fire button of the ZX-joystick played by the button nr. ... of the
'PC-joystick  nr. 2
Public lPCJoystick2Fire As PCJOYBUTTONS
'Number of the buttons by PC-joystick nr. 2
'Remarks: up to 8 PC-joystick buttons supported
Public lPCJoystick2Buttons As Long
'Supported joystick directions
Public Const zxjdUp As Long = 1
Public Const zxjdDown As Long = 2
Public Const zxjdLeft As Long = 4
Public Const zxjdRight As Long = 8
Public Const zxjdFire As Long = 16
'MM JD
Public Const zxjdButtonBase As Long = 32
'Kempston joystick directions
Public Const KEMPSTON_UP As Long = 8
Public Const KEMPSTON_PUP As Long = 31
Public Const KEMPSTON_DOWN As Long = 4
Public Const KEMPSTON_PDOWN As Long = 31
Public Const KEMPSTON_LEFT As Long = 2
Public Const KEMPSTON_PLEFT As Long = 31
Public Const KEMPSTON_RIGHT As Long = 1
Public Const KEMPSTON_PRIGHT As Long = 31
Public Const KEMPSTON_FIRE As Long = 16
Public Const KEMPSTON_PFIRE As Long = 31
'Cursor joystick directions
Public Const CURSOR_UP As Long = 183
Public Const CURSOR_PUP As Long = 61438
Public Const CURSOR_DOWN As Long = 175
Public Const CURSOR_PDOWN As Long = 61438
Public Const CURSOR_LEFT As Long = 175
Public Const CURSOR_PLEFT As Long = 63486
Public Const CURSOR_RIGHT As Long = 187
Public Const CURSOR_PRIGHT As Long = 61438
Public Const CURSOR_FIRE As Long = 190
Public Const CURSOR_PFIRE As Long = 61438
'Sinclair1 joystick directions
Public Const SINCLAIR1_UP As Long = 189
Public Const SINCLAIR1_PUP As Long = 61438
Public Const SINCLAIR1_DOWN As Long = 187
Public Const SINCLAIR1_PDOWN As Long = 61438
Public Const SINCLAIR1_LEFT As Long = 175
Public Const SINCLAIR1_PLEFT As Long = 61438
Public Const SINCLAIR1_RIGHT As Long = 183
Public Const SINCLAIR1_PRIGHT As Long = 61438
Public Const SINCLAIR1_FIRE As Long = 190
Public Const SINCLAIR1_PFIRE As Long = 61438
'Sinclair2 joystick directions
Public Const SINCLAIR2_UP As Long = 183
Public Const SINCLAIR2_PUP As Long = 63486
Public Const SINCLAIR2_DOWN As Long = 187
Public Const SINCLAIR2_PDOWN As Long = 63486
Public Const SINCLAIR2_LEFT As Long = 190
Public Const SINCLAIR2_PLEFT As Long = 63486
Public Const SINCLAIR2_RIGHT As Long = 189
Public Const SINCLAIR2_PRIGHT As Long = 63486
Public Const SINCLAIR2_FIRE As Long = 175
Public Const SINCLAIR2_PFIRE As Long = 63486
'Fuller Box directions
Public Const FULLER_UP As Long = 1
Public Const FULLER_PUP As Long = 127
Public Const FULLER_DOWN As Long = 2
Public Const FULLER_PDOWN As Long = 127
Public Const FULLER_LEFT As Long = 4
Public Const FULLER_PLEFT As Long = 127
Public Const FULLER_RIGHT As Long = 8
Public Const FULLER_PRIGHT As Long = 127
Public Const FULLER_FIRE As Long = 128
Public Const FULLER_PFIRE As Long = 127
'Last key
Private lKey1 As Long
Private lKey2 As Long
'MM 12.03.2003 - Joystick support code
'=================================================================================
'Joystick validity Flags
Public bJoystick1Valid As Boolean
Public bJoystick2Valid As Boolean
'=================================================================================
'MM 03.02.2003 - END

'MM JD
'Maximal number of buttons
Public Const JDT_MAXBUTTONS As Long = 8
'Joystick actions
Public Const JDT_UP As Long = 0
Public Const JDT_DOWN As Long = 1
Public Const JDT_LEFT As Long = 2
Public Const JDT_RIGHT As Long = 3
Public Const JDT_BUTTON1 As Long = 4
Public Const JDT_BUTTON2 As Long = 5
Public Const JDT_BUTTON3 As Long = 6
Public Const JDT_BUTTON4 As Long = 7
Public Const JDT_BUTTON5 As Long = 8
Public Const JDT_BUTTON6 As Long = 9
Public Const JDT_BUTTON7 As Long = 10
Public Const JDT_BUTTON8 As Long = 11
Public Const JDT_BUTTON_BASE As Long = 3
'Joystick numbers
Public Const JDT_JOYSTICK1 As Long = 0
Public Const JDT_JOYSTICK2 As Long = 1
'Joystick definition type
Public Type JOYSTICK_DEFINITION
    sKey As String
    lPort As Long
    lValue As Long
    lCumulate As Long
End Type
'Joystick definition table
'That is a 3D table, wich describes the joystick actions, even for standart joysticks
'- first dimension 1..0, the number of the joystick
'- the second dimension 0..5, the joystick type (Enum ZXJOYSTICKS)
'- the third dimension 0..11, the joystick actions (JDT_...)
Public aJoystikDefinitionTable(1, 5, 11) As JOYSTICK_DEFINITION

'MM 16.04.2003
'True if Speccy in Full-Screen-Modus
Public bFullScreen As Boolean

Public glNewBorder As Long
Public glLastBorder As Long

Public lZXPrinterX As Long
Public lZXPrinterY As Long
Public lZXPrinterEncoder As Long
Private lZXPrinterMotorOn As Long
Private lZXPrinterStylusOn As Long
Private lRevBitValues(7) As Long

Public glAmigaMouseY(0 To 3) As Long
Public glAmigaMouseX(0 To 3) As Long
Private lOldAmigaX As Long
Private lOldAmigaY As Long
Private lAmigaXPtr As Long
Private lAmigaYPtr As Long
Function doKey(down As Boolean, ascii As Integer, mods As Integer) As Boolean
    Dim CAPS As Boolean, SYMB As Boolean
    
    CAPS = mods And 1&
    SYMB = mods And 2&
    
    ' // Change control versions of keys to lower case
    If (ascii >= 1&) And (ascii <= &H27&) And SYMB Then
        ascii = ascii + Asc("a") - 1&
    End If
    
    If CAPS Then keyCAPS_V = (keyCAPS_V And (Not 1&)) Else keyCAPS_V = (keyCAPS_V Or 1&)
    If SYMB Then keyB_SPC = (keyB_SPC And (Not 2&)) Else keyB_SPC = (keyB_SPC Or 2&)

    Select Case ascii
    Case 8& ' Backspace
        If down Then
            key6_0 = (key6_0 And (Not 1&))
            keyCAPS_V = (keyCAPS_V And (Not 1&))
        Else
            key6_0 = (key6_0 Or 1&)
            If Not CAPS Then
                keyCAPS_V = (keyCAPS_V Or 1&)
            End If
        End If
    Case 65& ' A
        If down Then keyA_G = (keyA_G And (Not 1&)) Else keyA_G = (keyA_G Or 1&)
    Case 66& ' B
        If down Then keyB_SPC = (keyB_SPC And (Not 16&)) Else keyB_SPC = (keyB_SPC Or 16&)
    Case 67& ' C
        If down Then keyCAPS_V = (keyCAPS_V And (Not 8&)) Else keyCAPS_V = (keyCAPS_V Or 8&)
    Case 68& ' D
        If down Then keyA_G = (keyA_G And (Not 4&)) Else keyA_G = (keyA_G Or 4&)
    Case 69& ' E
        If down Then keyQ_T = (keyQ_T And (Not 4&)) Else keyQ_T = (keyQ_T Or 4&)
    Case 70& ' F
        If down Then keyA_G = (keyA_G And (Not 8&)) Else keyA_G = (keyA_G Or 8&)
    Case 71& ' G
        If down Then keyA_G = (keyA_G And (Not 16&)) Else keyA_G = (keyA_G Or 16&)
    Case 72& ' H
        If down Then keyH_ENT = (keyH_ENT And (Not 16&)) Else keyH_ENT = (keyH_ENT Or 16&)
    Case 73& ' I
        If down Then keyY_P = (keyY_P And (Not 4&)) Else keyY_P = (keyH_ENT Or 4&)
    Case 74& ' J
        If down Then keyH_ENT = (keyH_ENT And (Not 8&)) Else keyH_ENT = (keyH_ENT Or 8&)
    Case 75& ' K
        If down Then keyH_ENT = (keyH_ENT And (Not 4&)) Else keyH_ENT = (keyH_ENT Or 4&)
    Case 76& ' L
        If down Then keyH_ENT = (keyH_ENT And (Not 2&)) Else keyH_ENT = (keyH_ENT Or 2&)
    Case 77& ' M
        If down Then keyB_SPC = (keyB_SPC And (Not 4&)) Else keyB_SPC = (keyB_SPC Or 4&)
    Case 78& ' N
        If down Then keyB_SPC = (keyB_SPC And (Not 8&)) Else keyB_SPC = (keyB_SPC Or 8&)
    Case 79& ' O
        If down Then keyY_P = (keyY_P And (Not 2&)) Else keyY_P = (keyY_P Or 2&)
    Case 80& ' P
        If down Then keyY_P = (keyY_P And (Not 1&)) Else keyY_P = (keyY_P Or 1&)
    Case 81& ' Q
        If down Then keyQ_T = (keyQ_T And (Not 1&)) Else keyQ_T = (keyQ_T Or 1&)
    Case 82& ' R
        If down Then keyQ_T = (keyQ_T And (Not 8&)) Else keyQ_T = (keyQ_T Or 8&)
    Case 83& ' S
        If down Then keyA_G = (keyA_G And (Not 2&)) Else keyA_G = (keyA_G Or 2&)
    Case 84& ' T
        If down Then keyQ_T = (keyQ_T And (Not 16&)) Else keyQ_T = (keyQ_T Or 16&)
    Case 85& ' U
        If down Then keyY_P = (keyY_P And (Not 8&)) Else keyY_P = (keyY_P Or 8&)
    Case 86& ' V
        If down Then keyCAPS_V = (keyCAPS_V And (Not 16&)) Else keyCAPS_V = (keyCAPS_V Or 16&)
    Case 87& ' W
        If down Then keyQ_T = (keyQ_T And (Not 2&)) Else keyQ_T = (keyQ_T Or 2&)
    Case 88& ' X
        If down Then keyCAPS_V = (keyCAPS_V And (Not 4&)) Else keyCAPS_V = (keyCAPS_V Or 4&)
    Case 89& ' Y
        'MM 03.02.2003 - BEGIN
        If Not frmMainWnd.mnuOptions(7).Checked Then
            If down Then keyY_P = (keyY_P And (Not 16&)) Else keyY_P = (keyY_P Or 16&)
        Else
            If down Then keyCAPS_V = (keyCAPS_V And (Not 2&)) Else keyCAPS_V = (keyCAPS_V Or 2&)
        End If
        'MM 03.02.2003 - END
    Case 90& ' Z
        'MM 03.02.2003 - BEGIN
        If Not frmMainWnd.mnuOptions(7).Checked Then
            If down Then keyCAPS_V = (keyCAPS_V And (Not 2&)) Else keyCAPS_V = (keyCAPS_V Or 2&)
        Else
            If down Then keyY_P = (keyY_P And (Not 16&)) Else keyY_P = (keyY_P Or 16&)
        End If
        'MM 03.02.2003 - END
    Case 48& ' 0
        If down Then key6_0 = (key6_0 And (Not 1&)) Else key6_0 = (key6_0 Or 1&)
    Case 49& ' 1
        If down Then key1_5 = (key1_5 And (Not 1&)) Else key1_5 = (key1_5 Or 1&)
    Case 50& ' 2
        If down Then key1_5 = (key1_5 And (Not 2&)) Else key1_5 = (key1_5 Or 2&)
    Case 51& ' 3
        If down Then key1_5 = (key1_5 And (Not 4&)) Else key1_5 = (key1_5 Or 4&)
    Case 52& ' 4
        If down Then key1_5 = (key1_5 And (Not 8&)) Else key1_5 = (key1_5 Or 8&)
    Case 53& ' 5
        If down Then key1_5 = (key1_5 And (Not 16&)) Else key1_5 = (key1_5 Or 16&)
    Case 54& ' 6
        If down Then key6_0 = (key6_0 And (Not 16&)) Else key6_0 = (key6_0 Or 16&)
    Case 55& ' 7
        If down Then key6_0 = (key6_0 And (Not 8&)) Else key6_0 = (key6_0 Or 8&)
    Case 56& ' 8
        If down Then key6_0 = (key6_0 And (Not 4&)) Else key6_0 = (key6_0 Or 4&)
    Case 57& ' 9
        If down Then key6_0 = (key6_0 And (Not 2&)) Else key6_0 = (key6_0 Or 2&)
    Case 96& ' Keypad 0
        If down Then key6_0 = (key6_0 And (Not 1&)) Else key6_0 = (key6_0 Or 1&)
    Case 97& ' Keypad 1
        If down Then key1_5 = (key1_5 And (Not 1&)) Else key1_5 = (key1_5 Or 1&)
    Case 98& ' Keypad 2
        If down Then key1_5 = (key1_5 And (Not 2&)) Else key1_5 = (key1_5 Or 2&)
    Case 99& ' Keypad 3
        If down Then key1_5 = (key1_5 And (Not 4&)) Else key1_5 = (key1_5 Or 4&)
    Case 100& ' Keypad 4
        If down Then key1_5 = (key1_5 And (Not 8&)) Else key1_5 = (key1_5 Or 8&)
    Case 101& ' Keypad 5
        If down Then key1_5 = (key1_5 And (Not 16&)) Else key1_5 = (key1_5 Or 16&)
    Case 102& ' Keypad 6
        If down Then key6_0 = (key6_0 And (Not 16&)) Else key6_0 = (key6_0 Or 16&)
    Case 103& ' Keypad 7
        If down Then key6_0 = (key6_0 And (Not 8&)) Else key6_0 = (key6_0 Or 8&)
    Case 104& ' Keypad 8
        If down Then key6_0 = (key6_0 And (Not 4&)) Else key6_0 = (key6_0 Or 4&)
    Case 105& ' Keypad 9
        If down Then key6_0 = (key6_0 And (Not 2&)) Else key6_0 = (key6_0 Or 2&)
    Case 106& ' Keypad *
        If down Then
            keyB_SPC = (keyB_SPC And Not (18&))
        Else
            If SYMB Then
                keyB_SPC = (keyB_SPC Or 16&)
            Else
                keyB_SPC = (keyB_SPC Or 18&)
            End If
        End If
    Case 107& ' Keypad +
        If down Then
            keyH_ENT = (keyH_ENT And (Not 4&))
            keyB_SPC = (keyB_SPC And (Not 2&))
        Else
            keyH_ENT = (keyH_ENT Or 4&)
            If Not SYMB Then
                keyB_SPC = (keyB_SPC Or 2&)
            End If
        End If
    Case 109& ' Keypad -
        If down Then
            keyH_ENT = (keyH_ENT And (Not 8&))
            keyB_SPC = (keyB_SPC And (Not 2&))
        Else
            keyH_ENT = (keyH_ENT Or 8&)
            If Not SYMB Then
                keyB_SPC = (keyB_SPC Or 2&)
            End If
        End If
    Case 110& ' Keypad .
        If down Then
            keyB_SPC = (keyB_SPC And (Not 6&))
        Else
            If SYMB Then
                keyB_SPC = (keyB_SPC Or 4&)
            Else
                keyB_SPC = (keyB_SPC Or 6&)
            End If
        End If
    Case 111& ' Keypad /
        If down Then
            keyCAPS_V = (keyCAPS_V And (Not 16&))
            keyB_SPC = (keyB_SPC And (Not 2&))
        Else
            keyCAPS_V = (keyCAPS_V Or 16&)
            If Not SYMB Then
                keyB_SPC = (keyB_SPC Or 2&)
            End If
        End If
    Case 37 ' Left
        If down Then
            key1_5 = (key1_5 And (Not 16&))
            keyCAPS_V = (keyCAPS_V And (Not 1&))
        Else
            key1_5 = (key1_5 Or 16&)
            If Not SYMB Then
                keyB_SPC = (keyB_SPC Or 2&)
            End If
        End If
    Case 38 ' Up
        If down Then
            key6_0 = (key6_0 And (Not 8&))
            keyCAPS_V = (keyCAPS_V And (Not 1&))
        Else
            key6_0 = (key6_0 Or 8&)
            If Not CAPS Then
                keyCAPS_V = (keyCAPS_V Or 1&)
            End If
        End If
    Case 39 ' Right
        If down Then
            key6_0 = (key6_0 And (Not 4&))
            keyCAPS_V = (keyCAPS_V And (Not 1&))
        Else
            key6_0 = (key6_0 Or 4&)
            If Not CAPS Then
                keyCAPS_V = (keyCAPS_V Or 1&)
            End If
        End If
    Case 40 ' Down
        If down Then
            key6_0 = (key6_0 And (Not 16&))
            keyCAPS_V = (keyCAPS_V And (Not 1&))
        Else
            key6_0 = (key6_0 Or 16&)
            If Not CAPS Then
                keyCAPS_V = (keyCAPS_V Or 1&)
            End If
        End If
    Case 13 ' RETURN
        If down Then keyH_ENT = (keyH_ENT And (Not 1&)) Else keyH_ENT = (keyH_ENT Or 1&)
    Case 32 ' SPACE BAR
        If down Then keyB_SPC = (keyB_SPC And (Not 1&)) Else keyB_SPC = (keyB_SPC Or 1&)
    Case 187 ' =/+ key
        If down Then
            If CAPS Then
                keyH_ENT = (keyH_ENT And (Not 4&))
            Else
                keyH_ENT = (keyH_ENT And (Not 2&))
            End If
            keyB_SPC = (keyB_SPC And (Not 2&))
            keyCAPS_V = (keyCAPS_V Or 1&)
        Else
            keyH_ENT = (keyH_ENT Or 4&)
            keyH_ENT = (keyH_ENT Or 2&)
            keyB_SPC = (keyB_SPC Or 2&)
        End If
    Case 189 ' -/_ key
        If down Then
            If CAPS Then
                key6_0 = (key6_0 And (Not 1&))
            Else
                keyH_ENT = (keyH_ENT And (Not 8&))
            End If
            keyB_SPC = (keyB_SPC And (Not 2&))
            keyCAPS_V = (keyCAPS_V Or 1&)
        Else
            key6_0 = (key6_0 Or 1&)     ' // Release the Spectrum's '0' key
            keyH_ENT = (keyH_ENT Or 8&) ' // Release the Spectrum's 'J' key
            keyB_SPC = (keyB_SPC Or 2&) ' // Release the Symbol Shift key
        End If
    Case 186 ' ;/: keys
        If down Then
            If CAPS Then
                keyCAPS_V = (keyCAPS_V And (Not 2&))
            Else
                keyY_P = (keyY_P And (Not 2&))
            End If
            keyB_SPC = (keyB_SPC And (Not 2&))
            keyCAPS_V = (keyCAPS_V Or 1&)
        Else
            keyCAPS_V = (keyCAPS_V Or 2&)
            keyY_P = (keyY_P Or 2&)
            keyB_SPC = (keyB_SPC Or 2&)
        End If
    Case Else
        doKey = False
    End Select

    doKey = True
End Function

Public Sub Hook_LDBYTES()
    Dim l As Long
    
    If LoadTAP(glMemAddrDiv256(regAF_), regIX, regDE) Then
        regAF_ = regAF_ Or 64   ' // Congraturation Load Sucsess!
    Else
        regAF_ = regAF_ And 190 ' // Load failed
    End If
    
    l = getAF()
    setAF regAF_
    regAF_ = l
    
    regPC = 1506
End Sub




Public Sub Hook_SABYTES()
    SaveTAPFileDlg
End Sub


Public Sub InitReverseBitValues()
    lRevBitValues(0) = 128
    lRevBitValues(1) = 64
    lRevBitValues(2) = 32
    lRevBitValues(3) = 16
    lRevBitValues(4) = 8
    lRevBitValues(5) = 4
    lRevBitValues(6) = 2
    lRevBitValues(7) = 1
End Sub

Sub plot(addr As Long)
    Dim lne As Long, i As Long, X As Long
    
    If addr < 22528& Then
        ' // Alter a pixel
        lne = (glMemAddrDiv256(addr) And &H7&) Or _
                  (glMemAddrDiv4(addr) And &H38&) Or _
                  (glMemAddrDiv32(addr) And &HC0&)
        ScrnLines(lne, 32&) = True
        ScrnLines(lne, addr And 31&) = True
    Else
        ' // Alter an attribute
        lne = glMemAddrDiv32(addr - 22528&)
        X = addr Mod 32&
        For i = lne * 8& To lne * 8& + 7&
            ScrnLines(i, 32&) = True
            ScrnLines(i, X) = True
        Next i
    End If
    If glUseScreen >= 1000& Then ScrnNeedRepaint = True
End Sub

Function inb(port As Long) As Long
    Dim p As POINTAPI
    Dim bPortDefined As Boolean
    Dim lCounter As Long
    
    'MM JD
    'Init inb with the joystick values
    inb = JoystickInitIN(port)
    
    If (port And &HFF&) = 254& Then
        Dim res As Long
        
        res = &HFF&
        
        If (port And &H8000&) = 0& Then
            res = res And keyB_SPC
        End If
        If (port And &H4000&) = 0& Then
            res = res And keyH_ENT
        End If
        If (port And &H2000&) = 0& Then
            res = res And keyY_P
        End If
        If (port And &H1000&) = 0& Then
            res = res And key6_0
        End If
        If (port And &H800&) = 0& Then
            res = res And key1_5
        End If
        If (port And &H400&) = 0& Then
            res = res And keyQ_T
        End If
        If (port And &H200&) = 0& Then
            res = res And keyA_G
        End If
        If (port And &H100&) = 0& Then
            res = res And keyCAPS_V
        End If
        If inb <> 0 Then
            inb = inb And (res And glKeyPortMask) Or glEarBit ' glEarBit holds tape state (0 or 64 only)
        Else
            inb = (res And glKeyPortMask) Or glEarBit ' glEarBit holds tape state (0 or 64 only)
        End If
        glTStates = glTStates + glContentionTable(-glTStates)
    ElseIf port = &HFFFD& Then
        If (glEmulatedModel And 3&) <> 0& Then
            inb = AYPSG.Regs(glSoundRegister)
        End If
    ElseIf (port And &HFF&) = &HFF& Then
        If glEmulatedModel = 5& Then
            ' // TC2048
            inb = glTC2048LastFFOut
        Else
            If (glTStates >= glTStatesAtTop) And (glTStates <= glTStatesAtBottom) Then
                inb = 0&
            Else
                inb = 255&
            End If
        End If
    ElseIf (port And &HFF&) = 31& Then
        If glMouseType = MOUSE_AMIGA Then
            GetCursorPos p
            If p.X > lOldAmigaX Then
                lAmigaXPtr = lAmigaXPtr + 1
                If lAmigaXPtr = 4 Then lAmigaXPtr = 0
            ElseIf p.X < lOldAmigaX Then
                lAmigaXPtr = lAmigaXPtr - 1
                If lAmigaXPtr = -1 Then lAmigaXPtr = 3
            End If
            If p.y < lOldAmigaY Then
                lAmigaYPtr = lAmigaYPtr + 1
                If lAmigaYPtr = 4 Then lAmigaYPtr = 0
            ElseIf p.y > lOldAmigaY Then
                lAmigaYPtr = lAmigaYPtr - 1
                If lAmigaYPtr = -1 Then lAmigaYPtr = 3
            End If
            lOldAmigaX = p.X
            lOldAmigaY = p.y
            inb = glAmigaMouseX(lAmigaXPtr) Or glAmigaMouseY(lAmigaYPtr)
            
            If gbMouseGlobal Then
                inb = inb Or (-((GetKeyState(VK_LBUTTON) And 256) = 256) * 16) Or (-((GetKeyState(VK_RBUTTON) And 256) = 256) * 32)
            Else
                inb = inb Or (glMouseBtn * 16)
            End If
        Else
            'MM JD
            'The emulator does not need this line anymore
            'inb = 0&
        End If
    ElseIf (port And 4) = 0 Then
        inb = ZXPrinterIn
    ' // Kempston Mouse Interface
    ElseIf glMouseType = MOUSE_KEMPSTON Then
        GetCursorPos p
        If (port = 64479) Then
            inb = p.X Mod 256
        ElseIf (port = 65503) Then
            inb = (4000 - p.y) Mod 256
        ElseIf (port = 64223) Then
            If gbMouseGlobal Then
                inb = 255 + ((GetKeyState(VK_RBUTTON) And 256) = 256)
                inb = inb + (((GetKeyState(VK_LBUTTON) And 256) = 256) * 2)
            Else
                inb = 255 - (glMouseBtn And 1) * 2
                inb = inb - (glMouseBtn And 2) \ 2
            End If
        End If
    Else
        ' // Unconnected port
        If (glTStates >= glTStatesAtTop) And (glTStates <= glTStatesAtBottom) And _
           (glEmulatedModel <> 5&) Then
            'If inb is not 0 here, this means, that that is a joystick port an this port
            'is connected. Othervise is inb alread 0, so the Speccy does not need this
            'line anymore
            'inb = 0& '// This should return a floating bus value, but zero suffices
            '         '// for commericial games that depend on the floating bus such
            '         '// as Cobra and Arkanoid
        Else
            'Only if the port ist'n the actual joystick port
            'IF the actual port is a valid joystick port and there is some move on the
            'joystick, the inb is not 0.
            If inb = 0 Then
                inb = 255&
            End If
        End If
    End If
End Function

Sub outb(port As Long, outbyte As Long)
    If (port And 1&) = 0& Then
        glLastFEOut = outbyte And &HFF&
        If glUseScreen <> 1006& Then
            glNewBorder = glNormalColor(outbyte And &H7&)
        End If

        If (outbyte And 16&) Then
            glBeeperVal = 159&
        Else
            glBeeperVal = 128&
        End If
        glTStates = glTStates + glContentionTable(-glTStates)
        Exit Sub
    ElseIf glMemPagingType <> 0& Then
        ' // 128/+2 memory page operation
        If (port And 32770) = 0& Then
            ' // RAM page
            glPageAt(3&) = (outbyte And 7&)
            ' // Screen page
            If (outbyte And 8&) Then
                If glUseScreen = 5& Then
                    glUseScreen = 7&
                    initscreen
                End If
            Else
                If glUseScreen = 7& Then
                    glUseScreen = 5&
                    initscreen
                End If
            End If
            ' // ROM
            If (outbyte And 16&) Then
                glPageAt(0&) = 9&
                glPageAt(4&) = 9&
            Else
                glPageAt(0&) = 8&
                glPageAt(4&) = 8&
            End If
            
            If (outbyte And 32&) Then glMemPagingType = 0&
            glLastOut7FFD = outbyte
        ElseIf (port And &HC002&) = &HC000& Then
            glSoundRegister = outbyte And &HF
            Exit Sub
        ElseIf (port And &HC002&) = &H8000& Then
            AYWriteReg glSoundRegister, outbyte
            Exit Sub
        'ElseIf port = &HBEFD& Then
        '    AYWriteReg glSoundRegister, outbyte
        ' ElseIf port = &H1FFD& Then
            ' // +2A/+3 special paging mode
        End If
    ElseIf (port And &H4&) = 0& Then
        ZXPrinterOut outbyte
    ElseIf glEmulatedModel = 5& Then
        ' // TC2048=May slow things down :(
        If (port And &HFF&) = &HFF& Then
            glTC2048LastFFOut = outbyte And &HFF&
            If (outbyte And 7&) = 0& Then
                ' // screen 0
                glUseScreen = 5&
                bmiBuffer.bmiHeader.biWidth = 256&
                
                glNewBorder = glNormalColor(glLastFEOut And 7&)
            ElseIf (outbyte And 7&) = 1& Then
                ' // screen 1
                glUseScreen = 1001&
                bmiBuffer.bmiHeader.biWidth = 256&
                
                glNewBorder = glNormalColor(glLastFEOut And 7&)
            ElseIf (outbyte And 7&) = 2& Then
                ' // hi-colour
                glUseScreen = 1002&
                bmiBuffer.bmiHeader.biWidth = 256&
                
                glNewBorder = glNormalColor(glLastFEOut And 7&)
            ElseIf (outbyte And 7&) = 6& Then
                ' // hi-res
                glUseScreen = 1006&
                bmiBuffer.bmiHeader.biWidth = 512&
            
                If (outbyte And 56&) = 0& Then
                    ' // Black on white
                    glTC2048HiResColour = 120&
                    glNewBorder = glBrightColor(7&)
                ElseIf (outbyte And 56&) = 8& Then
                    ' // Blue on yellow
                    glTC2048HiResColour = 113&
                    glNewBorder = glBrightColor(6&)
                ElseIf (outbyte And 56&) = 16& Then
                    ' // Red on cyan
                    glTC2048HiResColour = 106&
                    glNewBorder = glBrightColor(5&)
                ElseIf (outbyte And 56&) = 24& Then
                    ' // Magenta on green
                    glTC2048HiResColour = 99&
                    glNewBorder = glBrightColor(4&)
                ElseIf (outbyte And 56&) = 32& Then
                    ' // Green on magenta
                    glTC2048HiResColour = 92&
                    glNewBorder = glBrightColor(3&)
                ElseIf (outbyte And 56&) = 40& Then
                    ' // Cyan on red
                    glTC2048HiResColour = 85&
                    glNewBorder = glBrightColor(2&)
                ElseIf (outbyte And 56&) = 48& Then
                    ' // Yellow on blue
                    glTC2048HiResColour = 78&
                    glNewBorder = glBrightColor(1&)
                ElseIf (outbyte And 56) = 56 Then
                    ' // White on black
                    glTC2048HiResColour = 71
                    glNewBorder = glBrightColor(0)
                End If
            End If
            initscreen
        End If
    End If
End Sub

Sub plotTC2048HiResHiArea(addr As Long)
    Dim lne As Long
    
    ' // Alter a pixel in the higher screen (odd columns)
    lne = (glMemAddrDiv256(addr) And &H7&) Or _
              (glMemAddrDiv4(addr) And &H38&) Or _
              (glMemAddrDiv32(addr) And &HC0&)
    ScrnLines(lne, 64) = True
    ScrnLines(lne, ((addr And 31) * 2) + 1) = True
    
    ScrnNeedRepaint = True
End Sub

Sub plotTC2048HiResLowArea(addr As Long)
    Dim lne As Long
    
    ' // Alter a pixel in the lower screen (even columns)
    lne = (glMemAddrDiv256(addr) And &H7&) Or _
              (glMemAddrDiv4(addr) And &H38&) Or _
              (glMemAddrDiv32(addr) And &HC0&)
    ScrnLines(lne, 64) = True
    ScrnLines(lne, (addr And 31) * 2) = True
    
    ScrnNeedRepaint = True
End Sub

Public Sub refreshFlashChars()
    If glUseScreen > 8& Then
        TC2048refreshFlashChars
        Exit Sub
    End If
        
    Dim addr As Long, lne As Long, i As Long

    bFlashInverse = Not (bFlashInverse)
    
    For addr = 6144& To 6911&
        If gRAMPage(glUseScreen, addr) And 128& Then
            lne = glMemAddrDiv32(addr - 6144&)
            For i = lne * 8& To lne * 8& + 7&
                ScrnLines(i, 32&) = True
                ScrnLines(i, addr And 31&) = True
            Next i
        End If
    Next addr
End Sub




Public Sub ScanlinePaint(lne As Long)
    If glUseScreen >= 1000& Then Exit Sub
    
    Dim lLneIndex As Long, lColIndex As Long, X As Long, sbyte As Long, abyte As Long, lIndex As Long
        
    If ScrnLines(lne, 32&) = True Then
        If lne < glTopMost Then glTopMost = lne
        If lne > glBottomMost Then glBottomMost = lne

        lLneIndex = glRowIndex(lne)
        lColIndex = glColIndex(lne)
        For X = 0& To 31&
            If ScrnLines(lne, X) = True Then
                If X < glLeftMost Then glLeftMost = X
                If X > glRightMost Then glRightMost = X
                
                sbyte = gRAMPage(glUseScreen, glScreenMem(lne, X))
                abyte = gRAMPage(glUseScreen, (lLneIndex + X))
                
                If (abyte And 128&) And (bFlashInverse) Then
                    ' // Swap fore- and back-colours
                    abyte = abyte Xor 128&
                End If
                
                lIndex = (lColIndex + X + X)
                glBufferBits(lIndex) = gtBitTable(sbyte, abyte).dw0
                glBufferBits(lIndex + 1&) = gtBitTable(sbyte, abyte).dw1
                
                ScrnLines(lne, X) = False
            End If
        Next X
        ScrnLines(lne, 32&) = False ' // Flag indicates this line has been rendered on the bitmap
        ScrnNeedRepaint = True
    End If
End Sub

Sub screenPaint()
    ' // Only update screen if necessary
    If ScrnNeedRepaint = False Then Exit Sub
    
    ' // TC2048=May slow things down :(
    If glUseScreen >= 1000& Then
        TC2048screenPaint
        Exit Sub
    End If
    
    glLeftMost = glLeftMost * 8&
    glRightMost = glRightMost * 8&
    
    'gpicDisplay.Visible = False
    StretchDIBits gpicDC, _
                  glLeftMost * glDisplayXMultiplier, _
                  (glBottomMost + 1&) * glDisplayYMultiplier - 1&, _
                  (glRightMost - glLeftMost + 8&) * glDisplayXMultiplier, _
                  -(glBottomMost - glTopMost + 1&) * glDisplayYMultiplier, _
                  glLeftMost, _
                  glTopMost, _
                  (glRightMost - glLeftMost) + 8&, _
                  glBottomMost - glTopMost + 1&, _
                  glBufferBits(0&), _
                  bmiBuffer, _
                  DIB_RGB_COLORS, _
                  SRCCOPY
    gpicDisplay.REFRESH
    'gpicDisplay.Visible = True
    
    glTopMost = 191&
    glBottomMost = 0&
    glLeftMost = 31&
    glRightMost = 0&
    
    ScrnNeedRepaint = False
End Sub


 
Private Sub SetZXPrinterPixel(X As Long, y As Long)
    Dim lElement As Long
    If X = -1 Then X = 256
    
    lElement = y * 32 + X \ 8
    gcZXPrinterBits(lElement) = gcZXPrinterBits(lElement) Or lRevBitValues(X Mod 8)
End Sub

Sub TC2048PaintHiRes()
    Dim lne As Long, X As Long
    Dim sbyte As Long
    Dim lLeftMost As Long, lRightMost As Long, lTopMost As Long, lBottomMost As Long
        
    ' // Bob Woodring's (RGW) improvements to display speed (lookup table of colour values)
    Dim lIndex    As Long
    Dim lLneIndex As Long
    Dim lColIndex As Long
    
    'gpicDisplay.Visible = False
    
    lTopMost = 191
    lBottomMost = 0
    lLeftMost = 63
    lRightMost = 0
    

    For lne = 0 To 191
        If ScrnLines(lne, 64) = True Then
            If lne < lTopMost Then lTopMost = lne
            If lne > lBottomMost Then lBottomMost = lne
            ' // RGW: Get line and column indexes from a lookup table for speed
            lLneIndex = glRowIndex(lne)
            lColIndex = glColIndex(lne) * 2
            For X = 0 To 63
                If ScrnLines(lne, X) = True Then
                    If X < lLeftMost Then lLeftMost = X
                    If X > lRightMost Then lRightMost = X
                    
                    sbyte = gRAMPage(5, glScreenMemTC2048HiRes(lne, X))
                    
                    lIndex = (lColIndex + X + X)
                    glBufferBits(lIndex) = gtBitTable(sbyte, glTC2048HiResColour).dw0
                    glBufferBits(lIndex + 1) = gtBitTable(sbyte, glTC2048HiResColour).dw1
                    ScrnLines(lne, X) = False
                End If
            Next X
            ScrnLines(lne, 64) = False
        End If
    Next lne
    
    lLeftMost = lLeftMost * 8
    lRightMost = lRightMost * 8
    
    StretchDIBits gpicDC, lLeftMost * (glDisplayXMultiplier / 2), (lBottomMost + 1) * glDisplayYMultiplier - 1, (lRightMost - lLeftMost + 8) * (glDisplayXMultiplier / 2), -(lBottomMost - lTopMost + 1) * glDisplayYMultiplier, lLeftMost, lTopMost, (lRightMost - lLeftMost) + 8, lBottomMost - lTopMost + 1, glBufferBits(0), bmiBuffer, DIB_RGB_COLORS, SRCCOPY
    gpicDisplay.REFRESH
    'gpicDisplay.Visible = True
    ScrnNeedRepaint = False
End Sub

Public Sub TC2048refreshFlashChars()
    Dim addr As Long, lne As Long, i As Long, lScrn As Long, lOffset As Long
    
    If glUseScreen = 1006 Then Exit Sub
    
    If glUseScreen = 1001 Then
        lOffset = 8192
    ElseIf glUseScreen = 1002 Then
        ' // HiColour
        bFlashInverse = Not (bFlashInverse)
        
        For addr = 8192 To 14335
            If gRAMPage(5, addr) And 128 Then
                lne = glMemAddrDiv32(addr - 8192)
                For i = lne * 8 To lne * 8 + 7
                    ScrnLines(i, 32) = True
                    ScrnLines(i, addr And 31) = True
                Next i
                ScrnNeedRepaint = True
            End If
        Next addr
        
        Exit Sub
    End If
      
    bFlashInverse = Not (bFlashInverse)
    
    For addr = 6144 To 6911
        If gRAMPage(5, addr + lOffset) And 128 Then
            lne = glMemAddrDiv32(addr - 6144)
            For i = lne * 8 To lne * 8 + 7
                ScrnLines(i, 32) = True
                ScrnLines(i, addr And 31) = True
            Next i
            ScrnNeedRepaint = True
        End If
    Next addr
End Sub

Sub TC2048screenPaint()
    If glUseScreen = 1001 Then
        TC2048ScreenPaintScrn1
    ElseIf glUseScreen = 1002 Then
        TC2048PaintHiColour
    ElseIf glUseScreen = 1006 Then
        TC2048PaintHiRes
    End If
End Sub

Sub TC2048PaintHiColour()
    Dim lne As Long, X As Long
    Dim sbyte As Long, abyte As Long
    Dim lLeftMost As Long, lRightMost As Long, lTopMost As Long, lBottomMost As Long
        
    ' // Bob Woodring's (RGW) improvements to display speed (lookup table of colour values)
    Dim lIndex    As Long
    Dim lLneIndex As Long
    Dim lColIndex As Long
    
    'gpicDisplay.Visible = False
    
    lTopMost = 191
    lBottomMost = 0
    lLeftMost = 31
    lRightMost = 0
    
    For lne = 0 To 191
        If ScrnLines(lne, 32) = True Then
            If lne < lTopMost Then lTopMost = lne
            If lne > lBottomMost Then lBottomMost = lne
            ' // RGW: Get line and column indexes from a lookup table for speed
            lLneIndex = glRowIndex(lne)
            lColIndex = glColIndex(lne)
            For X = 0 To 31
                If ScrnLines(lne, X) = True Then
                    If X < lLeftMost Then lLeftMost = X
                    If X > lRightMost Then lRightMost = X
                    
                    ' // All screen memory is in the bottom 16K of RAM (page 5)
                    sbyte = gRAMPage(5, glScreenMem(lne, X))
                    abyte = gRAMPage(5, glScreenMem(lne, X) + 8192)
                    
                    If (abyte And 128) And (bFlashInverse) Then
                        ' // Swap fore- and back-colours
                        abyte = abyte Xor 128
                    End If
                    
                    lIndex = (lColIndex + X + X)
                    glBufferBits(lIndex) = gtBitTable(sbyte, abyte).dw0
                    glBufferBits(lIndex + 1) = gtBitTable(sbyte, abyte).dw1
                    ScrnLines(lne, X) = False
                End If
            Next X
            ScrnLines(lne, 32) = False
        End If
    Next lne
    
    lLeftMost = lLeftMost * 8
    lRightMost = lRightMost * 8
    StretchDIBits gpicDC, lLeftMost * glDisplayXMultiplier, (lBottomMost + 1) * glDisplayYMultiplier - 1, (lRightMost - lLeftMost + 8) * glDisplayXMultiplier, -(lBottomMost - lTopMost + 1) * glDisplayYMultiplier, lLeftMost, lTopMost, (lRightMost - lLeftMost) + 8, lBottomMost - lTopMost + 1, glBufferBits(0), bmiBuffer, DIB_RGB_COLORS, SRCCOPY
    gpicDisplay.REFRESH
    'gpicDisplay.Visible = True
    ScrnNeedRepaint = False
End Sub

Sub TC2048ScreenPaintScrn1()
    Dim lne As Long, X As Long
    Dim sbyte As Long, abyte As Long
    Dim lLeftMost As Long, lRightMost As Long, lTopMost As Long, lBottomMost As Long
        
    ' // Bob Woodring's (RGW) improvements to display speed (lookup table of colour values)
    Dim lIndex    As Long
    Dim lLneIndex As Long
    Dim lColIndex As Long
    
    lTopMost = 191
    lBottomMost = 0
    lLeftMost = 31
    lRightMost = 0
    
    For lne = 0 To 191
        If ScrnLines(lne, 32) = True Then
            If lne < lTopMost Then lTopMost = lne
            If lne > lBottomMost Then lBottomMost = lne
            ' // RGW: Get line and column indexes from a lookup table for speed
            lLneIndex = glRowIndex(lne)
            lColIndex = glColIndex(lne)
            For X = 0 To 31
                If ScrnLines(lne, X) = True Then
                    If X < lLeftMost Then lLeftMost = X
                    If X > lRightMost Then lRightMost = X
                    
                    ' // All screen memory is in the bottom 16K of RAM (page 5)
                    sbyte = gRAMPage(5, glScreenMem(lne, X) + 8192)
                    abyte = gRAMPage(5, (lLneIndex + X + 8192))
                    
                    If (abyte And 128) And (bFlashInverse) Then
                        ' // Swap fore- and back-colours
                        abyte = abyte Xor 128
                    End If
                    
                    lIndex = (lColIndex + X + X)
                    glBufferBits(lIndex) = gtBitTable(sbyte, abyte).dw0
                    glBufferBits(lIndex + 1) = gtBitTable(sbyte, abyte).dw1
                    ScrnLines(lne, X) = False
                End If
            Next X
            ScrnLines(lne, 32) = False
        End If
    Next lne
    
    lLeftMost = lLeftMost * 8
    lRightMost = lRightMost * 8
    StretchDIBits gpicDC, lLeftMost * glDisplayXMultiplier, (lBottomMost + 1) * glDisplayYMultiplier - 1, (lRightMost - lLeftMost + 8) * glDisplayXMultiplier, -(lBottomMost - lTopMost + 1) * glDisplayYMultiplier, lLeftMost, lTopMost, (lRightMost - lLeftMost) + 8, lBottomMost - lTopMost + 1, glBufferBits(0), bmiBuffer, DIB_RGB_COLORS, SRCCOPY
    gpicDisplay.REFRESH

    ScrnNeedRepaint = False
End Sub


Private Function ZXPrinterIn() As Long
    ' //  (64) D6 = 0 if ZXPrinter is present, else 1
    ' // (128) D7 = 1 if the stylus in on the paper
    ' //   (1) D0 = 0/1 toggle from the encoder disk
    
    If frmZXPrinter.Visible = False Then
        ' // Unconnected port
        ZXPrinterIn = 64&
        Exit Function
    End If
    
    If lZXPrinterMotorOn Then
        ' // Flip the encoder disk bit
        If lZXPrinterEncoder = 1 Then lZXPrinterEncoder = 0 Else lZXPrinterEncoder = 1
        ' // For every 0>1 cycle of the encoder disk, draw a pixel if the stylus is on
        If (lZXPrinterEncoder = 0) Or (lZXPrinterX >= 257) Then
            If lZXPrinterStylusOn Then
                If lZXPrinterX > 128 Then
                    SetZXPrinterPixel lZXPrinterX - 2, lZXPrinterY
                Else
                    SetZXPrinterPixel lZXPrinterX - 1, lZXPrinterY
                End If
            End If
            lZXPrinterX = lZXPrinterX + 1
        End If
    End If
    
    If lZXPrinterX >= 0 And lZXPrinterX < 128 Then
        ' // Stylus 1 is over paper
        ZXPrinterIn = 128 Or lZXPrinterEncoder
    ElseIf lZXPrinterX = 128 Then
        ' // Stylus 1 has left the paper
        ZXPrinterIn = lZXPrinterEncoder
        lZXPrinterX = 129
    ElseIf lZXPrinterX > 128 And lZXPrinterX <= 256 Then
        ' // Stylus 2 is over paper
        ZXPrinterIn = 128 Or lZXPrinterEncoder
    Else
        ' // Stylus 2 has left the paper, advance the paper one pixel row
        ' // and set the position of Stylus 1 to the start of the next row
        lZXPrinterEncoder = 0
        ZXPrinterIn = lZXPrinterEncoder
        lZXPrinterX = 0
        lZXPrinterY = lZXPrinterY + 1
        If lZXPrinterY >= glZXPrinterBMPHeight Then
            glZXPrinterBMPHeight = lZXPrinterY + 32
            ReDim Preserve gcZXPrinterBits(glZXPrinterBMPHeight * 32)
            bmiZXPrinter.bmiHeader.biHeight = glZXPrinterBMPHeight
        End If
        
        If (lZXPrinterY Mod 8) = 0 Then
            ' // Every 8 rows, update the visible display
            If frmZXPrinter.picView.Height > lZXPrinterY Then
                StretchDIBitsMono frmZXPrinter.picView.hdc, 0, frmZXPrinter.picView.Height, 256, -lZXPrinterY - 1, 0, 0, 256, lZXPrinterY + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
            Else
                StretchDIBitsMono frmZXPrinter.picView.hdc, 0, frmZXPrinter.picView.Height, 256, -frmZXPrinter.picView.Height - 1, 0, lZXPrinterY - frmZXPrinter.picView.Height, 256, frmZXPrinter.picView.Height + 1, gcZXPrinterBits(0&), bmiZXPrinter, DIB_RGB_COLORS, SRCCOPY
            End If
            frmZXPrinter.picView.REFRESH
            
            ' // Set up the scroll bar properties for the visible display
            ' // to allow scrolling back over the material printed so far
            If lZXPrinterY > frmZXPrinter.picView.Height Then
                frmZXPrinter.vs.Min = frmZXPrinter.picView.Height \ 8
                frmZXPrinter.vs.Max = lZXPrinterY \ 8
                frmZXPrinter.vs.Value = lZXPrinterY \ 8
            Else
                frmZXPrinter.vs.Min = 0
                frmZXPrinter.vs.Max = 0
            End If
        End If
    End If
End Function

Private Sub ZXPrinterOut(b As Long)
    If (b And 4&) Then
        lZXPrinterMotorOn = False
    Else
        lZXPrinterMotorOn = True
    End If
    
    If (b And 128&) Then
        lZXPrinterStylusOn = True
    Else
        lZXPrinterStylusOn = False
    End If
End Sub


'MM 03.02.2003 - BEGIN
Private Function PCJoystickToZXJoystick(ByVal lPCJoystick As Long, _
                                        ByVal lPCJoystickFire As Long) As Long
        
    'API return value
    Dim lResult As Long
    'Joystick capablities
    Dim uJoyCaps As JOYCAPS
    'Extended joystick-info
    Dim uJoyInfoEx As JOYINFOEX
    'Minimal Up-Value
    Dim lMinimalUp As Long
    'Minimal Down-Value
    Dim lMinimalDown As Long
    'Minimal Left-Value
    Dim lMinimalLeft As Long
    'Minimal Right-Value
    Dim lMinimalRight As Long
    'Middle X
    Dim lMiddleX As Long
    'Middle Y
    Dim lMiddleY As Long
    'Joystick-Value
    Dim lJoyRes As Long
    
    'Init
    Let PCJoystickToZXJoystick = 0
       
    'vbSpec reads the joystick capabilities -- the user could theoretically
    'unplugg a joystick and plug another with another capabilities. This
    'code supports analog joysticks only -- I didn't have any digital :)
    Let lResult = joyGetDevCaps(lPCJoystick, uJoyCaps, Len(uJoyCaps))
    'The uJoyInfoEx will be preared to get the joystick informations
    'The lenght of the structure muss be transfered
    Let uJoyInfoEx.dwSize = Len(uJoyInfoEx)
    'All joystick information should be got
    Let uJoyInfoEx.dwFlags = JOY_RETURNALL
    'Joystickinfos will be loaded into uJoyInfoEx
    'Remarks: the joystick must be pluggen as Joystick1
    Let lResult = joyGetPosEx(lPCJoystick, uJoyInfoEx)
    'Calculate the middle X and y
    Let lMiddleX = Fix((uJoyCaps.wXmax - uJoyCaps.wXmin) / 2)
    Let lMiddleY = Fix((uJoyCaps.wYmax - uJoyCaps.wYmin) / 2)
    'Calculate minimal values for the directions
    Let lMinimalUp = lMiddleY - Fix((uJoyCaps.wYmax - uJoyCaps.wYmin) / 4)
    Let lMinimalDown = lMiddleY + Fix((uJoyCaps.wYmax - uJoyCaps.wYmin) / 4)
    Let lMinimalLeft = lMiddleX - Fix((uJoyCaps.wXmax - uJoyCaps.wXmin) / 4)
    Let lMinimalRight = lMiddleX + Fix((uJoyCaps.wXmax - uJoyCaps.wXmin) / 4)
    'Init joystick value
    Let lJoyRes = 0
    'Up
    If uJoyInfoEx.dwYpos < lMinimalUp Then
        Let lJoyRes = CLng(lJoyRes Or zxjdUp)
    End If
    'Down
    If uJoyInfoEx.dwYpos > lMinimalDown Then
        Let lJoyRes = CLng(lJoyRes Or zxjdDown)
    End If
    'Left
    If uJoyInfoEx.dwXpos < lMinimalLeft Then
        Let lJoyRes = CLng(lJoyRes Or zxjdLeft)
    End If
    'Right
    If uJoyInfoEx.dwXpos > lMinimalRight Then
        Let lJoyRes = CLng(lJoyRes Or zxjdRight)
    End If
    'Fire
    If (uJoyInfoEx.dwButtons And CLng(2 ^ lPCJoystickFire)) > 0 Then
        Let lJoyRes = CLng(lJoyRes Or zxjdFire)
    End If
    'Button
    If uJoyInfoEx.dwButtons <> 0 Then
        Let lJoyRes = CLng(lJoyRes Or (uJoyInfoEx.dwButtons * zxjdButtonBase))
    End If
    'Return value
    Let PCJoystickToZXJoystick = lJoyRes
End Function
'MM 03.02.2003 - END

'MM JD
Private Function JoystickInitIN(ByVal port As Long) As Long
    
    'ZX-joystick position
    Dim lZXJoystickPosition As Long
    'Joystick value (as the result of IN)
    Dim lJoyRes As Long
    'Joystick definition
    Dim jdAction As JOYSTICK_DEFINITION
    'Counter
    Dim lCounter As Long
    'TMP-Value
    Dim lTMP As Long
    Dim sTMP As String
    'Result
    Dim lRes As Long
    
    'Initialise
    lRes = 0
    
    'If Joystick 1 is valid then
    If bJoystick1Valid Then
        'Get joystick position
        Let lZXJoystickPosition = PCJoystickToZXJoystick(JOYSTICKID1, lPCJoystick1Fire)
        'Up
        If (lZXJoystickPosition And zxjdUp) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_UP).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Down
        If (lZXJoystickPosition And zxjdDown) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_DOWN).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Left
        If (lZXJoystickPosition And zxjdLeft) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_LEFT).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Right
        If (lZXJoystickPosition And zxjdRight) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_RIGHT).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Buttons
        If (lZXJoystickPosition And &HFFFFFFE0) > 0 Then
            'Get button number
            lCounter = CLng((lZXJoystickPosition And &HFFFFFFE0) / 32)
            lCounter = CLng(Log(lCounter) / Log(2)) + 1
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK1, lPCJoystick1Is, JDT_BUTTON_BASE + lCounter).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Finish
        If lZXJoystickPosition = 0 Then
            If lKey1 > 0 Then
                lRes = 0
                Let lKey1 = 0
            End If
        End If
    End If
       
    'If Joystick 2 is valid then
    If bJoystick2Valid Then
        'Get joystick position
        Let lZXJoystickPosition = PCJoystickToZXJoystick(JOYSTICKID2, lPCJoystick2Fire)
        'Up
        If (lZXJoystickPosition And zxjdUp) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_UP).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Down
        If (lZXJoystickPosition And zxjdDown) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_DOWN).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Left
        If (lZXJoystickPosition And zxjdLeft) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_LEFT).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Right
        If (lZXJoystickPosition And zxjdRight) > 0 Then
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_RIGHT).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Buttons
        If (lZXJoystickPosition And &HFFFFFFE0) > 0 Then
            'Get button number
            lCounter = CLng((lZXJoystickPosition And &HFFFFFFE0) / 32)
            lCounter = CLng(Log(lCounter) / Log(2)) + 1
            'Get Joystick definition
            jdAction = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter)
            'If the joystick had been
            If ActionDefined(jdAction) Then
                'If the actual port is the up-port of the joystikc
                If (((port And &HFF&) = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).sKey = vbNullString)) Or _
                   ((port = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).lPort) And (aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).sKey <> vbNullString)) Then
                    'Set value
                    lTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).lValue
                    sTMP = aJoystikDefinitionTable(JDT_JOYSTICK2, lPCJoystick2Is, JDT_BUTTON_BASE + lCounter).sKey
                    If sTMP <> vbNullString Then
                        If lRes = 0 Then
                            lRes = lTMP
                        Else
                            lRes = lRes And lTMP
                        End If
                    Else
                        lRes = lRes Or lTMP
                    End If
                    Let lKey1 = 1
                End If
            End If
        End If
        'Finish
        If lZXJoystickPosition = 0 Then
            If lKey1 > 0 Then
                lRes = 0
                Let lKey1 = 0
            End If
        End If
    End If
       
    'Return value
    JoystickInitIN = lRes
End Function

'MM JD
'Get if the actual action has been definet
Private Function ActionDefined(ByRef jdAction As JOYSTICK_DEFINITION) As Boolean
    'An action is defined if you have port-and-value-pair or a key
    ActionDefined = ((jdAction.lPort <> 0) And (jdAction.lValue <> 0))
End Function

'MM JD
'Translate key strokes in port and in values
Public Sub KeyStrokeToPortValue(ByVal lKeyCode As Long, ByVal lShift As Long, _
                                ByRef sKey As String, ByRef lPort As Long, _
                                ByRef lValue As Long)

    'Initialise
    sKey = vbNullString
    lPort = 0
    lValue = 0
    
    'Delete current definition
    If lKeyCode = 46 And lShift = 0 Then
        sKey = vbNullString
        lPort = 0
        lValue = 0
        Exit Sub
    End If
    'Symbol shift
    If lKeyCode = 17 And lShift = 2 Then
        sKey = "SYM"
        lPort = 32766
        lValue = 189
        Exit Sub
    End If
    'Caps shift
    If lKeyCode = 16 And lShift = 1 Then
        sKey = "CAP"
        lPort = 65278
        lValue = 190
        Exit Sub
    End If
    'Space
    If lKeyCode = 32 And lShift = 0 Then
        sKey = "SPA"
        lPort = 32766
        lValue = 190
        Exit Sub
    End If
    'The normal keys
    sKey = Chr(lKeyCode)
    'First row
    If sKey = "1" Then
        lPort = 63486
        lValue = 190
        Exit Sub
    End If
    If sKey = "2" Then
        lPort = 63486
        lValue = 189
        Exit Sub
    End If
    If sKey = "3" Then
        lPort = 63486
        lValue = 187
        Exit Sub
    End If
    If sKey = "4" Then
        lPort = 63486
        lValue = 183
        Exit Sub
    End If
    If sKey = "5" Then
        lPort = 63486
        lValue = 175
        Exit Sub
    End If
    If sKey = "6" Then
        lPort = 61438
        lValue = 175
        Exit Sub
    End If
    If sKey = "7" Then
        lPort = 61438
        lValue = 183
        Exit Sub
    End If
    If sKey = "8" Then
        lPort = 61438
        lValue = 187
        Exit Sub
    End If
    If sKey = "9" Then
        lPort = 61438
        lValue = 189
        Exit Sub
    End If
    If sKey = "0" Then
        lPort = 61438
        lValue = 190
        Exit Sub
    End If
    'Second row
    If sKey = "Q" Then
        lPort = 64510
        lValue = 190
        Exit Sub
    End If
    If sKey = "W" Then
        lPort = 64510
        lValue = 189
        Exit Sub
    End If
    If sKey = "E" Then
        lPort = 64510
        lValue = 187
        Exit Sub
    End If
    If sKey = "R" Then
        lPort = 64510
        lValue = 183
        Exit Sub
    End If
    If sKey = "T" Then
        lPort = 64510
        lValue = 175
        Exit Sub
    End If
    If sKey = "Y" Then
        lPort = 57342
        lValue = 175
        Exit Sub
    End If
    If sKey = "U" Then
        lPort = 57342
        lValue = 183
        Exit Sub
    End If
    If sKey = "I" Then
        lPort = 57342
        lValue = 187
        Exit Sub
    End If
    If sKey = "O" Then
        lPort = 57342
        lValue = 189
        Exit Sub
    End If
    If sKey = "P" Then
        lPort = 57342
        lValue = 190
        Exit Sub
    End If
    'Third row
    If sKey = "A" Then
        lPort = 65022
        lValue = 190
        Exit Sub
    End If
    If sKey = "S" Then
        lPort = 65022
        lValue = 189
        Exit Sub
    End If
    If sKey = "D" Then
        lPort = 65022
        lValue = 187
        Exit Sub
    End If
    If sKey = "F" Then
        lPort = 65022
        lValue = 183
        Exit Sub
    End If
    If sKey = "G" Then
        lPort = 65022
        lValue = 175
        Exit Sub
    End If
    If sKey = "H" Then
        lPort = 49150
        lValue = 175
        Exit Sub
    End If
    If sKey = "J" Then
        lPort = 49150
        lValue = 183
        Exit Sub
    End If
    If sKey = "K" Then
        lPort = 49150
        lValue = 187
        Exit Sub
    End If
    If sKey = "L" Then
        lPort = 49150
        lValue = 189
        Exit Sub
    End If
    'Fourth row
    If sKey = "Z" Then
        lPort = 65278
        lValue = 189
        Exit Sub
    End If
    If sKey = "X" Then
        lPort = 65278
        lValue = 187
        Exit Sub
    End If
    If sKey = "C" Then
        lPort = 65278
        lValue = 183
        Exit Sub
    End If
    If sKey = "V" Then
        lPort = 65278
        lValue = 175
        Exit Sub
    End If
    If sKey = "B" Then
        lPort = 32766
        lValue = 175
        Exit Sub
    End If
    If sKey = "N" Then
        lPort = 32766
        lValue = 183
        Exit Sub
    End If
    If sKey = "M" Then
        lPort = 32766
        lValue = 187
        Exit Sub
    End If
    'This key is not a Speccy-key
    sKey = vbNullString
    lPort = 0
    lValue = 0
End Sub

'MM JD
'Translate ports and values in keystrokes
Public Sub PortValueToKeyStroke(ByVal lPort As Long, ByVal lValue As Long, _
                                ByRef sKey As String)

    'Initialise
    sKey = vbNullString
    
    'First row, left
    If lPort = 63486 Then
        If lValue = 190 Then
            sKey = "1"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "2"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "3"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "4"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "5"
            Exit Sub
        End If
    End If
    'First row, right
    If lPort = 61438 Then
        If lValue = 190 Then
            sKey = "0"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "9"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "8"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "7"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "6"
            Exit Sub
        End If
    End If
    'Second row, left
    If lPort = 64510 Then
        If lValue = 190 Then
            sKey = "Q"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "W"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "E"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "R"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "T"
            Exit Sub
        End If
    End If
    'Second row, right
    If lPort = 57342 Then
        If lValue = 190 Then
            sKey = "P"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "O"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "I"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "U"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "Y"
            Exit Sub
        End If
    End If
    'Third row, left
    If lPort = 65022 Then
        If lValue = 190 Then
            sKey = "A"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "S"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "D"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "F"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "G"
            Exit Sub
        End If
    End If
    'Third row, right
    If lPort = 49150 Then
        If lValue = 190 Then
            sKey = "RET"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "L"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "K"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "J"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "H"
            Exit Sub
        End If
    End If
    'Fourth row, left
    If lPort = 65278 Then
        If lValue = 190 Then
            sKey = "CAP"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "Z"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "X"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "C"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "V"
            Exit Sub
        End If
    End If
    'Fourth row, right
    If lPort = 32766 Then
        If lValue = 190 Then
            sKey = "SPA"
            Exit Sub
        End If
        If lValue = 189 Then
            sKey = "SYM"
            Exit Sub
        End If
        If lValue = 187 Then
            sKey = "M"
            Exit Sub
        End If
        If lValue = 183 Then
            sKey = "N"
            Exit Sub
        End If
        If lValue = 175 Then
            sKey = "B"
            Exit Sub
        End If
    End If
    'That is not a key
    sKey = vbNullString
End Sub
