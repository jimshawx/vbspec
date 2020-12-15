Attribute VB_Name = "modAY8912"
' /*******************************************************************************
'   modAY8912.bas within vbSpec.vbp
'
'   Routines for emulating the 128K Spectrum's AY-3-8912 sound generator
'
'   Written using technical information and techniques gleaned from
'   the MAME AY-8910 module "src/sound/ay8910.c". See http://www.mame.net/
'   for further information.
'
'   This version of the code is offered under the GPL with the kind
'   permission of Nicola Salmoria and the MAME authors.
'
'   Original Visual Basic Port Written By:
'          James Bagg <chipmunk_uk_1@hotmail.com>
'
'   Minor VB-specific optimisations and mods by
'          Chris Cowley <ccowley@grok.co.uk>
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

Public Const MAX_OUTPUT As Long = 63
Public Const AY_STEP As Long = 32768
Public Const MAXVOL As Long = &H1F

' // AY register ID's
Public Const AY_AFINE As Long = 0
Public Const AY_ACOARSE As Long = 1
Public Const AY_BFINE As Long = 2
Public Const AY_BCOARSE As Long = 3
Public Const AY_CFINE As Long = 4
Public Const AY_CCOARSE As Long = 5
Public Const AY_NOISEPER As Long = 6
Public Const AY_ENABLE As Long = 7
Public Const AY_AVOL As Long = 8
Public Const AY_BVOL As Long = 9
Public Const AY_CVOL As Long = 10
Public Const AY_EFINE As Long = 11
Public Const AY_ECOARSE As Long = 12
Public Const AY_ESHAPE As Long = 13
Public Const AY_PORTA As Long = 14
Public Const AY_PORTB As Long = 15

Public AYPSG As AY8912

Type AY8912
    sampleRate As Long
    register_latch As Long
    Regs(16) As Long
    UpdateStep As Double
    PeriodA As Long
    PeriodB As Long
    PeriodC As Long
    PeriodN As Long
    PeriodE As Long
    CountA As Long
    CountB As Long
    CountC As Long
    CountN As Long
    CountE As Long
    VolA As Long
    VolB As Long
    VolC As Long
    VolE As Long
    EnvelopeA As Long
    EnvelopeB As Long
    EnvelopeC As Long
    OutputA As Long
    OutputB As Long
    OutputC As Long
    OutputN As Long
    CountEnv As Long
    Hold As Long
    Alternate As Long
    Attack As Long
    Holding As Long
    VolTable2(64) As Long
End Type

Public AY_OutNoise As Long
Public VolA As Long, VolB As Long, VolC As Long
Private lOut1 As Long, lOut2 As Long, lOut3 As Long
Public AY_Left As Long
Public AY_NextEvent As Long
Public Buffer_Length As Long

Public Sub AY8912_reset()
    Dim i As Long
    
    With AYPSG
        .register_latch = 0
        .OutputA = 0
        .OutputB = 0
        .OutputC = 0
        .OutputN = &HFF
        .PeriodA = 0
        .PeriodB = 0
        .PeriodC = 0
        .PeriodN = 0
        .PeriodE = 0
        .CountA = 0
        .CountB = 0
        .CountC = 0
        .CountN = 0
        .CountE = 0
        .VolA = 0
        .VolB = 0
        .VolC = 0
        .VolE = 0
        .EnvelopeA = 0
        .EnvelopeB = 0
        .EnvelopeC = 0
        .CountEnv = 0
        .Hold = 0
        .Alternate = 0
        .Holding = 0
        .Attack = 0
    End With
    
    For i = 0 To AY_PORTA
        AYWriteReg i, 0
    Next i
End Sub

Public Sub AY8912_set_clock(clock As Double)
    Dim t1 As Double
    
    ' // Calculate the number of AY_STEPs which happen during one sample
    ' // at the given sample rate. No. of events = sample rate / (clock/8).
    ' // AY_STEP is a multiplier used to turn the fraction into a fixed point
    ' // number.
    t1 = CDbl(AY_STEP) * CDbl(AYPSG.sampleRate) * CDbl(8)
    
    AYPSG.UpdateStep = t1 / clock
End Sub


' // AY8912_set_volume()
' //
' // Initialize the volume table
Public Sub AY8912InitVolumeTable()
    ' // The following volume levels are taken from the sound.c & sound.h files
    ' // in the FUSE emulator (suitably rescaled to 00-3F from 0000-FFFF) and
    ' // apparently more accurately represent real volume levels as measured
    ' // from a 128K Spectrum than the original algorithm used in previous
    ' // versions of vbSpec.
    AYPSG.VolTable2(0) = 0: AYPSG.VolTable2(1) = 0: AYPSG.VolTable2(2) = 1: AYPSG.VolTable2(3) = 1
    AYPSG.VolTable2(4) = 1: AYPSG.VolTable2(5) = 1: AYPSG.VolTable2(6) = 2: AYPSG.VolTable2(7) = 2
    AYPSG.VolTable2(8) = 3: AYPSG.VolTable2(9) = 3: AYPSG.VolTable2(10) = 4: AYPSG.VolTable2(11) = 4
    AYPSG.VolTable2(12) = 5: AYPSG.VolTable2(13) = 5: AYPSG.VolTable2(14) = 9: AYPSG.VolTable2(15) = 9
    
    AYPSG.VolTable2(16) = 11: AYPSG.VolTable2(17) = 11: AYPSG.VolTable2(18) = 17: AYPSG.VolTable2(19) = 17
    AYPSG.VolTable2(20) = 23: AYPSG.VolTable2(21) = 23: AYPSG.VolTable2(22) = 29: AYPSG.VolTable2(23) = 29
    AYPSG.VolTable2(24) = 37: AYPSG.VolTable2(25) = 37: AYPSG.VolTable2(26) = 44: AYPSG.VolTable2(27) = 44
    AYPSG.VolTable2(28) = 54: AYPSG.VolTable2(29) = 54: AYPSG.VolTable2(30) = 63: AYPSG.VolTable2(31) = 63
End Sub

Public Sub AYWriteReg(r As Long, v As Long)
    Dim old As Long
    
    AYPSG.Regs(r) = v
    
    ' // A note about the period of tones, noise and envelope: for speed reasons,
    ' // we count down from the period to 0, but careful studies of the chip
    ' // output prove that it instead counts up from 0 until the counter becomes
    ' // greater or equal to the period. This is an important difference when the
    ' // program is rapidly changing the period to modulate the sound.
    ' // To compensate for the difference, when the period is changed we adjust
    ' // our internal counter.
    ' // Also, note that period = 0 is the same as period = 1. This is mentioned
    ' // in the YM2203 data sheets. However, this does NOT apply to the Envelope
    ' // period. In that case, period = 0 is half as period = 1.
    Select Case r
        Case AY_AFINE, AY_ACOARSE
            AYPSG.Regs(AY_ACOARSE) = AYPSG.Regs(AY_ACOARSE) And &HF
            old = AYPSG.PeriodA
            AYPSG.PeriodA = (AYPSG.Regs(AY_AFINE) + (256 * AYPSG.Regs(AY_ACOARSE))) * AYPSG.UpdateStep
            If (AYPSG.PeriodA = 0) Then AYPSG.PeriodA = AYPSG.UpdateStep
            AYPSG.CountA = AYPSG.CountA + (AYPSG.PeriodA - old)
            If (AYPSG.CountA <= 0) Then AYPSG.CountA = 1
        Case AY_BFINE, AY_BCOARSE
            AYPSG.Regs(AY_BCOARSE) = AYPSG.Regs(AY_BCOARSE) And &HF
            old = AYPSG.PeriodB
            AYPSG.PeriodB = (AYPSG.Regs(AY_BFINE) + (256 * AYPSG.Regs(AY_BCOARSE))) * AYPSG.UpdateStep
            If (AYPSG.PeriodB = 0) Then AYPSG.PeriodB = AYPSG.UpdateStep
            AYPSG.CountB = AYPSG.CountB + AYPSG.PeriodB - old
            If (AYPSG.CountB <= 0) Then AYPSG.CountB = 1
        Case AY_CFINE, AY_CCOARSE
            AYPSG.Regs(AY_CCOARSE) = AYPSG.Regs(AY_CCOARSE) And &HF
            old = AYPSG.PeriodC
            AYPSG.PeriodC = (AYPSG.Regs(AY_CFINE) + (256 * AYPSG.Regs(AY_CCOARSE))) * AYPSG.UpdateStep
            If (AYPSG.PeriodC = 0) Then AYPSG.PeriodC = AYPSG.UpdateStep
            AYPSG.CountC = AYPSG.CountC + (AYPSG.PeriodC - old)
            If (AYPSG.CountC <= 0) Then AYPSG.CountC = 1
        Case AY_NOISEPER
            AYPSG.Regs(AY_NOISEPER) = AYPSG.Regs(AY_NOISEPER) And &H1F
            old = AYPSG.PeriodN
            AYPSG.PeriodN = AYPSG.Regs(AY_NOISEPER) * AYPSG.UpdateStep
            If (AYPSG.PeriodN = 0) Then AYPSG.PeriodN = AYPSG.UpdateStep
            AYPSG.CountN = AYPSG.CountN + (AYPSG.PeriodN - old)
            If (AYPSG.CountN <= 0) Then AYPSG.CountN = 1
        Case AY_AVOL
            AYPSG.Regs(AY_AVOL) = AYPSG.Regs(AY_AVOL) And &H1F
            AYPSG.EnvelopeA = AYPSG.Regs(AY_AVOL) And &H10
            If AYPSG.EnvelopeA <> 0 Then
                AYPSG.VolA = AYPSG.VolE
            Else
                If AYPSG.Regs(AY_AVOL) <> 0 Then
                    AYPSG.VolA = AYPSG.VolTable2(AYPSG.Regs(AY_AVOL) * 2 + 1)
                Else
                    AYPSG.VolA = AYPSG.VolTable2(0)
                End If
            End If
        Case AY_BVOL
            AYPSG.Regs(AY_BVOL) = AYPSG.Regs(AY_BVOL) And &H1F
            AYPSG.EnvelopeB = AYPSG.Regs(AY_BVOL) And &H10
            If AYPSG.EnvelopeB <> 0 Then
                AYPSG.VolB = AYPSG.VolE
            Else
                If AYPSG.Regs(AY_BVOL) <> 0 Then
                    AYPSG.VolB = AYPSG.VolTable2(AYPSG.Regs(AY_BVOL) * 2 + 1)
                Else
                    AYPSG.VolB = AYPSG.VolTable2(0)
                End If
            End If
        Case AY_CVOL
            AYPSG.Regs(AY_CVOL) = AYPSG.Regs(AY_CVOL) And &H1F
            AYPSG.EnvelopeC = AYPSG.Regs(AY_CVOL) And &H10
            If AYPSG.EnvelopeC <> 0 Then
                AYPSG.VolC = AYPSG.VolE
            Else
                If AYPSG.Regs(AY_CVOL) <> 0 Then
                    AYPSG.VolC = AYPSG.VolTable2(AYPSG.Regs(AY_CVOL) * 2 + 1)
                Else
                    AYPSG.VolC = AYPSG.VolTable2(0)
                End If
            End If
        Case AY_EFINE, AY_ECOARSE
            old = AYPSG.PeriodE
            AYPSG.PeriodE = ((AYPSG.Regs(AY_EFINE) + (256 * AYPSG.Regs(AY_ECOARSE)))) * AYPSG.UpdateStep
            If (AYPSG.PeriodE = 0) Then AYPSG.PeriodE = AYPSG.UpdateStep \ 2
            AYPSG.CountE = AYPSG.CountE + (AYPSG.PeriodE - old)
            If (AYPSG.CountE <= 0) Then AYPSG.CountE = 1
        Case AY_ESHAPE
            ' // envelope shapes:
            ' // C AtAlH
            ' // 0 0 x x  \___
            ' //
            ' // 0 1 x x  /___
            ' //
            ' // 1 0 0 0  \\\\
            ' //
            ' // 1 0 0 1  \___
            ' //
            ' // 1 0 1 0  \/\/
            ' //          ___
            ' // 1 0 1 1  \
            ' //
            ' // 1 1 0 0  ////
            ' //           ___
            ' // 1 1 0 1  /
            ' //
            ' // 1 1 1 0  /\/\
            ' //
            ' // 1 1 1 1  /___
            ' //
            ' // The envelope counter on the AY-3-8910 has 16 AY_STEPs. On the YM2149 it
            ' // has twice the AY_STEPs, happening twice as fast. Since the end result is
            ' // just a smoother curve, we always use the YM2149 behaviour.
            If (AYPSG.Regs(AY_ESHAPE) <> &HFF) Then
                AYPSG.Regs(AY_ESHAPE) = AYPSG.Regs(AY_ESHAPE) And &HF
                If ((AYPSG.Regs(AY_ESHAPE) And &H4) = &H4) Then
                    AYPSG.Attack = MAXVOL
                Else
                    AYPSG.Attack = &H0
                End If
                
                AYPSG.Hold = AYPSG.Regs(AY_ESHAPE) And &H1
                AYPSG.Alternate = AYPSG.Regs(AY_ESHAPE) And &H2
                
                AYPSG.CountE = AYPSG.PeriodE
                
                AYPSG.CountEnv = MAXVOL ' // &h1f
                AYPSG.Holding = 0
                AYPSG.VolE = AYPSG.VolTable2(AYPSG.CountEnv Xor AYPSG.Attack)
                If (AYPSG.EnvelopeA <> 0) Then AYPSG.VolA = AYPSG.VolE
                If (AYPSG.EnvelopeB <> 0) Then AYPSG.VolB = AYPSG.VolE
                If (AYPSG.EnvelopeC <> 0) Then AYPSG.VolC = AYPSG.VolE
            End If
    End Select
End Sub

Public Function AY8912_init(clock As Double, sample_rate As Long, sample_bits As Long) As Long
    AYPSG.sampleRate = sample_rate
    
    AY8912_set_clock clock
    AY8912InitVolumeTable
    AY8912_reset
    
    AY8912_init = 0
End Function

Public Sub AY8912Update_8()
    Dim Buffer_Length As Long
    
    Buffer_Length = 400
    
    ' // The 8910 has three outputs, each output is the mix of one of the three
    ' // tone generators and of the (single) noise generator. The two are mixed
    ' // BEFORE going into the DAC. The formula to mix each channel is:
    ' // (ToneOn | ToneDisable) & (NoiseOn | NoiseDisable).
    ' // Note that this means that if both tone and noise are disabled, the output
    ' // is 1, not 0, and can be modulated changing the volume.

    ' // If the channels are disabled, set their output to 1, and increase the
    ' // counter, if necessary, so they will not be inverted during this update.
    ' // Setting the output to 1 is necessary because a disabled channel is locked
    ' // into the ON state (see above); and it has no effect if the volume is 0.
    ' // If the volume is 0, increase the counter, but don't touch the output.
    If (AYPSG.Regs(AY_ENABLE) And &H1) = &H1 Then
        If AYPSG.CountA <= (Buffer_Length * AY_STEP) Then AYPSG.CountA = AYPSG.CountA + (Buffer_Length * AY_STEP)
        AYPSG.OutputA = 1
    ElseIf (AYPSG.Regs(AY_AVOL) = 0) Then
        ' // note that I do count += Buffer_Length, NOT count = Buffer_Length + 1. You might think
        ' // it's the same since the volume is 0, but doing the latter could cause
        ' // interference when the program is rapidly modulating the volume.
        If AYPSG.CountA <= (Buffer_Length * AY_STEP) Then AYPSG.CountA = AYPSG.CountA + (Buffer_Length * AY_STEP)
    End If
    
    If (AYPSG.Regs(AY_ENABLE) And &H2) = &H2 Then
        If AYPSG.CountB <= (Buffer_Length * AY_STEP) Then AYPSG.CountB = AYPSG.CountB + (Buffer_Length * AY_STEP)
        AYPSG.OutputB = 1
    ElseIf AYPSG.Regs(AY_BVOL) = 0 Then
        If AYPSG.CountB <= (Buffer_Length * AY_STEP) Then AYPSG.CountB = AYPSG.CountB + (Buffer_Length * AY_STEP)
    End If
    
    If (AYPSG.Regs(AY_ENABLE) And &H4) = &H4 Then
        If AYPSG.CountC <= (Buffer_Length * AY_STEP) Then AYPSG.CountC = AYPSG.CountC + (Buffer_Length * AY_STEP)
        AYPSG.OutputC = 1
    ElseIf (AYPSG.Regs(AY_CVOL) = 0) Then
        If AYPSG.CountC <= (Buffer_Length * AY_STEP) Then AYPSG.CountC = AYPSG.CountC + (Buffer_Length * AY_STEP)
    End If
    
    ' // for the noise channel we must not touch OutputN - it's also not necessary
    ' //since we use AY_OutNoise.
    If ((AYPSG.Regs(AY_ENABLE) And &H38) = &H38) Then ' // all off
        If AYPSG.CountN <= (Buffer_Length * AY_STEP) Then AYPSG.CountN = AYPSG.CountN + (Buffer_Length * AY_STEP)
    End If
    
    AY_OutNoise = (AYPSG.OutputN Or AYPSG.Regs(AY_ENABLE))
End Sub

Public Function RenderByte() As Long
    VolA = 0: VolB = 0: VolC = 0

    ' // VolA, VolB and VolC keep track of how long each square wave stays
    ' // in the 1 position during the sample period.

    AY_Left = AY_STEP

    Do
        AY_NextEvent = 0
        
        If (AYPSG.CountN < AY_Left) Then
            AY_NextEvent = AYPSG.CountN
        Else
            AY_NextEvent = AY_Left
        End If

        If (AY_OutNoise And &H8) = &H8 Then
            If (AYPSG.OutputA = 1) Then VolA = VolA + AYPSG.CountA
            AYPSG.CountA = AYPSG.CountA - AY_NextEvent
            ' // PeriodA is the half period of the square wave. Here, in each
            ' // loop I add PeriodA twice, so that at the end of the loop the
            ' // square wave is in the same status (0 or 1) it was at the start.
            ' // vola is also incremented by PeriodA, since the wave has been 1
            ' // exactly half of the time, regardless of the initial position.
            ' // If we exit the loop in the middle, OutputA has to be inverted
            ' // and vola incremented only if the exit status of the square
            ' // wave is 1.

            Do While (AYPSG.CountA <= 0)
                AYPSG.CountA = AYPSG.CountA + AYPSG.PeriodA
                If (AYPSG.CountA > 0) Then
                    If (AYPSG.Regs(AY_ENABLE) And 1) = 0 Then AYPSG.OutputA = AYPSG.OutputA Xor 1
                    If (AYPSG.OutputA) Then VolA = VolA + AYPSG.PeriodA
                    Exit Do
                End If
                AYPSG.CountA = AYPSG.CountA + AYPSG.PeriodA
                VolA = VolA + AYPSG.PeriodA
            Loop
            If (AYPSG.OutputA = 1) Then VolA = VolA - AYPSG.CountA
        Else
            AYPSG.CountA = AYPSG.CountA - AY_NextEvent
            Do While (AYPSG.CountA <= 0)
                AYPSG.CountA = AYPSG.CountA + AYPSG.PeriodA
                If (AYPSG.CountA > 0) Then
                    AYPSG.OutputA = AYPSG.OutputA Xor 1
                    Exit Do
                End If
                AYPSG.CountA = AYPSG.CountA + AYPSG.PeriodA
            Loop
        End If

        If (AY_OutNoise And &H10) = &H10 Then
            If (AYPSG.OutputB = 1) Then VolB = VolB + AYPSG.CountB
            AYPSG.CountB = AYPSG.CountB - AY_NextEvent
            Do While (AYPSG.CountB <= 0)
                AYPSG.CountB = AYPSG.CountB + AYPSG.PeriodB
                If (AYPSG.CountB > 0) Then
                    If (AYPSG.Regs(AY_ENABLE) And 2) = 0 Then AYPSG.OutputB = AYPSG.OutputB Xor 1
                    If (AYPSG.OutputB) Then VolB = VolB + AYPSG.PeriodB
                    Exit Do
                End If
                AYPSG.CountB = AYPSG.CountB + AYPSG.PeriodB
                VolB = VolB + AYPSG.PeriodB
            Loop
            If (AYPSG.OutputB = 1) Then VolB = VolB - AYPSG.CountB
        Else
            AYPSG.CountB = AYPSG.CountB - AY_NextEvent
            Do While (AYPSG.CountB <= 0)
                AYPSG.CountB = AYPSG.CountB + AYPSG.PeriodB
                If (AYPSG.CountB > 0) Then
                    AYPSG.OutputB = AYPSG.OutputB Xor 1
                    Exit Do
                End If
                AYPSG.CountB = AYPSG.CountB + AYPSG.PeriodB
            Loop
        End If
        
        If (AY_OutNoise And &H20) = &H20 Then
            If (AYPSG.OutputC = 1) Then VolC = VolC + AYPSG.CountC
            AYPSG.CountC = AYPSG.CountC - AY_NextEvent
            Do While (AYPSG.CountC <= 0)
                AYPSG.CountC = AYPSG.CountC + AYPSG.PeriodC
                If (AYPSG.CountC > 0) Then
                    If (AYPSG.Regs(AY_ENABLE) And 4) = 0 Then AYPSG.OutputC = AYPSG.OutputC Xor 1
                    If (AYPSG.OutputC) Then VolC = VolC + AYPSG.PeriodC
                    Exit Do
                End If
                AYPSG.CountC = AYPSG.CountC + AYPSG.PeriodC
                VolC = VolC + AYPSG.PeriodC
            Loop
            If (AYPSG.OutputC = 1) Then VolC = VolC - AYPSG.CountC
        Else
            AYPSG.CountC = AYPSG.CountC - AY_NextEvent
            Do While (AYPSG.CountC <= 0)
                AYPSG.CountC = AYPSG.CountC + AYPSG.PeriodC
                If (AYPSG.CountC > 0) Then
                    AYPSG.OutputC = AYPSG.OutputC Xor 1
                    Exit Do
                End If
                AYPSG.CountC = AYPSG.CountC + AYPSG.PeriodC
            Loop
        End If

        AYPSG.CountN = AYPSG.CountN - AY_NextEvent
        If (AYPSG.CountN <= 0) Then
            ' // Is noise output going to change?
            AYPSG.OutputN = Int(Rnd(1) * 2) * 255
            AY_OutNoise = (AYPSG.OutputN Or AYPSG.Regs(AY_ENABLE))

            AYPSG.CountN = AYPSG.CountN + AYPSG.PeriodN
        End If

        AY_Left = AY_Left - AY_NextEvent
    Loop While (AY_Left > 0)

    If (AYPSG.Holding = 0) Then
        AYPSG.CountE = AYPSG.CountE - AY_STEP
        If (AYPSG.CountE <= 0) Then
            Do
                AYPSG.CountEnv = AYPSG.CountEnv - 1
                AYPSG.CountE = AYPSG.CountE + AYPSG.PeriodE
            Loop While (AYPSG.CountE <= 0)

            ' // check envelope current position
            If (AYPSG.CountEnv < 0) Then
                If (AYPSG.Hold) Then
                    If (AYPSG.Alternate) Then
                        AYPSG.Attack = AYPSG.Attack Xor MAXVOL '&h1f
                    End If
                    AYPSG.Holding = 1
                    AYPSG.CountEnv = 0
                Else
                    ' // if CountEnv has looped an odd number of times (usually 1),
                    ' // invert the output.
                    If (AYPSG.Alternate And ((AYPSG.CountEnv And &H20) = &H20)) Then
                        AYPSG.Attack = AYPSG.Attack Xor MAXVOL ' // &h1f
                    End If

                    AYPSG.CountEnv = AYPSG.CountEnv And MAXVOL ' // &h1f
                End If
            End If

            AYPSG.VolE = AYPSG.VolTable2(AYPSG.CountEnv Xor AYPSG.Attack)
            ' // reload volume
            If (AYPSG.EnvelopeA <> 0) Then AYPSG.VolA = AYPSG.VolE
            If (AYPSG.EnvelopeB <> 0) Then AYPSG.VolB = AYPSG.VolE
            If (AYPSG.EnvelopeC <> 0) Then AYPSG.VolC = AYPSG.VolE
        End If
    End If

    lOut1 = (VolA * AYPSG.VolA) \ 65535
    lOut2 = (VolB * AYPSG.VolB) \ 65535
    lOut3 = (VolC * AYPSG.VolC) \ 65535
    
    RenderByte = lOut1 + lOut2 + lOut3
End Function
