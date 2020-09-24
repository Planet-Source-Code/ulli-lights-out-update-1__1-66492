Attribute VB_Name = "modIdle"
Option Explicit


Public Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hDC As Long, lpv As Any) As Long
Public Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hDC As Long, lpv As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public RampHi(1 To 768) As Integer
Public RampLo(1 To 768) As Integer
Public TimeOut  As Long
Public Reason   As Long '1 manual; 2 by idle timer

Private Declare Function BeginIdleDetection Lib "Msidle.dll" Alias "#3" (ByVal pfnCallback As Long, ByVal dwIdleMin As Long, ByVal dwReserved As Long) As Long
Private Declare Function EndIdleDetection Lib "Msidle.dll" Alias "#4" (ByVal dwReserved As Long) As Long
Private Const USER_IDLE_BEGIN       As Long = 1
Private Const USER_IDLE_END         As Long = 2

Public Sub IdleBeginDetection(Optional ByVal IdleMinutes As Long = 5)

    BeginIdleDetection AddressOf IdleCallBack, IdleMinutes, 0&

End Sub

Private Sub IdleCallBack(ByVal dwState As Long)

    Select Case dwState
      Case USER_IDLE_BEGIN
        If Reason = 0 Then
            frmMain.mnuLo_Click
            Reason = 2
        End If
      Case USER_IDLE_END
        If Reason = 2 Then
            frmMain.mnuHi_Click
        End If
    End Select

End Sub

Public Sub IdleStopDetection()

    EndIdleDetection 0&

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Sep-07 21:22)  Decl: 18  Code: 32  Total: 50 Lines
':) CommentOnly: 0 (0%)  Commented: 1 (2%)  Empty: 13 (26%)  Max Logic Depth: 3
