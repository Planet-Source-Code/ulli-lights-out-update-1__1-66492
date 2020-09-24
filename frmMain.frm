VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   1020
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   930
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   930
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.Menu mnuTray 
      Caption         =   "*"
      Visible         =   0   'False
      Begin VB.Menu mnuLo 
         Caption         =   "Lower Brightness"
      End
      Begin VB.Menu mnuHi 
         Caption         =   "Restore Brightness"
      End
      Begin VB.Menu mnuTimeout 
         Caption         =   "Set Timeout"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Menu"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Found at PSC a few days ago and havily modified - looked for it again just now to give proper credit
'to the original author but can't find it no more. anyway, tnx for showing how it goes.

Private WithEvents Systray As clsSystray
Attribute Systray.VB_VarHelpID = -1

Private Sub AdjTimeout(ByVal Value As Long)

    Select Case Value
      Case Is > 15
        Value = 15
      Case Is < 0
        Value = 0
    End Select
    TimeOut = Value
    mnuTimeout.Caption = "Set Timeout (currently " & IIf(TimeOut = 0, "disabled", TimeOut & " minute" & IIf(TimeOut = 1, vbNullString, "s")) & ")"
    IdleStopDetection
    If TimeOut Then
        IdleBeginDetection TimeOut
    End If

End Sub

Private Sub Form_Load()

  Dim t     As Date

    GetDeviceGammaRamp hDC, RampHi(1)
    Set Systray = New clsSystray
    If App.PrevInstance Then
        Beeper 440, 30
        Unload Me
      Else 'APP.PREVINSTANCE = FALSE/0
        AdjTimeout 2
        With Systray
            .SetOwner Me
            .AddIconToTray Icon.Handle, , True
            .Tooltip = "Monitor Gamma Control" & vbCrLf & vbCrLf & "Right Click For Menu..."
            .ShowBalloon "I am here...", , InfoIcon Or SoundOff
            t = Now
            Do
                DoEvents
            Loop Until Now > t + 0.00004
            .HideBalloon
        End With 'SYSTRAY
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    SetDeviceGammaRamp hDC, RampHi(1)
    IdleStopDetection
    With Systray
        If .IsIconInTray Then
            .RemoveIconFromTray
        End If
    End With 'SYSTRAY

End Sub

Private Function Int2Lng(IntVal As Integer) As Long

    CopyMemory Int2Lng, IntVal, 2

End Function

Private Function Lng2Int(LngVal As Long) As Integer

    CopyMemory Lng2Int, LngVal, 2

End Function

Private Sub mnuExit_Click()

    Unload Me

End Sub

Public Sub mnuHi_Click()

  Dim Div As Single

    If Reason Then
        For Div = 2 To 1 Step -0.05
            ModRamp Div
        Next Div
    End If
    Reason = 0

End Sub

Public Sub mnuLo_Click()

  Dim Div As Single

    If Reason = 0 Then
        For Div = 1 To 2 Step 0.005
            ModRamp Div
        Next Div
    End If
    Reason = 1

End Sub

Private Sub mnuTimeout_Click()

  Dim sTO As String

    sTO = InputBox("Enter timeout value in minutes for automatic reduction of brightness (zero to disable) :", App.ProductName, TimeOut, Screen.Width - 6000, Screen.Height - 2700)
    Select Case True
      Case StrPtr(sTO) = 0
        'do nothing
      Case CBool(Len(sTO))
        AdjTimeout Val(sTO)
    End Select

End Sub

Private Sub ModRamp(Div As Single)

  Dim i As Long

    For i = 1 To 768
        RampLo(i) = Lng2Int(Int2Lng(RampHi(i)) / Div)
    Next i
    SetDeviceGammaRamp hDC, RampLo(1)
    DoEvents

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Sep-08 10:32)  Decl: 6  Code: 128  Total: 134 Lines
':) CommentOnly: 3 (2,2%)  Commented: 3 (2,2%)  Empty: 37 (27,6%)  Max Logic Depth: 4
