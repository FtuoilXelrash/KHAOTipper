VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Timers"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   LinkTopic       =   "Timers"
   ScaleHeight     =   1140
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerTipCows 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer LogOffDeath 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LogOffDeath_Timer()

Timers.LogOffDeath.Enabled = False
hook.achooks.Logout

End Sub

Public Sub TimerTipCows_Timer()

If hook.achooks.IsValidObject(CowGUID) Then
hook.achooks.UseItem CowGUID, 0
Else
    Timers.TimerTipCows.Enabled = False
    control.btnStartTips.Text = "START"
    control.btnStartTips.TextColor = "65280"
    Tippin_OK = False
    WriteToChat "KHAO TIPPING STOPPED!! THE COW HAS DISSAPEARED!", 10
End If

End Sub
