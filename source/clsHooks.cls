VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents achooks As Decal.achooks
Attribute achooks.VB_VarHelpID = -1


Private Sub AChooks_StatusTextIntercept(ByVal bstrText As String, bEat As Boolean)
On Error GoTo err_dead
     
'   LOOKING FOR THE DAMN RED TEXT CRAP
            
    If control.chkEatRedMessages.Checked = True Then
'    WriteToChat "RED TEXT SPAM: " & bstrText, 2
    bEat = True
    End If
       
    Exit Sub
    
err_dead:
    Hub.SendText "err_StatusTextIntercept: " & Err.Description, 1
End Sub

Private Sub AChooks_ChatTextIntercept(ByVal bstrText As String, _
ByVal lColor As Long, ByVal lTarget As Long, bEat As Boolean)

On Error GoTo err_dead


    If lColor = 0 Then
    
'   --50    Tips
    If (InStr(1, bstrText, "The cows lips do not move but you would almost swear you just heard", vbTextCompare) <> 0) Then
    WriteToChat "50 Tips Completed!", 8
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "moomoo.wav")
    End If
    End If
    
'   --100   Tips
    If (InStr(1, bstrText, "Maybe you can ask Dwennon about it some day", vbTextCompare) <> 0) Then
    WriteToChat "100 Tips Completed!", 8
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "moomoo.wav")
    End If
    End If
    
'   --150  Tips
    If (InStr(1, bstrText, "I must be at least half done", vbTextCompare) <> 0) Then
    WriteToChat "150 Tips Completed!", 8
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "moomoo.wav")
    End If
    End If

'   --200 Tips
    If (InStr(1, bstrText, "I must have tipped at least 200 times by now", vbTextCompare) <> 0) Then
    WriteToChat "200 Tips Completed!", 8
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "moomoo.wav")
    End If
    End If
    
'   --250  Tips
    If (InStr(1, bstrText, "you believe that's your 250th tip", vbTextCompare) <> 0) Then
    WriteToChat "250 Tips Completed!", 8
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "moomoo.wav")
    End If
    End If
    
'   --300 Tips
    If (InStr(1, bstrText, "seems to you like you have done enough tipping for now", vbTextCompare) <> 0) Then
    
    Timers.TimerTipCows.Enabled = False
    control.btnStartTips.Text = "START"
    control.btnStartTips.TextColor = "65280"
    Tippin_OK = False
    WriteToChat "300 Tips Completed!", 8
    
    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "done.wav")
    End If
    
    End If

    End If

    Exit Sub
    
err_dead:
    Hub.SendText "err_AChooks_ChatTextIntercept: " & Err.Description, 1
End Sub

