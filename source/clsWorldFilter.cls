VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorldFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents wfilter As DecalFilters.World
Attribute wfilter.VB_VarHelpID = -1

Public Function NameFromGuid(ByVal ObjectGUID As Long) As String
On Error GoTo err_dead

    Dim wObj As WorldObject
    Set wObj = wfilter(ObjectGUID)

    NameFromGuid = wObj.Strings(keyName)

    Exit Function
err_dead:
    Hub.SendText "NameFromGuid: " & Err.Description, 1
End Function

