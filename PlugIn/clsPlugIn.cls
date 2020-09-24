VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents CPlugIn As VBControlExtender
Attribute CPlugIn.VB_VarHelpID = -1

Public Sub ControlLoad(CTest As PictureBox, CName As String, CID As String)
On Error Resume Next
Set CPlugIn = CTest.Parent.Controls.Add(CName, CID, CTest)
With CPlugIn
    .Left = 30
    .Top = 30
    .Visible = True
    On Error Resume Next
    .Width = 900
    .Height = 600
End With
Exit Sub
err:
    MsgBox "Error : " & err.Description
End Sub

Public Sub PlugInWidth(Width As String)
With CPlugIn
On Error Resume Next
    .Width = Width * 15 - 90
End With
End Sub

Public Sub PlugInHeight(Height As String)
With CPlugIn
On Error Resume Next
    .Height = Height * 15 - 90
End With
End Sub

Public Sub PlugInVisible(Visible As Boolean)
With CPlugIn
On Error Resume Next
    .Visible = Visible
End With
End Sub



