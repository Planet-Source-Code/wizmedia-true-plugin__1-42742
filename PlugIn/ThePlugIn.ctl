VERSION 5.00
Begin VB.UserControl ThePlugIn 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton Command1 
      Caption         =   "MsgBox"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "ThePlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub Command1_Click()
MsgBox "Hello From PlugIn", vbOKOnly, "WizMedia"
End Sub

Private Sub UserControl_Resize()
Command1.Width = UserControl.ScaleWidth
Command1.Height = UserControl.ScaleHeight
End Sub
