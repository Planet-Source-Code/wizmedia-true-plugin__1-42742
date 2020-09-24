VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PlugIn"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox Picture1 
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   4635
         TabIndex        =   5
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "ThePlugIn"
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "PlugIn.ThePlugIn"
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Unload"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Load"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Object Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OCX Name . UsercontrolName"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PlugIn As New clsPlugIn

Private Sub Command1_Click()
PlugIn.ControlLoad Picture1, Text1.Text, Text2.Text
End Sub

Private Sub Command2_Click()
PlugIn.PlugInVisible False
End Sub

Private Sub Form_Load()
MsgBox "Please Compile 'PlugIn.vbp' To 'PlugIn.ocx'", vbInformation, "WizMedia"
End Sub
