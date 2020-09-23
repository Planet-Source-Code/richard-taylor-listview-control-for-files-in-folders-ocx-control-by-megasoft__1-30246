VERSION 5.00
Object = "{E6330B58-FE11-11D5-B02C-00A0242DDE13}#6.0#0"; "MegaSoft_Explorer.ocx"
Begin VB.Form frmmain 
   Caption         =   "MegaSoft Explorer Example"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin MegaSoft_Explorer.Explore Explore 
      Height          =   4575
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8070
   End
   Begin VB.DirListBox Dir1 
      Height          =   4590
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Explore.Path Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub
