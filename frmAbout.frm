VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Format Drive A: Ver. 5.1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.OLE OLE1 
      Class           =   "WordPad.Document.1"
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmAbout.frx":0442
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmAbout.Visible = False
frmFormat.Visible = True

End Sub

Private Sub Document1_GotFocus()

End Sub

