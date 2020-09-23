VERSION 5.00
Begin VB.Form frmDisclaimer 
   Caption         =   "Disclaimer..."
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmDisclaimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDisclaimer.frx":0442
   MousePointer    =   4  'Icon
   Moveable        =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OLE OLE1 
      Class           =   "WordPad.Document.1"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   360
      OleObjectBlob   =   "frmDisclaimer.frx":0884
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
frmDisclaimer.Visible = False
frmAbout.Visible = False
frmFormat.Visible = True

End Sub

Private Sub Form_Load()
frmAbout.Visible = False
frmFormat.Visible = True


End Sub

