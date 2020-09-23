VERSION 5.00
Begin VB.Form frmFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format Drive A: Ver. 5.1"
   ClientHeight    =   3060
   ClientLeft      =   30
   ClientTop       =   630
   ClientWidth     =   5670
   Icon            =   "frmFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Full Fo&rmat"
         Height          =   372
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1332
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Q&uit"
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Quick Format"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.OLE OLE1 
         Class           =   "WordPad.Document.1"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   1920
         OleObjectBlob   =   "frmFormat.frx":0442
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number After Format:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number Before Format: "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   855
         Left            =   120
         Top             =   2040
         Width           =   3675
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu format 
         Caption         =   "Q&uickFormat"
      End
      Begin VB.Menu Fullformat1 
         Caption         =   "Full F&ormat"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu disclaimer 
         Caption         =   "&Disclaimer"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Format Drive A: Ver. 5.1
'Copyright Matt Cheche
'Matt Cheche Productions 2002
Dim fso As New FileSystemObject
Dim dr As Drive
Dim AppPath As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub About_Click()
frmAbout.Visible = True
frmFormat.Visible = True
End Sub

Private Sub Command1_Click()

    FullFormatDriveA
  If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
End If
  DoEvents
    If dr.IsReady Then
        Sleep (15000)
        Label2.Caption = "Serial Number After Format: " & dr.SerialNumber
    End If
    If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
 End If
    
End Sub

Public Sub WriteResponseFile()
    Dim F As Integer
    Dim resp As String
    resp = AppPath & "resp.txt"
    
    F = FreeFile
    Open resp For Output As #F
        Print #F, ""
        Print #F, ""
        Print #F, "n"
        
    Close #F
    
End Sub

Public Sub FullFormatDriveA()
    Set dr = fso.GetDrive("A:")
    If dr.IsReady Then
        WriteResponseFile
        Dim F As Integer
        Dim batch As String
        batch = AppPath & "qf_a.bat"
        F = FreeFile
        
        Open batch For Output As #F
            Print #F, "Format A: /U < resp.txt"
            Print #F, "del resp.txt"
            Print #F, "del %0"
        Close #F
        Shell batch, vbHide
    Else
        MsgBox "Drive A: has no floppy in it!", vbOKOnly
    
    End If
End Sub

Private Sub Command2_Click()
    End
        
End Sub

Private Sub Command3_Click()
    QuickFormatDriveA
      If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
End If
      DoEvents
    If dr.IsReady Then
        Sleep (15000)
        Label2.Caption = "Serial Number After Format: " & dr.SerialNumber
    End If
    If Command3.Enabled = True Then
        Command3.Enabled = False
        Command1.Enabled = False
        format.Enabled = False
        Fullformat1.Enabled = False
        
        End If
        
    
End Sub

Private Sub disclaimer_Click()
frmDisclaimer.Visible = True
frmFormat.Visible = True

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()

    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then
        AppPath = AppPath & "\"
    End If
    
    Set dr = fso.GetDrive("a:")
    If dr.IsReady Then
        frmFormat.Visible = True
   
          End If
           
    If dr.IsReady Then
        Label1.Caption = Label1.Caption & dr.SerialNumber
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmFormat = Nothing
    
    
End Sub

Private Sub QuickFormatDriveA()
  Set dr = fso.GetDrive("A:")
    If dr.IsReady Then
        WriteResponseFile
        Dim F As Integer
        Dim batch As String
        batch = AppPath & "qf_a.bat"
        F = FreeFile
        
        Open batch For Output As #F
            Print #F, "Format A: /Q < resp.txt"
            Print #F, "del resp.txt"
            Print #F, "del %0"
        Close #F
        Shell batch, vbHide
    Else
        MsgBox "Drive A: has no floppy in it!", vbOKOnly
       
      End If
        
End Sub
Private Sub format_Click()
    
    If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
End If
DoEvents
    If dr.IsReady Then
        Sleep (15000)
        Label2.Caption = "Serial Number After Format: " & dr.SerialNumber
    End If
  If dr.IsReady Then
        WriteResponseFile
        Dim F As Integer
        Dim batch As String
        batch = AppPath & "qf_a.bat"
        F = FreeFile
        
        Open batch For Output As #F
            Print #F, "Format A: /Q < resp.txt"
            Print #F, "del resp.txt"
            Print #F, "del %0"
        Close #F
        Shell batch, vbHide
    Else
        MsgBox "Drive A: has no floppy in it!", vbOKOnly
    End If
If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
End If


End Sub

Private Sub Fullformat1_Click()
    If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
        End If
    DoEvents
    If dr.IsReady Then
        Sleep (15000)
        Label2.Caption = "Serial Number After Format: " & dr.SerialNumber
    End If
  If dr.IsReady Then
        WriteResponseFile
        Dim F As Integer
        Dim batch As String
        batch = AppPath & "qf_a.bat"
        F = FreeFile
        
        Open batch For Output As #F
            Print #F, "Format A: /U < resp.txt"
            Print #F, "del resp.txt"
            Print #F, "del %0"
        Close #F
        Shell batch, vbHide
    Else
        MsgBox "Drive A: has no floppy in it!", vbOKOnly
    End If
If Command1.Enabled = True Then
        Command1.Enabled = False
        Command3.Enabled = False
         format.Enabled = False
        Fullformat1.Enabled = False
        End If
        
End Sub

