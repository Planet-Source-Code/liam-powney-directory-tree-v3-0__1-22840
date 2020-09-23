VERSION 5.00
Object = "{0400D0DC-DCE2-4635-8197-CED742D90A82}#2.0#0"; "DirectoryTreeV3.ocx"
Begin VB.Form frmMain 
   Caption         =   "Directory Control V3 Test Program"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin DirectoryTreeV3.DirTreeV3 DirTree 
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
   End
   Begin VB.Label lblPath 
      Caption         =   " "
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   4320
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdTest_Click()
    
    'This is the only line of code you need to
    'get output from the Directory Tree control
    
    lblPath.Caption = DirTree.FullPath
    
End Sub
