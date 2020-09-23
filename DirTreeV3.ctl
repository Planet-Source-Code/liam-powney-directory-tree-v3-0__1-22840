VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl DirTreeV3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   ScaleHeight     =   3810
   ScaleWidth      =   5955
   Begin ComctlLib.TreeView tvwDirectory 
      Height          =   3810
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6720
      _Version        =   327682
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilsTVWDirectory"
      Appearance      =   1
   End
   Begin VB.FileListBox filFiles 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DirListBox DirFolders 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DriveListBox driDrives 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label DirPath 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin ComctlLib.ImageList ilsTVWDirectory 
      Left            =   3960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DirTreeV3.ctx":09A2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DirTreeV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'** Directory Tree Control **
'****************************
'** Shows A Tree View of   **
'** Drives and Folders on  **
'** the Current Computer   **
'****************************


Option Explicit
Public FullPath As String



Private Sub DirFolders_Change()

    filFiles.Path = DirFolders.Path

End Sub


Private Sub driDrives_Change()

    
    DirFolders.Path = driDrives '.Path

End Sub

Private Sub tvwDirectory_NodeClick(ByVal Node As ComctlLib.Node)

    FullPath = tvwDirectory.SelectedItem.FullPath
    
End Sub



Private Sub UserControl_Initialize()
    
    Call BuildDriveList
    
End Sub

Private Sub BuildDriveList()

    Dim i As Integer
    Dim strPath As String
    Dim intIcon As Integer

    tvwDirectory.Nodes.Clear
    
    For i = 0 To driDrives.ListCount - 1
    
        strPath = UCase(Left(driDrives.List(i), 1)) & ":\"
        
        Select Case strPath
        
            Case "A:\", "B:\" ' Diskette drive.
                intIcon = 1
                
            Case "D:\"
                intIcon = 3     ' CD drive.
                
            Case Else           ' Hard drive.
                intIcon = 2
        
        End Select
        
        tvwDirectory.Nodes.Add , , strPath, driDrives.List(i), intIcon
        tvwDirectory.Nodes.Add strPath, tvwChild, ""
            
    Next

End Sub

Private Sub tvwDirectory_Expand(ByVal Node As ComctlLib.Node)

    On Error GoTo ErrorTrapping
    
    Dim i As Integer
    Dim strRelative As String
    Dim strFolderName As String
    Dim intFolderPos As Integer
    Dim intIcon As Integer
    Dim strNewPath As String
    Dim strExt As String
    Dim intExtPos As Integer
        
    MousePointer = vbHourglass
        
    If Node.Child.Text = "" Then
                
        tvwDirectory.Nodes.Remove Node.Child.Index
        strRelative = Node.Key
        DirFolders.Path = strRelative
        intFolderPos = Len(strRelative) + 1
                
        ' Add folders
        For i = 0 To DirFolders.ListCount - 1
        
            strFolderName = Mid(DirFolders.List(i), intFolderPos)
            
            strNewPath = strRelative & strFolderName & "\"
            tvwDirectory.Nodes.Add strRelative, tvwChild, strNewPath, strFolderName, 4
            
            DirFolders.Path = strNewPath
            
            If (filFiles.ListCount > 0) Or (DirFolders.ListCount > 0) Then
            
                tvwDirectory.Nodes.Add strNewPath, tvwChild, , ""
                tvwDirectory.Nodes(strNewPath).ExpandedImage = 5
                            
            End If
            
            DirFolders.Path = strRelative
                        
        Next
        
        
    End If
    
    GoTo EndSub
    
ErrorTrapping:
    ' An error occurs when you try reading on a not ready drive
    
    ' re-add the precedent removed item
    tvwDirectory.Nodes.Add Node.Key, tvwChild, , ""
    Resume EndSub
    
EndSub:
    MousePointer = vbDefault
End Sub

