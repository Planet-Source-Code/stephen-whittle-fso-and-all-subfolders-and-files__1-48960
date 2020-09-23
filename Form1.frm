VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FFFF&
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose folder"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private getdir As String
Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, lParam As Any) As Long
Const LB_SETTABSTOPS = &H192



Private Sub Command1_Click()
getdir = BrowseForFolder(Me, "Select A Directory", "c:\")
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    
 List1.AddItem getdir
 
 ShowAllFiles getdir
 
End Sub

Private Sub ShowAllFiles(ByVal sPath As String)

    Dim fso As New FileSystemObject
    Dim fil As File
    Dim fol As Folder
    Dim sub1 As Folder
    
    Set fol = fso.GetFolder(sPath)
    
    List1.AddItem (vbTab & fol)
    
    For Each fil In fol.Files
       List1.AddItem (vbTab & vbTab & fil.Name)
    Next
    
    For Each sub1 In fol.SubFolders
        ShowAllFiles sub1.Path
    Next
    
    Set fil = Nothing
    Set sub1 = Nothing
    Set fol = Nothing
    Set fso = Nothing
    
End Sub



Private Sub Command2_Click()
If List1.ListCount = 0 Then
MsgBox "Nothing to clear", vbInformation, "Clear"
Else
List1.Clear
End If

End Sub


Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Dim LBTab(1) As Long
LBTab(0) = 30
LBTab(1) = 60
SendMessageArray List1.hWnd, LB_SETTABSTOPS, 2, LBTab(0)
End Sub


