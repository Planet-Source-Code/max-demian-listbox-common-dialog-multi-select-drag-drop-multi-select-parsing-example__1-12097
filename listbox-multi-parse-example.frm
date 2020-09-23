VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listbox, Common Dialog/Multi Select, Drag Drop/Multi Select, Parsing Example"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7050
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      Height          =   2205
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   4800
      Width           =   7455
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   2520
      Width           =   7455
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag Files Onto One of The List Boxes"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo errrhandler
    Dim vFiles As Variant
    Dim lFile As Long
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select File(s)..."
    CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    CommonDialog1.Filter = "All files (*.*)|*.*"
    CommonDialog1.ShowOpen
    vFiles = Split(CommonDialog1.FileName, Chr(0))
    If UBound(vFiles) = 0 Then
        List1.AddItem CommonDialog1.FileName
        List2.AddItem CommonDialog1.FileTitle
    Else
        For lFile = 1 To UBound(vFiles)
            List1.AddItem Left(vFiles(0) + "\" & vFiles(lFile), InStrRev(vFiles(0) + "\" & vFiles(lFile), "\"))
            List2.AddItem vFiles(lFile)
            List3.AddItem vFiles(0) + "\" & vFiles(lFile)
        Next
    End If
    Exit Sub
errrhandler:
    Exit Sub
End Sub

Private Sub Command2_Click()
    List1.Clear
    List2.Clear
    List3.Clear
End Sub

Private Sub List1_Click()
    On Error Resume Next
    List2.ListIndex = List1.ListIndex
    List3.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
    On Error Resume Next
    List1.ListIndex = List2.ListIndex
    List3.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
    On Error Resume Next
    List1.ListIndex = List3.ListIndex
    List2.ListIndex = List3.ListIndex
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    For i = 1 To Data.Files.Count
    List1.AddItem Left(Data.Files(i), InStrRev(Data.Files(i), "\"))
    List2.AddItem Mid(Data.Files(i), InStrRev(Data.Files(i), "\") + 1)
    List3.AddItem Data.Files(i)
    Next i
End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    For i = 1 To Data.Files.Count
    List1.AddItem Left(Data.Files(i), InStrRev(Data.Files(i), "\"))
    List2.AddItem Mid(Data.Files(i), InStrRev(Data.Files(i), "\") + 1)
    List3.AddItem Data.Files(i)
    Next i
End Sub

Private Sub List3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    For i = 1 To Data.Files.Count
    List1.AddItem Left(Data.Files(i), InStrRev(Data.Files(i), "\"))
    List2.AddItem Mid(Data.Files(i), InStrRev(Data.Files(i), "\") + 1)
    List3.AddItem Data.Files(i)
    Next i
End Sub
