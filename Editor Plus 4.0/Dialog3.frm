VERSION 5.00
Begin VB.Form Dialog3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabs"
   ClientHeight    =   2445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3255
   Icon            =   "Dialog3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear &All"
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      ItemData        =   "Dialog3.frx":014A
      Left            =   120
      List            =   "Dialog3.frx":014C
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tabs :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Dialog3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Option Base 1   ' set the default lower array limit

Dim mblnLoading As Boolean

Private Sub cmdClear_Click()
    List1.RemoveItem List1.ListIndex
    cmdClear.Enabled = False
    cmdSet.Enabled = True
    cmdSet.SetFocus
End Sub

Private Sub cmdClearAll_Click()
    List1.Clear ' clear them
    Text1.Text = "" 'clear the tab text
    Text1.SetFocus 'set focus to the new tab box
End Sub

Private Sub cmdSet_Click()
    List1.AddItem Text1.Text
    cmdClear.Enabled = True
    cmdSet.Enabled = False
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
            mblnLoading = True
End Sub

Private Sub Command1_Click()
    Dim X As Integer
    If List1.Text <> "" Then
        With Form1.RichTextBox1
            .SelTabCount = List1.Text
            For X = 0 To .SelTabCount - 1
                .SelTabs(X) = List1.Text * X
            Next X
        End With
    Unload Me
    Else
        Me.Visible = False
        MsgBox "Please select a tab.", vbInformation, "Alconsoft - Editor Plus"
        Me.Visible = True
    End If
End Sub
Private Sub Form_Paint()
    If mblnLoading = True Then
        Dim intI As Integer
            For intI = 0 To Form1.RichTextBox1.SelTabCount - 1
                Dim sglTabValue As Single
                sglTabValue = Form1.RichTextBox1.SelTabs(intI) / 1440#
                sglTabValue = CInt(sglTabValue * 100) / 100#
                List1.AddItem Str(sglTabValue)
            Next intI
        mblnLoading = False
    End If
End Sub

Private Sub Text1_Change()
    cmdSet.Enabled = True
End Sub
