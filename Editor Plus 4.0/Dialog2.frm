VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove & Replace"
   ClientHeight    =   1020
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   Icon            =   "Dialog2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUnwanted 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "~`!@#$%^&*()_+-=|\?/.>,<'"""
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtRepl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Remove"
      Height          =   300
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "Remove and Replace  unwanted charecters"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Unwanted characters"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Replace with"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub OKButton_Click()
    Form1.RichTextBox1.Text = ReplaceCharacters(Form1.RichTextBox1.Text, txtUnwanted.Text, txtRepl.Text)
End Sub
Public Function ReplaceCharacters(ByRef strText As String, ByRef strUnwanted As String, ByRef strRepl As String) As String
Dim i As Integer
Dim ch As String

    For i = 1 To Len(strUnwanted)
        Form1.RichTextBox1.Text = Replace(Form1.RichTextBox1.Text, Mid$(strUnwanted, i, 1), strRepl)
    Next

    ReplaceCharacters = Form1.RichTextBox1.Text
End Function
