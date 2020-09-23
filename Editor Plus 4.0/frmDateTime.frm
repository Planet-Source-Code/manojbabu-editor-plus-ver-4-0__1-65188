VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date & Time"
   ClientHeight    =   2745
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4215
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      WhatsThisHelpID =   3
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Available Formats :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
     SetWindowPos hwnd, _
        HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
    List1.AddItem Format$(Date, "d/m/yy")
    List1.AddItem Format$(Date, "d/m/yyyy")
    List1.AddItem Format$(Date, "dd/mm/yy")
    List1.AddItem Format$(Date, "dd/mm/yyyy")
    
    List1.AddItem Format$(Date, "dd-mmm-yy")
    List1.AddItem Format$(Date, "dd-mmm-yyyy")
    
    List1.AddItem Format$(Date, "dd-mmmm-yy")
    List1.AddItem Format$(Date, "dd-mmmm-yyyy")
    
    List1.AddItem Format$(Date, "ddd, mmm dd , yy")
    List1.AddItem Format$(Date, "ddd, mmmm dd, yyyy")
    List1.AddItem Format$(Date, "dddd, mmmm dd, yyyy")
    
    List1.AddItem Format$(Date, "dd, mmmm, yyyy")
    
    List1.AddItem Format$(Time, "h:m:s AM/PM")
    List1.AddItem Format$(Time, "hh:mm:ss AM/PM")
    List1.AddItem Format$(Time, "hh:mm:ss")
End Sub

Private Sub List1_DblClick()
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    If List1.Text <> "" Then
    Form1.RichTextBox1.SelText = List1.Text
    Unload Me
    Else
        Me.Visible = False
        MsgBox "Please select a Date or Time format.", vbInformation, "Alconsoft - Editor Plus"
        Me.Visible = True
    End If
End Sub
