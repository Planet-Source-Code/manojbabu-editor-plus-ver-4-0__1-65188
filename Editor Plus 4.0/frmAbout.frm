VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Editor Plus 4.0"
   ClientHeight    =   3255
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2246.659
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.761
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1800
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   4665
      TabIndex        =   6
      Top             =   720
      Width           =   4695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7000
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   4695
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@lconsoft"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editor Plus"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   1785
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   -120
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.02.2006"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c)  2006 , AlconSoft Allright Reserved."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID : AS-EP-VER-2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Website  : http://www.geocities.com/alconsoft/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Sub Form_Load()
    SetWindowPos hwnd, _
        HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
    Label1.Caption = "Copyright (c)  2006 , AlconSoft Allright Reserved." & _
    vbCrLf & "Version 4.05.2006" & _
    vbCrLf & "Product ID : AS-EP-VER-4.0" & vbCrLf & vbCrLf & "Created By" & _
    vbCrLf & "MANOJBABU" & _
    vbCrLf & vbCrLf & "Editor Plus 4.0" & _
    vbCrLf & "This software is freeware or shareware.  You may use it" & _
    vbCrLf & "free of charge as it fits your own projects but may not" & _
    vbCrLf & "sell the original product or the source code.  If you " & _
    vbCrLf & "decide to distribute these files you must include this" & _
    vbCrLf & "disclaimer and all the original copyright notices from the" & _
    vbCrLf & "original owner. Alconsoft take no responsibility of how" & _
    vbCrLf & "you use and modify this program." & _
    vbCrLf & _
    vbCrLf & vbCrLf & "Features Include" & _
    vbCrLf & "Create RTF,TXT,DOC,etc.. Files,Encrypt & Decrypt,Advanced Find," & _
    vbCrLf & "Remove and Replace,Insert Current System Date & Time" & _
    vbCrLf & "Embeded Files,Text Formating (Bold,Italic,Underline," & _
    vbCrLf & "Strike,Bullets,Alignmets,Fore Color,Back Color,Fonts and Etc..)" & _
    vbCrLf & "Open Other Applications (MS-Paint , MS-Calculator)." & _
    vbCrLf & vbCrLf & "For more information Please visit:" & _
    vbCrLf & "http://geocities.com/alconsoft/legal.html" & _
    vbCrLf & "E-mail: manojabu@sancharnet.in" & _
    vbCrLf & _
    vbCrLf & vbCrLf & "--- Thank you for using this Software ---"
End Sub

Private Sub Timer1_Timer()
    Label1.Top = Label1.Top - 5
    If Label1.Top < Picture1.Top - Label1.Height Then Label1.Top = 735
End Sub

Private Sub Label22_Click(Index As Integer)
Shell "rundll32.exe url.dll,FileProtocolHandler http://geocities.com/alconsoft/", 3
End Sub


