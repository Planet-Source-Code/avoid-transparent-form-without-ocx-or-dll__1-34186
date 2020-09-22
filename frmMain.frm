VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   120
      ScaleHeight     =   2805
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "question, suggestion: stealth@vlz.ru"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   720
         MouseIcon       =   "frmMain.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2400
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simple Transparent Form without OCX or DLL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()
'mail to
Call MailTo("mailto:stealth@vlz.ru?Subject=About%20transparent%20form")
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
            'easy move
            Call ReleaseCapture
            Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        Case vbRightButton
            'unloading
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
'move picMainturebox into corner of window
    Me.picMain.Move 0, 0
End Sub


