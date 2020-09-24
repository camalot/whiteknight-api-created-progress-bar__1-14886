VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "email - witenite87@excite.com"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Progress Bar Created by Win32 API"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Visit camalot.virtualave.net"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Progress 32 v1 by whiteknight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pg As New CDProgress
Dim x As Integer
Private Sub Form_Load()
Set pg = New CDProgress
Timer1.Enabled = True
pg.Create
pg.Min = 1
pg.Max = 100
pg.BorderType = PBB_Raised
pg.OwnerhWnd = Me.hWnd
pg.ProgressStyle = PBS_SmoothBar
pg.Top = Label1.Top + Label1.Height
pg.Left = 1
pg.Width = Me.ScaleWidth - 2
pg.Height = Me.ScaleHeight - (Label1.Top + Label1.Height)
'pg.BarColor = vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
pg.Destroy
End Sub

Private Sub Timer1_Timer()
If x >= pg.Max Then x = 1
x = x + 1
pg.Value = x
End Sub
