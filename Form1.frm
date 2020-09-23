VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cool Form close"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Cool Close"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'The Matrix
'Qais Ghalib
'Qais60@hotmail.com
'Thanks 4 Download code

Private Sub Command1_Click()
'You Can Chang speed
Call coolClose(Me, 6)
End Sub

Public Function coolClose(FormClose As Form, speed As Integer)
Do Until FormClose.Height <= 405
DoEvents
FormClose.Height = FormClose.Height - speed * 9
FormClose.Top = FormClose.Top + speed * 5
Loop
Do Until FormClose.Width <= 1680
DoEvents
FormClose.Width = FormClose.Width - speed * 9
FormClose.Left = FormClose.Left + speed * 5
Loop
Unload FormClose
End Function
