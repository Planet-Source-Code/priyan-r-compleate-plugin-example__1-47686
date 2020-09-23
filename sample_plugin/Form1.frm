VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's test plugin"
   ClientHeight    =   4350
   ClientLeft      =   5655
   ClientTop       =   2325
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4710
   Begin VB.CommandButton Command6 
      Caption         =   "End"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Text"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get Text"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Text"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "http://www.Priyan.tk"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Caption"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "Type new caption here..."
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------
'please vote for me and visit my home page at www.priyan.tk
'----------------------------------------------
Public formobj As Form 'application forms handle
Option Explicit

Private Sub Command1_Click()
formobj.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
formobj.Text1.Text = formobj.Text1 + Text2.Text
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Hide" Then
    formobj.Hide
    Command3.Caption = "Show"
Else
    formobj.Show
    Command3.Caption = "Hide"
End If
End Sub

Private Sub Command4_Click()
MsgBox formobj.Text1.Text, vbInformation, "From Plugin"
End Sub

Private Sub Command5_Click()
formobj.Text1.Text = ""
End Sub

Private Sub Command6_Click()
Unload formobj
End Sub

