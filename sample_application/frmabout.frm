VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Caption         =   "Priyan R"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmabout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "I think this example is useful to you . If you like this program please vote for me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   120
      Picture         =   "frmabout.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Label2_Click()
ShellExecute 0, "open", Label2.caption, "", "", 1

End Sub
