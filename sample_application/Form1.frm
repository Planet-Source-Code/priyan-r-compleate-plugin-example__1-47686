VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's Plugin Sample"
   ClientHeight    =   4020
   ClientLeft      =   1590
   ClientTop       =   2055
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4890
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   4095
   End
   Begin VB.Menu mnuplugins 
      Caption         =   "Plugins"
      Begin VB.Menu mnupluginlist 
         Caption         =   "No Plugins"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
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
Option Explicit
Private Sub Form_Load()
loadplugins 'loads all the plugins in the pllugins directory
End Sub
Public Sub addpluginmenu(ByVal plugname$, ByVal caption$)
Dim mindex%
If mnupluginlist(0).Enabled = True Then
    mindex = mindex + 1
    Load mnupluginlist(mindex)
Else
    mnupluginlist(0).Enabled = True
End If
mnupluginlist(mindex).caption = caption
mnupluginlist(mindex).Tag = plugname
End Sub
Public Sub loadplugins()
Dim fld$, temp$, obj
fld = Dir(addstrap(App.Path, "plugins\*.dll"), vbNormal) 'gets the first dll in the plugins directory
Do Until fld = ""
 '   Shell "regsvr32 /" & addstrap(App.Path, "plugins\" & fld)  'register the plugin
    fld = Left(fld, Len(fld) - 4) 'removess the .dll from file name
    temp = fld & "." & "plugin"
Set obj = CreateObject(temp) 'creates the plugin object
    addpluginmenu temp, obj.pluginname 'Adds to the plugin menu
    fld = Dir()
Loop
End Sub
Public Function addstrap(ByVal path1 As String, ByVal path2 As String) As String
If Right$(path1, 1) = "\" Then
     addstrap = path1 & path2
Else
         addstrap = path1 & "\" & path2
End If
End Function

Private Sub mnuabout_Click()
frmabout.Show vbModal
End Sub

Private Sub mnupluginlist_Click(Index As Integer)
Dim obj
Set obj = CreateObject(mnupluginlist(Index).Tag)
obj.openplugin Me
End Sub

