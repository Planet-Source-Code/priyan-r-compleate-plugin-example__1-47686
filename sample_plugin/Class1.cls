VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'----------------------------------------------
'please vote for me and visit my home page at www.priyan.tk
'----------------------------------------------

'--------------------------------------------------------
' This class contains objects that can be called from the application
' From the application call openplugin function with the form handle
' Once we got the form handle we can use all the public objects in that form
'--------------------------------------------------------
Public Property Get pluginname() As String 'Returns the name of the plugin
pluginname = "Priyan's test plugin"

End Property

Public Sub openplugin(Form As Object) 'Function that called from the application
Set Form1.formobj = Form      'Saves application's form handle to form1's public object formobj
Form1.Show 'show the plugin intreface
End Sub
'-------------------------------------------------------------
'            '****Important****
'The project name (project->options->project name) and compiled dll's name must be same
'eg: The this project name is testplugin so the compailed dll must be testplugin.dll
'
'It is because in the appication we call the dll name and .plugin in the createobject method
'-------------------------------------------------------------
