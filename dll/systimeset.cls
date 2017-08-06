VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim MyApplication As Application
Dim MyRequest As Request
Dim MyResponse As Response
Dim MyServer As Server
Dim MySession As Session

Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
Set MyScriptingContext = PassedScriptingContext
Set MyApplication = MyScriptingContext.Application
Set MyRequest = MyScriptingContext.Request
Set MyResponse = MyScriptingContext.Response
Set MyServer = MyScriptingContext.Server
Set MySession = MyScriptingContext.Session
End Sub

Public Sub OnEndPage()
Set MyScriptingContext = Nothing
Set MyApplication = Nothing
Set MyRequest = Nothing
Set MyResponse = Nothing
Set MyServer = Nothing
Set MySession = Nothing
End Sub


Public Sub hellow()
     MyResponse.Write ("Hello World")
End Sub

Public Sub google()
     MyResponse.Write ("google!")
End Sub
