VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HelloWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
Set MyScriptingContext = PassedScriptingContext
Set MyApplication = MyScriptingContext.Application
Set MyRequest = MyScriptingContext.Request
Set MyResponse = MyScriptingContext.Response
Set MyServer = MyScriptingContext.Server
Set MySession = MyScriptingContext.Session
End Sub
