VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGLExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HeaderFilePath As String
Public ExtensionString As String
Public Prefix As String 'GLEW, WGLEW
Public StartLine As Long
Public EndLine As Long
Public TypeDefs As New Dictionary
Public MacroDefs As New Dictionary
Public API_DllName As String
Public APIs As New Dictionary
Public FuncTypeDef As New Dictionary
Public FuncPtrs As New Dictionary

Private Sub Class_Initialize()
API_DllName = "opengl32.dll"
End Sub
