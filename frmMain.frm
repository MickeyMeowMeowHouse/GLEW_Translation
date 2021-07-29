VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   13680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picBottomPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   0
      ScaleHeight     =   2040
      ScaleWidth      =   13680
      TabIndex        =   1
      Top             =   7065
      Width           =   13680
      Begin VB.TextBox txtErrReport 
         Height          =   1575
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox picLPanel 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7065
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin MSComctlLib.TreeView tvExtension 
         Height          =   2295
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4048
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MIT License
'
'Copyright (c) 2020 0xAA55-Official-Org
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Option Explicit

Public HeaderParse As New clsParser

Private Sub Form_Load()
Dim NoGUI As Boolean
If Command = "nogui" Then NoGUI = True
If NoGUI = False Then Show

'Parse the header files immediately
HeaderParse.ParseHeaderFile App.Path & "\glew.h", "glew"
HeaderParse.ParseHeaderFile App.Path & "\wglew.h", "wglew"

'Export the parsed C code to VB.NET
ExportVB_NET HeaderParse, App.Path
ExportVB_NET2 HeaderParse, App.Path
ExportCSharp HeaderParse, App.Path

If NoGUI Then
    Unload Me
    Exit Sub
End If

'Show in the TreeView
Dim ExtensionString
Dim Extension As clsGLExtension

tvExtension.Nodes.Clear
tvExtension.Nodes.Add , , "G", "<Total>"
tvExtension.Nodes.Add "G", tvwChild, "M", "MacroDefs"
tvExtension.Nodes.Add "G", tvwChild, "T", "TypeDefs"

For Each ExtensionString In HeaderParse.GLExtension.Keys
    Set Extension = HeaderParse.GLExtension(ExtensionString)
    AddExtensionToNode Extension.ExtensionString
Next

'If there's any errors, also show them
Dim ErrKeys
For Each ErrKeys In HeaderParse.ErrorReport.Keys
    txtErrReport.SelText = HeaderParse.ErrorReport(ErrKeys) & vbCrLf
Next
End Sub

Private Sub AddFuncToNode(IsAPI As Boolean, ExtensionString As String, NodeKey As String)
Dim DictKey
Dim I As Long
Dim FuncData() As String
Dim Param() As String
Dim CurExt As clsGLExtension
Set CurExt = HeaderParse.GLExtension(ExtensionString)
Dim FuncDataDict As Dictionary
If IsAPI Then
    Set FuncDataDict = CurExt.APIs
Else
    Set FuncDataDict = CurExt.FuncTypeDef
End If

For Each DictKey In FuncDataDict.Keys
    FuncData = Split(FuncDataDict(DictKey), ":")
    tvExtension.Nodes.Add ExtensionString & ":" & NodeKey, tvwChild, ExtensionString & ":" & DictKey, DictKey
    If IsAPI = False And CurExt.FuncPtrs.Exists(DictKey) Then
        tvExtension.Nodes.Add ExtensionString & ":" & DictKey, tvwChild, ExtensionString & ":" & DictKey & ":N", "Name: " & CurExt.FuncPtrs(DictKey)
    End If
    tvExtension.Nodes.Add ExtensionString & ":" & DictKey, tvwChild, ExtensionString & ":" & DictKey & ":R", "Returns: " & FuncData(0)
    Param = Split(FuncData(1), ",")
    tvExtension.Nodes.Add ExtensionString & ":" & DictKey, tvwChild, ExtensionString & ":" & DictKey & ":P", "Parameters: (" & CStr(UBound(Param) + 1) & ")"
    For I = 0 To UBound(Param)
        tvExtension.Nodes.Add ExtensionString & ":" & DictKey & ":P", tvwChild, ExtensionString & ":" & DictKey & ":P" & I, Param(I)
    Next
Next
End Sub

Private Sub AddExtensionToNode(ExtensionString As String)
On Error Resume Next
Dim Extension As clsGLExtension
Set Extension = HeaderParse.GLExtension(ExtensionString)

tvExtension.Nodes.Add , , ExtensionString, ExtensionString
tvExtension.Nodes.Add ExtensionString, tvwChild, ExtensionString & ":M", "MacroDefs"
tvExtension.Nodes.Add ExtensionString, tvwChild, ExtensionString & ":T", "TypeDefs"
tvExtension.Nodes.Add ExtensionString, tvwChild, ExtensionString & ":A", "APIs"
tvExtension.Nodes.Add ExtensionString, tvwChild, ExtensionString & ":F", "FuncTypeDef"

Dim DictKey

For Each DictKey In Extension.MacroDefs.Keys
    tvExtension.Nodes.Add ExtensionString & ":M", tvwChild, ExtensionString & ":" & DictKey, DictKey & " => " & Extension.MacroDefs(DictKey)
    tvExtension.Nodes.Add "M", tvwChild, "Total:" & DictKey, DictKey & " => " & Extension.MacroDefs(DictKey)
Next

For Each DictKey In Extension.TypeDefs.Keys
    tvExtension.Nodes.Add ExtensionString & ":T", tvwChild, ExtensionString & ":" & DictKey, DictKey & " => " & Extension.TypeDefs(DictKey)
    tvExtension.Nodes.Add "T", tvwChild, "Total:" & DictKey, DictKey & " => " & Extension.TypeDefs(DictKey)
Next

AddFuncToNode True, ExtensionString, "A"
AddFuncToNode False, ExtensionString, "F"
End Sub

Private Sub Form_Resize()
On Error Resume Next
picLPanel.Width = ScaleWidth
End Sub

Private Sub picBottomPanel_Resize()
txtErrReport.Move 0, 0, picBottomPanel.ScaleWidth, picBottomPanel.ScaleHeight
End Sub

Private Sub picLPanel_Resize()
tvExtension.Move 0, 0, picLPanel.ScaleWidth, picLPanel.ScaleHeight
End Sub

Private Sub tvExtension_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub tvExtension_DblClick()
txtErrReport.SelText = tvExtension.SelectedItem.Text & vbCrLf
End Sub
