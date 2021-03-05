Attribute VB_Name = "modExporter"
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

'Merge two dictionaries
Function MergeDict(Dict1 As Dictionary, Dict2 As Dictionary) As Dictionary
Dim Ret As New Dictionary
Dim Key

For Each Key In Dict1.Keys
    Ret.Item(Key) = Dict1(Key)
Next

For Each Key In Dict2.Keys
    Ret.Item(Key) = Dict2(Key)
Next

Set MergeDict = Ret
End Function

'Do type conversions
Function DoTypeConv(TypeDesc As String) As String
Select Case TypeDesc
Case "GLenum"
    DoTypeConv = "UInt32"
Case "GLbitfield"
    DoTypeConv = "UInt32"
Case "GLuint"
    DoTypeConv = "UInt32"
Case "GLint"
    DoTypeConv = "Int32"
Case "GLsizei"
    DoTypeConv = "Int32"
Case "GLboolean"
    DoTypeConv = "Boolean"
Case "GLbyte"
    DoTypeConv = "SByte"
Case "GLshort"
    DoTypeConv = "Int16"
Case "GLubyte"
    DoTypeConv = "Byte"
Case "GLushort"
    DoTypeConv = "UInt16"
Case "GLulong"
    DoTypeConv = "UIntPtr"
Case "GLfloat"
    DoTypeConv = "Single"
Case "GLclampf"
    DoTypeConv = "Single"
Case "GLdouble"
    DoTypeConv = "Double"
Case "GLclampd"
    DoTypeConv = "Double"
Case "GLvoid"
    DoTypeConv = "IntPtr"
Case "GLint64EXT"
    DoTypeConv = "Int64"
Case "GLuint64EXT"
    DoTypeConv = "UInt64"
Case "GLint64"
    DoTypeConv = "Int64"
Case "GLuint64"
    DoTypeConv = "UInt64"
Case "GLsync"
    DoTypeConv = "IntPtr"
Case "GLchar"
    DoTypeConv = "Byte"
Case "GLintptr"
    DoTypeConv = "IntPtr"
Case "GLsizeiptr"
    DoTypeConv = "IntPtr"
Case "GLfixed"
    DoTypeConv = "Int32"
Case "cl_context"
    DoTypeConv = "IntPtr"
Case "cl_event"
    DoTypeConv = "IntPtr"
Case "GLcharARB"
    DoTypeConv = "Byte"
Case "GLhandleARB"
    DoTypeConv = "Int32"
Case "GLintptrARB"
    DoTypeConv = "IntPtr"
Case "GLsizeiptrARB"
    DoTypeConv = "IntPtr"
Case "GLeglClientBufferEXT"
    DoTypeConv = "UIntPtr"
Case "GLhalf"
    DoTypeConv = "UInt16"
Case "GLvdpauSurfaceNV"
    DoTypeConv = "IntPtr"
Case "GLclampx"
    DoTypeConv = "Int32"
Case "HPBUFFERARB"
    DoTypeConv = "IntPtr"
Case "HPBUFFEREXT"
    DoTypeConv = "IntPtr"
Case "HGPUNV"
    DoTypeConv = "IntPtr"
Case "HVIDEOOUTPUTDEVICENV"
    DoTypeConv = "IntPtr"
Case "HVIDEOINPUTDEVICENV"
    DoTypeConv = "IntPtr"
Case "HPVIDEODEV"
    DoTypeConv = "IntPtr"
Case "FLOAT"
    DoTypeConv = "Single"
Case "float"
    DoTypeConv = "Single"
Case "UINT"
    DoTypeConv = "UInt32"
Case "int"
    DoTypeConv = "Int32"
Case "INT"
    DoTypeConv = "Int32"
Case "unsigned int"
    DoTypeConv = "UInt32"
Case "unsigned long"
    DoTypeConv = "UIntPtr"
Case "BOOL"
    DoTypeConv = "Boolean"
Case "USHORT"
    DoTypeConv = "UInt16"
Case "DWORD"
    DoTypeConv = "UInt32"
Case "HDC"
    DoTypeConv = "IntPtr"
Case "HGLRC"
    DoTypeConv = "IntPtr"
Case "HANDLE"
    DoTypeConv = "IntPtr"
Case "LPVOID"
    DoTypeConv = "IntPtr"
Case "PGPU_DEVICE"
    DoTypeConv = "IntPtr"
Case "GLUquadric"
    DoTypeConv = "IntPtr"
Case "GLUtesselator"
    DoTypeConv = "IntPtr"
Case "GLUnurbs"
    DoTypeConv = "IntPtr"
Case Else
    DoTypeConv = TypeDesc
    Debug.Print TypeDesc
End Select
End Function

'Do type conversion but only for the return type
Function DoRetTypeConv(RetType As String) As String
Dim Trimmed As String
Trimmed = Trim$(Replace(RetType, "const ", ""))

If InStr(Trimmed, "*") Then
    Trimmed = Replace(Trimmed, " *", "*")
    Select Case Trimmed
    'You don't want the memory allocated from an API as a return value freed by CLR
    Case "GLubyte*"
        DoRetTypeConv = "IntPtr"
    Case "GLchar*"
        DoRetTypeConv = "IntPtr"
    Case "char*"
        DoRetTypeConv = "IntPtr"
    Case "void*"
        DoRetTypeConv = "IntPtr"
    Case "wchar_t*"
        DoRetTypeConv = "IntPtr"
    Case Else
        DoRetTypeConv = "IntPtr"
        Debug.Print Trimmed
    End Select
Else
    DoRetTypeConv = DoTypeConv(Trimmed)
End If
End Function

Function GetPointerLevel(TypeDesc As String) As Long
Dim I As Long
Do
    I = InStr(I + 1, TypeDesc, "*")
    GetPointerLevel = GetPointerLevel + 1
Loop While I
GetPointerLevel = GetPointerLevel - 1
End Function

'Rename the parameters if necessary
Function DoParamRename(Param As String, Optional DefaultName As String = "param") As String
Dim Trimmed As String
Dim HaveConst As Boolean
If InStr(Param, "const ") Then HaveConst = True
If InStr(Param, "const*") Then HaveConst = True
Trimmed = Trim$(Replace(Param, "const ", ""))
Trimmed = Trim$(Replace(Trimmed, "const*", "*"))
Trimmed = Replace(Trimmed, "*", "* ")
Trimmed = Replace(Trimmed, "  ", " ")
Trimmed = Replace(Trimmed, " *", "*")

Dim PassType As String, ParamType As String, ParamName As String
Dim DelimPos As Long, IsReserved As Boolean
DelimPos = InStrRev(Trimmed, " ")
If DelimPos Then
    ParamType = Left$(Trimmed, DelimPos - 1)
    ParamName = Mid$(Trimmed, DelimPos + 1)
    Select Case LCase$(ParamName)
    Case "": ParamName = DefaultName 'The param only has it's type
    Case "type": IsReserved = True
    Case "end": IsReserved = True
    Case "string": IsReserved = True
    Case "object": IsReserved = True
    Case "option": IsReserved = True
    Case "event": IsReserved = True
    Case "in": IsReserved = True
    Case "error": IsReserved = True
    Case "property": IsReserved = True
    End Select
    If IsReserved Then ParamName = ParamName & "_"
Else
    ParamType = Trimmed
    ParamName = DefaultName
End If
If InStr(ParamName, "[") Then
    ParamType = ParamType & "*"
    ParamName = Left$(ParamName, InStr(ParamName, "[") - 1)
End If

'Determine if using IntPtr is needed or just use ByRef
Select Case GetPointerLevel(ParamType)
Case 0
    PassType = "ByVal "
    ParamType = DoTypeConv(ParamType)
Case 1
    Select Case ParamType
    Case "void*"
        PassType = "ByVal "
        ParamType = "IntPtr"
    Case "GLvoid*"
        PassType = "ByVal "
        ParamType = "IntPtr"
    Case "char*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPStr)> ByVal "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPStr)> ByRef "
            ParamType = "String"
        End If
    Case "GLchar*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPStr)> ByVal "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPStr)> ByRef "
            ParamType = "String"
        End If
    Case "wchar_t*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPWStr)> ByVal "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPWStr)> ByRef "
            ParamType = "String"
        End If
    Case Else
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByVal "
            ParamType = DoTypeConv(Replace(ParamType, "*", "")) & "()"
        Else
            PassType = "ByRef "
            ParamType = DoTypeConv(Replace(ParamType, "*", ""))
        End If
    End Select
Case 2
    Select Case ParamType
    Case "char**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByVal "
            ParamType = "String()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "String()"
        End If
    Case "GLchar**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByVal "
            ParamType = "String()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "String()"
        End If
    Case "wchar_t**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByVal "
            ParamType = "IntPtr()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "IntPtr()"
        End If
    Case Else
        If HaveConst Then
            PassType = "ByVal "
            ParamType = "IntPtr"
        Else
            PassType = "ByRef "
            ParamType = "IntPtr"
        End If
    End Select
Case Else
    PassType = "ByRef "
    ParamType = "IntPtr"
End Select

Select Case ParamType
Case "Object"
    PassType = "<MarshalAs(UnmanagedType.AsAny)> ByVal "
Case "Boolean"
    PassType = "<MarshalAs(UnmanagedType.Bool)> " & PassType
End Select

DoParamRename = PassType & ParamName & " As " & ParamType
End Function

Function CHex2VBHex(CHexStr As String) As String
If Left$(CHexStr, 2) = "0x" Then
    If Len(CHexStr) = 6 Then
        CHex2VBHex = "&H0" & Mid$(CHexStr, 3)
    Else
        CHex2VBHex = "&H" & Mid$(CHexStr, 3)
    End If
Else
    CHex2VBHex = CHexStr
End If
If UCase$(Right$(CHex2VBHex, 3)) = "ULL" Then CHex2VBHex = Left$(CHex2VBHex, Len(CHex2VBHex) - 3)
If UCase$(Right$(CHex2VBHex, 2)) = "LL" Then CHex2VBHex = Left$(CHex2VBHex, Len(CHex2VBHex) - 2)
If UCase$(Right$(CHex2VBHex, 2)) = "UL" Then CHex2VBHex = Left$(CHex2VBHex, Len(CHex2VBHex) - 2)
If UCase$(Right$(CHex2VBHex, 1)) = "L" Then CHex2VBHex = Left$(CHex2VBHex, Len(CHex2VBHex) - 1)
If UCase$(Right$(CHex2VBHex, 1)) = "U" Then CHex2VBHex = Left$(CHex2VBHex, Len(CHex2VBHex) - 1)
End Function

Sub ExportVB_NET(Parser As clsParser, ExportTo As String)
Dim GLExtString
Dim Ext As clsGLExtension
Dim MacroName
Dim FuncTypeDefName
Dim FuncPtrType
Dim FuncName
Dim I As Long, ExtCount As Long
Dim FuncData() As String
Dim Param() As String
Dim Tail As String
Dim AlreadyDefinedMacros As New Dictionary

Dim FN As Integer
FN = FreeFile
Open ExportTo For Output As #FN
Print #FN, "Imports System.Runtime.InteropServices"
Print #FN,
Print #FN, "Public Module GL_API"
Print #FN,

Print #FN, "#Region ""OpenGL Extension Declerations"""
    
For Each GLExtString In Parser.GLExtension.Keys
    Dim HasMacro As Boolean, HasAPI As Boolean, HasFuncPtr As Boolean
    HasMacro = False
    HasAPI = False
    HasFuncPtr = False
    Set Ext = Parser.GLExtension(GLExtString)

    Print #FN, "#Region """; GLExtString; """"
    Print #FN, vbTab; "' ----------------------------- "; GLExtString; " -----------------------------"
    Print #FN,
    
    'A variable tells you if this extension is available
    Print #FN, vbTab; "Public "; GLExtString; " As Boolean = False"
    Print #FN,
    
    'Macro definitions
    For Each MacroName In Ext.MacroDefs.Keys
        HasMacro = True
        'Make sure no duplicated
        If AlreadyDefinedMacros.Exists(MacroName) = False Then
            AlreadyDefinedMacros.Add MacroName, Ext.MacroDefs(MacroName)
            Print #FN, vbTab; "Public Const "; MacroName; " = "; CHex2VBHex(Ext.MacroDefs(MacroName))
        Else
            Print #FN, vbTab; "'Public Const "; MacroName; " = "; CHex2VBHex(Ext.MacroDefs(MacroName))
        End If
    Next
    If HasMacro Then Print #FN,
    
    'Function typedefs
    For Each FuncTypeDefName In Ext.FuncTypeDef.Keys
        Print #FN, vbTab; "<System.Security.SuppressUnmanagedCodeSecurity()>"
        
        'Function or Sub
        FuncData = Split(Ext.FuncTypeDef(FuncTypeDefName), ":")
        If LCase$(FuncData(0)) = "void" Then
            Tail = ")"
            Print #FN, vbTab; "Public Delegate Sub ";
        Else
            Tail = ") As " & DoRetTypeConv(FuncData(0))
            Print #FN, vbTab; "Public Delegate Function ";
        End If
        
        'Function name
        Print #FN, FuncTypeDefName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRename(Param(I), "param" & I);
            Next
        End If
        
        Print #FN, Tail
        Print #FN,
    Next
    
    'Function pointer declaration
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        HasFuncPtr = True
        Print #FN, vbTab; "Public "; Ext.FuncPtrs(FuncPtrType); " As "; FuncPtrType
    Next
    If HasFuncPtr Then Print #FN,
    
    'API declaration
    For Each FuncName In Ext.APIs.Keys
        HasAPI = True
        FuncData = Split(Ext.APIs(FuncName), ":")
        
        'Function or Sub
        If LCase$(FuncData(0)) = "void" Then
            Tail = ")"
            Print #FN, vbTab; "Declare Sub ";
        Else
            Tail = ") As " & DoRetTypeConv(FuncData(0))
            Print #FN, vbTab; "Declare Function ";
        End If
        
        'Function name and Dll name
        Print #FN, FuncName; " Lib """ & Ext.API_DllName & """ (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRename(Param(I), "param" & I);
            Next
        End If
        
        Print #FN, Tail
    Next
    If HasAPI Then Print #FN,
    
    'Function pointer assignment code generation
    Print #FN, vbTab; "Public Function GLAPI_Init_"; GLExtString; " () As Boolean"
    If HasFuncPtr Then
        Print #FN, vbTab; vbTab; "Dim FuncPtr As IntPtr"
        Print #FN,
        
        'This is only for Windows. using wglGetProcAddress to retrieve the function pointer, and convert to a Delegate
        For Each FuncPtrType In Ext.FuncPtrs.Keys
            FuncName = Ext.FuncPtrs(FuncPtrType)
            Print #FN, vbTab; vbTab; "FuncPtr = wglGetProcAddress("""; FuncName; """)"
            Print #FN, vbTab; vbTab; "If FuncPtr = 0 Then Return False"
            Print #FN, vbTab; vbTab; FuncName; " = Marshal.GetDelegateForFunctionPointer(FuncPtr, GetType("; FuncPtrType; "))"
            Print #FN,
        Next
    End If
    Print #FN, vbTab; vbTab; "Return True"
    Print #FN, vbTab; "End Function"
    Print #FN,
    
    Print #FN, "#End Region"
    ExtCount = ExtCount + 1
Next
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension Strings"""
Print #FN,
Print #FN, vbTab; "Public ReadOnly GLAPI_Extensions() As String ="
Print #FN, vbTab; "{"
Dim LinePos As Long
LinePos = 0
I = 0
Print #FN, vbTab; vbTab;
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, """"; GLExtString;
    If I < ExtCount - 1 Then
        Print #FN, """, ";
    Else
        Print #FN, """";
    End If
    LinePos = LinePos + Len(GLExtString)
    If LinePos > 80 Then
        Print #FN,
        Print #FN, vbTab; vbTab;
        LinePos = 0
    End If
    I = I + 1
Next
Print #FN,
Print #FN, vbTab; "}"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension String Indices"""
Print #FN,
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; "Private Const GLAPI_IndexOf_"; GLExtString; " = "; CStr(I)
    I = I + 1
Next
Print #FN,
Print #FN, vbTab; "Private Function GLAPI_GetIndexOfExtension(ByVal ExtString As String) As Integer"
Print #FN, vbTab; vbTab; "Dim I As Integer"
Print #FN, vbTab; vbTab; "For I = 0 To "; CStr(ExtCount - 1)
Print #FN, vbTab; vbTab; vbTab; "If GLAPI_Extensions(I) = ExtString Then Return I"
Print #FN, vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Return -1"
Print #FN, vbTab; "End Function"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension Initialize"""
Print #FN,
Print #FN, vbTab; "Public Sub GLAPI_Init()"
Print #FN, vbTab; vbTab; "Dim I As Integer"
Print #FN, vbTab; vbTab; "Dim ExtIndex As Integer"
Print #FN, vbTab; vbTab; "Dim ExtCount As Integer"
Print #FN, vbTab; vbTab; "Dim ExtString As String"
Print #FN, vbTab; vbTab; "Dim VersionString As String"
Print #FN, vbTab; vbTab; "Dim VendorSplit() As String"
Print #FN, vbTab; vbTab; "Dim VersionSplit() As String"
Print #FN, vbTab; vbTab; "Dim Major As Integer, Minor As Integer"
Print #FN, vbTab; vbTab; "Dim Extensions_Supported("; CStr(ExtCount - 1); ") As Boolean"
Print #FN,
Print #FN, vbTab; vbTab; "VersionString = Marshal.PtrToStringAnsi(glGetString(GL_VERSION))"
Print #FN, vbTab; vbTab; "VendorSplit = Split(VersionString)"
Print #FN, vbTab; vbTab; "VersionSplit = Split(VendorSplit(0), ""."")"
Print #FN, vbTab; vbTab; "Major = VersionSplit(0)"
Print #FN, vbTab; vbTab; "Minor = VersionSplit(1)"
Print #FN, vbTab; vbTab; "If Major = 1 And Minor = 0 Then Return"
Print #FN,
Print #FN, vbTab; vbTab; "GL_VERSION_4_6 = (Major > 4) Or ((Major = 4) And (Minor >= 6))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_5 = GL_VERSION_4_6 Or ((Major = 4) And (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_4 = GL_VERSION_4_5 Or ((Major = 4) And (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_3 = GL_VERSION_4_4 Or ((Major = 4) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_2 = GL_VERSION_4_3 Or ((Major = 4) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_1 = GL_VERSION_4_2 Or ((Major = 4) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_0 = GL_VERSION_4_1 Or ((Major = 4) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_3 = GL_VERSION_4_0 Or ((Major = 3) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_2 = GL_VERSION_3_3 Or ((Major = 3) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_1 = GL_VERSION_3_2 Or ((Major = 3) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_0 = GL_VERSION_3_1 Or ((Major = 3) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_1 = GL_VERSION_3_0 Or ((Major = 2) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_0 = GL_VERSION_2_1 Or ((Major = 2) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_5 = GL_VERSION_2_0 Or ((Major = 1) And (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_4 = GL_VERSION_1_5 Or ((Major = 1) And (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_3 = GL_VERSION_1_4 Or ((Major = 1) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2_1 = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2 = GL_VERSION_1_2_1 Or ((Major = 1) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_1 = GL_VERSION_1_2 Or ((Major = 1) And (Minor >= 1))"
Print #FN,
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_6) = GL_VERSION_4_6"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_5) = GL_VERSION_4_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_4) = GL_VERSION_4_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_3) = GL_VERSION_4_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_2) = GL_VERSION_4_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_1) = GL_VERSION_4_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_0) = GL_VERSION_4_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_3) = GL_VERSION_3_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_2) = GL_VERSION_3_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_1) = GL_VERSION_3_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_0) = GL_VERSION_3_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_2_1) = GL_VERSION_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_2_0) = GL_VERSION_2_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_5) = GL_VERSION_1_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_4) = GL_VERSION_1_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_3) = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_2_1) = GL_VERSION_1_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_2) = GL_VERSION_1_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_1) = GL_VERSION_1_1"
Print #FN,
Print #FN, vbTab; vbTab; "If GL_VERSION_3_0 Then"
Print #FN, vbTab; vbTab; vbTab; "glGetStringi = Marshal.GetDelegateForFunctionPointer(wglGetProcAddress(""glGetStringi""), GetType(PFNGLGETSTRINGIPROC))"
Print #FN, vbTab; vbTab; vbTab; "glGetIntegerv(GL_NUM_EXTENSIONS, ExtCount)"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtString = Marshal.PtrToStringAnsi(glGetStringi(GL_EXTENSIONS, I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = GLAPI_GetIndexOfExtension(ExtString)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; vbTab; "End If"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Else"
Print #FN, vbTab; vbTab; vbTab; "Dim ExtStrings() As String"
Print #FN, vbTab; vbTab; vbTab; "ExtStrings = Split(Marshal.PtrToStringAnsi(glGetString(GL_EXTENSIONS)))"
Print #FN, vbTab; vbTab; vbTab; "ExtCount = UBound(ExtStrings) + 1"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = GLAPI_GetIndexOfExtension(ExtStrings(I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; vbTab; "End If"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "End If"
Print #FN,
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; vbTab; "If Extensions_Supported(GLAPI_IndexOf_"; GLExtString; ") Then "; GLExtString; " = GLAPI_Init_"; GLExtString; "() Else "; GLExtString; " = False"
    I = I + 1
Next
Print #FN, vbTab; "End Sub"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Context Related"""
Print #FN,
Print #FN, vbTab; "Public Const PFD_TYPE_RGBA = 0"
Print #FN, vbTab; "Public Const PFD_TYPE_COLORINDEX = 1"
Print #FN, vbTab; "Public Const PFD_MAIN_PLANE = 0"
Print #FN, vbTab; "Public Const PFD_OVERLAY_PLANE = 1"
Print #FN, vbTab; "Public Const PFD_UNDERLAY_PLANE = (-1)"
Print #FN, vbTab; "Public Const PFD_DOUBLEBUFFER As Integer = &H1"
Print #FN, vbTab; "Public Const PFD_STEREO As Integer = &H2"
Print #FN, vbTab; "Public Const PFD_DRAW_TO_WINDOW As Integer = &H4"
Print #FN, vbTab; "Public Const PFD_DRAW_TO_BITMAP As Integer = &H8"
Print #FN, vbTab; "Public Const PFD_SUPPORT_GDI As Integer = &H10"
Print #FN, vbTab; "Public Const PFD_SUPPORT_OPENGL As Integer = &H20"
Print #FN, vbTab; "Public Const PFD_GENERIC_FORMAT As Integer = &H40"
Print #FN, vbTab; "Public Const PFD_NEED_PALETTE As Integer = &H80"
Print #FN, vbTab; "Public Const PFD_NEED_SYSTEM_PALETTE As Integer = &H100"
Print #FN, vbTab; "Public Const PFD_SWAP_EXCHANGE As Integer = &H200"
Print #FN, vbTab; "Public Const PFD_SWAP_COPY As Integer = &H400"
Print #FN, vbTab; "Public Const PFD_SWAP_LAYER_BUFFERS As Integer = &H800"
Print #FN, vbTab; "Public Const PFD_GENERIC_ACCELERATED As Integer = &H1000"
Print #FN, vbTab; "Public Const PFD_SUPPORT_DIRECTDRAW As Integer = &H2000"
Print #FN, vbTab; "Public Const PFD_DEPTH_DONTCARE As Integer = &H20000000"
Print #FN, vbTab; "Public Const PFD_DOUBLEBUFFER_DONTCARE As Integer = &H40000000"
Print #FN, vbTab; "Public Const PFD_STEREO_DONTCARE As Integer = &H80000000"
Print #FN,
Print #FN, vbTab; "Structure PIXELFORMATDESCRIPTOR"
Print #FN, vbTab; vbTab; "Public nSize As UInt16"
Print #FN, vbTab; vbTab; "Public nVersion As UInt16"
Print #FN, vbTab; vbTab; "Public dwFlags As UInt32"
Print #FN, vbTab; vbTab; "Public iPixelType As Byte"
Print #FN, vbTab; vbTab; "Public cColorBits As Byte"
Print #FN, vbTab; vbTab; "Public cRedBits As Byte"
Print #FN, vbTab; vbTab; "Public cRedShift As Byte"
Print #FN, vbTab; vbTab; "Public cGreenBits As Byte"
Print #FN, vbTab; vbTab; "Public cGreenShift As Byte"
Print #FN, vbTab; vbTab; "Public cBlueBits As Byte"
Print #FN, vbTab; vbTab; "Public cBlueShift As Byte"
Print #FN, vbTab; vbTab; "Public cAlphaBits As Byte"
Print #FN, vbTab; vbTab; "Public cAlphaShift As Byte"
Print #FN, vbTab; vbTab; "Public cAccumBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumRedBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumGreenBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumBlueBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumAlphaBits As Byte"
Print #FN, vbTab; vbTab; "Public cDepthBits As Byte"
Print #FN, vbTab; vbTab; "Public cStencilBits As Byte"
Print #FN, vbTab; vbTab; "Public cAuxBuffers As Byte"
Print #FN, vbTab; vbTab; "Public iLayerType As Byte"
Print #FN, vbTab; vbTab; "Public bReserved As Byte"
Print #FN, vbTab; vbTab; "Public dwLayerMask As UInt32"
Print #FN, vbTab; vbTab; "Public dwVisibleMask As UInt32"
Print #FN, vbTab; vbTab; "Public dwDamageMask As UInt32"
Print #FN, vbTab; "End Structure"
Print #FN,
Print #FN, vbTab; "Declare Function ChoosePixelFormat Lib ""gdi32.dll"" (ByVal hDC As IntPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Int32"
Print #FN, vbTab; "Declare Function SetPixelFormat Lib ""gdi32.dll"" (ByVal hDC As IntPtr, ByVal pm As Int32, ByRef pfd As PIXELFORMATDESCRIPTOR) As Boolean"
Print #FN, vbTab; "Declare Function wglCreateContext Lib ""opengl32.dll"" (ByVal hDC As IntPtr) As IntPtr"
Print #FN, vbTab; "Declare Function wglMakeCurrent Lib ""opengl32.dll"" (ByVal hDC As IntPtr, ByVal hGLRC As IntPtr) As Boolean"
Print #FN, vbTab; "Declare Function wglDeleteContext Lib ""opengl32.dll"" (ByVal hGLRC As IntPtr) As Boolean"
Print #FN, vbTab; "Declare Function wglSwapBuffers Lib ""opengl32.dll"" (ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Declare Function GetDC Lib ""user32.dll"" (ByVal hWnd As IntPtr) As IntPtr"
Print #FN, vbTab; "Declare Function ReleaseDC Lib ""user32.dll"" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Declare Function wglUseFontBitmaps Lib ""opengl32.dll"" Alias ""wglUseFontBitmapsA"" (ByVal hDC As IntPtr, ByVal first As UInt32, ByVal count As UInt32, ByVal listBase As UInt32) As Boolean"
Print #FN, vbTab; "Declare Function SwapBuffers Lib ""gdi32.dll"" (ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Declare Function wglGetProcAddress Lib ""opengl32.dll"" (ByVal ProcName As String) As IntPtr"
Print #FN,
Print #FN, "#End Region"
Print #FN, "End Module"
Print #FN,
Close #FN
End Sub

Sub ExportVB_NET2(Parser As clsParser, ExportTo As String)
Dim GLExtString
Dim Ext As clsGLExtension
Dim MacroName
Dim FuncTypeDefName
Dim FuncPtrType
Dim FuncName
Dim I As Long, ExtCount As Long
Dim FuncData() As String
Dim Param() As String
Dim Tail As String
Dim AlreadyDefinedMacros As New Dictionary

Dim FN As Integer
FN = FreeFile
Open ExportTo For Output As #FN
Print #FN, "Imports System.Runtime.InteropServices"
Print #FN,
Print #FN, "Class GL_API"
Print #FN,

Print #FN, "#Region ""OpenGL Extension Declerations"""
    
For Each GLExtString In Parser.GLExtension.Keys
    Dim HasMacro As Boolean, HasAPI As Boolean, HasFuncPtr As Boolean
    HasMacro = False
    HasAPI = False
    HasFuncPtr = False
    Set Ext = Parser.GLExtension(GLExtString)

    Print #FN, "#Region """; GLExtString; """"
    Print #FN, vbTab; "' ----------------------------- "; GLExtString; " -----------------------------"
    Print #FN,
    
    'A variable tells you if this extension is available
    Print #FN, vbTab; "Public ReadOnly "; GLExtString; " As Boolean = False"
    Print #FN,
    
    'Macro definitions
    For Each MacroName In Ext.MacroDefs.Keys
        HasMacro = True
        'Make sure no duplicated
        If AlreadyDefinedMacros.Exists(MacroName) = False Then
            AlreadyDefinedMacros.Add MacroName, Ext.MacroDefs(MacroName)
            Print #FN, vbTab; "Public Const "; MacroName; " = "; CHex2VBHex(Ext.MacroDefs(MacroName))
        Else
            Print #FN, vbTab; "'Public Const "; MacroName; " = "; CHex2VBHex(Ext.MacroDefs(MacroName))
        End If
    Next
    If HasMacro Then Print #FN,
    
    'Function typedefs
    For Each FuncTypeDefName In Ext.FuncTypeDef.Keys
        Print #FN, vbTab; "<System.Security.SuppressUnmanagedCodeSecurity()>"
        
        'Function or Sub
        FuncData = Split(Ext.FuncTypeDef(FuncTypeDefName), ":")
        If LCase$(FuncData(0)) = "void" Then
            Tail = ")"
            Print #FN, vbTab; "Public Delegate Sub ";
        Else
            Tail = ") As " & DoRetTypeConv(FuncData(0))
            Print #FN, vbTab; "Public Delegate Function ";
        End If
        
        'Function name
        Print #FN, FuncTypeDefName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRename(Param(I), "param" & I);
            Next
        End If
        
        Print #FN, Tail
        Print #FN,
    Next
    
    'Function pointer declaration
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        HasFuncPtr = True
        Print #FN, vbTab; "Public ReadOnly "; Ext.FuncPtrs(FuncPtrType); " As "; FuncPtrType
    Next
    If HasFuncPtr Then Print #FN,
    
    'API declaration
    For Each FuncName In Ext.APIs.Keys
        HasAPI = True
        FuncData = Split(Ext.APIs(FuncName), ":")
        
        'Function or Sub
        If LCase$(FuncData(0)) = "void" Then
            Tail = ")"
            Print #FN, vbTab; "Declare Sub ";
        Else
            Tail = ") As " & DoRetTypeConv(FuncData(0))
            Print #FN, vbTab; "Declare Function ";
        End If
        
        'Function name and Dll name
        Print #FN, FuncName; " Lib """ & Ext.API_DllName & """ (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRename(Param(I), "param" & I);
            Next
        End If
        
        Print #FN, Tail
    Next
    If HasAPI Then Print #FN,
    
    'Function pointer assignment code generation
    'Not good for a class. We move all the code into Sub New
'    Print #FN, vbTab; "Private Function GLAPI_Init_"; GLExtString; " () As Boolean"
'    If HasFuncPtr Then
'        Print #FN, vbTab; vbTab; "Dim FuncPtr As IntPtr"
'        Print #FN,
'
'        'This is only for Windows. using wglGetProcAddress to retrieve the function pointer, and convert to a Delegate
'        For Each FuncPtrType In Ext.FuncPtrs.Keys
'            FuncName = Ext.FuncPtrs(FuncPtrType)
'            Print #FN, vbTab; vbTab; "FuncPtr = wglGetProcAddress("""; FuncName; """)"
'            Print #FN, vbTab; vbTab; "If FuncPtr = 0 Then Return False"
'            Print #FN, vbTab; vbTab; FuncName; " = Marshal.GetDelegateForFunctionPointer(FuncPtr, GetType("; FuncPtrType; "))"
'            Print #FN,
'        Next
'    End If
'    Print #FN, vbTab; vbTab; "Return True"
'    Print #FN, vbTab; "End Function"
'    Print #FN,
    
    Print #FN, "#End Region"
    ExtCount = ExtCount + 1
Next
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension Strings"""
Print #FN,
Print #FN, vbTab; "Public ReadOnly GLAPI_Extensions() As String ="
Print #FN, vbTab; "{"
Dim LinePos As Long
LinePos = 0
I = 0
Print #FN, vbTab; vbTab;
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, """"; GLExtString;
    If I < ExtCount - 1 Then
        Print #FN, """, ";
    Else
        Print #FN, """";
    End If
    LinePos = LinePos + Len(GLExtString)
    If LinePos > 80 Then
        Print #FN,
        Print #FN, vbTab; vbTab;
        LinePos = 0
    End If
    I = I + 1
Next
Print #FN,
Print #FN, vbTab; "}"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension String Indices"""
Print #FN,
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; "Private Const GLAPI_IndexOf_"; GLExtString; " = "; CStr(I)
    I = I + 1
Next
Print #FN,
Print #FN, vbTab; "Private Function GLAPI_GetIndexOfExtension(ByVal ExtString As String) As Integer"
Print #FN, vbTab; vbTab; "Dim I As Integer"
Print #FN, vbTab; vbTab; "For I = 0 To "; CStr(ExtCount - 1)
Print #FN, vbTab; vbTab; vbTab; "If GLAPI_Extensions(I) = ExtString Then Return I"
Print #FN, vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Return -1"
Print #FN, vbTab; "End Function"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Extension Initialize"""
Print #FN,
Print #FN, vbTab; "Sub New()"
Print #FN, vbTab; vbTab; "Dim I As Integer"
Print #FN, vbTab; vbTab; "Dim ExtIndex As Integer"
Print #FN, vbTab; vbTab; "Dim ExtCount As Integer"
Print #FN, vbTab; vbTab; "Dim ExtString As String"
Print #FN, vbTab; vbTab; "Dim VersionString As String"
Print #FN, vbTab; vbTab; "Dim VendorSplit() As String"
Print #FN, vbTab; vbTab; "Dim VersionSplit() As String"
Print #FN, vbTab; vbTab; "Dim Major As Integer, Minor As Integer"
Print #FN, vbTab; vbTab; "Dim Extensions_Supported("; CStr(ExtCount - 1); ") As Boolean"
Print #FN,
Print #FN, vbTab; vbTab; "VersionString = Marshal.PtrToStringAnsi(glGetString(GL_VERSION))"
Print #FN, vbTab; vbTab; "VendorSplit = Split(VersionString)"
Print #FN, vbTab; vbTab; "VersionSplit = Split(VendorSplit(0), ""."")"
Print #FN, vbTab; vbTab; "Major = VersionSplit(0)"
Print #FN, vbTab; vbTab; "Minor = VersionSplit(1)"
Print #FN, vbTab; vbTab; "If Major = 1 And Minor = 0 Then Return"
Print #FN,
Print #FN, vbTab; vbTab; "GL_VERSION_4_6 = (Major > 4) Or ((Major = 4) And (Minor >= 6))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_5 = GL_VERSION_4_6 Or ((Major = 4) And (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_4 = GL_VERSION_4_5 Or ((Major = 4) And (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_3 = GL_VERSION_4_4 Or ((Major = 4) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_2 = GL_VERSION_4_3 Or ((Major = 4) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_1 = GL_VERSION_4_2 Or ((Major = 4) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_0 = GL_VERSION_4_1 Or ((Major = 4) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_3 = GL_VERSION_4_0 Or ((Major = 3) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_2 = GL_VERSION_3_3 Or ((Major = 3) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_1 = GL_VERSION_3_2 Or ((Major = 3) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_0 = GL_VERSION_3_1 Or ((Major = 3) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_1 = GL_VERSION_3_0 Or ((Major = 2) And (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_0 = GL_VERSION_2_1 Or ((Major = 2) And (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_5 = GL_VERSION_2_0 Or ((Major = 1) And (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_4 = GL_VERSION_1_5 Or ((Major = 1) And (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_3 = GL_VERSION_1_4 Or ((Major = 1) And (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2_1 = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2 = GL_VERSION_1_2_1 Or ((Major = 1) And (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_1 = GL_VERSION_1_2 Or ((Major = 1) And (Minor >= 1))"
Print #FN,
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_6) = GL_VERSION_4_6"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_5) = GL_VERSION_4_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_4) = GL_VERSION_4_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_3) = GL_VERSION_4_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_2) = GL_VERSION_4_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_1) = GL_VERSION_4_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_4_0) = GL_VERSION_4_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_3) = GL_VERSION_3_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_2) = GL_VERSION_3_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_1) = GL_VERSION_3_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_3_0) = GL_VERSION_3_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_2_1) = GL_VERSION_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_2_0) = GL_VERSION_2_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_5) = GL_VERSION_1_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_4) = GL_VERSION_1_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_3) = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_2_1) = GL_VERSION_1_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_2) = GL_VERSION_1_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(GLAPI_IndexOf_GL_VERSION_1_1) = GL_VERSION_1_1"
Print #FN,
Print #FN, vbTab; vbTab; "If GL_VERSION_3_0 Then"
Print #FN, vbTab; vbTab; vbTab; "glGetStringi = Marshal.GetDelegateForFunctionPointer(wglGetProcAddress(""glGetStringi""), GetType(PFNGLGETSTRINGIPROC))"
Print #FN, vbTab; vbTab; vbTab; "glGetIntegerv(GL_NUM_EXTENSIONS, ExtCount)"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtString = Marshal.PtrToStringAnsi(glGetStringi(GL_EXTENSIONS, I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = GLAPI_GetIndexOfExtension(ExtString)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; vbTab; "End If"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Else"
Print #FN, vbTab; vbTab; vbTab; "Dim ExtStrings() As String"
Print #FN, vbTab; vbTab; vbTab; "ExtStrings = Split(Marshal.PtrToStringAnsi(glGetString(GL_EXTENSIONS)))"
Print #FN, vbTab; vbTab; vbTab; "ExtCount = UBound(ExtStrings) + 1"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = GLAPI_GetIndexOfExtension(ExtStrings(I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; vbTab; "End If"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "End If"
Print #FN,
I = 0
Print #FN, vbTab; "Dim FuncPtr As IntPtr"
For Each GLExtString In Parser.GLExtension.Keys
    Set Ext = Parser.GLExtension(GLExtString)

    Print #FN, "#Region """; GLExtString; "_Initialize"""
    Print #FN, vbTab; "' ----------------------------- "; GLExtString; " -----------------------------"
    Print #FN,
    
    'Function pointer assignment code generation
    Print #FN, vbTab; GLExtString; " = False"
    Print #FN, vbTab; "If Extensions_Supported(GLAPI_IndexOf_"; GLExtString; ") Then"
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        FuncName = Ext.FuncPtrs(FuncPtrType)
        Print #FN, vbTab; vbTab; "FuncPtr = wglGetProcAddress("""; FuncName; """)"
        Print #FN, vbTab; vbTab; "If FuncPtr = 0 Then Goto EndOf_"; GLExtString
        Print #FN, vbTab; vbTab; FuncName; " = Marshal.GetDelegateForFunctionPointer(FuncPtr, GetType("; FuncPtrType; "))"
        Print #FN,
    Next
    Print #FN, vbTab; vbTab; GLExtString; " = True"
    Print #FN, vbTab; "End If"
    Print #FN, vbTab; "EndOf_"; GLExtString; ":"
    Print #FN,
    
    Print #FN, "#End Region"
    I = I + 1
Next
Print #FN, vbTab; "End Sub"
Print #FN,
Print #FN, "#End Region"
Print #FN,
Print #FN, "#Region ""OpenGL Context Related"""
Print #FN,
Print #FN, vbTab; "Public Const PFD_TYPE_RGBA = 0"
Print #FN, vbTab; "Public Const PFD_TYPE_COLORINDEX = 1"
Print #FN, vbTab; "Public Const PFD_MAIN_PLANE = 0"
Print #FN, vbTab; "Public Const PFD_OVERLAY_PLANE = 1"
Print #FN, vbTab; "Public Const PFD_UNDERLAY_PLANE = (-1)"
Print #FN, vbTab; "Public Const PFD_DOUBLEBUFFER As Integer = &H1"
Print #FN, vbTab; "Public Const PFD_STEREO As Integer = &H2"
Print #FN, vbTab; "Public Const PFD_DRAW_TO_WINDOW As Integer = &H4"
Print #FN, vbTab; "Public Const PFD_DRAW_TO_BITMAP As Integer = &H8"
Print #FN, vbTab; "Public Const PFD_SUPPORT_GDI As Integer = &H10"
Print #FN, vbTab; "Public Const PFD_SUPPORT_OPENGL As Integer = &H20"
Print #FN, vbTab; "Public Const PFD_GENERIC_FORMAT As Integer = &H40"
Print #FN, vbTab; "Public Const PFD_NEED_PALETTE As Integer = &H80"
Print #FN, vbTab; "Public Const PFD_NEED_SYSTEM_PALETTE As Integer = &H100"
Print #FN, vbTab; "Public Const PFD_SWAP_EXCHANGE As Integer = &H200"
Print #FN, vbTab; "Public Const PFD_SWAP_COPY As Integer = &H400"
Print #FN, vbTab; "Public Const PFD_SWAP_LAYER_BUFFERS As Integer = &H800"
Print #FN, vbTab; "Public Const PFD_GENERIC_ACCELERATED As Integer = &H1000"
Print #FN, vbTab; "Public Const PFD_SUPPORT_DIRECTDRAW As Integer = &H2000"
Print #FN, vbTab; "Public Const PFD_DEPTH_DONTCARE As Integer = &H20000000"
Print #FN, vbTab; "Public Const PFD_DOUBLEBUFFER_DONTCARE As Integer = &H40000000"
Print #FN, vbTab; "Public Const PFD_STEREO_DONTCARE As Integer = &H80000000"
Print #FN,
Print #FN, vbTab; "Public Structure PIXELFORMATDESCRIPTOR"
Print #FN, vbTab; vbTab; "Public nSize As UInt16"
Print #FN, vbTab; vbTab; "Public nVersion As UInt16"
Print #FN, vbTab; vbTab; "Public dwFlags As UInt32"
Print #FN, vbTab; vbTab; "Public iPixelType As Byte"
Print #FN, vbTab; vbTab; "Public cColorBits As Byte"
Print #FN, vbTab; vbTab; "Public cRedBits As Byte"
Print #FN, vbTab; vbTab; "Public cRedShift As Byte"
Print #FN, vbTab; vbTab; "Public cGreenBits As Byte"
Print #FN, vbTab; vbTab; "Public cGreenShift As Byte"
Print #FN, vbTab; vbTab; "Public cBlueBits As Byte"
Print #FN, vbTab; vbTab; "Public cBlueShift As Byte"
Print #FN, vbTab; vbTab; "Public cAlphaBits As Byte"
Print #FN, vbTab; vbTab; "Public cAlphaShift As Byte"
Print #FN, vbTab; vbTab; "Public cAccumBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumRedBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumGreenBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumBlueBits As Byte"
Print #FN, vbTab; vbTab; "Public cAccumAlphaBits As Byte"
Print #FN, vbTab; vbTab; "Public cDepthBits As Byte"
Print #FN, vbTab; vbTab; "Public cStencilBits As Byte"
Print #FN, vbTab; vbTab; "Public cAuxBuffers As Byte"
Print #FN, vbTab; vbTab; "Public iLayerType As Byte"
Print #FN, vbTab; vbTab; "Public bReserved As Byte"
Print #FN, vbTab; vbTab; "Public dwLayerMask As UInt32"
Print #FN, vbTab; vbTab; "Public dwVisibleMask As UInt32"
Print #FN, vbTab; vbTab; "Public dwDamageMask As UInt32"
Print #FN, vbTab; "End Structure"
Print #FN,
Print #FN, vbTab; "Public Declare Function ChoosePixelFormat Lib ""gdi32.dll"" (ByVal hDC As IntPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Int32"
Print #FN, vbTab; "Public Declare Function SetPixelFormat Lib ""gdi32.dll"" (ByVal hDC As IntPtr, ByVal pm As Int32, ByRef pfd As PIXELFORMATDESCRIPTOR) As Boolean"
Print #FN, vbTab; "Public Declare Function wglCreateContext Lib ""opengl32.dll"" (ByVal hDC As IntPtr) As IntPtr"
Print #FN, vbTab; "Public Declare Function wglGetCurrentDC Lib ""opengl32.dll"" () As IntPtr"
Print #FN, vbTab; "Public Declare Function wglGetCurrentContext Lib ""opengl32.dll"" () As IntPtr"
Print #FN, vbTab; "Public Declare Function wglMakeCurrent Lib ""opengl32.dll"" (ByVal hDC As IntPtr, ByVal hGLRC As IntPtr) As Boolean"
Print #FN, vbTab; "Public Declare Function wglDeleteContext Lib ""opengl32.dll"" (ByVal hGLRC As IntPtr) As Boolean"
Print #FN, vbTab; "Public Declare Function wglSwapBuffers Lib ""opengl32.dll"" (ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Public Declare Function GetDC Lib ""user32.dll"" (ByVal hWnd As IntPtr) As IntPtr"
Print #FN, vbTab; "Public Declare Function ReleaseDC Lib ""user32.dll"" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Public Declare Function wglUseFontBitmaps Lib ""opengl32.dll"" Alias ""wglUseFontBitmapsA"" (ByVal hDC As IntPtr, ByVal first As UInt32, ByVal count As UInt32, ByVal listBase As UInt32) As Boolean"
Print #FN, vbTab; "Public Declare Function SwapBuffers Lib ""gdi32.dll"" (ByVal hDC As IntPtr) As Boolean"
Print #FN, vbTab; "Public Declare Function wglGetProcAddress Lib ""opengl32.dll"" (ByVal ProcName As String) As IntPtr"
Print #FN,
Print #FN, "#End Region"
Print #FN, "End Class"
Print #FN,
Close #FN
End Sub

