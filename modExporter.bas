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

Function MatchPrefix(ByVal Str_ As String, ByVal Prefix As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
If CaseSensitive Then
    If Left$(Str_, Len(Prefix)) = Prefix Then MatchPrefix = True
Else
    If Left$(UCase$(Str_), Len(Prefix)) = UCase$(Prefix) Then MatchPrefix = True
End If
End Function

'Do type conversions
Function DoTypeConvVB_NET(TypeDesc As String) As String
Select Case TypeDesc
Case "GLenum"
    DoTypeConvVB_NET = "UInt32"
Case "GLbitfield"
    DoTypeConvVB_NET = "UInt32"
Case "GLuint"
    DoTypeConvVB_NET = "UInt32"
Case "GLint"
    DoTypeConvVB_NET = "Int32"
Case "GLsizei"
    DoTypeConvVB_NET = "Int32"
Case "GLboolean"
    DoTypeConvVB_NET = "Boolean"
Case "GLbyte"
    DoTypeConvVB_NET = "Byte"
Case "GLshort"
    DoTypeConvVB_NET = "Int16"
Case "GLubyte"
    DoTypeConvVB_NET = "Byte"
Case "GLushort"
    DoTypeConvVB_NET = "UInt16"
Case "GLulong"
    DoTypeConvVB_NET = "UIntPtr"
Case "GLfloat"
    DoTypeConvVB_NET = "Single"
Case "GLclampf"
    DoTypeConvVB_NET = "Single"
Case "GLdouble"
    DoTypeConvVB_NET = "Double"
Case "GLclampd"
    DoTypeConvVB_NET = "Double"
Case "GLvoid"
    DoTypeConvVB_NET = "IntPtr"
Case "GLint64EXT"
    DoTypeConvVB_NET = "Int64"
Case "GLuint64EXT"
    DoTypeConvVB_NET = "UInt64"
Case "GLint64"
    DoTypeConvVB_NET = "Int64"
Case "GLuint64"
    DoTypeConvVB_NET = "UInt64"
Case "GLsync"
    DoTypeConvVB_NET = "IntPtr"
Case "GLchar"
    DoTypeConvVB_NET = "Byte"
Case "GLintptr"
    DoTypeConvVB_NET = "IntPtr"
Case "GLsizeiptr"
    DoTypeConvVB_NET = "IntPtr"
Case "GLfixed"
    DoTypeConvVB_NET = "Int32"
Case "cl_context"
    DoTypeConvVB_NET = "IntPtr"
Case "cl_event"
    DoTypeConvVB_NET = "IntPtr"
Case "GLcharARB"
    DoTypeConvVB_NET = "Byte"
Case "GLhandleARB"
    DoTypeConvVB_NET = "Int32"
Case "GLintptrARB"
    DoTypeConvVB_NET = "IntPtr"
Case "GLsizeiptrARB"
    DoTypeConvVB_NET = "IntPtr"
Case "GLeglClientBufferEXT"
    DoTypeConvVB_NET = "UIntPtr"
Case "GLhalf"
    DoTypeConvVB_NET = "UInt16"
Case "GLvdpauSurfaceNV"
    DoTypeConvVB_NET = "IntPtr"
Case "GLclampx"
    DoTypeConvVB_NET = "Int32"
Case "HPBUFFERARB"
    DoTypeConvVB_NET = "IntPtr"
Case "HPBUFFEREXT"
    DoTypeConvVB_NET = "IntPtr"
Case "HGPUNV"
    DoTypeConvVB_NET = "IntPtr"
Case "HVIDEOOUTPUTDEVICENV"
    DoTypeConvVB_NET = "IntPtr"
Case "HVIDEOINPUTDEVICENV"
    DoTypeConvVB_NET = "IntPtr"
Case "HPVIDEODEV"
    DoTypeConvVB_NET = "IntPtr"
Case "FLOAT"
    DoTypeConvVB_NET = "Single"
Case "float"
    DoTypeConvVB_NET = "Single"
Case "UINT"
    DoTypeConvVB_NET = "UInt32"
Case "int"
    DoTypeConvVB_NET = "Int32"
Case "INT"
    DoTypeConvVB_NET = "Int32"
Case "unsigned int"
    DoTypeConvVB_NET = "UInt32"
Case "unsigned long"
    DoTypeConvVB_NET = "UIntPtr"
Case "BOOL"
    DoTypeConvVB_NET = "Boolean"
Case "USHORT"
    DoTypeConvVB_NET = "UInt16"
Case "DWORD"
    DoTypeConvVB_NET = "UInt32"
Case "HDC"
    DoTypeConvVB_NET = "IntPtr"
Case "HGLRC"
    DoTypeConvVB_NET = "IntPtr"
Case "HANDLE"
    DoTypeConvVB_NET = "IntPtr"
Case "LPVOID"
    DoTypeConvVB_NET = "IntPtr"
Case "PGPU_DEVICE"
    DoTypeConvVB_NET = "IntPtr"
Case "GLUquadric"
    DoTypeConvVB_NET = "IntPtr"
Case "GLUtesselator"
    DoTypeConvVB_NET = "IntPtr"
Case "GLUnurbs"
    DoTypeConvVB_NET = "IntPtr"
Case Else
    DoTypeConvVB_NET = TypeDesc
    Debug.Print TypeDesc
End Select
End Function

'Do type conversion but only for the return type
Function DoRetTypeConvVB_NET(RetType As String) As String
Dim Trimmed As String
Trimmed = Trim$(Replace(RetType, "const ", ""))

If InStr(Trimmed, "*") Then
    Trimmed = Replace(Trimmed, " *", "*")
    Select Case Trimmed
    'You don't want the memory allocated from an API as a return value freed by CLR
    Case "GLubyte*"
        DoRetTypeConvVB_NET = "IntPtr"
    Case "GLchar*"
        DoRetTypeConvVB_NET = "IntPtr"
    Case "char*"
        DoRetTypeConvVB_NET = "IntPtr"
    Case "void*"
        DoRetTypeConvVB_NET = "IntPtr"
    Case "wchar_t*"
        DoRetTypeConvVB_NET = "IntPtr"
    Case Else
        DoRetTypeConvVB_NET = "IntPtr"
        Debug.Print Trimmed
    End Select
Else
    DoRetTypeConvVB_NET = DoTypeConvVB_NET(Trimmed)
End If
End Function

'Do type conversions
Function DoTypeConvCSharp(TypeDesc As String) As String
Select Case TypeDesc
Case "void"
    DoTypeConvCSharp = "void"
Case "VOID"
    DoTypeConvCSharp = "void"
Case "GLenum"
    DoTypeConvCSharp = "uint"
Case "GLbitfield"
    DoTypeConvCSharp = "uint"
Case "GLuint"
    DoTypeConvCSharp = "uint"
Case "GLint"
    DoTypeConvCSharp = "int"
Case "GLsizei"
    DoTypeConvCSharp = "int"
Case "GLboolean"
    DoTypeConvCSharp = "bool"
Case "GLbyte"
    DoTypeConvCSharp = "byte"
Case "GLshort"
    DoTypeConvCSharp = "short"
Case "GLubyte"
    DoTypeConvCSharp = "byte"
Case "GLushort"
    DoTypeConvCSharp = "ushort"
Case "GLulong"
    DoTypeConvCSharp = "UIntPtr"
Case "GLfloat"
    DoTypeConvCSharp = "float"
Case "GLclampf"
    DoTypeConvCSharp = "float"
Case "GLdouble"
    DoTypeConvCSharp = "double"
Case "GLclampd"
    DoTypeConvCSharp = "double"
Case "GLvoid"
    DoTypeConvCSharp = "IntPtr"
Case "GLint64EXT"
    DoTypeConvCSharp = "long"
Case "GLuint64EXT"
    DoTypeConvCSharp = "ulong"
Case "GLint64"
    DoTypeConvCSharp = "long"
Case "GLuint64"
    DoTypeConvCSharp = "ulong"
Case "GLsync"
    DoTypeConvCSharp = "IntPtr"
Case "GLchar"
    DoTypeConvCSharp = "byte"
Case "GLintptr"
    DoTypeConvCSharp = "IntPtr"
Case "GLsizeiptr"
    DoTypeConvCSharp = "IntPtr"
Case "GLfixed"
    DoTypeConvCSharp = "int"
Case "cl_context"
    DoTypeConvCSharp = "IntPtr"
Case "cl_event"
    DoTypeConvCSharp = "IntPtr"
Case "GLcharARB"
    DoTypeConvCSharp = "byte"
Case "GLhandleARB"
    DoTypeConvCSharp = "uint"
Case "GLintptrARB"
    DoTypeConvCSharp = "IntPtr"
Case "GLsizeiptrARB"
    DoTypeConvCSharp = "IntPtr"
Case "GLeglClientBufferEXT"
    DoTypeConvCSharp = "UIntPtr"
Case "GLhalf"
    DoTypeConvCSharp = "ushort"
Case "GLvdpauSurfaceNV"
    DoTypeConvCSharp = "IntPtr"
Case "GLclampx"
    DoTypeConvCSharp = "int"
Case "HPBUFFERARB"
    DoTypeConvCSharp = "IntPtr"
Case "HPBUFFEREXT"
    DoTypeConvCSharp = "IntPtr"
Case "HGPUNV"
    DoTypeConvCSharp = "IntPtr"
Case "HVIDEOOUTPUTDEVICENV"
    DoTypeConvCSharp = "IntPtr"
Case "HVIDEOINPUTDEVICENV"
    DoTypeConvCSharp = "IntPtr"
Case "HPVIDEODEV"
    DoTypeConvCSharp = "IntPtr"
Case "FLOAT"
    DoTypeConvCSharp = "float"
Case "float"
    DoTypeConvCSharp = "float"
Case "UINT"
    DoTypeConvCSharp = "uint"
Case "int"
    DoTypeConvCSharp = "int"
Case "INT"
    DoTypeConvCSharp = "int"
Case "unsigned int"
    DoTypeConvCSharp = "uint"
Case "unsigned long"
    DoTypeConvCSharp = "UIntPtr"
Case "BOOL"
    DoTypeConvCSharp = "bool"
Case "USHORT"
    DoTypeConvCSharp = "ushort"
Case "DWORD"
    DoTypeConvCSharp = "uint"
Case "HDC"
    DoTypeConvCSharp = "IntPtr"
Case "HGLRC"
    DoTypeConvCSharp = "IntPtr"
Case "HANDLE"
    DoTypeConvCSharp = "IntPtr"
Case "LPVOID"
    DoTypeConvCSharp = "IntPtr"
Case "PGPU_DEVICE"
    DoTypeConvCSharp = "IntPtr"
Case "GLUquadric"
    DoTypeConvCSharp = "IntPtr"
Case "GLUtesselator"
    DoTypeConvCSharp = "IntPtr"
Case "GLUnurbs"
    DoTypeConvCSharp = "IntPtr"
Case Else
    DoTypeConvCSharp = TypeDesc
    Debug.Print TypeDesc
End Select
End Function

'Do type conversion but only for the return type
Function DoRetTypeConvCSharp(RetType As String) As String
Dim Trimmed As String
Trimmed = Trim$(Replace(RetType, "const ", ""))

If InStr(Trimmed, "*") Then
    Trimmed = Replace(Trimmed, " *", "*")
    Select Case Trimmed
    'You don't want the memory allocated from an API as a return value freed by CLR
    Case "GLubyte*"
        DoRetTypeConvCSharp = "IntPtr"
    Case "GLchar*"
        DoRetTypeConvCSharp = "IntPtr"
    Case "char*"
        DoRetTypeConvCSharp = "IntPtr"
    Case "void*"
        DoRetTypeConvCSharp = "IntPtr"
    Case "wchar_t*"
        DoRetTypeConvCSharp = "IntPtr"
    Case Else
        DoRetTypeConvCSharp = "IntPtr"
        Debug.Print Trimmed
    End Select
Else
    DoRetTypeConvCSharp = DoTypeConvCSharp(Trimmed)
End If
End Function

'Do type conversions
Function DoTypeConvCSharpUnsafe(TypeDesc As String) As String
Select Case TypeDesc
Case "void"
    DoTypeConvCSharpUnsafe = "void"
Case "VOID"
    DoTypeConvCSharpUnsafe = "void"
Case "GLenum"
    DoTypeConvCSharpUnsafe = "uint"
Case "GLbitfield"
    DoTypeConvCSharpUnsafe = "uint"
Case "GLuint"
    DoTypeConvCSharpUnsafe = "uint"
Case "GLint"
    DoTypeConvCSharpUnsafe = "int"
Case "GLsizei"
    DoTypeConvCSharpUnsafe = "int"
Case "GLboolean"
    DoTypeConvCSharpUnsafe = "bool"
Case "GLbyte"
    DoTypeConvCSharpUnsafe = "byte"
Case "GLshort"
    DoTypeConvCSharpUnsafe = "short"
Case "GLubyte"
    DoTypeConvCSharpUnsafe = "byte"
Case "GLushort"
    DoTypeConvCSharpUnsafe = "ushort"
Case "GLulong"
    DoTypeConvCSharpUnsafe = "UIntPtr"
Case "GLfloat"
    DoTypeConvCSharpUnsafe = "float"
Case "GLclampf"
    DoTypeConvCSharpUnsafe = "float"
Case "GLdouble"
    DoTypeConvCSharpUnsafe = "double"
Case "GLclampd"
    DoTypeConvCSharpUnsafe = "double"
Case "GLvoid"
    DoTypeConvCSharpUnsafe = "void"
Case "GLint64EXT"
    DoTypeConvCSharpUnsafe = "long"
Case "GLuint64EXT"
    DoTypeConvCSharpUnsafe = "ulong"
Case "GLint64"
    DoTypeConvCSharpUnsafe = "long"
Case "GLuint64"
    DoTypeConvCSharpUnsafe = "ulong"
Case "GLsync"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLchar"
    DoTypeConvCSharpUnsafe = "byte"
Case "GLintptr"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "GLsizeiptr"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "GLfixed"
    DoTypeConvCSharpUnsafe = "int"
Case "cl_context"
    DoTypeConvCSharpUnsafe = "void*"
Case "cl_event"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLcharARB"
    DoTypeConvCSharpUnsafe = "byte"
Case "GLhandleARB"
    DoTypeConvCSharpUnsafe = "uint"
Case "GLintptrARB"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "GLsizeiptrARB"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "GLeglClientBufferEXT"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLhalf"
    DoTypeConvCSharpUnsafe = "ushort"
Case "GLvdpauSurfaceNV"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "GLclampx"
    DoTypeConvCSharpUnsafe = "int"
Case "HPBUFFERARB"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HPBUFFEREXT"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HGPUNV"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HVIDEOOUTPUTDEVICENV"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HVIDEOINPUTDEVICENV"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HPVIDEODEV"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "FLOAT"
    DoTypeConvCSharpUnsafe = "float"
Case "float"
    DoTypeConvCSharpUnsafe = "float"
Case "UINT"
    DoTypeConvCSharpUnsafe = "uint"
Case "int"
    DoTypeConvCSharpUnsafe = "int"
Case "INT"
    DoTypeConvCSharpUnsafe = "int"
Case "unsigned int"
    DoTypeConvCSharpUnsafe = "uint"
Case "unsigned long"
    DoTypeConvCSharpUnsafe = "UIntPtr"
Case "BOOL"
    DoTypeConvCSharpUnsafe = "bool"
Case "USHORT"
    DoTypeConvCSharpUnsafe = "ushort"
Case "DWORD"
    DoTypeConvCSharpUnsafe = "uint"
Case "HDC"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HGLRC"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "HANDLE"
    DoTypeConvCSharpUnsafe = "IntPtr"
Case "LPVOID"
    DoTypeConvCSharpUnsafe = "void*"
Case "PGPU_DEVICE"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLUquadric"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLUtesselator"
    DoTypeConvCSharpUnsafe = "void*"
Case "GLUnurbs"
    DoTypeConvCSharpUnsafe = "void*"
Case Else
    DoTypeConvCSharpUnsafe = TypeDesc
    Debug.Print TypeDesc
End Select
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
Function DoParamRenameVB_NET(Param As String, Optional DefaultName As String = "param") As String
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
    PassType = ""
    ParamType = DoTypeConvVB_NET(ParamType)
Case 1
    Select Case ParamType
    Case "void*"
        PassType = ""
        ParamType = "IntPtr"
    Case "GLvoid*"
        PassType = ""
        ParamType = "IntPtr"
    Case "char*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPStr)> "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPStr)> "
            ParamType = "StringBuilder"
        End If
    Case "GLchar*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPStr)> "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPStr)> "
            ParamType = "StringBuilder"
        End If
    Case "wchar_t*"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPWStr)> "
            ParamType = "String"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPWStr)> "
            ParamType = "StringBuilder"
        End If
    Case Else
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> "
            ParamType = DoTypeConvVB_NET(Replace(ParamType, "*", "")) & "()"
        Else
            PassType = "ByRef "
            ParamType = DoTypeConvVB_NET(Replace(ParamType, "*", ""))
        End If
    End Select
Case 2
    Select Case ParamType
    Case "char**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> "
            ParamType = "String()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "String()"
        End If
    Case "GLchar**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> "
            ParamType = "String()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "String()"
        End If
    Case "wchar_t**"
        If HaveConst Then
            PassType = "<MarshalAs(UnmanagedType.LPArray)> "
            ParamType = "IntPtr()"
        Else
            PassType = "<MarshalAs(UnmanagedType.LPArray)> ByRef "
            ParamType = "IntPtr()"
        End If
    Case Else
        If HaveConst Then
            PassType = ""
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
    PassType = "<MarshalAs(UnmanagedType.AsAny)> "
Case "Boolean"
    PassType = "<MarshalAs(UnmanagedType.Bool)> " & PassType
End Select

DoParamRenameVB_NET = PassType & ParamName & " As " & ParamType
End Function

Function TheFunctionIsglGetSomething(ByVal FunctionName As String) As Boolean
If MatchPrefix(FunctionName, "PFNGLGET", False) Or _
   MatchPrefix(FunctionName, "PFNWGLGET", False) Or _
   MatchPrefix(FunctionName, "glGet", True) Or _
   MatchPrefix(FunctionName, "wglGet", True) Then TheFunctionIsglGetSomething = True
End Function

Function TheFunctionIsglGenSomething(ByVal FunctionName As String) As Boolean
If MatchPrefix(FunctionName, "PFNGLGEN", False) Or _
   MatchPrefix(FunctionName, "PFNWGLGEN", False) Or _
   MatchPrefix(FunctionName, "glGen", True) Or _
   MatchPrefix(FunctionName, "wglGen", True) Then TheFunctionIsglGenSomething = True
End Function

'Rename the parameters if necessary
Function DoParamRenameCSharp(Param As String, Optional DefaultName As String = "param", Optional ByVal FunctionName As String = "", Optional OutCallingParam As String = "", Optional ByVal NoMarshal As Boolean = False) As String
Dim Trimmed As String
Dim HaveConst As Boolean
If InStr(Param, "const ") Then HaveConst = True
If InStr(Param, "const*") Then HaveConst = True
Trimmed = Trim$(Replace(Param, "const ", ""))
Trimmed = Trim$(Replace(Trimmed, "const*", "*"))
Trimmed = Replace(Trimmed, "*", "* ")
Trimmed = Replace(Trimmed, "  ", " ")
Trimmed = Replace(Trimmed, " *", "*")

Dim MarshalType As String, PassType As String, ParamType As String, ParamName As String
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
    Case "ref": IsReserved = True
    Case "params": IsReserved = True
    Case "base": IsReserved = True
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

'Determine if using IntPtr is needed or just use ref
Select Case GetPointerLevel(ParamType)
Case 0
    ParamType = DoTypeConvCSharp(ParamType)
Case 1
    Select Case ParamType
    Case "void*"
        ParamType = "IntPtr"
    Case "GLvoid*"
        ParamType = "IntPtr"
    Case "char*"
        MarshalType = "[MarshalAs(UnmanagedType.LPStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case "GLchar*"
        MarshalType = "[MarshalAs(UnmanagedType.LPStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case "wchar_t*"
        MarshalType = "[MarshalAs(UnmanagedType.LPWStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case Else
        If HaveConst Then
            MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
            ParamType = DoTypeConvCSharp(Replace(ParamType, "*", "")) & "[]"
        Else
            PassType = "ref "
            ParamType = DoTypeConvCSharp(Replace(ParamType, "*", ""))
        End If
    End Select
Case 2
    Select Case ParamType
    Case "char**"
        ParamType = "string[]"
        If HaveConst = False Then PassType = "ref "
    Case "GLchar**"
        MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
        ParamType = "string[]"
        If HaveConst = False Then PassType = "ref "
    Case "wchar_t**"
        MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
        ParamType = "IntPtr[]"
        If HaveConst = False Then PassType = "ref "
    Case Else
        ParamType = "IntPtr"
        If HaveConst = False Then PassType = "ref "
    End Select
Case Else
    PassType = "ref "
    ParamType = "IntPtr"
End Select

Select Case ParamType
Case "Object"
    MarshalType = "[MarshalAs(UnmanagedType.AsAny)] "
Case "bool"
    MarshalType = "[MarshalAs(UnmanagedType.Bool)] "
End Select

OutCallingParam = PassType & ParamName
If NoMarshal = False Then DoParamRenameCSharp = MarshalType
DoParamRenameCSharp = DoParamRenameCSharp & PassType & ParamType & " " & ParamName
End Function

'Rename the parameters if necessary
Function DoParamRenameCSharpUnsafe(Param As String, Optional DefaultName As String = "param", Optional ByVal FunctionName As String = "", Optional OutCallingParam As String = "", Optional ByVal NoMarshal As Boolean = False) As String
Dim Trimmed As String
Dim HaveConst As Boolean
If InStr(Param, "const ") Then HaveConst = True
If InStr(Param, "const*") Then HaveConst = True
Trimmed = Trim$(Replace(Param, "const ", ""))
Trimmed = Trim$(Replace(Trimmed, "const*", "*"))
Trimmed = Replace(Trimmed, "*", "* ")
Trimmed = Replace(Trimmed, "  ", " ")
Trimmed = Replace(Trimmed, " *", "*")

Dim MarshalType As String, PassType As String, ParamType As String, ParamName As String
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
    Case "ref": IsReserved = True
    Case "params": IsReserved = True
    Case "base": IsReserved = True
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
    ParamType = DoTypeConvCSharpUnsafe(ParamType)
Case 1
    Select Case ParamType
    Case "void*"
        ParamType = "void*"
    Case "GLvoid*"
        ParamType = "void*"
    Case "char*"
        MarshalType = "[MarshalAs(UnmanagedType.LPStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case "GLchar*"
        MarshalType = "[MarshalAs(UnmanagedType.LPStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case "wchar_t*"
        MarshalType = "[MarshalAs(UnmanagedType.LPWStr)] "
        If HaveConst Then
            ParamType = "string"
        Else
            ParamType = "StringBuilder"
        End If
    Case Else
        ParamType = DoTypeConvCSharpUnsafe(Replace(ParamType, "*", "")) & "*"
    End Select
Case 2
    Select Case ParamType
    Case "char**"
        MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
        ParamType = "string[]"
        If HaveConst = False Then PassType = "ref "
    Case "GLchar**"
        MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
        ParamType = "string[]"
        If HaveConst = False Then PassType = "ref "
    Case "wchar_t**"
        MarshalType = "[MarshalAs(UnmanagedType.LPArray)] "
        ParamType = "IntPtr[]"
        If HaveConst = False Then PassType = "ref "
    Case Else
        ParamType = DoTypeConvCSharpUnsafe(Replace(ParamType, "*", "")) & "**"
    End Select
Case Else
    ParamType = DoTypeConvCSharpUnsafe(Replace(ParamType, "*", "")) & String(GetPointerLevel(ParamType), "*")
End Select

Select Case ParamType
Case "Object"
    MarshalType = "[MarshalAs(UnmanagedType.AsAny)] "
Case "bool"
    MarshalType = "[MarshalAs(UnmanagedType.Bool)] "
End Select

OutCallingParam = PassType & ParamName
If NoMarshal = False Then DoParamRenameCSharpUnsafe = MarshalType
DoParamRenameCSharpUnsafe = DoParamRenameCSharpUnsafe & PassType & ParamType & " " & ParamName
End Function

'Do type conversion but only for the return type
Function DoRetTypeConvCSharpUnsafe(RetType As String) As String
Dim Trimmed As String
Trimmed = Trim$(Replace(RetType, "const ", ""))

If InStr(Trimmed, "*") Then
    Trimmed = Replace(Trimmed, " *", "*")
    Select Case Trimmed
    'You don't want the memory allocated from an API as a return value freed by CLR
    Case "GLubyte*"
        DoRetTypeConvCSharpUnsafe = "byte*"
    Case "GLchar*"
        DoRetTypeConvCSharpUnsafe = "byte*"
    Case "char*"
        DoRetTypeConvCSharpUnsafe = "byte*"
    Case "void*"
        DoRetTypeConvCSharpUnsafe = "void*"
    Case "wchar_t*"
        DoRetTypeConvCSharpUnsafe = "char*"
    Case Else
        DoRetTypeConvCSharpUnsafe = "void*"
        Debug.Print Trimmed
    End Select
Else
    DoRetTypeConvCSharpUnsafe = DoTypeConvCSharpUnsafe(Trimmed)
End If
End Function

Function GuessConstMacroDefType(MacroDefs As Dictionary, MacroName, OutMacroDef As String) As String
Dim MacroDef As String
MacroDef = MacroDefs(MacroName)
OutMacroDef = MacroDef
Do While MacroDefs.Exists(MacroDef)
    MacroDef = MacroDefs(MacroDef)
Loop
If IsNumeric(MacroDef) Then
    If Int(MacroDef) = Val(MacroDef) And Abs(MacroDef) <= 4294967295# Then
        GuessConstMacroDefType = "int "
    Else
        GuessConstMacroDefType = "double "
    End If
ElseIf Left$(MacroDef, 2) = "0x" Then
    If Right$(MacroDef, 3) = "ull" Then
        GuessConstMacroDefType = "ulong "
        If OutMacroDef = MacroDef Then OutMacroDef = Left$(OutMacroDef, Len(OutMacroDef) - 3)
    ElseIf Right$(MacroDef, 1) = "u" Or Right$(MacroDef, 2) = "ul" Then
        GuessConstMacroDefType = "uint "
    Else
        GuessConstMacroDefType = "int "
        Dim Value As Long
        Value = CLng("&H" & Mid$(MacroDef, 3))
        If Value < 0 Then OutMacroDef = "unchecked((int)" & MacroDef & ")"
    End If
Else
    GuessConstMacroDefType = "string "
End If
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
Dim LabelUsed As Boolean

Dim FN As Integer
FN = FreeFile
Open ExportTo & "\modGL_API.vb" For Output As #FN
Print #FN, "Imports System.Runtime.InteropServices"
Print #FN, "Imports System.Text"
Print #FN,
Print #FN, "Module GL_API"
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
            Tail = ") As " & DoRetTypeConvVB_NET(FuncData(0))
            Print #FN, vbTab; "Public Delegate Function ";
        End If
        
        'Function name
        Print #FN, FuncTypeDefName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRenameVB_NET(Param(I), "param" & I);
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
            Tail = ") As " & DoRetTypeConvVB_NET(FuncData(0))
            Print #FN, vbTab; "Declare Function ";
        End If
        
        'Function name and Dll name
        Print #FN, FuncName; " Lib """ & Ext.API_DllName & """ (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRenameVB_NET(Param(I), "param" & I);
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
Print #FN,
Print #FN, vbTab; "Private Function GLAPI_CreateIndexOfExtensionDict() As Dictionary(Of String, UInteger)"
Print #FN, vbTab; vbTab; "Dim RetDict As New Dictionary(Of String, UInteger)"
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; vbTab; "RetDict.Add("""; GLExtString; """, "; CStr(I); ")"
    I = I + 1
Next
Print #FN, vbTab; vbTab; "Return RetDict"
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
Print #FN, vbTab; vbTab; "Dim IndexOfExtensionDict As Dictionary(Of String, UInteger) = GLAPI_CreateIndexOfExtensionDict()"
Print #FN, vbTab; vbTab; "Dim Extensions_Supported("; CStr(ExtCount - 1); ") As Boolean"
Print #FN,
Print #FN, vbTab; vbTab; "VersionString = Marshal.PtrToStringAnsi(glGetString(GL_VERSION))"
Print #FN, vbTab; vbTab; "VendorSplit = Split(VersionString)"
Print #FN, vbTab; vbTab; "VersionSplit = Split(VendorSplit(0), ""."")"
Print #FN, vbTab; vbTab; "Major = VersionSplit(0)"
Print #FN, vbTab; vbTab; "Minor = VersionSplit(1)"
Print #FN, vbTab; vbTab; "If Major = 1 And Minor = 0 Then Return"
Print #FN,
Print #FN, vbTab; vbTab; "GL_VERSION_4_6 = (Major > 4) OrElse ((Major = 4) AndAlso (Minor >= 6))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_5 = GL_VERSION_4_6 OrElse ((Major = 4) AndAlso (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_4 = GL_VERSION_4_5 OrElse ((Major = 4) AndAlso (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_3 = GL_VERSION_4_4 OrElse ((Major = 4) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_2 = GL_VERSION_4_3 OrElse ((Major = 4) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_1 = GL_VERSION_4_2 OrElse ((Major = 4) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_0 = GL_VERSION_4_1 OrElse ((Major = 4) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_3 = GL_VERSION_4_0 OrElse ((Major = 3) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_2 = GL_VERSION_3_3 OrElse ((Major = 3) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_1 = GL_VERSION_3_2 OrElse ((Major = 3) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_0 = GL_VERSION_3_1 OrElse ((Major = 3) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_1 = GL_VERSION_3_0 OrElse ((Major = 2) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_0 = GL_VERSION_2_1 OrElse ((Major = 2) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_5 = GL_VERSION_2_0 OrElse ((Major = 1) AndAlso (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_4 = GL_VERSION_1_5 OrElse ((Major = 1) AndAlso (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_3 = GL_VERSION_1_4 OrElse ((Major = 1) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2_1 = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2 = GL_VERSION_1_2_1 OrElse ((Major = 1) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_1 = GL_VERSION_1_2 OrElse ((Major = 1) AndAlso (Minor >= 1))"
Print #FN,
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_6"")) = GL_VERSION_4_6"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_5"")) = GL_VERSION_4_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_4"")) = GL_VERSION_4_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_3"")) = GL_VERSION_4_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_2"")) = GL_VERSION_4_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_1"")) = GL_VERSION_4_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_0"")) = GL_VERSION_4_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_3"")) = GL_VERSION_3_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_2"")) = GL_VERSION_3_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_1"")) = GL_VERSION_3_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_0"")) = GL_VERSION_3_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_2_1"")) = GL_VERSION_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_2_0"")) = GL_VERSION_2_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_5"")) = GL_VERSION_1_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_4"")) = GL_VERSION_1_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_3"")) = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_2_1"")) = GL_VERSION_1_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_2"")) = GL_VERSION_1_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_1"")) = GL_VERSION_1_1"
Print #FN,
Print #FN, vbTab; vbTab; "If GL_VERSION_3_0 Then"
Print #FN, vbTab; vbTab; vbTab; "glGetStringi = Marshal.GetDelegateForFunctionPointer(wglGetProcAddress(""glGetStringi""), GetType(PFNGLGETSTRINGIPROC))"
Print #FN, vbTab; vbTab; vbTab; "glGetIntegerv(GL_NUM_EXTENSIONS, ExtCount)"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtString = Marshal.PtrToStringAnsi(glGetStringi(GL_EXTENSIONS, I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If IndexOfExtensionDict.ContainsKey(ExtString) = False Then Continue For"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict(ExtString)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Else"
Print #FN, vbTab; vbTab; vbTab; "Dim ExtStrings() As String"
Print #FN, vbTab; vbTab; vbTab; "ExtStrings = Split(Marshal.PtrToStringAnsi(glGetString(GL_EXTENSIONS)))"
Print #FN, vbTab; vbTab; vbTab; "ExtCount = UBound(ExtStrings) + 1"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If IndexOfExtensionDict.ContainsKey(ExtStrings(I)) = False Then Continue For"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict(ExtStrings(I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "End If"
Print #FN,
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; vbTab; "If Extensions_Supported(IndexOfExtensionDict("""; GLExtString; """)) Then "; GLExtString; " = GLAPI_Init_"; GLExtString; "() Else "; GLExtString; " = False"
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
Dim LabelUsed As Boolean

Dim FN As Integer
FN = FreeFile
Open ExportTo & "\GL_API.vb" For Output As #FN
Print #FN, "Imports System.Runtime.InteropServices"
Print #FN, "Imports System.Text"
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
            Tail = ") As " & DoRetTypeConvVB_NET(FuncData(0))
            Print #FN, vbTab; "Public Delegate Function ";
        End If
        
        'Function name
        Print #FN, FuncTypeDefName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRenameVB_NET(Param(I), "param" & I);
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
            Tail = ") As " & DoRetTypeConvVB_NET(FuncData(0))
            Print #FN, vbTab; "Declare Function ";
        End If
        
        'Function name and Dll name
        Print #FN, FuncName; " Lib """ & Ext.API_DllName & """ (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRenameVB_NET(Param(I), "param" & I);
            Next
        End If
        
        Print #FN, Tail
    Next
    If HasAPI Then Print #FN,
    
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
Print #FN,
Print #FN, vbTab; "Private Function GLAPI_CreateIndexOfExtensionDict() As Dictionary(Of String, UInteger)"
Print #FN, vbTab; vbTab; "Dim RetDict As New Dictionary(Of String, UInteger)"
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    Print #FN, vbTab; vbTab; "RetDict.Add("""; GLExtString; """, "; CStr(I); ")"
    I = I + 1
Next
Print #FN, vbTab; vbTab; "Return RetDict"
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
Print #FN, vbTab; vbTab; "Dim IndexOfExtensionDict As Dictionary(Of String, UInteger) = GLAPI_CreateIndexOfExtensionDict()"
Print #FN, vbTab; vbTab; "Dim Extensions_Supported("; CStr(ExtCount - 1); ") As Boolean"
Print #FN,
Print #FN, vbTab; vbTab; "VersionString = Marshal.PtrToStringAnsi(glGetString(GL_VERSION))"
Print #FN, vbTab; vbTab; "VendorSplit = Split(VersionString)"
Print #FN, vbTab; vbTab; "VersionSplit = Split(VendorSplit(0), ""."")"
Print #FN, vbTab; vbTab; "Major = VersionSplit(0)"
Print #FN, vbTab; vbTab; "Minor = VersionSplit(1)"
Print #FN, vbTab; vbTab; "If Major = 1 And Minor = 0 Then Return"
Print #FN,
Print #FN, vbTab; vbTab; "GL_VERSION_4_6 = (Major > 4) OrElse ((Major = 4) AndAlso (Minor >= 6))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_5 = GL_VERSION_4_6 OrElse ((Major = 4) AndAlso (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_4 = GL_VERSION_4_5 OrElse ((Major = 4) AndAlso (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_3 = GL_VERSION_4_4 OrElse ((Major = 4) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_2 = GL_VERSION_4_3 OrElse ((Major = 4) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_1 = GL_VERSION_4_2 OrElse ((Major = 4) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_4_0 = GL_VERSION_4_1 OrElse ((Major = 4) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_3 = GL_VERSION_4_0 OrElse ((Major = 3) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_2 = GL_VERSION_3_3 OrElse ((Major = 3) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_1 = GL_VERSION_3_2 OrElse ((Major = 3) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_3_0 = GL_VERSION_3_1 OrElse ((Major = 3) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_1 = GL_VERSION_3_0 OrElse ((Major = 2) AndAlso (Minor >= 1))"
Print #FN, vbTab; vbTab; "GL_VERSION_2_0 = GL_VERSION_2_1 OrElse ((Major = 2) AndAlso (Minor >= 0))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_5 = GL_VERSION_2_0 OrElse ((Major = 1) AndAlso (Minor >= 5))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_4 = GL_VERSION_1_5 OrElse ((Major = 1) AndAlso (Minor >= 4))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_3 = GL_VERSION_1_4 OrElse ((Major = 1) AndAlso (Minor >= 3))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2_1 = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "GL_VERSION_1_2 = GL_VERSION_1_2_1 OrElse ((Major = 1) AndAlso (Minor >= 2))"
Print #FN, vbTab; vbTab; "GL_VERSION_1_1 = GL_VERSION_1_2 OrElse ((Major = 1) AndAlso (Minor >= 1))"
Print #FN,
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_6"")) = GL_VERSION_4_6"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_5"")) = GL_VERSION_4_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_4"")) = GL_VERSION_4_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_3"")) = GL_VERSION_4_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_2"")) = GL_VERSION_4_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_1"")) = GL_VERSION_4_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_4_0"")) = GL_VERSION_4_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_3"")) = GL_VERSION_3_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_2"")) = GL_VERSION_3_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_1"")) = GL_VERSION_3_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_3_0"")) = GL_VERSION_3_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_2_1"")) = GL_VERSION_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_2_0"")) = GL_VERSION_2_0"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_5"")) = GL_VERSION_1_5"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_4"")) = GL_VERSION_1_4"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_3"")) = GL_VERSION_1_3"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_2_1"")) = GL_VERSION_1_2_1"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_2"")) = GL_VERSION_1_2"
Print #FN, vbTab; vbTab; "Extensions_Supported(IndexOfExtensionDict(""GL_VERSION_1_1"")) = GL_VERSION_1_1"
Print #FN,
Print #FN, vbTab; vbTab; "If GL_VERSION_3_0 Then"
Print #FN, vbTab; vbTab; vbTab; "glGetStringi = Marshal.GetDelegateForFunctionPointer(wglGetProcAddress(""glGetStringi""), GetType(PFNGLGETSTRINGIPROC))"
Print #FN, vbTab; vbTab; vbTab; "glGetIntegerv(GL_NUM_EXTENSIONS, ExtCount)"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtString = Marshal.PtrToStringAnsi(glGetStringi(GL_EXTENSIONS, I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If IndexOfExtensionDict.ContainsKey(ExtString) = False Then Continue For"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict(ExtString)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then Extensions_Supported(ExtIndex) = True"
Print #FN, vbTab; vbTab; vbTab; "Next"
Print #FN, vbTab; vbTab; "Else"
Print #FN, vbTab; vbTab; vbTab; "Dim ExtStrings() As String"
Print #FN, vbTab; vbTab; vbTab; "ExtStrings = Split(Marshal.PtrToStringAnsi(glGetString(GL_EXTENSIONS)))"
Print #FN, vbTab; vbTab; vbTab; "ExtCount = UBound(ExtStrings) + 1"
Print #FN, vbTab; vbTab; vbTab; "For I = 0 To ExtCount - 1"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If IndexOfExtensionDict.ContainsKey(ExtStrings(I)) = False Then Continue For"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict(ExtStrings(I))"
Print #FN, vbTab; vbTab; vbTab; vbTab; "If ExtIndex >= 0 Then Extensions_Supported(ExtIndex) = True"
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
    Print #FN, vbTab; "If Extensions_Supported(IndexOfExtensionDict("""; GLExtString; """)) Then"
    LabelUsed = False
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        FuncName = Ext.FuncPtrs(FuncPtrType)
        Print #FN, vbTab; vbTab; "FuncPtr = wglGetProcAddress("""; FuncName; """)"
        Print #FN, vbTab; vbTab; "If FuncPtr = 0 Then Goto EndOf_"; GLExtString
        Print #FN, vbTab; vbTab; FuncName; " = Marshal.GetDelegateForFunctionPointer(FuncPtr, GetType("; FuncPtrType; "))"
        Print #FN,
        LabelUsed = True
    Next
    Print #FN, vbTab; vbTab; GLExtString; " = True"
    Print #FN, vbTab; "End If"
    If LabelUsed Then Print #FN, vbTab; "EndOf_"; GLExtString; ":"
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

Function SuffixToCSharpType(ByVal Suffix As String) As String
Select Case Suffix
Case "b": SuffixToCSharpType = "sbyte"
Case "ub": SuffixToCSharpType = "byte"
Case "s": SuffixToCSharpType = "short"
Case "us": SuffixToCSharpType = "ushort"
Case "i": SuffixToCSharpType = "int"
Case "ui": SuffixToCSharpType = "uint"
Case "i64": SuffixToCSharpType = "long"
Case "ui64": SuffixToCSharpType = "ulong"
Case "f": SuffixToCSharpType = "float"
Case "d": SuffixToCSharpType = "double"

Case "bv": SuffixToCSharpType = "sbyte*"
Case "ubv": SuffixToCSharpType = "byte*"
Case "sv": SuffixToCSharpType = "short*"
Case "usv": SuffixToCSharpType = "ushort*"
Case "iv": SuffixToCSharpType = "int*"
Case "uiv": SuffixToCSharpType = "uint*"
Case "i64v": SuffixToCSharpType = "long*"
Case "ui64v": SuffixToCSharpType = "ulong*"
Case "fv": SuffixToCSharpType = "float*"
Case "dv": SuffixToCSharpType = "double*"

Case Else
    Debug.Assert False
End Select
End Function

Sub ExportOverloads(ByVal FN As Integer, ByVal IsStatic As Boolean, ReturnType As String, FuncName As String, PreParamDef As String, PreParamCall As String, CallParamVariant As String, Suffixes As String, Optional ByVal Indents As Long = 2, Optional ByVal NoV As Boolean = False, Optional ByVal HideVarCount As Boolean = False, Optional ByVal RefParam As Boolean = False)
Dim Tabs As String, I As Long, J As Long, K As Long, L As Long, U As Long, HasUnsafe As Boolean, ParamType As String
Dim ParamVarCount As Long, CallParamVarList() As String
Dim ParamCount As Long, CallParamLists(), ParamCountStr As String
Dim SuffixList() As String
Dim FuncMod As String
Dim ParamMod As String
Tabs = String(Indents, vbTab)
CallParamVarList = Split(CallParamVariant, "|")
ParamVarCount = UBound(CallParamVarList) + 1
ReDim CallParamLists(ParamVarCount - 1)
For I = 0 To ParamVarCount - 1
    CallParamLists(I) = Split(CallParamVarList(I), ",")
Next

SuffixList = Split(Suffixes, ",")
If NoV Then
    For I = 0 To UBound(SuffixList)
        If Right$(SuffixList(I), 1) = "v" Then SuffixList(I) = Left$(SuffixList(I), Len(SuffixList(I)) - 1)
    Next
End If

If IsStatic Then FuncMod = "static "

For I = 0 To UBound(SuffixList)
    ParamType = SuffixToCSharpType(SuffixList(I))
    HasUnsafe = (Right$(ParamType, 1) = "*") Or NoV
    If NoV Then ParamType = ParamType & "*"
    If HasUnsafe = False Then
        For K = 0 To ParamVarCount - 1
            ParamCount = UBound(CallParamLists(K)) + 1
            Print #FN, Tabs; "public "; FuncMod; ReturnType; " "; FuncName; "("; PreParamDef;
            For J = 0 To ParamCount - 1
                Print #FN, ParamType; " "; CallParamLists(K)(J);
                If J < ParamCount - 1 Then Print #FN, ", ";
            Next
            Print #FN, ")"
            Print #FN, Tabs; "{"
            Print #FN, Tabs; vbTab;
            If ReturnType <> "void" Then Print #FN, "return ";
            Print #FN, FuncName;
            If HideVarCount = False Then Print #FN, CStr(ParamCount);
            Print #FN, SuffixList(I); "("; PreParamCall;
            For J = 0 To ParamCount - 1
                If ParamType = "sbyte" Then Print #FN, "(byte)";
                Print #FN, CallParamLists(K)(J);
                If J < ParamCount - 1 Then Print #FN, ", ";
            Next
            Print #FN, ");"
            Print #FN, Tabs; "}"
        Next
    Else
        ParamType = Left$(ParamType, Len(ParamType) - 1)
        If RefParam Then ParamMod = "ref "
        
        Print #FN, Tabs; "public "; FuncMod; ReturnType; " "; FuncName; "("; PreParamDef; ParamMod; ParamType; "[] v)"
        Print #FN, Tabs; "{"
        If ParamVarCount Then
            Print #FN, Tabs; vbTab; "switch(v.Length)"
            Print #FN, Tabs; vbTab; "{"
            For J = 1 To ParamVarCount
                ParamCount = UBound(CallParamLists(J - 1)) + 1
                Print #FN, Tabs; vbTab; "case "; CStr(ParamCount); ": ";
                If ReturnType <> "void" Then Print #FN, "return ";
                Print #FN, FuncName;
                If HideVarCount = False Then Print #FN, CStr(ParamCount);
                Print #FN, SuffixList(I); "("; PreParamCall; ParamMod;
                If RefParam Then
                    If ParamType = "sbyte" Then Print #FN, "(byte)";
                    Print #FN, "v[0]);";
                Else
                    If ParamType = "sbyte" Then Print #FN, "(byte[])(Array)";
                    Print #FN, "v);";
                End If
                If ReturnType <> "void" Then Print #FN, Else Print #FN, " break;"
            Next
            Print #FN, Tabs; vbTab; "default: System.Diagnostics.Debug.Assert(false); break;"
            Print #FN, Tabs; vbTab; "}"
        Else
            ParamCount = UBound(CallParamLists(0)) + 1
            Print #FN, Tabs; vbTab;
            If ReturnType <> "void" Then Print #FN, "return ";
            If RefParam Then
                Print #FN, FuncName; CStr(ParamCount); SuffixList(I); "("; PreParamCall; ParamMod; "v[0]);"
            Else
                Print #FN, FuncName; CStr(ParamCount); SuffixList(I); "("; PreParamCall; ParamMod; "v);"
            End If
        End If
        Print #FN, Tabs; "}"
        
        Print #FN, Tabs; "#if CSHARPGLEW_ALLOW_UNSAFE"
        For J = 1 To ParamVarCount
            ParamCount = UBound(CallParamLists(J - 1)) + 1
            Print #FN, Tabs; "public "; FuncMod; "unsafe "; ReturnType; " "; FuncName;
            If HideVarCount = False Then Print #FN, CStr(ParamCount);
            Print #FN, "("; PreParamDef; ParamType; "* v)"
            Print #FN, Tabs; "{"
            Print #FN, Tabs; vbTab;
            If ReturnType <> "void" Then Print #FN, "return ";
            Print #FN, FuncName;
            If HideVarCount = False Then Print #FN, CStr(ParamCount);
            Print #FN, SuffixList(I); "("; PreParamCall;
            If ParamType = "sbyte" Then Print #FN, "(byte*)";
            Print #FN, "v);"
            Print #FN, Tabs; "}"
        Next
        Print #FN, Tabs; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
    End If
Next
End Sub

Sub ExportCSharp(Parser As clsParser, ExportTo As String)
Dim GLExtString
Dim Ext As clsGLExtension
Dim MacroName
Dim MacroDef As String
Dim FuncTypeDefName
Dim FuncPtrType
Dim FuncName
Dim ConvertedParam As String
Dim I As Long, ExtCount As Long
Dim FuncData() As String
Dim Param() As String
Dim ParamName As String
Dim AlreadyDefinedMacros As New Dictionary
Dim LabelUsed As Boolean
Dim HasUnsafe As New Dictionary

Dim FN As Integer
FN = FreeFile
Open ExportTo & "\GLAPI.cs" For Output As #FN
Print #FN, "#define CSHARPGLEW_ALLOW_UNSAFE"
Print #FN, "#undef CSHARPGLEW_ALLOW_UNSAFE"
Print #FN,
Print #FN, "using System;"
Print #FN, "using System.Text;"
Print #FN, "using System.Collections.Generic;"
Print #FN, "using System.Runtime.InteropServices;"
Print #FN,
Print #FN, "#pragma warning disable IDE1006, CS0649"
Print #FN,
Print #FN, "namespace CSharpGLEW"
Print #FN, "{"
Print #FN,
Print #FN, vbTab; "partial class GLAPI"
Print #FN, vbTab; "{"

Print #FN, vbTab; vbTab; "#region ""OpenGL Extension Declerations"""
    
For Each GLExtString In Parser.GLExtension.Keys
    Dim HasMacro As Boolean, HasAPI As Boolean, HasFuncPtr As Boolean
    HasMacro = False
    HasAPI = False
    HasFuncPtr = False
    Set Ext = Parser.GLExtension(GLExtString)

    Print #FN, vbTab; vbTab; "#region """; GLExtString; """"
    Print #FN, vbTab; vbTab; "// ----------------------------- "; GLExtString; " -----------------------------"
    Print #FN,
    
    'A variable tells you if this extension is available
    Print #FN, vbTab; vbTab; "public readonly bool "; GLExtString; " = false;"
    Print #FN,
    
    'Macro definitions
    For Each MacroName In Ext.MacroDefs.Keys
        HasMacro = True
        'Make sure no duplicated
        If AlreadyDefinedMacros.Exists(MacroName) = False Then
            AlreadyDefinedMacros.Add MacroName, Ext.MacroDefs(MacroName)
            Print #FN, vbTab; vbTab; "public const "; GuessConstMacroDefType(AlreadyDefinedMacros, MacroName, MacroDef); MacroName; " = "; MacroDef; ";"
        Else
            Print #FN, vbTab; vbTab; "//public const "; GuessConstMacroDefType(AlreadyDefinedMacros, MacroName, MacroDef); MacroName; " = "; MacroDef; ";"
        End If
    Next
    If HasMacro Then Print #FN,
    
    'Function typedefs
    For Each FuncTypeDefName In Ext.FuncTypeDef.Keys
        FuncData = Split(Ext.FuncTypeDef(FuncTypeDefName), ":")
        Print #FN, vbTab; vbTab; "[System.Security.SuppressUnmanagedCodeSecurity()]"
        Print #FN, vbTab; vbTab; "public delegate "; DoRetTypeConvCSharp(FuncData(0)); " "; FuncTypeDefName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                ConvertedParam = DoParamRenameCSharp(Param(I), "param" & I, FuncTypeDefName)
                Print #FN, ConvertedParam;
                If InStr(ConvertedParam, "*") <> 0 Or InStr(DoParamRenameCSharpUnsafe(Param(I), "param" & I, FuncTypeDefName), "*") <> 0 Then HasUnsafe(FuncTypeDefName) = True
            Next
        End If
        Print #FN, ");"
        
        If HasUnsafe(FuncTypeDefName) = True Then
            Print #FN, vbTab; vbTab; "#if CSHARPGLEW_ALLOW_UNSAFE"
            Print #FN, vbTab; vbTab; "[System.Security.SuppressUnmanagedCodeSecurity()]"
            Print #FN, vbTab; vbTab; "public unsafe delegate "; DoRetTypeConvCSharpUnsafe(FuncData(0)); " "; FuncTypeDefName; "_UNSAFE (";
            
            'Parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    Print #FN, DoParamRenameCSharpUnsafe(Param(I), "param" & I, FuncTypeDefName);
                Next
            End If
            Print #FN, ");"
            Print #FN, vbTab; vbTab; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
        End If
        Print #FN,
    Next
    
    'Function pointer declaration
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        HasFuncPtr = True
        FuncName = Ext.FuncPtrs(FuncPtrType)
        If HasUnsafe(FuncPtrType) Then
            Print #FN, vbTab; vbTab; "public readonly "; FuncPtrType; " "; FuncName; "_Safe;"
            Print #FN, vbTab; vbTab; "#if CSHARPGLEW_ALLOW_UNSAFE"
            Print #FN, vbTab; vbTab; "public readonly "; FuncPtrType; "_UNSAFE "; FuncName; "_Unsafe;"
            Print #FN, vbTab; vbTab; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
            
            FuncData = Split(Ext.FuncTypeDef(FuncPtrType), ":")
            
            Print #FN, vbTab; vbTab; "public "; DoRetTypeConvCSharp(FuncData(0)); " "; FuncName; " (";
            
            'Parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    Print #FN, DoParamRenameCSharp(Param(I), "param" & I, FuncName, , True);
                Next
            End If
            
            Print #FN, ")"
            Print #FN, vbTab; vbTab; "{ ";
            If LCase$(FuncData(0)) <> "void" Then
                Print #FN, "return ";
            End If
            Print #FN, FuncName; "_Safe (";
            'Calling parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    DoParamRenameCSharp Param(I), "param" & I, FuncName, ParamName
                    Print #FN, ParamName;
                Next
            End If
            Print #FN, "); }"
            
            
            Print #FN, vbTab; vbTab; "#if CSHARPGLEW_ALLOW_UNSAFE"
            Print #FN, vbTab; vbTab; "public unsafe "; DoRetTypeConvCSharpUnsafe(FuncData(0)); " "; FuncName; " (";
            
            'Parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    Print #FN, DoParamRenameCSharpUnsafe(Param(I), "param" & I, FuncName, , True);
                Next
            End If
            
            Print #FN, ")"
            Print #FN, vbTab; vbTab; "{ ";
            If LCase$(FuncData(0)) <> "void" Then
                Print #FN, "return ";
            End If
            Print #FN, FuncName; "_Unsafe (";
            'Calling parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    DoParamRenameCSharpUnsafe Param(I), "param" & I, FuncName, ParamName
                    Print #FN, ParamName;
                Next
            End If
            Print #FN, "); }"
            Print #FN, vbTab; vbTab; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
            
        Else
            Print #FN, vbTab; vbTab; "public readonly "; FuncPtrType; " "; FuncName; ";"
        End If
    Next
    If HasFuncPtr Then Print #FN,
    
    'API declaration
    For Each FuncName In Ext.APIs.Keys
        HasAPI = True
        FuncData = Split(Ext.APIs(FuncName), ":")
        Print #FN, vbTab; vbTab; "[DllImport("""; Ext.API_DllName; """)]"
        Print #FN, vbTab; vbTab; "public static extern "; DoRetTypeConvCSharp(FuncData(0)); " "; FuncName; " (";
        
        'Parameter list
        If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
            Param = Split(FuncData(1), ",")
            For I = 0 To UBound(Param)
                If I Then Print #FN, ", ";
                Print #FN, DoParamRenameCSharp(Param(I), "param" & I, FuncTypeDefName);
            Next
        End If
        Print #FN, ");"
        
        If InStr(FuncData(1), "*") Then
            Print #FN, vbTab; vbTab; "#if CSHARPGLEW_ALLOW_UNSAFE"
            Print #FN, vbTab; vbTab; "[DllImport("""; Ext.API_DllName; """)]"
            Print #FN, vbTab; vbTab; "public static extern unsafe "; DoRetTypeConvCSharpUnsafe(FuncData(0)); " "; FuncName; " (";
            
            'Parameter list
            If Len(FuncData(1)) > 0 And LCase$(FuncData(1)) <> "void" Then
                Param = Split(FuncData(1), ",")
                For I = 0 To UBound(Param)
                    If I Then Print #FN, ", ";
                    Print #FN, DoParamRenameCSharpUnsafe(Param(I), "param" & I, FuncTypeDefName);
                Next
            End If
            
            Print #FN, ");"
            Print #FN, vbTab; vbTab; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
        End If
    Next
    If HasAPI Then Print #FN,
    
    Print #FN, vbTab; vbTab; "#endregion"
    ExtCount = ExtCount + 1
Next
Print #FN, vbTab; vbTab; "#endregion"
Print #FN,
Print #FN, vbTab; vbTab; "#region ""OpenGL Extension Strings"""
Print #FN, vbTab; vbTab; "public readonly string[] GLAPI_Extensions ="
Print #FN, vbTab; vbTab; "{"
Dim LinePos As Long
LinePos = 0
I = 0
Print #FN, vbTab; vbTab; vbTab;
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
        Print #FN, vbTab; vbTab; vbTab;
        LinePos = 0
    End If
    I = I + 1
Next
Print #FN, vbTab; vbTab; "};"
Print #FN, vbTab; vbTab; "#endregion"
Print #FN,
Print #FN, vbTab; vbTab; "#region ""OpenGL Extension String Indices"""
Print #FN, vbTab; vbTab; "private Dictionary<string, uint> GLAPI_CreateIndexOfExtensionDict()"
Print #FN, vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; "var RetDict = new Dictionary<string, uint>()"
Print #FN, vbTab; vbTab; vbTab; "{"
I = 0
For Each GLExtString In Parser.GLExtension.Keys
    If I Then Print #FN, ", ";
    If I Mod 16 = 0 Then
        If I Then Print #FN,
        Print #FN, vbTab; vbTab; vbTab; vbTab;
    End If
    Print #FN, "{"""; GLExtString; """, "; CStr(I); "}";
    I = I + 1
Next
Print #FN,
Print #FN, vbTab; vbTab; vbTab; "};"
Print #FN, vbTab; vbTab; vbTab; "return RetDict;"
Print #FN, vbTab; vbTab; "}"
Print #FN, vbTab; vbTab; "#endregion"
Print #FN,
Print #FN, vbTab; vbTab; "#region ""OpenGL Extension Initialize"""
Print #FN,
Print #FN, vbTab; vbTab; "public GLAPI()"
Print #FN, vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; "uint i;"
Print #FN, vbTab; vbTab; vbTab; "int ExtCount = 0;"
Print #FN, vbTab; vbTab; vbTab; "uint ExtIndex;"
Print #FN, vbTab; vbTab; vbTab; "string ExtString, VersionString;"
Print #FN, vbTab; vbTab; vbTab; "string[] VendorSplit;"
Print #FN, vbTab; vbTab; vbTab; "string[] VersionSplit;"
Print #FN, vbTab; vbTab; vbTab; "int Major, Minor;"
Print #FN, vbTab; vbTab; vbTab; "Dictionary<string, uint> IndexOfExtensionDict = GLAPI_CreateIndexOfExtensionDict();"
Print #FN, vbTab; vbTab; vbTab; "bool[] Extensions_Supported = new bool["; CStr(ExtCount); "];"
Print #FN,
Print #FN, vbTab; vbTab; vbTab; "VersionString = Marshal.PtrToStringAnsi(glGetString(GL_VERSION));"
Print #FN, vbTab; vbTab; vbTab; "VendorSplit = VersionString.Split();"
Print #FN, vbTab; vbTab; vbTab; "VersionSplit = VendorSplit[0].Split('.');"
Print #FN, vbTab; vbTab; vbTab; "Major = Convert.ToInt32(VersionSplit[0]);"
Print #FN, vbTab; vbTab; vbTab; "Minor = Convert.ToInt32(VersionSplit[1]);"
Print #FN, vbTab; vbTab; vbTab; "if (Major <= 1 && Minor == 0 ) return;"
Print #FN,
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_6 = (Major > 4) || ((Major == 4) && (Minor >= 6));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_5 = GL_VERSION_4_6 || ((Major == 4) && (Minor >= 5));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_4 = GL_VERSION_4_5 || ((Major == 4) && (Minor >= 4));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_3 = GL_VERSION_4_4 || ((Major == 4) && (Minor >= 3));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_2 = GL_VERSION_4_3 || ((Major == 4) && (Minor >= 2));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_1 = GL_VERSION_4_2 || ((Major == 4) && (Minor >= 1));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_4_0 = GL_VERSION_4_1 || ((Major == 4) && (Minor >= 0));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_3_3 = GL_VERSION_4_0 || ((Major == 3) && (Minor >= 3));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_3_2 = GL_VERSION_3_3 || ((Major == 3) && (Minor >= 2));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_3_1 = GL_VERSION_3_2 || ((Major == 3) && (Minor >= 1));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_3_0 = GL_VERSION_3_1 || ((Major == 3) && (Minor >= 0));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_2_1 = GL_VERSION_3_0 || ((Major == 2) && (Minor >= 1));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_2_0 = GL_VERSION_2_1 || ((Major == 2) && (Minor >= 0));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_5 = GL_VERSION_2_0 || ((Major == 1) && (Minor >= 5));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_4 = GL_VERSION_1_5 || ((Major == 1) && (Minor >= 4));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_3 = GL_VERSION_1_4 || ((Major == 1) && (Minor >= 3));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_2_1 = GL_VERSION_1_3;"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_2 = GL_VERSION_1_2_1 || ((Major == 1) && (Minor >= 2));"
Print #FN, vbTab; vbTab; vbTab; "GL_VERSION_1_1 = GL_VERSION_1_2 || ((Major == 1) && (Minor >= 1));"
Print #FN,
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_6""]] = GL_VERSION_4_6;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_5""]] = GL_VERSION_4_5;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_4""]] = GL_VERSION_4_4;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_3""]] = GL_VERSION_4_3;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_2""]] = GL_VERSION_4_2;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_1""]] = GL_VERSION_4_1;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_4_0""]] = GL_VERSION_4_0;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_3_3""]] = GL_VERSION_3_3;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_3_2""]] = GL_VERSION_3_2;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_3_1""]] = GL_VERSION_3_1;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_3_0""]] = GL_VERSION_3_0;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_2_1""]] = GL_VERSION_2_1;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_2_0""]] = GL_VERSION_2_0;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_5""]] = GL_VERSION_1_5;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_4""]] = GL_VERSION_1_4;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_3""]] = GL_VERSION_1_3;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_2_1""]] = GL_VERSION_1_2_1;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_2""]] = GL_VERSION_1_2;"
Print #FN, vbTab; vbTab; vbTab; "Extensions_Supported[IndexOfExtensionDict[""GL_VERSION_1_1""]] = GL_VERSION_1_1;"
Print #FN,
Print #FN, vbTab; vbTab; vbTab; "if (GL_VERSION_3_0)"
Print #FN, vbTab; vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; vbTab; "glGetStringi = (PFNGLGETSTRINGIPROC)Marshal.GetDelegateForFunctionPointer(wglGetProcAddress(""glGetStringi""), typeof(PFNGLGETSTRINGIPROC));"
Print #FN, vbTab; vbTab; vbTab; vbTab; "glGetIntegerv(GL_NUM_EXTENSIONS, ref ExtCount);"
Print #FN, vbTab; vbTab; vbTab; vbTab; "for (i = 0; i < ExtCount; i++)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "ExtString = Marshal.PtrToStringAnsi(glGetStringi(GL_EXTENSIONS, i));"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "if (!IndexOfExtensionDict.ContainsKey(ExtString)) continue;"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict[ExtString];"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "if (ExtIndex >= 0) Extensions_Supported[ExtIndex] = true;"
Print #FN, vbTab; vbTab; vbTab; vbTab; "}"
Print #FN, vbTab; vbTab; vbTab; "}"
Print #FN, vbTab; vbTab; vbTab; "else"
Print #FN, vbTab; vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; vbTab; "string[] ExtStrings;"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtStrings = Marshal.PtrToStringAnsi(glGetString(GL_EXTENSIONS)).Split();"
Print #FN, vbTab; vbTab; vbTab; vbTab; "ExtCount = ExtStrings.Length;"
Print #FN, vbTab; vbTab; vbTab; vbTab; "for (i = 0; i < ExtCount; i++)"
Print #FN, vbTab; vbTab; vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "if (!IndexOfExtensionDict.ContainsKey(ExtStrings[i])) continue;"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "ExtIndex = IndexOfExtensionDict[ExtStrings[i]];"
Print #FN, vbTab; vbTab; vbTab; vbTab; vbTab; "if (ExtIndex >= 0) Extensions_Supported[ExtIndex] = true;"
Print #FN, vbTab; vbTab; vbTab; vbTab; "}"
Print #FN, vbTab; vbTab; vbTab; "}"
Print #FN,
I = 0
Print #FN, vbTab; vbTab; vbTab; "IntPtr FuncPtr;"
For Each GLExtString In Parser.GLExtension.Keys
    Set Ext = Parser.GLExtension(GLExtString)

    Print #FN, vbTab; vbTab; vbTab; "#region """; GLExtString; "_Initialize"""
    Print #FN, vbTab; vbTab; vbTab; "// ----------------------------- "; GLExtString; " -----------------------------"
    Print #FN,
    
    'Function pointer assignment code generation
    Print #FN, vbTab; vbTab; vbTab; GLExtString; " = false;"
    Print #FN, vbTab; vbTab; vbTab; "if (Extensions_Supported[IndexOfExtensionDict["""; GLExtString; """]])"
    Print #FN, vbTab; vbTab; vbTab; "{"
    LabelUsed = False
    For Each FuncPtrType In Ext.FuncPtrs.Keys
        FuncName = Ext.FuncPtrs(FuncPtrType)
        Print #FN, vbTab; vbTab; vbTab; vbTab; "FuncPtr = wglGetProcAddress("""; FuncName; """);"
        Print #FN, vbTab; vbTab; vbTab; vbTab; "if (FuncPtr == IntPtr.Zero) goto EndOf_"; GLExtString; ";"
        If HasUnsafe(FuncPtrType) Then
            Print #FN, vbTab; vbTab; vbTab; vbTab; FuncName; "_Safe = ("; FuncPtrType; ")Marshal.GetDelegateForFunctionPointer(FuncPtr, typeof("; FuncPtrType; "));"
            Print #FN, vbTab; vbTab; "#if CSHARPGLEW_ALLOW_UNSAFE"
            Print #FN, vbTab; vbTab; vbTab; vbTab; FuncName; "_Unsafe = ("; FuncPtrType; "_UNSAFE)Marshal.GetDelegateForFunctionPointer(FuncPtr, typeof("; FuncPtrType; "_UNSAFE));"
            Print #FN, vbTab; vbTab; "#endif // CSHARPGLEW_ALLOW_UNSAFE"
        Else
            Print #FN, vbTab; vbTab; vbTab; vbTab; FuncName; " = ("; FuncPtrType; ")Marshal.GetDelegateForFunctionPointer(FuncPtr, typeof("; FuncPtrType; "));"
        End If
        Print #FN,
        LabelUsed = True
    Next
    Print #FN, vbTab; vbTab; vbTab; vbTab; GLExtString; " = true;"
    Print #FN, vbTab; vbTab; vbTab; "}"
    If LabelUsed Then Print #FN, vbTab; vbTab; vbTab; "EndOf_"; GLExtString; ":;"
    Print #FN,
    
    Print #FN, vbTab; vbTab; vbTab; "#endregion"
    I = I + 1
Next
Print #FN, vbTab; vbTab; "}"
Print #FN,
Print #FN, vbTab; vbTab; "#endregion"
Print #FN,
Print #FN, vbTab; vbTab; "#region ""OpenGL Context Related"""
Print #FN, vbTab; vbTab; "public const uint PFD_TYPE_RGBA = 0;"
Print #FN, vbTab; vbTab; "public const uint PFD_TYPE_COLORINDEX = 1;"
Print #FN, vbTab; vbTab; "public const uint PFD_MAIN_PLANE = 0;"
Print #FN, vbTab; vbTab; "public const uint PFD_OVERLAY_PLANE = 1;"
Print #FN, vbTab; vbTab; "public const uint PFD_UNDERLAY_PLANE = 0xffffffff;"
Print #FN, vbTab; vbTab; "public const uint PFD_DOUBLEBUFFER = 0x1;"
Print #FN, vbTab; vbTab; "public const uint PFD_STEREO = 0x2;"
Print #FN, vbTab; vbTab; "public const uint PFD_DRAW_TO_WINDOW = 0x4;"
Print #FN, vbTab; vbTab; "public const uint PFD_DRAW_TO_BITMAP = 0x8;"
Print #FN, vbTab; vbTab; "public const uint PFD_SUPPORT_GDI = 0x10;"
Print #FN, vbTab; vbTab; "public const uint PFD_SUPPORT_OPENGL = 0x20;"
Print #FN, vbTab; vbTab; "public const uint PFD_GENERIC_FORMAT = 0x40;"
Print #FN, vbTab; vbTab; "public const uint PFD_NEED_PALETTE = 0x80;"
Print #FN, vbTab; vbTab; "public const uint PFD_NEED_SYSTEM_PALETTE = 0x100;"
Print #FN, vbTab; vbTab; "public const uint PFD_SWAP_EXCHANGE = 0x200;"
Print #FN, vbTab; vbTab; "public const uint PFD_SWAP_COPY = 0x400;"
Print #FN, vbTab; vbTab; "public const uint PFD_SWAP_LAYER_BUFFERS = 0x800;"
Print #FN, vbTab; vbTab; "public const uint PFD_GENERIC_ACCELERATED = 0x1000;"
Print #FN, vbTab; vbTab; "public const uint PFD_SUPPORT_DIRECTDRAW = 0x2000;"
Print #FN, vbTab; vbTab; "public const uint PFD_DEPTH_DONTCARE = 0x20000000;"
Print #FN, vbTab; vbTab; "public const uint PFD_DOUBLEBUFFER_DONTCARE = 0x40000000;"
Print #FN, vbTab; vbTab; "public const uint PFD_STEREO_DONTCARE = 0x80000000;"
Print #FN,
Print #FN, vbTab; vbTab; "public struct PIXELFORMATDESCRIPTOR"
Print #FN, vbTab; vbTab; "{"
Print #FN, vbTab; vbTab; vbTab; "public UInt16 nSize;"
Print #FN, vbTab; vbTab; vbTab; "public UInt16 nVersion;"
Print #FN, vbTab; vbTab; vbTab; "public UInt32 dwFlags;"
Print #FN, vbTab; vbTab; vbTab; "public byte iPixelType;"
Print #FN, vbTab; vbTab; vbTab; "public byte cColorBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cRedBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cRedShift;"
Print #FN, vbTab; vbTab; vbTab; "public byte cGreenBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cGreenShift;"
Print #FN, vbTab; vbTab; vbTab; "public byte cBlueBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cBlueShift;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAlphaBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAlphaShift;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAccumBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAccumRedBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAccumGreenBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAccumBlueBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAccumAlphaBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cDepthBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cStencilBits;"
Print #FN, vbTab; vbTab; vbTab; "public byte cAuxBuffers;"
Print #FN, vbTab; vbTab; vbTab; "public byte iLayerType;"
Print #FN, vbTab; vbTab; vbTab; "public byte bReserved;"
Print #FN, vbTab; vbTab; vbTab; "public UInt32 dwLayerMask;"
Print #FN, vbTab; vbTab; vbTab; "public UInt32 dwVisibleMask;"
Print #FN, vbTab; vbTab; vbTab; "public UInt32 dwDamageMask;"
Print #FN, vbTab; vbTab; "}"
Print #FN,
Print #FN, vbTab; vbTab; "[DllImport(""gdi32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern Int32 ChoosePixelFormat(IntPtr hDC, ref PIXELFORMATDESCRIPTOR pfd);"
Print #FN, vbTab; vbTab; "[DllImport(""gdi32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool SetPixelFormat(IntPtr hDC, Int32 pm, ref PIXELFORMATDESCRIPTOR pfd);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern IntPtr wglCreateContext(IntPtr hDC);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern IntPtr wglGetCurrentDC();"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern IntPtr wglGetCurrentContext();"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool wglMakeCurrent(IntPtr hDC, IntPtr hGLRC);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool wglDeleteContext(IntPtr hGLRC);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool wglSwapBuffers(IntPtr hDC);"
Print #FN, vbTab; vbTab; "[DllImport(""user32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern IntPtr GetDC(IntPtr hWnd);"
Print #FN, vbTab; vbTab; "[DllImport(""user32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool wglUseFontBitmaps(IntPtr hDC, UInt32 first, UInt32 count, UInt32 listBase);"
Print #FN, vbTab; vbTab; "[DllImport(""gdi32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern bool SwapBuffers(IntPtr hDC);"
Print #FN, vbTab; vbTab; "[DllImport(""opengl32.dll"")]"
Print #FN, vbTab; vbTab; "public static extern IntPtr wglGetProcAddress(string ProcName);"
Print #FN,
Print #FN, vbTab; vbTab; "#endregion"
Print #FN, vbTab; "}"
Print #FN, "}"
Print #FN,
Print #FN, "#pragma warning restore IDE1006, CS0649"
Print #FN,
Close #FN

Open ExportTo & "\GLAPI_Overloads.cs" For Output As #FN
Print #FN, "#define CSHARPGLEW_ALLOW_UNSAFE"
Print #FN, "#undef CSHARPGLEW_ALLOW_UNSAFE"
Print #FN,
Print #FN, "using System;"
Print #FN,
Print #FN, "#pragma warning disable IDE1006"
Print #FN,
Print #FN, "namespace CSharpGLEW"
Print #FN, "{"
Print #FN, vbTab; "partial class GLAPI"
Print #FN, vbTab; "{"
ExportOverloads FN, True, "void", "glColor", "", "", "red,green,blue|red,green,blue,alpha", "b,ub,s,us,i,ui,f,d,bv,ubv,sv,usv,iv,uiv,fv,dv"
ExportOverloads FN, True, "void", "glVertex", "", "", "x,y|x,y,z|x,y,z,w", "s,i,f,d,sv,iv,fv,dv"
ExportOverloads FN, True, "void", "glTexCoord", "", "", "x,y|x,y,z|x,y,z,w", "s,i,f,d,sv,iv,fv,dv"
ExportOverloads FN, True, "void", "glNormal", "", "", "x,y,z", "b,s,i,f,d,bv,sv,iv,fv,dv"
ExportOverloads FN, True, "void", "glEvalCoord", "", "", "u|u,v", "f,d,fv,dv"
ExportOverloads FN, True, "void", "glFog", "uint pname, ", "pname, ", "param", "f,i,fv,iv", , , True
ExportOverloads FN, True, "void", "glGetLight", "uint light, uint pname, ", "light, pname, ", "ref params", "fv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetMap", "uint target, uint query, ", "target, query, ", "ref v", "fv,dv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetMaterial", "uint face, uint pname, ", "face, pname, ", "ref params", "fv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetPixelMap", "uint map, ", "map, ", "ref values", "fv,uiv,usv", , , True, True
ExportOverloads FN, True, "void", "glGetTexEnv", "uint target, uint pname, ", "target, pname, ", "ref params", "fv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetTexGen", "uint coord, uint pname, ", "coord, pname, ", "ref params", "dv,fv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetTexLevelParameter", "uint target, int level, uint pname, ", "target, level, pname, ", "ref params", "fv,iv", , , True, True
ExportOverloads FN, True, "void", "glGetTexParameter", "uint target, uint pname, ", "target, pname, ", "ref params", "fv,iv", , , True, True
ExportOverloads FN, True, "void", "glIndex", "", "", "c", "d,dv,f,fv,i,iv,s,sv,ub,ubv", , , True
ExportOverloads FN, True, "void", "glLightModel", "uint pname, ", "pname, ", "param", "f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glLight", "uint light, uint pname, ", "light, pname, ", "param", "f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glLoadMatrix", "", "", "m", "d,f", , True, True
ExportOverloads FN, True, "void", "glMaterial", "uint face, uint pname, ", "face, pname, ", "param", "f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glMultMatrix", "", "", "m", "d,f", , True, True
ExportOverloads FN, True, "void", "glPixelMap", "uint map, int mapsize, ", "map, mapsize, ", "values", "fv,uiv,usv", , , True
ExportOverloads FN, True, "void", "glPixelStore", "uint pname, ", "pname, ", "param", "f,i", , , True
ExportOverloads FN, True, "void", "glPixelTransfer", "uint pname, ", "pname, ", "param", "f,i", , , True
ExportOverloads FN, True, "void", "glRasterPos", "", "", "x,y|x,y,z|x,y,z,w", "d,dv,f,fv,i,iv,s,sv"
ExportOverloads FN, True, "void", "glRect", "", "", "x1,y1,x2,y2", "d,f,i,s", , , True
ExportOverloads FN, True, "void", "glRotate", "", "", "angle,x,y,z", "d,f", , , True
ExportOverloads FN, True, "void", "glScale", "", "", "x,y,z", "d,f", , , True
ExportOverloads FN, True, "void", "glTexEnv", "uint target, uint pname, ", "target, pname, ", "param", "f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glTexGen", "uint coord, uint pname, ", "coord, pname, ", "param", "d,dv,f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glTexParameter", "uint target, uint pname, ", "target, pname, ", "param", "f,fv,i,iv", , , True
ExportOverloads FN, True, "void", "glTranslate", "", "", "x,y,z", "d,f", , , True
Print #FN, vbTab; "}"
Print #FN, "}"
Print #FN,
Print #FN, "#pragma warning restore IDE1006"
Print #FN,
Close #FN

End Sub

