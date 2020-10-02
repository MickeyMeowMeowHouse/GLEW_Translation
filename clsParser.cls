VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GLExtension As New Dictionary
Public ErrorReport As New Dictionary

Private m_CurPrefix As String
Private m_CurPrefix2 As String

'合并空格
Private Sub UnifySpaces(ToUnify As String)
ToUnify = Replace(ToUnify, vbTab, " ")
UnifyChar ToUnify, " "
End Sub

'合并字符
Private Sub UnifyChar(ToUnify As String, Char As String)
Dim Length As Long, NewLength As Long
Dim CharChar As String
CharChar = Char & Char
Do
    Length = NewLength
    ToUnify = Replace(ToUnify, CharChar, Char)
    NewLength = Len(ToUnify)
Loop While NewLength <> Length
End Sub

'是否字母
Private Function IsAlpha(ByVal CharCode As Long) As Boolean
Select Case CharCode
Case &H41 To &H5A
    IsAlpha = True
Case &H61 To &H7A
    IsAlpha = True
End Select
End Function

'是否英数字
Private Function IsAlNum(ByVal CharCode As Long) As Boolean
Select Case CharCode
Case &H30 To &H39
    IsAlNum = True
Case &H41 To &H5A
    IsAlNum = True
Case &H61 To &H7A
    IsAlNum = True
End Select
End Function

'是否C标识符
Private Function IsCSym(Expression As String) As Boolean
Dim I As Long, CS As Long
If Len(Expression) = 0 Then Exit Function
CS = AscW(Left$(Expression, 1))
If IsAlpha(CS) = False And CS <> &H5F Then Exit Function
IsCSym = True
For I = 2 To Len(Expression)
    CS = AscW(Mid$(Expression, I, 1))
    Select Case CS
    Case &H30 To &H39
    Case &H41 To &H5A
    Case &H5F
    Case &H61 To &H7A
    Case Else
        IsCSym = False
        Exit Function
    End Select
Next
End Function

'是否宏
Private Function IsMacro(Expression As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(Expression)

If Left$(Trimmed, 1) = "#" Then IsMacro = True
End Function

'对比字符串前缀
Private Function HasPrefix(Expression As String, Prefix As String) As Boolean
If Left$(Expression, Len(Prefix)) = Prefix Then HasPrefix = True
End Function

'对比字符串后缀
Private Function HasSuffix(Expression As String, Suffix As String) As Boolean
If Right$(Expression, Len(Suffix)) = Suffix Then HasSuffix = True
End Function

'对比并清除前缀
Private Function DropPrefix(Expression As String, Prefix As String) As Boolean
If Left$(Expression, Len(Prefix)) = Prefix Then
    Expression = Mid$(Expression, Len(Prefix) + 1)
    DropPrefix = True
End If
End Function

'对比并清除后缀
Private Function DropSuffix(Expression As String, Suffix As String) As Boolean
If Right$(Expression, Len(Suffix)) = Suffix Then
    Expression = Left$(Expression, Len(Expression) - Len(Suffix))
    DropSuffix = True
End If
End Function

'是否某版本的开始行
Private Function LooksLikeStartOfExtension(ReadLine As String, Optional ExtensionIndicated As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)

'必须是注释
If DropPrefix(Trimmed, "/*") = False Then Exit Function
If DropSuffix(Trimmed, "*/") = False Then Exit Function

'使用查找替换合并减号
UnifyChar Trimmed, "-"
Trimmed = Trim$(Trimmed)

'现在的Trimmed应该是这个样子的：
'- GL_VERSION_X_X -

If DropPrefix(Trimmed, "-") = False Then Exit Function
If DropSuffix(Trimmed, "-") = False Then Exit Function
Trimmed = Trim$(Trimmed)
If IsCSym(Trimmed) = False Then Exit Function

ExtensionIndicated = Trimmed
LooksLikeStartOfExtension = True
End Function

'是否某版本的结束行
Private Function LooksLikeEndOfExtension(ReadLine As String, Optional ExtensionIndicated As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)

If DropPrefix(Trimmed, "#endif ") = False Then Exit Function
If DropPrefix(Trimmed, "/*") = False Then Exit Function
If DropSuffix(Trimmed, "*/") = False Then Exit Function
If IsCSym(Trimmed) = False Then Exit Function

ExtensionIndicated = Trimmed
LooksLikeEndOfExtension = True
End Function

'是否宏定义
Private Function LooksLikeMacroDef(ReadLine As String, Optional MacroName As String, Optional MacroDef As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)

UnifySpaces Trimmed

Dim Splitted() As String
Splitted = Split(Trimmed, " ", 3)

If UBound(Splitted) <= 0 Then Exit Function
If Splitted(0) <> "#define" Then Exit Function

MacroName = Splitted(1)
MacroDef = Splitted(2)

LooksLikeMacroDef = True
End Function

'是否句柄定义
Private Function IsDeclareHandle(ReadLine As String, Optional HandleName As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)
If DropPrefix(Trimmed, "DECLARE_HANDLE(") = False Then Exit Function
Trimmed = Split(Trimmed, ";", 2)(0) '去掉分号和其后的部分

HandleName = Left$(Trimmed, Len(Trimmed) - 1)
IsDeclareHandle = True
End Function

'是否类型定义
Private Function LooksLikeNonFuncPtrTypeDef(ReadLine As String, Optional TypeName As String, Optional TypeAliases As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)

If DropPrefix(Trimmed, "typedef ") = False Then Exit Function
Trimmed = Split(Trimmed, ";", 2)(0) '去掉分号和其后的部分
If InStr(Trimmed, "(") Then Exit Function '有括号的都视为函数指针定义
'此时得到typedef    aaa bbb * ccc, ddd, * * eee , fff

Trimmed = Trim$(Trimmed) '去掉前后空格
UnifySpaces Trimmed '合并空格
'此时得到aaa bbb * ccc, ddd, * * eee , fff

'处理指针星号
Trimmed = Replace(Trimmed, "*", " *") '在指针星号的左边插入空格，使其和左边的标识符分开
Trimmed = Replace(Trimmed, "  ", " ") '合并多余空格
Trimmed = Replace(Trimmed, "* ", "*") '让指针星号直接连接右边的星号或者标识符
Trimmed = Replace(Trimmed, " ,", ",") '去掉逗号左右的空格
Trimmed = Replace(Trimmed, ", ", ",")
'此时得到aaa bbb *ccc,ddd,**eee,fff

Dim LastSpace As Long
LastSpace = InStrRev(Trimmed, " ")

TypeName = Left$(Trimmed, LastSpace - 1)
TypeAliases = Mid$(Trimmed, LastSpace + 1)

LooksLikeNonFuncPtrTypeDef = True
End Function

'是否结构体定义
Private Function LooksLikeTypedefStruct(ReadLine As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)
UnifySpaces Trimmed '合并空格

If HasPrefix(Trimmed, "typedef struct ") = False Then Exit Function
If InStr(Trimmed, "{") = 0 Then Exit Function
LooksLikeTypedefStruct = True
End Function

'是否API
Private Function LooksLikeAPIDef(ReadLine As String, RetType As String, FuncName As String, Params As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)

Dim LBr As Long, RBr As Long, SpcPos As Long

If DropPrefix(Trimmed, "GLAPI ") = False Then Exit Function
Trimmed = Replace(Trimmed, "GLAPIENTRY", "") '去掉开头和调用约定部分
Trimmed = Split(Trimmed, ";", 2)(0) '去掉分号和其后的部分
Trimmed = Replace(Trimmed, "*", "* ") '在指针星号的右边插入空格，使其和右边的标识符分开
UnifySpaces Trimmed '合并空格
Trimmed = Replace(Trimmed, " *", "*") '让指针星号直接连接左边的星号或者标识符
Trimmed = Replace(Trimmed, " ,", ",") '去掉逗号左右的空格
Trimmed = Replace(Trimmed, ", ", ",")

LBr = InStr(Trimmed, "(")
RBr = InStrRev(Trimmed, ")")
If LBr = 0 Or RBr = 0 Then Exit Function 'API必须有括弧

'参数部分
Params = Trim$(Mid$(Trimmed, LBr + 1, RBr - LBr - 1))
Trimmed = Trim$(Left$(Trimmed, LBr - 1))

SpcPos = InStrRev(Trimmed, " ")
RetType = Left$(Trimmed, SpcPos - 1)
FuncName = Mid$(Trimmed, SpcPos + 1)

LooksLikeAPIDef = True
End Function

'是否函数指针
Private Function LooksLikeFuncPtrDef(ReadLine As String, RetType As String, FuncPtrTypeName As String, Params As String) As Boolean
Dim Trimmed As String
Trimmed = Trim$(ReadLine)
'假定ReadLine包含完整的一行typedef语句，中途没有换行
'glew.h文件本身是脚本生成的，它里面的typedef不会中途换行

Dim LBr As Long, RBr As Long, SpcPos As Long

If DropPrefix(Trimmed, "typedef ") = False Then Exit Function
Trimmed = Split(Trimmed, ";", 2)(0) '去掉分号和其后的部分

LBr = InStr(Trimmed, "(")
RBr = InStr(LBr + 1, Trimmed, ")")
If LBr = 0 Or RBr = 0 Then Exit Function '函数指针定义必须有括弧

RetType = Trim$(Left$(Trimmed, LBr - 1))

FuncPtrTypeName = Mid$(Trimmed, LBr + 1, RBr - LBr - 1)
FuncPtrTypeName = Trim$(Mid$(FuncPtrTypeName, InStrRev(FuncPtrTypeName, "*") + 1))
'If DropPrefix(FuncPtrTypeName, "PFN" & UCase$(m_CurPrefix2)) = False Then Exit Function
'If DropSuffix(FuncPtrTypeName, "PROC") = False Then Exit Function

LBr = InStr(RBr + 1, Trimmed, "(")
RBr = InStr(LBr + 1, Trimmed, ")")
If LBr = 0 Or RBr = 0 Then Exit Function '函数指针定义必须有两对括弧

Params = Trim$(Mid$(Trimmed, LBr + 1, RBr - LBr - 1))
Params = Replace(Params, "*", "* ") '在指针星号的右边插入空格，使其和右边的标识符分开
UnifySpaces Params '合并空格
Params = Replace(Params, " *", "*") '让指针星号直接连接左边的星号或者标识符
Params = Replace(Params, " ,", ",") '去掉逗号左右的空格
Params = Replace(Params, ", ", ",")

LooksLikeFuncPtrDef = True
End Function

Private Function SearchInString(ByVal Start As Long, String1 As String, String2 As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
If Start = 0 Then Start = 1
SearchInString = InStr(Start, String1, String2, Compare)
If SearchInString = 0 Then SearchInString = Len(String1) + 1
End Function

Private Function Min(ByVal A As Long, ByVal B As Long) As Long
If A > B Then Min = B Else Min = A
End Function

Private Function Max(ByVal A As Long, ByVal B As Long) As Long
If A > B Then Max = A Else Max = B
End Function

Function RemoveComments(StringWithComments As String) As String
Dim Quote As Long
Dim ComOpen As Long
Dim ComClose As Long
Dim Nesting As Long
Dim CurPos As Long

RemoveComments = StringWithComments

Do
    '先跳过所有双引号内的内容
    Quote = SearchInString(CurPos, StringWithComments, """")
    ComOpen = Min(SearchInString(CurPos, StringWithComments, "/*"), SearchInString(CurPos, StringWithComments, "//"))
    
    Do While Quote < ComOpen
        '跳过双引号内容，重新搜索注释开头
        CurPos = SearchInString(Quote + 1, StringWithComments, """") + 1
        Quote = SearchInString(CurPos, StringWithComments, """")
        ComOpen = Min(SearchInString(CurPos, StringWithComments, "/*"), SearchInString(CurPos, StringWithComments, "//"))
    Loop
    '到达结尾
    If Quote = ComOpen Then Exit Do

    '然后根据注释的起始方式判断注释范围
    If Mid$(StringWithComments, ComOpen, 2) = "/*" Then
        ComClose = ComOpen
        Nesting = 1
        Do
            ComClose = Min(SearchInString(ComClose + 2, StringWithComments, "/*"), SearchInString(ComClose + 2, StringWithComments, "*/"))
            If Mid$(StringWithComments, ComClose, 2) = "/*" Then '检测套娃
                Nesting = Nesting + 1
            Else
                Nesting = Nesting - 1
            End If
        Loop While Nesting > 0
    Else
        ComClose = SearchInString(ComOpen + 2, StringWithComments, vbCrLf)
    End If
    '先将注释区域的字符设为空字符
    Mid$(RemoveComments, ComOpen, ComClose + 2 - ComOpen) = String$(ComClose + 2 - ComOpen, vbNullChar)
    CurPos = ComClose + 2
Loop While CurPos <= Len(StringWithComments)

'然后再干掉空字符
RemoveComments = Replace(RemoveComments, vbNullChar, "")
End Function

'宏查找
Private Function MacroLookup(MacroName As String, OutMacroDef As String) As Boolean
Dim ExtensionString
Dim ExtData As clsGLExtension
For Each ExtensionString In GLExtension
    Set ExtData = GLExtension(ExtensionString)
    If ExtData.MacroDefs.Exists(MacroName) Then
        OutMacroDef = ExtData.MacroDefs(MacroName)
        MacroLookup = True
        Exit Function
    End If
Next
End Function

'分析头文件
Sub ParseHeaderFile(HeaderFilePath As String, Prefix As String)
Dim FN As Integer
FN = FreeFile

Dim LineNum As Long
Dim ReadLine As String
Dim ExtensionString As String
Dim ReadingExtension As Boolean
Dim CurExtension As clsGLExtension

Dim HandleName As String

Dim MacroName As String
Dim MacroDef As String
Dim MacroDefClean As String
Dim MacroToLookup As New Dictionary
Dim MacroLookingUp

Dim TypeName As String
Dim TypeAliases As String
Dim TypeAlias() As String
Dim TypeAliasTrueName As String
Dim TypeNamePointer As String

Dim ReadingStruct As Boolean
Dim StructGuts As String

Dim RetType As String
Dim FuncName As String
Dim FuncParams As String

Dim I As Long

m_CurPrefix = LCase$(Prefix)
m_CurPrefix2 = Left$(m_CurPrefix, Len(m_CurPrefix) - 2) '去掉EW结尾
'Set MacroToLookup = CreateObject("Scripting.Dictionary")

Open HeaderFilePath For Input As #FN
Do
    LineNum = LineNum + 1
    Line Input #FN, ReadLine
    If LooksLikeStartOfExtension(ReadLine, ExtensionString) Then '/* ----------------------------- GL_xxx_x_x ---------------------------- */
        If ReadingExtension Then
            GLExtension.Add CurExtension.ExtensionString, CurExtension
            Set CurExtension = New clsGLExtension
            CurExtension.EndLine = LineNum - 1
        Else
            ReadingExtension = True
            Set CurExtension = New clsGLExtension
        End If
        CurExtension.HeaderFilePath = HeaderFilePath
        CurExtension.ExtensionString = ExtensionString
        CurExtension.Prefix = m_CurPrefix
        CurExtension.StartLine = LineNum
        If ExtensionString = "GLU" Then
            ThisIsGLU CurExtension
        End If
    ElseIf LooksLikeEndOfExtension(ReadLine, ExtensionString) Then ' #endif /* GL_xxx_x_x */
        If ReadingExtension Then
            CurExtension.EndLine = LineNum
            GLExtension.Add CurExtension.ExtensionString, CurExtension
            Set CurExtension = Nothing
            ReadingExtension = False
        End If
    ElseIf ReadingExtension Then
        '宏
        If IsMacro(ReadLine) Then
            If LooksLikeMacroDef(ReadLine, MacroName, MacroDef) Then
                '首先去掉那些#define GL_xxx_x_x
                If MacroName <> ExtensionString Then
                    '其次判断是不是 GLEW_GET_FUN
                    If DropPrefix(MacroDef, UCase$(m_CurPrefix) & "_GET_FUN(__" & m_CurPrefix) Then
                        MacroDef = Left$(MacroDef, Len(MacroDef) - 1)
                        CurExtension.FuncPtrs.Add "PFN" & UCase$(m_CurPrefix2 & MacroDef) & "PROC", m_CurPrefix2 & MacroDef
                    ElseIf DropPrefix(UCase$(MacroDef), UCase$(m_CurPrefix) & "_GET_VAR(") Then
                    ElseIf HasPrefix(UCase$(MacroName), UCase$(m_CurPrefix)) Then
                    ElseIf InStr(MacroName, "(") Then
                    Else
                        CurExtension.MacroDefs.Add MacroName, MacroDef
                        MacroDefClean = Trim$(RemoveComments(MacroDef)) '去掉宏定义内容中的注释用于判断宏定义具体是个啥
                        If IsNumeric(MacroDefClean) = False And HasPrefix(MacroDefClean, "0x") = False Then
                            MacroToLookup.Add MacroName, ExtensionString
                        End If
                    End If
                End If
            End If
        Else 'IsMacro(ReadLine) = False
            If ReadingStruct = False Then
                '句柄定义
                If IsDeclareHandle(ReadLine, HandleName) Then
                    CurExtension.TypeDefs.Add HandleName, "HANDLE"
                '结构体定义
                ElseIf LooksLikeTypedefStruct(ReadLine) Then
                    StructGuts = ReadLine & vbCrLf
                    ReadingStruct = True
                
                '类型定义（非函数指针）
                ElseIf LooksLikeNonFuncPtrTypeDef(ReadLine, TypeName, TypeAliases) Then
                    TypeAlias = Split(TypeAliases, ",")
                    For I = 0 To UBound(TypeAlias)
                        TypeAliasTrueName = Replace(TypeAlias(I), "*", "")
                        TypeNamePointer = TypeName & String$(Len(TypeAlias(I)) - Len(TypeAliasTrueName), "*")
                        If CurExtension.TypeDefs.Exists(TypeAliasTrueName) = False Then
                            CurExtension.TypeDefs.Add TypeAliasTrueName, TypeNamePointer
                        End If
                    Next
                'API函数
                ElseIf LooksLikeAPIDef(ReadLine, RetType, FuncName, FuncParams) Then
                    CurExtension.APIs.Add FuncName, RetType & ":" & FuncParams
                '函数指针类型声明
                ElseIf LooksLikeFuncPtrDef(ReadLine, RetType, FuncName, FuncParams) Then
                    CurExtension.FuncTypeDef.Add FuncName, RetType & ":" & FuncParams
                End If
            Else 'ReadingStruct = True
                '我赌你的结构体里没有别的结构体或者联合体定义
                StructGuts = StructGuts & ReadLine & vbCrLf
                If InStr(ReadLine, "}") Then
                    StructGuts = Trim$(RemoveComments(StructGuts))
                    DropPrefix StructGuts, "typedef "
                    StructGuts = Left$(StructGuts, InStr(StructGuts, "}"))
                    
                    ReadLine = Split(ReadLine, ";", 2)(0) '去掉分号和其后的部分
                    ReadLine = Trim$(Mid$(ReadLine, InStr(ReadLine, "}") + 1))
                    TypeAlias = Split(ReadLine, ",")
                    For I = 0 To UBound(TypeAlias)
                        TypeAlias(I) = Trim$(TypeAlias(I))
                        TypeAlias(I) = Replace(TypeAlias(I), "*", " *") '在指针星号的左边插入空格，使其和左边的标识符分开
                        TypeAlias(I) = Replace(TypeAlias(I), "  ", " ") '合并多余空格
                        TypeAlias(I) = Replace(TypeAlias(I), "* ", "*") '让指针星号直接连接右边的星号或者标识符
                        TypeAlias(I) = Trim$(TypeAlias(I))
                        TypeAliasTrueName = Replace(TypeAlias(I), "*", "")
                        TypeNamePointer = StructGuts & String$(Len(TypeAlias(I)) - Len(TypeAliasTrueName), "*")
                        If CurExtension.TypeDefs.Exists(TypeAliasTrueName) = False Then
                            CurExtension.TypeDefs.Add TypeAliasTrueName, TypeNamePointer
                        End If
                    Next
                    
                    ReadingStruct = False
                End If
            End If 'ReadingStruct
        End If 'IsMacro(ReadLine)
    End If
    
Loop While Not EOF(FN)
Close #FN

'宏别名处理
For Each MacroLookingUp In MacroToLookup.Keys
    MacroName = MacroLookingUp
    
    ExtensionString = MacroToLookup(MacroName)
    Set CurExtension = GLExtension(ExtensionString)
    
    '取要展开的宏的定义内容
    MacroDefClean = Trim$(RemoveComments(CurExtension.MacroDefs(MacroName)))
    If MacroLookup(MacroDefClean, MacroDef) Then '如果宏定义的内容是另外一个宏，则查找那个宏的值
        CurExtension.MacroDefs.Item(MacroName) = MacroDef
    Else
        ErrorReport.Add ErrorReport.Count, HeaderFilePath & ": Macro definition isn't a constant integer." & vbCrLf & vbTab & "#define " & MacroName & " " & CurExtension.MacroDefs(MacroName)
    End If
Next

'Set MacroToLookup = Nothing
End Sub

Private Sub ThisIsGLU(CurExtension As clsGLExtension)
CurExtension.API_DllName = "glu32.dll"
CurExtension.APIs.Add "gluErrorString", "const GLubyte*:GLenum errCode"
CurExtension.APIs.Add "gluErrorUnicodeStringEXT", "const wchar_t*:GLenum errCode"
CurExtension.APIs.Add "gluGetString", "const GLubyte*:GLenum name"
CurExtension.APIs.Add "gluOrtho2D", "void:GLdouble left,GLdouble right,GLdouble bottom,GLdouble top"
CurExtension.APIs.Add "gluPerspective", "void:GLdouble fovy,GLdouble aspect,GLdouble zNear,GLdouble zFar"
CurExtension.APIs.Add "gluPickMatrix", "void:GLdouble x,GLdouble y,GLdouble width,GLdouble height,GLint* viewport"
CurExtension.APIs.Add "gluLookAt", "void:GLdouble eyex,GLdouble eyey,GLdouble eyez,GLdouble centerx,GLdouble centery,GLdouble centerz,GLdouble upx,GLdouble upy,GLdouble upz"
CurExtension.APIs.Add "gluProject", "int:GLdouble objx,GLdouble objy,GLdouble objz,const GLdouble* modelMatrix,const GLdouble* projMatrix,const GLint* viewport,GLdouble* winx,GLdouble* winy,GLdouble* winz"
CurExtension.APIs.Add "gluUnProject", "int:GLdouble objx,GLdouble objy,GLdouble objz,const GLdouble* modelMatrix,const GLdouble* projMatrix,const GLint* viewport,GLdouble* winx,GLdouble* winy,GLdouble* winz"
CurExtension.APIs.Add "gluScaleImage", "int:GLenum format,GLint widthin,GLint heightin,GLenum typein,const void* datain,GLint widthout,GLint heightout,GLenum typeout,void* dataout"
CurExtension.APIs.Add "gluBuild1DMipmaps", "int:GLenum target,GLint components,GLint width,GLenum format,GLenum type,const void* data"
CurExtension.APIs.Add "gluBuild2DMipmaps", "int:GLenum target,GLint components,GLint width,GLint height,GLenum format,GLenum type,const void* data"
CurExtension.APIs.Add "gluNewQuadric", "GLUquadric*:"
CurExtension.APIs.Add "gluDeleteQuadric", "void:GLUquadric* state"
CurExtension.APIs.Add "gluQuadricNormals", "void:GLUquadric* quadObject,GLenum normals"
CurExtension.APIs.Add "gluQuadricTexture", "void:GLUquadric* quadObject,GLboolean textureCoords"
CurExtension.APIs.Add "gluQuadricOrientation", "void:GLUquadric* quadObject,GLenum orientation"
CurExtension.APIs.Add "gluQuadricDrawStyle", "void:GLUquadric* quadObject,GLenum drawStyle"
CurExtension.APIs.Add "gluCylinder", "void:GLUquadric* qobj,GLdouble baseRadius,GLdouble topRadius,GLdouble height,GLint slices,GLint stacks"
CurExtension.APIs.Add "gluDisk", "void:GLUquadric* qobj,GLdouble innerRadius,GLdouble outerRadius,GLint slices,GLint loops"
CurExtension.APIs.Add "gluPartialDisk", "void:GLUquadric* qobj,GLdouble innerRadius,GLdouble outerRadius,GLint slices,GLint loops,GLdouble startAngle,GLdouble sweepAngle"
CurExtension.APIs.Add "gluSphere", "void:GLUquadric* qobj,GLdouble radius,GLint slices,GLint stacks"
CurExtension.FuncTypeDef.Add "fnGLUquadricCallback", "void:"
CurExtension.APIs.Add "gluQuadricCallback", "void:GLUquadric* qobj,GLenum which,fnGLUquadricCallback fn"
CurExtension.APIs.Add "gluNewTess", "GLUtesselator*:"
CurExtension.APIs.Add "gluDeleteTess", "void:GLUtesselator* tess"
CurExtension.APIs.Add "gluTessBeginPolygon", "void:GLUtesselator* tess,void* polygon_data"
CurExtension.APIs.Add "gluTessBeginContour", "void:GLUtesselator* tess"
CurExtension.APIs.Add "gluTessVertex", "void:GLUtesselator* tess,GLdouble* coords,void* data"
CurExtension.APIs.Add "gluTessEndContour", "void:GLUtesselator* tess"
CurExtension.APIs.Add "gluTessEndPolygon", "void:GLUtesselator* tess"
CurExtension.APIs.Add "gluTessProperty", "void:GLUtesselator* tess,GLenum which,GLdouble value"
CurExtension.APIs.Add "gluTessNormal", "void:GLUtesselator* tess,GLdouble x,GLdouble y,GLdouble z"
CurExtension.FuncTypeDef.Add "fnGLUtessCallback", "void:"
CurExtension.APIs.Add "gluTessCallback", "void:GLUtesselator* tess,GLenum which,fnGLUtessCallback fn"
CurExtension.APIs.Add "gluGetTessProperty", "void:GLUtesselator* tess,GLenum which,GLdouble* value"
CurExtension.APIs.Add "gluNewNurbsRenderer", "GLUnurbs*:"
CurExtension.APIs.Add "gluDeleteNurbsRenderer", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluBeginSurface", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluBeginCurve", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluEndCurve", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluEndSurface", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluBeginTrim", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluEndTrim", "void:GLUnurbs* nobj"
CurExtension.APIs.Add "gluPwlCurve", "void:GLUnurbs* nobj,GLint count,GLfloat* array,GLint stride,GLenum type"
CurExtension.APIs.Add "gluNurbsCurve", "void:GLUnurbs* nobj,GLint nknots,GLfloat* knot,GLint stride,GLfloat* ctlarray,GLint order,GLenum type"
CurExtension.APIs.Add "gluNurbsSurface", "void:GLUnurbs* nobj,GLint sknot_count,float* sknot,GLint tknot_count,GLfloat* tknot,GLint s_stride,GLint t_stride,GLfloat* ctlarray,GLint sorder,GLint torder,GLenum type"
CurExtension.APIs.Add "gluLoadSamplingMatrices", "void:GLUnurbs* nobj,const GLfloat* modelMatrix,const GLfloat* projMatrix,const GLint* viewport"
CurExtension.APIs.Add "gluNurbsProperty", "void:GLUnurbs* nobj,GLenum property,GLfloat value"
CurExtension.APIs.Add "gluGetNurbsProperty", "void:GLUnurbs* nobj,GLenum property,GLfloat* value"
CurExtension.FuncTypeDef.Add "fnGLUnurbsCallback", "void:"
CurExtension.APIs.Add "gluNurbsCallback", "void:GLUnurbs* nobj,GLenum which,fnGLUnurbsCallback fn"
CurExtension.FuncTypeDef.Add "fnGLUquadricErrorProc", "void:GLenum"
CurExtension.FuncTypeDef.Add "fnGLUtessEdgeFlagProc", "void:GLboolean"
CurExtension.FuncTypeDef.Add "fnGLUtessVertexProc", "void:void*"
CurExtension.FuncTypeDef.Add "fnGLUtessEndProc", "void:"
CurExtension.FuncTypeDef.Add "fnGLUtessErrorProc", "void:GLenum"
CurExtension.FuncTypeDef.Add "fnGLUtessCombineProc", "void:GLdouble*,void**,GLfloat*,void**"
CurExtension.FuncTypeDef.Add "fnGLUtessBeginDataProc", "void:GLenum,void*"
CurExtension.FuncTypeDef.Add "fnGLUtessEdgeFlagDataProc", "void:GLboolean,void*"
CurExtension.FuncTypeDef.Add "fnGLUtessVertexDataProc", "void:void*,void*"
CurExtension.FuncTypeDef.Add "fnGLUtessEndDataProc", "void:void*"
CurExtension.FuncTypeDef.Add "fnGLUtessErrorDataProc", "void:GLenum,void*"
CurExtension.FuncTypeDef.Add "fnGLUtessCombineDataProc", "void:GLdouble*,void**,GLfloat*,void**,void*"
CurExtension.FuncTypeDef.Add "fnGLUnurbsErrorProc", "void:GLenum"
CurExtension.MacroDefs.Add "GLU_VERSION_1_1", "1"
CurExtension.MacroDefs.Add "GLU_VERSION_1_2", "1"
CurExtension.MacroDefs.Add "GLU_INVALID_ENUM", "100900"
CurExtension.MacroDefs.Add "GLU_INVALID_VALUE", "100901"
CurExtension.MacroDefs.Add "GLU_OUT_OF_MEMORY", "100902"
CurExtension.MacroDefs.Add "GLU_INCOMPATIBLE_GL_VERSION", "100903"
CurExtension.MacroDefs.Add "GLU_VERSION", "100800"
CurExtension.MacroDefs.Add "GLU_EXTENSIONS", "100801"
CurExtension.MacroDefs.Add "GLU_TRUE", "GL_TRUE"
CurExtension.MacroDefs.Add "GLU_FALSE", "GL_FALSE"
CurExtension.MacroDefs.Add "GLU_SMOOTH", "100000"
CurExtension.MacroDefs.Add "GLU_FLAT", "100001"
CurExtension.MacroDefs.Add "GLU_NONE", "100002"
CurExtension.MacroDefs.Add "GLU_POINT", "100010"
CurExtension.MacroDefs.Add "GLU_LINE", "100011"
CurExtension.MacroDefs.Add "GLU_FILL", "100012"
CurExtension.MacroDefs.Add "GLU_SILHOUETTE", "100013"
CurExtension.MacroDefs.Add "GLU_OUTSIDE", "100020"
CurExtension.MacroDefs.Add "GLU_INSIDE", "100021"
CurExtension.MacroDefs.Add "GLU_TESS_MAX_COORD", "1.0e150"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_RULE", "100140"
CurExtension.MacroDefs.Add "GLU_TESS_BOUNDARY_ONLY", "100141"
CurExtension.MacroDefs.Add "GLU_TESS_TOLERANCE", "100142"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_ODD", "100130"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_NONZERO", "100131"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_POSITIVE", "100132"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_NEGATIVE", "100133"
CurExtension.MacroDefs.Add "GLU_TESS_WINDING_ABS_GEQ_TWO", "100134"
CurExtension.MacroDefs.Add "GLU_TESS_BEGIN", "100100"
CurExtension.MacroDefs.Add "GLU_TESS_VERTEX", "100101"
CurExtension.MacroDefs.Add "GLU_TESS_END", "100102"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR", "100103"
CurExtension.MacroDefs.Add "GLU_TESS_EDGE_FLAG", "100104"
CurExtension.MacroDefs.Add "GLU_TESS_COMBINE", "100105"
CurExtension.MacroDefs.Add "GLU_TESS_BEGIN_DATA", "100106"
CurExtension.MacroDefs.Add "GLU_TESS_VERTEX_DATA", "100107"
CurExtension.MacroDefs.Add "GLU_TESS_END_DATA", "100108"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR_DATA", "100109"
CurExtension.MacroDefs.Add "GLU_TESS_EDGE_FLAG_DATA", "100110"
CurExtension.MacroDefs.Add "GLU_TESS_COMBINE_DATA", "100111"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR1", "100151"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR2", "100152"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR3", "100153"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR4", "100154"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR5", "100155"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR6", "100156"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR7", "100157"
CurExtension.MacroDefs.Add "GLU_TESS_ERROR8", "100158"
CurExtension.MacroDefs.Add "GLU_TESS_MISSING_BEGIN_POLYGON", "GLU_TESS_ERROR1"
CurExtension.MacroDefs.Add "GLU_TESS_MISSING_BEGIN_CONTOUR", "GLU_TESS_ERROR2"
CurExtension.MacroDefs.Add "GLU_TESS_MISSING_END_POLYGON", "GLU_TESS_ERROR3"
CurExtension.MacroDefs.Add "GLU_TESS_MISSING_END_CONTOUR", "GLU_TESS_ERROR4"
CurExtension.MacroDefs.Add "GLU_TESS_COORD_TOO_LARGE", "GLU_TESS_ERROR5"
CurExtension.MacroDefs.Add "GLU_TESS_NEED_COMBINE_CALLBACK", "GLU_TESS_ERROR6"
CurExtension.MacroDefs.Add "GLU_AUTO_LOAD_MATRIX", "100200"
CurExtension.MacroDefs.Add "GLU_CULLING", "100201"
CurExtension.MacroDefs.Add "GLU_SAMPLING_TOLERANCE", "100203"
CurExtension.MacroDefs.Add "GLU_DISPLAY_MODE", "100204"
CurExtension.MacroDefs.Add "GLU_PARAMETRIC_TOLERANCE", "100202"
CurExtension.MacroDefs.Add "GLU_SAMPLING_METHOD", "100205"
CurExtension.MacroDefs.Add "GLU_U_STEP", "100206"
CurExtension.MacroDefs.Add "GLU_V_STEP", "100207"
CurExtension.MacroDefs.Add "GLU_PATH_LENGTH", "100215"
CurExtension.MacroDefs.Add "GLU_PARAMETRIC_ERROR", "100216"
CurExtension.MacroDefs.Add "GLU_DOMAIN_DISTANCE", "100217"
CurExtension.MacroDefs.Add "GLU_MAP1_TRIM_2", "100210"
CurExtension.MacroDefs.Add "GLU_MAP1_TRIM_3", "100211"
CurExtension.MacroDefs.Add "GLU_OUTLINE_POLYGON", "100240"
CurExtension.MacroDefs.Add "GLU_OUTLINE_PATCH", "100241"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR1", "100251"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR2", "100252"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR3", "100253"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR4", "100254"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR5", "100255"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR6", "100256"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR7", "100257"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR8", "100258"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR9", "100259"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR10", "100260"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR11", "100261"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR12", "100262"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR13", "100263"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR14", "100264"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR15", "100265"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR16", "100266"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR17", "100267"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR18", "100268"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR19", "100269"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR20", "100270"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR21", "100271"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR22", "100272"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR23", "100273"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR24", "100274"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR25", "100275"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR26", "100276"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR27", "100277"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR28", "100278"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR29", "100279"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR30", "100280"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR31", "100281"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR32", "100282"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR33", "100283"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR34", "100284"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR35", "100285"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR36", "100286"
CurExtension.MacroDefs.Add "GLU_NURBS_ERROR37", "100287"
CurExtension.APIs.Add "gluBeginPolygon", "void:GLUtesselator* tess"
CurExtension.APIs.Add "gluNextContour", "void:GLUtesselator* tess,GLenum type"
CurExtension.APIs.Add "gluEndPolygon", "void:GLUtesselator* tess"
CurExtension.MacroDefs.Add "GLU_CW", "100120"
CurExtension.MacroDefs.Add "GLU_CCW", "100121"
CurExtension.MacroDefs.Add "GLU_INTERIOR", "100122"
CurExtension.MacroDefs.Add "GLU_EXTERIOR", "100123"
CurExtension.MacroDefs.Add "GLU_UNKNOWN", "100124"
CurExtension.MacroDefs.Add "GLU_BEGIN", "GLU_TESS_BEGIN"
CurExtension.MacroDefs.Add "GLU_VERTEX", "GLU_TESS_VERTEX"
CurExtension.MacroDefs.Add "GLU_END", "GLU_TESS_END"
CurExtension.MacroDefs.Add "GLU_ERROR", "GLU_TESS_ERROR"
CurExtension.MacroDefs.Add "GLU_EDGE_FLAG", "GLU_TESS_EDGE_FLAG"
End Sub
