# GLEW_Translation
The OpenGL Extension Library for VB.NET by translate glew.h using VB6

# How to use
Simply using VB6 to open the project, run it, or compile and run. Then you will get the file named `GL_API.vb`, which is the product.

# Using OpenGL in VB.NET
Currently this library only runs in Windows. It's basically translated from the original glew.c source code.
When you get `GL_API.vb`, add it to your VB.NET project, then you have access to OpenGL function and constant definitions.

## Initialize
The usage of `GL_API.vb` is very similar to [GLEW: The OpenGL Extension Wrangler Library](http://glew.sourceforge.net/). Make an OpenGL context, call `GLAPI_Init()`, then use the extensions just like using GLEW.

First we need an OpenGL Context which should be created by calling `SetPixelFormat()` on a HDC then call `wglCreateContext()` to create the context, then call `wglMakeCurrent()` to use it. Check out *[the tutorial of how to create the OpenGL context](https://www.khronos.org/opengl/wiki/Creating_an_OpenGL_Context_(WGL))*

Then call `GLAPI_Init()`, just like using GLEW and call `glewInit()`.

## Using OpenGL Extensions
`GL_API.vb` provides a Boolean variable for all the extensions which indicates the availability of the extension. Check the variable when you need it.

For example:

    Function CompileShader(ByVal VS_code As String, ByVal FS_code As String) As Long
        Dim VS_code_Array(0) As String
        Dim FS_code_Array(0) As String
        VS_code_Array(0) = VS_code
        FS_code_Array(0) = FS_code

        Dim name_VS As UInt32
        Dim name_FS As UInt32
        Dim ShaderProgram As UInt32

        If GL_VERSION_2_0 Then 'HERE! Check the variable to know the extension is available or not
            name_VS = glCreateShader(GL_VERTEX_SHADER)
            name_FS = glCreateShader(GL_FRAGMENT_SHADER)

            glShaderSource(name_VS, 1, VS_code_Array, Len(VS_code))
            glShaderSource(name_FS, 1, FS_code_Array, Len(FS_code))

            glCompileShader(name_VS)
            glCompileShader(name_FS)

            ShaderProgram = glCreateProgram()
            glAttachShader(ShaderProgram, name_VS)
            glAttachShader(ShaderProgram, name_FS)
            glLinkProgram(ShaderProgram)

            glDeleteShader(name_VS)
            glDeleteShader(name_FS)

            Return ShaderProgram
        Else
            Return 0
        End If
     End Function
