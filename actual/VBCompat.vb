'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' VBCompat                                                          (10/Sep/20)
' Clase con instrucciones para compatibilidad con .NET Standard 2.0
'
' Declaro algunas funciones de Microsoft.VisualBasic.Strings
' que no están en .NET Standard 2.0
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

Public Class VBCompat

    ''' <summary>
    ''' Devuelve los caracteres desde la posición (en base 1)
    ''' hasta el final.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="pos"></param>
    Public Shared Function Mid(str As String, pos As Integer) As String
        Return str.Substring(pos - 1)
    End Function


    ''' <summary>
    ''' Devuelve la cadena desde la posición indicada con len caracteres.
    ''' La posición del primer carácter es el 1.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="pos"></param>
    ''' <param name="len"></param>
    ''' <remarks>10/Sep/20/20</remarks>
    Public Shared Function Mid(str As String, pos As Integer, len As Integer) As String
        Return str.Substring(pos - 1, len)
    End Function


    ''' <summary>
    ''' Devuelve el número de caracteres.
    ''' Si es cadena vacía o nulo devuelve 0.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Shared Function Len(str As String) As Integer
        If String.IsNullOrEmpty(str) Then Return 0
        Return str.Length
    End Function

    ''' <summary>
    ''' Devuelve los primeros caracteres de la cadena.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="len"></param>
    ''' <returns></returns>
    Public Shared Function Left(str As String, len As Integer) As String
        Return str.Substring(0, If(len > str.Length, str.Length, len))
    End Function

    ''' <summary>
    ''' Devuelve los caracteres indicados desde la derecha.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="len"></param>
    ''' <returns></returns>
    Public Shared Function Right(str As String, len As Integer) As String
        Dim iPos = str.Length - len
        Return str.Substring(iPos, len)
    End Function

    ''' <summary>
    ''' Devuelve la posición (en base 1) de la segunda cadena en la primera
    ''' </summary>
    ''' <param name="str1"></param>
    ''' <param name="str2"></param>
    ''' <returns></returns>
    Public Shared Function InStr(str1 As String, str2 As String) As Integer
        Return str1.IndexOf(str2) + 1
    End Function

    ''' <summary>
    ''' Devuelve la posición (en base 1) de la segunda cadena en la primera 
    ''' empezando en la posición indicada.
    ''' </summary>
    ''' <param name="startPos"></param>
    ''' <param name="str1"></param>
    ''' <param name="str2"></param>
    ''' <returns></returns>
    Public Shared Function InStr(startPos As Integer, str1 As String, str2 As String) As Integer
        Return str1.IndexOf(str2, startPos - 1) + 1
    End Function

    ''' <summary>
    ''' Devuelve una cadena después de haber quitado 
    ''' los espacios delante y detrás.
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Shared Function Trim(str As String) As String
        Return str.Trim
    End Function

    ''' <summary>
    ''' Divide una cadena en elementos de un array.
    ''' Usando el delimitador indicado.
    ''' </summary>
    ''' <param name="Expression"></param>
    ''' <param name="Delimiter"></param>
    ''' <returns></returns>
    Public Shared Function Split(Expression As String, Optional Delimiter As String = " ") As String()
        Return Expression.Split(Delimiter.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    End Function

    ''
    '' El código IL de Prueba1 es más corto (y parece que eficiente) que el de Prueba2
    ''
    'Public Shared Function Prueba1(str As String, len As Integer) As String
    '    Return str.Substring(0, If(len > str.Length, str.Length, len))
    'End Function

    'Public Shared Function Prueba2(str As String, len As Integer) As String
    '    If len > str.Length Then
    '        len = str.Length
    '    End If
    '    Return str.Substring(0, len)
    'End Function

End Class
