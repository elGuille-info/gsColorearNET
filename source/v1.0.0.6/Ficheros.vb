'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' Comprobar el formato de un fichero                                (16/Nov/05)
'
' Primera versión: 21/Dic/03
'
' ©Guillermo 'guille' Som, 2003-2006, 2020
'------------------------------------------------------------------------------
Option Explicit On
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Text

''' <summary>
''' Clase con utilidades relacionadas con los ficheros
''' </summary>
Public NotInheritable Class Ficheros
    ' No permitir crear instancias de esta clase                (27/Ago/06)
    Private Sub New()
    End Sub

    '''<summary>
    ''' Averigua el formato del fichero indicado
    ''' Sólo se comprueban los formatos Unicode, UTF-8 y UTF-7 (manualmente)
    ''' si el formato no se puede averiguar, se devolverá ANSI (predeterminado de Windows)
    '''</summary>
    ''' <param name="fichero">
    ''' El fichero del que se va a comprobar el formato
    ''' </param>
    ''' <returns>
    ''' Devuelve un valor de tipo System.Text.Encoding
    ''' correspondiente al formato del fichero o 
    ''' System.Text.Encoding.Default si no se ha podido averiguar el formato
    ''' </returns>
    Public Shared Function FormatoFichero(fichero As String) As Encoding
        ' Por defecto devolver ANSI
        ' Incluso si es una cadena vacía o no existe            (26/Nov/05)
        ' Aunque se debería devolver Nothing, porque lo recomendable es
        ' comprobar que exista antes de llamar a esta función.
        If fichero Is Nothing OrElse fichero.Length = 0 Then
            Return Encoding.Default
        End If
        If File.Exists(fichero) = False Then
            Return Encoding.Default
        End If
        '
        ' Los ficheros Unicode tienen estos dos bytes: FF FE (normal o little-endian) o FE FF (big-endian)
        ' Los ficheros UTF-8 tienen estos tres bytes: EF BB BF
        Dim fs As System.IO.FileStream = Nothing
        Dim f = Encoding.Default
        Try
            ' Abrir el fichero y averiguar el formato
            ' Indicar que se abre para leer y se comparte       (13/Jun/06)
            ' para leer y escribir
            fs = New FileStream(fichero, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim c1 = fs.ReadByte
            Dim c2 = fs.ReadByte
            Dim c3 = fs.ReadByte

            If (c1 = &HFF AndAlso c2 = &HFE) OrElse (c1 = &HFE AndAlso c2 = &HFF) Then
                f = Encoding.Unicode
            ElseIf c1 = &HEF AndAlso c2 = &HBB AndAlso c3 = &HBF Then
                f = Encoding.UTF8
            Else
                ' comprobación del formato UTF-7                    (05/May/04)
                ' En el formato UTF-7 si tiene caracteres no ASCII contendrá: 2B 41 (+A)

                ' cerramos el fichero anterior, para poder abrirlo de nuevo
                fs.Close()
                Dim sr As New System.IO.StreamReader(fichero)
                Dim s As String
                ' El problema será cuando el fichero sea demasiado grande
                ' Nos arriesgamos a leer sólo una cantidad de caracteres, por ejemplo 100KB,
                ' sería mala suerte que en los primeros 100KB no hubiese caracteres "especiales"
                Const count As Integer = (100 * 1024)
                If sr.BaseStream.Length > count Then
                    Dim charBuffer(count - 1) As Char
                    sr.ReadBlock(charBuffer, 0, count)
                    s = charBuffer
                Else
                    s = sr.ReadToEnd()
                End If
                If s.IndexOf("+A") > -1 Then
#Disable Warning MSLIB0001 ' El tipo o el miembro están obsoletos
                    f = Encoding.UTF7
#Enable Warning MSLIB0001 ' El tipo o el miembro están obsoletos
                End If
                sr.Close()
            End If
        Catch 'ex As Exception
            ' si se produce algún error, asumimos que es ANSI
            f = Encoding.Default
        Finally
            ' Por si da error...                                (13/Jun/06)
            If fs IsNot Nothing Then
                fs.Close()
            End If
        End Try

        Return f
    End Function

End Class
