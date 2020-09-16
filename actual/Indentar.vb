'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' elGuille.UtilCode.Indentar                                        (10/Oct/04)
' Clase para indentar código BASIC
' Instrucciones soportadas de MSDOS-BASIC, VB6, VB.NET y VB2005
' En C# sólo se tendrá en cuenta el comentario de línea simple y
' los bloques encerrados entre llaves
'
' ©Guillermo 'guille' Som, 2004-2005, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.IO
Imports System.Text
Imports vb = gsColorearNET.VBCompat

''' <summary>
''' Clase para indentar el código
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Indentar

    Public Shared PrimerTab As Boolean = False      ' si se debe indentar desde el principio
    '                                                 esto es útil si se procesa un trozo de código
    Private Const espaciosTab As Integer = 4        ' espacios para cada tabulación
    Private Shared nTab As Integer                  ' Número de indentaciones en procesarLinea

    ' los métodos, etc. que estén dentro de Interface no se indentan
    Private Shared noEsInterfaz As Integer = 1

    ''' <summary>
    ''' Si es C# o C/C++ (o lenguajes que usen llaves)
    ''' </summary>
    Public Shared EsCSharp As Boolean               ' si es un fichero C#

    Private Shared instruccionesIni() As String
    Private Shared instruccionesFin() As String
    Private Shared instrucciones() As String

    ' Para poder indicar los espacios de indentación                (11/Oct/04)
    Private Shared _espaciosIndent As Integer = espaciosTab

    ''' <summary>
    ''' Los espacios que se van a usar para indentar
    ''' </summary>
    ''' <value>Un valor entre 0 y 8</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property EspaciosIndent() As Integer
        Get
            Return _espaciosIndent
        End Get
        Set(value As Integer)
            ' valores admitidos de 0 a 8
            If value >= 0 AndAlso value <= 8 Then
                _espaciosIndent = value
            End If
        End Set
    End Property

    Shared Sub New()
        ' asignar los valores predeterminados a usar en la clase
        ' No incluir las que tienen un trato especial, como:
        ' If, Select Case, Interface, End Select, End Interface

        ' Las instrucciones que tienen su equivalente con END instrucción
        instrucciones = New String() {
                    "NAMESPACE ", "CLASS ", "MODULE ", "STRUCTURE ",
                    "PROPERTY ", "SUB ", "FUNCTION ", "OPERATOR ",
                    "ENUM ", "USING ", "TYPE ",
                    "WHILE ", "WITH ", "TRY"}

        ' Las instrucciones para iniciar un bloque indentado,
        ' aquí se especifican las que no tienen equivalentes con END <instrucción>
        instruccionesIni = New String() {"FOR ", "DO"}

        ' instrucciones para terminar un bloque indentado,
        ' aquí se indican las que no tienen equivalentes con END <instrucción>
        instruccionesFin = New String() {"END IF", "END GET", "END SET", "NEXT", "LOOP", "WEND"}

        Dim n As Integer = instrucciones.Length - 1
        Dim m As Integer = instruccionesIni.Length
        Dim k As Integer = instruccionesFin.Length
        ReDim Preserve instruccionesIni(n + m)
        ReDim Preserve instruccionesFin(n + k)
        For i As Integer = 0 To n
            instruccionesIni(m + i) = instrucciones(i)
            instruccionesFin(k + i) = "END " & instrucciones(i).TrimEnd()
        Next
    End Sub

    '''<summary>
    ''' El método público para procesar el código a indentar
    '''</summary>
    ''' <param name="sr">
    ''' Un objeto TextReader con el fichero a indentar
    ''' </param>
    '''<returns>
    ''' Devuelve una cadena en la que cada línea termina con un retorno de carro
    '''</returns>
    ''' <remarks>
    ''' Existe tres sobrecargas que reciben parámetros de tipo
    ''' TextReader, array de String o una cadena
    ''' con cada línea separada con un retorno de carro.
    ''' </remarks>
    Public Shared Function ProcesarLineas(sr As TextReader) As String
        Dim sb As New StringBuilder
        nTab = 0
        noEsInterfaz = 1
        While sr.Peek <> -1
            sb.Append(String.Concat(procesarLinea(sr.ReadLine()), vbCrLf))
        End While

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' El método público para procesar el código a indentar
    ''' </summary>
    ''' <param name="lineas">
    ''' Un array de tipo String con las líneas a indentar
    ''' </param>
    ''' <returns>
    ''' Devuelve una cadena en la que cada línea termina con un retorno de carro
    ''' </returns>
    Public Shared Function ProcesarLineas(lineas() As String) As String
        Dim sb As New System.Text.StringBuilder
        nTab = 0
        noEsInterfaz = 1
        For i As Integer = 0 To lineas.Length - 1
            ' Usar vbCr en vez de vbCrLf                            (16/Sep/20)
            sb.Append(String.Concat(procesarLinea(lineas(i)), vbCr))
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' El método público para procesar el código a indentar
    ''' </summary>
    ''' <param name="sr">
    ''' Una cadena con el código a intentar.
    ''' Cada línea está separada con un retorno de carro.
    ''' </param>
    ''' <returns>
    ''' Devuelve una cadena en la que cada línea termina con un retorno de carro
    ''' </returns>
    Public Shared Function ProcesarLineas(sr As String) As String
        Dim lineas() As String = sr.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        Return ProcesarLineas(lineas)
    End Function

    '''<summary>
    ''' función interna para procesar cada línea
    ''' recibe como parámetro una cadena (una línea de código)
    '''</summary>
    '''<returns>
    ''' Devuelve una cadena ya indentada
    '''</returns>
    Private Shared Function procesarLinea(entrada As String) As String
        '------------------------------------------------------------------------
        ' Procesa los datos de entrada y los pone en salida.    (23.13 25/Nov/93)
        '
        ' Devuelve el valor de la salida
        '------------------------------------------------------------------------
        Dim indentar As Integer     ' Con cuantos espacios indentar esta línea
        Dim p As Integer
        Dim salida As String
        Dim hallado As Boolean
        '
        ' indentar siempre
        '
        indentar = nTab

        entrada = entrada.TrimStart(" "c, ChrW(9))
        If vb.Len(entrada) = 0 Then
            Return entrada
        End If

        ' si es un comentario o tiene comillas dobles, no comprobar nada
        If (EsCSharp = False AndAlso entrada.StartsWith("'")) OrElse
               (EsCSharp = True AndAlso (entrada.StartsWith("//") OrElse entrada.IndexOf(ChrW(34)) > -1)) Then
            If indentar > 0 Then
                entrada = New String(" "c, _espaciosIndent * indentar) & entrada
            End If
            Return entrada
        End If

        salida = entrada
        entrada = entrada.ToUpper

        ' En las comparaciones, sólo hay que asignar indentar,
        ' después de ajustar el valor de nTab,
        ' cuando esa instrucción deba indentarse

        ' Cuando se encuentra con estas instrucciones, se quitan los tabuladores
        ' y se deja para que la próxima línea se pueda indentar

        hallado = False

        If EsCSharp Then
            hallado = True
            ' sólo comprobar { y }
            If entrada.IndexOf("{") > -1 Then
                nTab += 1
            End If
            If entrada.IndexOf("}") > -1 Then
                nTab -= 1
                indentar = nTab
            End If
        End If

        If hallado = False Then
            hallado = True
            If vb.Left(entrada, 13) = "END INTERFACE" Then
                nTab -= 1
                indentar = nTab
                noEsInterfaz = 1
            ElseIf vb.InStr(" " & entrada, " INTERFACE ") > 0 Then
                nTab += 1
                noEsInterfaz = 0
                '
            ElseIf vb.Left(entrada, 3) = "GET" Then
                nTab += 1
            ElseIf vb.Left(entrada, 3) = "SET" Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PUBLIC GET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "FRIEND GET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PRIVATE GET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PROTECTED GET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PUBLIC SET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "FRIEND SET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PRIVATE SET") > 0 Then
                nTab += 1
            ElseIf vb.InStr(entrada, "PROTECTED SET") > 0 Then
                nTab += 1

                ' Instrucciones para BASIC de MS-DOS
            ElseIf vb.Left(entrada, 7) = "GLOBAL " Then
                indentar = 0
            ElseIf vb.InStr(entrada, "DECLARE ") > 0 Then
                ' Dejarlo como está
                indentar = 0
            ElseIf vb.Left(entrada, 5) = "BEGIN" Then
                nTab += 1
            ElseIf vb.Left(entrada, 9) = "ATTRIBUTE" Then
                indentar = 0
            ElseIf vb.Left(entrada, 6) = "COMMON" Then
                indentar = 0
            ElseIf vb.Left(entrada, 3) = "DEF" Then
                indentar = 0

                ' Estas instrucciones deben aparecer en la primera columna
            ElseIf vb.Left(entrada, 7) = "OPTION " Then
                indentar = 0
                'nTab = 0
            ElseIf vb.Left(entrada, 6) = "#CONST" Then
                indentar = 0
                'nTab = 0
            ElseIf vb.Left(entrada, 7) = "#REGION" Then
                indentar = 0
                nTab = 1
            ElseIf vb.Left(entrada, 11) = "#END REGION" Then
                nTab = 1
                indentar = 0

                ' Si se quiere que CASE esté al nivel de SELECT: cambiar el 2 por 1
            ElseIf vb.Left(entrada, 12) = "SELECT CASE " Then
                If vb.InStr(entrada, "END SELECT") = 0 Then
                    nTab = nTab + 2
                End If
            ElseIf vb.Left(entrada, 5) = "CASE " Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 10) = "END SELECT" Then
                nTab = nTab - 2
                indentar = nTab

            ElseIf vb.Left(entrada, 3) = "IF " OrElse vb.Left(entrada, 4) = "#IF " Then
                p = vb.InStr(entrada, "THEN")
                If p > 0 Then
                    If vb.Right(entrada, 4) = "THEN" Then
                        nTab += 1
                    Else
                        ' comprobar si hay un comentario
                        If vb.InStr(p + 1, entrada, "'") > 0 Then
                            nTab += 1
                        End If
                    End If
                End If
            ElseIf vb.Left(entrada, 4) = "ELSE" Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 7) = "ELSEIF " Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 5) = "#ELSE" Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 8) = "#ELSEIF " Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 7) = "#END IF" Then
                nTab -= 1
                indentar = nTab

            ElseIf vb.Left(entrada, 5) = "CATCH" Then
                indentar = indentar - 1
            ElseIf vb.Left(entrada, 7) = "FINALLY" Then
                indentar = indentar - 1

            Else
                hallado = False
            End If
        End If

        If hallado = False Then
            For i As Integer = 0 To instruccionesFin.Length - 1
                If (" " & entrada).IndexOf(" " & instruccionesFin(i)) > -1 Then
                    nTab -= 1 * noEsInterfaz
                    indentar = nTab
                    hallado = True
                    Exit For
                End If
            Next
        End If
        If hallado = False Then
            ' aquí puede darse que se encuentren los EXIT <instrucción>
            For i As Integer = 0 To instruccionesIni.Length - 1
                p = (" " & entrada).IndexOf(" " & instruccionesIni(i))
                Dim k As Integer = entrada.IndexOf("'")
                If k = -1 Then k = p + 1
                If (p > -1 AndAlso p < k) AndAlso
                       entrada.IndexOf("EXIT " & instruccionesIni(i)) = -1 Then
                    nTab += 1 * noEsInterfaz
                    hallado = True
                    Exit For
                End If
            Next
        End If

        If indentar = 0 Then
            If PrimerTab Then
                indentar = 1
                nTab = 1
            End If
        End If
        If indentar > 0 Then
            ' Se puede sustituir esto por CHR$(9)en lugar de 4 espacios
            ' es que yo suelo usar 4 espacios para cada tabulacion
            salida = New String(" "c, _espaciosIndent * indentar) & salida
        End If

        Return salida
    End Function
End Class
