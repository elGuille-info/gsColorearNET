'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' cWrap                                                             (13/Jun/98)
' Clase para efectuar "cortes" de palabras de forma apropiada
'
' Revisado el  4/Ene/1999
' Revisado el 20/Ago/2001   Nueva función: LoopPropperWrap
' Revisado el 08/Oct/2002   Algunos ajustes cuando la cadena contiene intro
' Revisado el 30/Nov/2005   Convertida a VB2005 y PropperText
'
' ©Guillermo 'guille' Som, 1998-2002, 2005, 2020
'
' Esta clase tiene los siguientes métodos (funciones)
'   Justificar      Justifica la cadena,
'                   añadiendo espacios hasta conseguir la longitud deseada
'   PropperJust     Justifica la cadena según los caracteres indicados
'                   Esto sólo será útil si el resultado se muestra con fuente
'                   no proporcional
'   PropperWrap     Es como las siguientes, pero se debe especificar por dónde
'                   empezar a contar los caracteres.
'   PropperLeft     como Left$(Cadena, longitud) pero sin cortar palabras
'   PropperMid      como Mid$(Cadena, longitud) pero sin cortar palabras
'   PropperRight    como Right$(Cadena, longitud) pero sin cortar palabras
'
'   LoopPropperWrap Bucle para desglosar un texto en trozos de
'                   la longitud indicada
'   PropperText     Devuelve un texto con líneas de la longitud indicada.
'                   Utiliza LoopPropperWrap
'
'   Separadores     Para indicar los separadores a usar
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports gsColorearNET.VBCompat
Imports System

''' <summary>
''' Clase para realizar cortes de palabras de forma apropiada
''' </summary>
Public Class Wrap
    Const cSeparadores As String = " ªº\!|@#$%&/()=?¿'¡[]*+{}<>,.-;:_"

    ''' <summary>
    ''' Alineación a usar con PropperWrap.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ePropperWrapConstants
        ''' <summary>
        ''' Para justificar desde la izquierda
        ''' </summary>
        pwLeft = 0
        ''' <summary>
        ''' Para justificar a partir de la posición que se indique
        ''' </summary>
        pwMid = 1
        ''' <summary>
        ''' Para justificar por la derecha
        ''' </summary>
        pwRight = 2
        '    pwIzquierda = 0
        '    pwCentro = 1
        '    pwDerecha = 2
    End Enum

    ''' <summary>
    ''' Devuelve la cadena que habría que imprimir para mostrar
    ''' los caracteres indicados, pero sin cortar una palabra.
    ''' </summary>
    ''' <param name="sCadena">La cadena a procesar</param>
    ''' <param name="nCaracteres">
    ''' Posición desde la que se justificará
    ''' </param>
    ''' <returns>
    ''' </returns>
    Public Shared Function PropperWrap(sCadena As String,
                                       nCaracteres As Integer
                                       ) As String
        Return PropperWrap(sCadena, nCaracteres, ePropperWrapConstants.pwLeft)
    End Function


    ''' <summary>
    ''' Devuelve la cadena que habría que imprimir para mostrar
    ''' los caracteres indicados, pero sin cortar una palabra.
    ''' </summary>
    ''' <param name="sCadena">La cadena a procesar</param>
    ''' <param name="nCaracteres">
    ''' Posición desde la que se justificará
    ''' </param>
    ''' <param name="desdeDonde">
    ''' Un valor de la enumeración 
    ''' <seealso cref="ePropperWrapConstants">ePropperWrapConstants</seealso>
    ''' </param>
    ''' <returns></returns>
    Public Shared Function PropperWrap(sCadena As String,
                                       nCaracteres As Integer,
                                       desdeDonde As ePropperWrapConstants
                                       ) As String
        ' Devuelve la cadena que habría que imprimir para mostrar los
        ' caracteres indicados, sin cortar una palabra.
        ' Esto es para los casos en los que se quiera usar:
        ' Left$(sCadena,nCaracteres) o Mid$/Right$(sCadena,nCaracteres)
        ' pero sin cortar una palabra
        Dim i As Integer
        '
        i = InStr(sCadena, vbCrLf)
        If i > 0 And i < nCaracteres Then
            sCadena = Left(sCadena, i + 1)
        ElseIf nCaracteres > Len(sCadena) Then
            i = InStr(sCadena, vbCrLf)
            If i > 0 Then
                sCadena = Left(sCadena, i - 1)
            End If
        Else
            For i = nCaracteres To 1 Step -1
                If InStr(Separadores, Mid(sCadena, i, 1)) > 0 Then
                    ' Si se especifica desde la izquierda
                    If desdeDonde = ePropperWrapConstants.pwLeft Then
                        sCadena = Left(sCadena, i)
                    Else
                        ' lo mismo da desde el centro que desde la derecha
                        sCadena = Mid(sCadena, i + 1)
                    End If
                    Exit For
                End If
            Next
        End If
        Return sCadena
    End Function

    ''' <summary>
    ''' Justifica la cadena (sin cortar palabras) a la derecha
    ''' </summary>
    ''' <param name="sCadena"></param>
    ''' <param name="nCaracteres"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PropperRight(sCadena As String, nCaracteres As Integer) As String
        Return PropperWrap(sCadena, nCaracteres, ePropperWrapConstants.pwRight)
    End Function

    ''' <summary>
    ''' Justifica la cadena (sin cortar palabras) 
    ''' </summary>
    ''' <param name="sCadena"></param>
    ''' <param name="nCaracteres"></param>
    ''' <param name="RestoNoUsado"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PropperMid(sCadena As String,
                                      nCaracteres As Integer,
                                      Optional RestoNoUsado As Integer = 0
                                      ) As String
        Return PropperWrap(sCadena, nCaracteres, ePropperWrapConstants.pwMid)
    End Function

    ''' <summary>
    ''' Justifica la cadena (sin cortar palabras) a la izquierda
    ''' </summary>
    ''' <param name="sCadena"></param>
    ''' <param name="nCaracteres"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PropperLeft(sCadena As String, nCaracteres As Integer) As String
        Return PropperWrap(sCadena, nCaracteres, ePropperWrapConstants.pwLeft)
    End Function

    ''' <summary>
    ''' Justifica la cadena según los caracteres indicados.
    ''' </summary>
    ''' <param name="cadena"></param>
    ''' <param name="longitud"></param>
    ''' <param name="justif"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PropperJust(cadena As String,
                                       Optional longitud As Integer = 70,
                                       Optional justif As Boolean = True) As String
        '--------------------------------------------------------------------------
        ' Justifica la cadena según los caracteres indicados            ( 3/Ene/99)
        ' Esto sólo será útil si el resultado se muestra con fuente no proporcional
        ' Valores de entrada:
        '   Cadena      Cadena a manipular
        '   Longitud    Longitud de cada línea, por defecto 70 caracteres
        '   Justificar  Si se justifica, rellenando con espacios, por defecto Si
        ' Devuelve:
        '   La cadena una vez manipulada
        '--------------------------------------------------------------------------
        Dim sTmp As String
        Dim sTmp2 = ""

        Do
            ' Los cambios de línea se consideran por separado
            Dim i = InStr(cadena, vbCrLf)
            If i > 0 Then
                sTmp = Left(cadena, i - 1)
                cadena = Mid(cadena, i + 2)
            Else
                sTmp = cadena
                cadena = ""
            End If
            Do
                Dim sLinea = PropperWrap(sTmp, longitud, ePropperWrapConstants.pwLeft)
                If sTmp = sLinea Then
                    ' no justificar cuando es el final de línea
                    sTmp = ""
                Else
                    sTmp = Mid(sTmp, Len(sLinea) + 1)
                    If justif Then
                        sLinea = Justificar(sLinea, longitud)
                    End If
                End If
                sTmp2 &= sLinea & vbCrLf
            Loop While Len(sTmp) > 0
        Loop While Len(cadena) > 0

        Return sTmp2
    End Function

    ''' <summary>
    ''' Justifica la cadena, añadiendo espacios
    ''' hasta conseguir la longitud deseada.
    ''' </summary>
    ''' <param name="cadena"></param>
    ''' <param name="longitud"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Justificar(cadena As String,
                                      Optional longitud As Integer = 70
                                      ) As String
        ' Justifica la cadena, añadiendo espacios hasta conseguir la longitud deseada
        cadena = Trim(cadena)
        If Len(cadena) < longitud Then
            Dim k = 1
            Dim n = 0

            Dim hallado = False
            Do
                For i = 1 To Len(Separadores)
                    Dim j = InStr(k, cadena, Mid(Separadores, i, 1))
                    If j > 0 Then
                        cadena = Left(cadena, j) & " " & Mid(cadena, j + 1)
                        k = j + 1
                        ' Buscar el siguiente caracter que no sea un separador
                        For j = k + 1 To Len(cadena)
                            If InStr(Separadores, Mid(cadena, j, 1)) = 0 Then
                                k = j
                                Exit For
                            End If
                        Next
                        hallado = True
                        n = n + 1
                        Exit For
                    Else
                        k = 1
                        hallado = False
                    End If
                Next
                If Not hallado Then
                    k = 1
                    If n = 0 Then
                        cadena &= " "
                    End If
                End If
            Loop While Len(cadena) < longitud
        End If
        Return Left(cadena, longitud)
    End Function

    ''' <summary>
    ''' Para indicar los separadores a usar
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property Separadores() As String = vbCr & vbLf & vbTab & cSeparadores & ChrW(34)

    ' Para usar LoopPropperWrap                                 (30/Nov/05)
    ' de esta forma devuelve el texto correcto de una vez

    ''' <summary>
    ''' Devuelve el código ajustado de una pasada.
    ''' Internamente usa <seealso cref="LoopPropperWrap">
    ''' LoopPropperWrap</seealso>.
    ''' </summary>
    ''' <param name="sCadena"></param>
    ''' <param name="nCaracteres"></param>
    ''' <param name="desdeDonde"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PropperText(sCadena As String,
                                       Optional nCaracteres As Integer = 70,
                                       Optional desdeDonde As ePropperWrapConstants = ePropperWrapConstants.pwLeft
                                       ) As String
        Dim sb As New System.Text.StringBuilder
        Dim s = sCadena
        sb.AppendFormat("{0}{1}", LoopPropperWrap(s, nCaracteres, desdeDonde), vbCrLf)
        While Len(s) > 0
            s = LoopPropperWrap()
            If Len(s) > 0 Then
                sb.AppendFormat("{0}{1}", s, vbCrLf)
            End If
        End While
        Return sb.ToString.TrimEnd()
    End Function

    ''' <summary>
    ''' Bucle para desglosar un texto en trozos de la longitud indicada.
    ''' En la primera llamada se indican todos los parámetros,
    ''' en las siguientes se dejan en blanco.
    ''' </summary>
    ''' <param name="sCadena"></param>
    ''' <param name="nCaracteres"></param>
    ''' <param name="desdeDonde"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoopPropperWrap(Optional sCadena As String = "",
                                           Optional nCaracteres As Integer = 70,
                                           Optional desdeDonde As ePropperWrapConstants = ePropperWrapConstants.pwLeft
                                           ) As String
        ' Repite la justificación hasta que la cadena esté vacia        (20/Ago/01)
        ' Devolviendo cada vez el número de caracteres indicados
        Static sCadenaCopia As String
        Static nCaracteresCopia As Integer
        Static desdeDondeCopia As ePropperWrapConstants

        ' Si la cadena es una cadena vacía, es que se continua "partiendo"
        ' sino, es la primera llamada
        If Len(sCadena) > 0 Then
            sCadenaCopia = sCadena
            nCaracteresCopia = nCaracteres
            desdeDondeCopia = desdeDonde
        Else
            ' Asignar los valores que había antes
            sCadena = sCadenaCopia
            nCaracteres = nCaracteresCopia
            desdeDonde = desdeDondeCopia
        End If

        Dim s = PropperWrap(sCadena, nCaracteres, desdeDonde)
        sCadenaCopia = Mid(sCadena, Len(s) + 1)

        Return s.TrimEnd
    End Function
End Class
