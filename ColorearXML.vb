'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' Clase para colorear los ficheros XML                              (28/Ago/06)
' También vale para HTML, ASP, etc.
'
' El coloreado de XML tiene un trato especial:
' Los signos se colorean en azul: < /> ? = <!-- --> [
' El texto entrecomillado (simple o doble) se colorea en azul
' Las palabras que están justo después de < y justo antes de >
'   se colorean en rojo "ladrillo"
' Las palabras que están después (atributos) en rojo
' Las palabras entre & y ; en rojo
' Las palabras que están entre < y : en color azul-verdoso
' Lo que esté entre <style> y </style> no se colorea.
' Tratar los tags de ASP y ASPX de forma especial:
'   <% %> <%@ marcando esa instrucciones con fondo amarillo.
' Código CDATA en gris
'   <![CDATA[ ... ]]>
'
' ©Guillermo 'guille' Som, 2005-2006
'------------------------------------------------------------------------------
Option Strict On

Imports Microsoft.VisualBasic
Imports vb = gsColorearNET.VBCompat ' Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

'Namespace elGuille.Util.Developer

''' <summary>
''' Clase para colorear los ficheros XML
''' </summary>
Public NotInheritable Class ColorearXML
    ' Constructor privado para evitar crear instancias
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Devuelve los colores RGB para la definición del RTF
    ''' </summary>
    ''' <param name="elColor">La cadena con el color de style</param>
    ''' <returns>La cadena para el formato de RTF</returns>
    Private Shared Function colorRTF_RGB(ByVal elColor As String) As String
        Dim r As Integer = CInt("&H" & vb.Mid(elColor, 2, 2))
        Dim g As Integer = CInt("&H" & vb.Mid(elColor, 4, 2))
        Dim b As Integer = CInt("&H" & vb.Mid(elColor, 6, 2))
        Return String.Format("\red{0}\green{1}\blue{2}", r, g, b)
    End Function

    Private Shared _colorTexto As String = "#0000FF" ' Textos
    Private Shared ReadOnly Property colorTexto() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorTexto
            Else
                Return colorRTF_RGB(_colorTexto)
            End If
        End Get
    End Property
    Private Shared _colorDelimitador As String = "#0000A0" ' Delimitadores
    Private Shared ReadOnly Property colorDelimitador() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorDelimitador
            Else
                Return colorRTF_RGB(_colorDelimitador)
            End If
        End Get
    End Property
    Private Shared _colorAtributo As String = "#FF0000" ' Atributos
    Private Shared ReadOnly Property colorAtributo() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorAtributo
            Else
                Return colorRTF_RGB(_colorAtributo)
            End If
        End Get
    End Property
    Private Shared _colorElemento As String = "#B22222" ' Elementos
    Private Shared ReadOnly Property colorElemento() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorElemento
            Else
                Return colorRTF_RGB(_colorElemento)
            End If
        End Get
    End Property
    Private Shared _colorCyan As String = "#2B91AF" ' Palabras clave XSLT <nnn:mmmm 
    Private Shared ReadOnly Property colorCyan() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorCyan
            Else
                Return colorRTF_RGB(_colorCyan)
            End If
        End Get
    End Property
    Private Shared _colorComentario As String = "#008000" ' Comentarios
    Private Shared ReadOnly Property colorComentario() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorComentario
            Else
                Return colorRTF_RGB(_colorComentario)
            End If
        End Get
    End Property
    Private Shared _colorCDATA As String = "#808080" ' "&H5C5C5C" ' Código CDATA
    Private Shared ReadOnly Property colorCDATA() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorCDATA
            Else
                Return colorRTF_RGB(_colorCDATA)
            End If
        End Get
    End Property
    Private Shared _colorAmarillo As String = "#FFFF00"
    Private Shared ReadOnly Property colorAmarillo() As String
        Get
            If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
                Return _colorAmarillo
            Else
                Return colorRTF_RGB(_colorAmarillo)
            End If
        End Get
    End Property
    '
    ' Métodos para devolver la cadena indicada en el color
    Private Shared Function colorearTexto(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorTexto,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf1 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearDelimitador(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorDelimitador,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf2 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearAtributo(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorAtributo,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf3 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearElemento(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorElemento,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf4 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearCyan(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorCyan,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf5 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearComentario(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<i><span style=""color:{0}"">{1}</span></i>",
                                colorComentario,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf6 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearCDATA(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""color:{0}"">{1}</span>",
                                colorCDATA,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf7 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function colorearFondoAmarillo(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return String.Format("<span style=""background-color:{0}"">{1}</span>",
                                colorAmarillo,
                                texto.Replace("<", "&lt;").Replace(">", "&gt;"))
        Else
            Return "\cf8 " & texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    Private Shared Function noColorear(ByVal texto As String) As String
        If FormatoColoreado = Colorear.FormatosColoreado.HTML Then
            Return texto.Replace("<", "&lt;").Replace(">", "&gt;")
        Else
            'Return String.Format("\cf0{0}", texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}"))
            Return texto.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
        End If
    End Function
    '
    '
    'Private Shared separadores As String = " " & vbTab
    '
    ''' <summary>
    ''' Los delimitadores usados en XML y HTML
    ''' </summary>
    Private Shared delimIni() As String = {"<style>",
                                               "sñ·", "sñç",
                                               "<!--",
                                               "<?xml",
                                               "<![CDATA[",
                                               "<%", "<!",
                                               "<"} ', _
    '
    Private Shared delimFin() As String = {"/>", ">"}
    '
    Public Shared FormatoColoreado As Colorear.FormatosColoreado = Colorear.FormatosColoreado.HTML
    '
    ''' <summary>
    ''' Colorea un texto con formato XML o HTML.
    ''' Cada línea estará separada con un retorno de carro.
    ''' </summary>
    ''' <param name="texto">El texto a colorear</param>
    ''' <returns>Una cadena con el texto una vez coloreado</returns>
    Public Shared Function ColorearXml(ByVal texto As String) As String
        Dim sb As New System.Text.StringBuilder
        '
        ' Para poder colorear también en RTF                    (17/Nov/06)
        If FormatoColoreado = Colorear.FormatosColoreado.RTF Then
            ' Cabecera del fichero RTF
            sb.AppendFormat("{{\rtf1\ansi\ansicpg1252\deff0{{\fonttbl{{\f0\fnil\fcharset0 {0};}}}}", Colorear.Fuente)
            'sb.Append(vbCrLf)
            ' Definición de los colores a usar
            ' Parece que solo se queda con los 4 primeros       (17/Nov/06)
            sb.AppendFormat("{{\colortbl ;{0};{1};{2};{3};{4};{5};{6};{7};}}", colorTexto, colorDelimitador, colorAtributo, colorElemento, colorCyan, colorComentario, colorCDATA, colorAmarillo)
            'sb.Append(vbCrLf)
            ' Ni idea de que es esto, pero...
            '------------------------------------------------------------------
            ' \viewkind4 es el estilo de visión del documento 4 = normal
            ' \pard Resets to default paragraph properties.
            ' Para más info del formato RTF_ http://www.biblioscape.com/rtf15_spec.htm
            '------------------------------------------------------------------
            sb.AppendFormat("\viewkind4\uc1\pard\lang3082\f0\fs{0} ", CInt(Colorear.FuenteTam) * 2)
        Else
            sb.AppendFormat("<style>pre{{font-family:{0}; font-size:{1}.0pt;}}", Colorear.Fuente, Colorear.FuenteTam)
            'sb.AppendFormat("p.Code{{font-family:{0}; font-size:{1}.0pt;}}", "Fixedsys", "9")
            sb.AppendFormat("</style>{0}", vbCrLf)
            sb.Append(Colorear.PreTag)
        End If
        '
        'sb.Append("<pre>")
        '
        ' Lo que haya entre > y < no analizarlo (si antes se ha encontrado un delimitador)
        Dim iDelim As Integer = -1
        Dim sDelim As String = ""
        Dim linea As String = ""
        Dim n As Integer = -1
        '
        Dim comillas(), comillas2() As String
        ' Las comillas dobles estarán entre sñ· y ·ñs
        ' Las comillas simples estarán entre sñç y çñs
        comillas = Utilidades.agruparPares(texto, "·"c, ChrW(34))
        comillas2 = Utilidades.agruparPares(texto, "ç"c, "'")
        '
        linea = texto
        While String.IsNullOrEmpty(linea) = False
            iDelim = -1
            sDelim = buscarDelimitador(linea, iDelim)
            If iDelim = -1 Then
                sb.AppendFormat("{0}{1}", noColorear(linea), vbCrLf)
                linea = ""
                Exit While
            End If
            ' sDelim tiene el delimitador
            ' Comprobar lo que hay que hacer:
            Select Case sDelim
                Case "<style>"
                    sb.Append(noColorear(linea.Substring(0, iDelim)))
                    'sb.Append(colorearElemento(sDelim))
                    sb.AppendFormat("{0}{1}{2}",
                                    colorearDelimitador("<"),
                                    colorearElemento("style"),
                                    colorearDelimitador(">"))
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Buscar </style> en esta línea
                    n = linea.IndexOf("</style>")
                    If n > -1 Then
                        sb.Append(noColorear(linea.Substring(0, n)))
                        sb.AppendFormat("{0}{1}{2}",
                                        colorearDelimitador("</"),
                                        colorearElemento("style"),
                                        colorearDelimitador(">"))
                        linea = linea.Substring(n + "</style>".Length)
                    Else
                        sb.AppendFormat("{0}{1}", noColorear(linea), vbCrLf)
                        linea = ""
                    End If
                Case "<!--"
                    'esRem = True
                    ' Buscar el final "-->" 
                    ' y pintarlo en verde
                    sb.Append(noColorear(linea.Substring(0, iDelim)))
                    sb.Append(colorearDelimitador(sDelim))
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Buscar --> en esta línea
                    n = linea.IndexOf("-->")
                    If n > -1 Then
                        sb.Append(colorearComentario(linea.Substring(0, n)))
                        sb.Append(colorearDelimitador("-->"))
                        linea = linea.Substring(n + "-->".Length)
                        'esRem = False
                    Else
                        sb.AppendFormat("{0}{1}", colorearComentario(linea), vbCrLf)
                        linea = ""
                        'esRem = True
                    End If
                Case "<?xml"
                    ' Buscar el par siguiente
                    ' y buscar la palabra, etc.
                    If iDelim > 0 Then
                        sb.Append(noColorear(linea.Substring(0, iDelim)))
                    End If
                    sb.Append(colorearDelimitador("<?"))
                    sb.Append(colorearElemento("xml"))
                    '
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Todo lo que haya hasta ?> serán atributos
                    sDelim = analizarAtributos(linea, "?>", False)
                    If String.IsNullOrEmpty(sDelim) = False Then
                        sb.Append(sDelim)
                        sb.Append(colorearDelimitador("?>"))
                    Else
                        Stop
                        sb.Append(colorearDelimitador("?>"))
                    End If
                        '
                Case "<%"
                    ' Buscar hasta el %> y debe ser atributos
                    If iDelim > 0 Then
                        sb.Append(noColorear(linea.Substring(0, iDelim)))
                    End If
                    sb.Append(colorearFondoAmarillo(sDelim))
                    '
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Todo lo que haya hasta %> serán atributos
                    ' Puede que haya un @, si es así,
                    ' ponerlo en azul y analizar el resto
                    If linea.TrimStart().StartsWith("@") Then
                        iDelim = linea.IndexOf("@")
                        sb.Append(colorearDelimitador(linea.Substring(0, iDelim + 1)))
                        linea = linea.Substring(iDelim + 1)
                        iDelim = 0
                        While linea(iDelim) = " "c '.StartsWith(" ")
                            iDelim += 1
                            sb.Append(" ")
                        End While
                        If iDelim > 0 Then
                            linea = linea.Substring(iDelim)
                        End If
                    End If
                    ' Aunque el primero será un elemento
                    'sDelim = analizarAtributos(linea, "%>", False)
                    sDelim = analizarAtributos(linea, "%>", True)
                    If String.IsNullOrEmpty(sDelim) = False Then
                        sb.Append(sDelim)
                        sb.Append(colorearFondoAmarillo("%>"))
                    Else
                        Stop
                        sb.Append(colorearFondoAmarillo("%>"))
                    End If
                Case "<", "<!"
                    If iDelim > 0 Then
                        sb.Append(noColorear(linea.Substring(0, iDelim)))
                    End If
                    sb.Append(colorearDelimitador(sDelim))
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Todo lo que haya hasta > o /> serán atributos
                    iDelim = -1
                    sDelim = buscarDelimitadorFin(linea, iDelim)
                    Dim sDelim2 As String = sDelim
                    If String.IsNullOrEmpty(sDelim) = False Then
                        sDelim = analizarAtributos(linea, sDelim, True)
                        If String.IsNullOrEmpty(sDelim) = False Then
                            sb.Append(sDelim)
                            sb.Append(colorearDelimitador(sDelim2))
                        Else
                            Stop
                            sb.Append(colorearDelimitador(sDelim2))
                        End If
                    Else
                        iDelim = iDelim
                        ' Seguir, pero dejar el aviso de error  (21/Dic/07)
                        sb.Append("<br/>ERROR delimitador no emparejado")
                        Exit While
                        'Stop
                    End If
                Case "<![CDATA["
                    ' Buscar el final "]]>"
                    ' y pintarlo en gris
                    If iDelim > 0 Then
                        sb.Append(noColorear(linea.Substring(0, iDelim)))
                    End If
                    sb.Append(colorearDelimitador(sDelim))
                    '
                    linea = linea.Substring(iDelim + sDelim.Length)
                    ' Todo lo que haya hasta ]]> serán código
                    n = linea.IndexOf("]]>")
                    If n = -1 Then
                        sb.Append("ERROR no hay ]]&gt;")
                        Exit While
                    End If
                    sb.Append(colorearCDATA(linea.Substring(0, n)))
                    sb.Append(colorearDelimitador("]]>"))
                    linea = linea.Substring(n + "]]>".Length)
                Case "sñ·", "sñç"
                    ' Si lo encuentra aquí, es que no es un atributo
                    ' por tanto dejarlo en el color que esté
                    '
                    'If iDelim > 0 Then
                    '    sb.Append(noColorear(linea.Substring(0, iDelim)))
                    'End If
                    'sb.Append(sDelim)
                    ''sb.Append(noColorear(linea.Substring(0, iDelim)))
                    'linea = linea.Substring(iDelim + 1)
                    Dim sDelim2 As String
                    If sDelim = "sñ·" Then
                        sDelim2 = "·ñs"
                    Else
                        sDelim2 = "çñs"
                    End If
                    n = linea.IndexOf(sDelim2)
                    'sb.Append(colorearTexto(linea.Substring(0, n)))
                    'sb.Append(sDelim2)
                    sb.Append(linea.Substring(0, n + sDelim2.Length))
                    linea = linea.Substring(n + sDelim2.Length)
                    '
                    'Case Else
                    '    ' A ver que es lo que llega aquí
                    '    System.Diagnostics.Debug.WriteLine(sDelim)
                    '
                    'Case ChrW(34), "'"
                    '    ' Buscar el siguiente
                    '    ' y pintarlo en azul
                    '    sb.Append(noColorear(linea.Substring(0, iDelim)))
                    '    linea = linea.Substring(iDelim + 1)
                    '    iDelim = linea.IndexOf(sDelim)
                    '    If iDelim = -1 Then
                    '        sb.Append("ERROR faltan comillas de cierre")
                    '        Exit While
                    '    End If
                    '    sb.Append(sDelim)
                    '    sb.Append(colorearTexto(linea.Substring(0, iDelim)))
                    '    sb.Append(sDelim)
                    '    linea = linea.Substring(iDelim + 1)
            End Select
        End While

        If FormatoColoreado = Colorear.FormatosColoreado.RTF Then
            ' Acaba con } y un valor nulo (Chrw(0))
            sb.Replace(vbLf, "\par")
            'sb.Replace(vbTab, "\tab")
            sb.Append("}")
            sb.Append(ChrW(0))
        Else
            sb.Append("</pre>")
        End If

        ' Quitarle los < y > que tenga
        For i As Integer = 0 To comillas.Length - 1
            comillas(i) = noColorear(comillas(i))
        Next
        For i As Integer = 0 To comillas2.Length - 1
            comillas2(i) = noColorear(comillas2(i))
        Next
        texto = Utilidades.reagruparPares(comillas, sb.ToString, "·"c)
        texto = Utilidades.reagruparPares(comillas2, texto, "ç"c)

        Return texto
    End Function

    ''' <summary>
    ''' Devuelve el primer delimitador o una cadena vacía si no hay
    ''' </summary>
    ''' <param name="texto">El texto a analizar</param>
    ''' <param name="index">El índice dentro del texto (por referencia)</param>
    ''' <returns>
    ''' Si se encuentra un delimitador,
    ''' devuelve el delimitador y la posición en index.
    ''' Si no hay, devuelve una cadena vacía y -1 en index.
    ''' </returns>
    Private Shared Function buscarDelimitador(ByVal texto As String, ByRef index As Integer) As String
        If String.IsNullOrEmpty(texto) Then
            index = -1
            Return ""
        End If
        '
        Dim ret As Integer = -1
        Dim n As Integer = 0
        Dim k As Integer = texto.Length
        Dim b As Boolean = False
        '
        For i As Integer = 0 To delimIni.Length - 1
            n = texto.IndexOf(delimIni(i))
            If n > -1 AndAlso n < k Then
                ret = i
                k = n
                b = True
            End If
        Next
        If b = False Then
            index = -1
            Return ""
        End If
        '
        index = k
        Return delimIni(ret)
    End Function

    ''' <summary>
    ''' Devuelve el primer delimitador final o una cadena vacía si no hay
    ''' </summary>
    ''' <param name="texto">El texto a analizar</param>
    ''' <param name="index">El índice dentro del texto (por referencia)</param>
    ''' <returns>
    ''' Si se encuentra un delimitador de término,
    ''' devuelve el delimitador y la posición en index.
    ''' Si no hay, devuelve una cadena vacía y -1 en index.
    ''' </returns>
    Private Shared Function buscarDelimitadorFin(ByVal texto As String, ByRef index As Integer) As String
        If String.IsNullOrEmpty(texto) Then
            index = -1
            Return ""
        End If
        '
        Dim ret As Integer = -1
        Dim n As Integer = 0
        Dim k As Integer = texto.Length
        Dim b As Boolean = False
        '
        '' No buscar dentro de comillas
        'Dim comillas As String = ""
        'n = texto.IndexOfAny((ChrW(34) & "'").ToCharArray)
        'If n > -1 Then
        '    k = texto.IndexOf(texto(n), n + 1)
        '    comillas = texto.Substring(n, k - n + 1)
        '    texto = texto.Substring(0, n) & "%%%COMILLAS%%%" & texto.Substring(k + 1)
        '    k = texto.Length
        'End If
        '
        For i As Integer = 0 To delimFin.Length - 1
            n = texto.IndexOf(delimFin(i))
            If n > -1 AndAlso n < k Then
                ret = i
                k = n
                b = True
            End If
        Next
        If b = False Then
            index = -1
            Return ""
        End If
        '
        index = k
        Return delimFin(ret)
    End Function


    ''' <summary>
    ''' Analiza el texto y procesa los atributos
    ''' </summary>
    ''' <param name="texto">
    ''' El texto a analizar.
    ''' Se devuelve lo que haya que seguir analizando,
    ''' que estará después del 
    ''' <paramref name="final">final</paramref> indicado.
    ''' </param>
    ''' <param name="final">
    ''' El final del elemento (hasta donde hay que analizar)
    ''' </param>
    ''' <returns>
    ''' Devuelve la cadena analizada hasta el atributo final.
    ''' En el parámetro <paramref name="texto">texto</paramref>
    ''' se devuelve el resto del texto a procesar.
    ''' </returns>
    Private Shared Function analizarAtributos(
                        ByRef texto As String,
                        ByVal final As String,
                        ByVal primerTag As Boolean) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        Dim linea As String
        '
        Dim n, k As Integer
        '
        '' No buscar dentro de comillas
        'Dim comillas As String = ""
        'n = texto.IndexOfAny((ChrW(34) & "'").ToCharArray)
        'If n > -1 Then
        '    k = texto.IndexOf(texto(n), n + 1)
        '    comillas = texto.Substring(n, k - n + 1)
        '    texto = texto.Substring(0, n) & "%%%COMILLAS%%%" & texto.Substring(k + 1)
        '    k = texto.Length
        'End If
        '
        i = texto.IndexOf(final)
        If i = -1 Then Return ""
        linea = texto.Substring(0, i)
        texto = texto.Substring(i + final.Length)
        '
        ' El texto será: <sep><atributo>={"|'}<texto>{"|'}
        If primerTag Then
            If linea.StartsWith("/") Then
                sb.Append(colorearDelimitador("/"))
                linea = linea.Substring(1)
            End If
            ' Buscar el primer elemento
            i = linea.IndexOf(" ")
            If i = -1 Then
                i = linea.Length ' - 1
                sb.Append(colorearElemento(linea.Substring(0, i)))
                i -= 1
                'Stop
            Else
                sb.Append(colorearElemento(linea.Substring(0, i)))
                sb.Append(" ")
            End If
            linea = linea.Substring(i + 1)
        End If
        While String.IsNullOrEmpty(linea) = False
            i = linea.IndexOf("=")
            If i > -1 Then
                sb.Append(colorearAtributo(linea.Substring(0, i)))
                sb.Append(colorearDelimitador("="))
                linea = linea.Substring(i + 1)
                'i = linea.IndexOfAny((ChrW(34) & "'").ToCharArray)
            Else
                ' No es un atributo,                            (04/Feb/07)
                ' pueden ser instrucciones al estilo de DOCTYPE html PUBLIC
                i = linea.IndexOf(" ")
                If i > -1 Then
                    ' Si antes del espacio hay unas comillas, 
                    ' colorearla en azul y seguir
                    k = linea.IndexOf("sñ·")
                    If k > -1 AndAlso k < i Then
                        Dim j As Integer
                        j = linea.IndexOf("·ñs", k)
                        If j = -1 Then
                            j = linea.Length - 1 - "·ñs".Length
                        End If
                        If j > -1 Then
                            sb.Append(colorearTexto(linea.Substring(k, j - k + "·ñs".Length)))
                            linea = linea.Substring(j + "·ñs".Length)
                            Continue While
                        End If
                        Stop 'Continue While
                    End If
                    sb.Append(colorearElemento(linea.Substring(0, i)))
                    sb.Append(" ")
                    linea = linea.Substring(i + 1)
                    Continue While
                End If
            End If
            '
            Dim sDelim, sDelim2 As String
            k = linea.Length
            sDelim = "sñ·"
            'sDelim2 = "·ñs"
            i = linea.IndexOf(sDelim)
            sDelim = "sñç"
            n = linea.IndexOf(sDelim)
            If i = -1 AndAlso n = -1 Then
                ' no hay comillas
                sb.Append(noColorear(linea))
                Exit While
            End If
            If i = -1 Then i = k + 1
            If n = -1 Then n = k + 1
            '
            If i < k Then
                sDelim = "sñ·"
                sDelim2 = "·ñs"
            Else
                sDelim = "sñç"
                sDelim2 = "çñs"
            End If
            i = linea.IndexOf(sDelim)
            '
            'i = linea.IndexOfAny((ChrW(34) & "'").ToCharArray)
            If i > -1 Then
                Dim j As Integer
                j = linea.IndexOf(sDelim2, i + sDelim.Length)
                If j = -1 Then
                    ' Puede que haya signos < y > entre comillas...
                    ' Esto no debe darse
                    j = linea.Length - 1
                    Stop
                End If
                sb.Append(colorearTexto(linea.Substring(i, j - i + sDelim2.Length)))
                linea = linea.Substring(j + sDelim2.Length)
            Else
                sb.Append(noColorear(linea))
                Exit While
            End If

        End While
        ' NO añadir el elemento final
        '
        Return sb.ToString
    End Function
End Class

'End Namespace
