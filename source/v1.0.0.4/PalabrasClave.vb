'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' Clase para contener las palabras clave a usar por gsEditor        (17/Nov/05)
'
' A�ado/quito palabras clave/tipos del fichero de sql (.40855)      (31/Mar/06)
'
' �Guillermo 'guille' Som, 2005-2006
'------------------------------------------------------------------------------
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

'Namespace elGuille.Util.Developer

''' <summary>
''' Enumeraci�n para los distintos lenguajes
''' </summary>
''' <remarks>
''' Se puede usar dotNet para todas las instrucciones de todos los lenguajes
''' Revisi�n 0.40789 del 04/Dic/2005, a�ado las instrucciones de CPP y Pascal.
''' Revisi�n 0.40800 del 14/Dic/2005, a�ado las instrucciones de IL (ILAsm).
''' Revisi�n 1.0.1.0 del 28/Ago/2006, colorear los ficheros XML (tratamiento especial)
''' </remarks>
<Flags()>
Public Enum Lenguajes As Integer
    ''' <summary>
    ''' Para indicar que no se coloreen las instrucciones
    ''' </summary>
    Ninguno = 0
    ''' <summary>
    ''' Visual Basic .NET
    ''' </summary>
    VB = 1
    ''' <summary>
    ''' C#
    ''' </summary>
    CS = 2
    ''' <summary>
    ''' F#
    ''' </summary>
    FSharp = 4
    ''' <summary>
    ''' Visual Basic, C# y F#
    ''' </summary>
    dotNet = VB + CS + FSharp
    ''' <summary>
    ''' C/C++
    ''' </summary>
    CPP = 8
    ''' <summary>
    ''' Java, J#
    ''' </summary>
    Java = 16
    ''' <summary>
    ''' Pascal, Delphi, Freya
    ''' </summary>
    Pascal = 32
    ''' <summary>
    ''' SQL Server
    ''' </summary>
    SQL = 64
    ''' <summary>
    ''' Visual Basic 6.0 o anterior
    ''' </summary>
    VB6 = 128
    ''' <summary>
    ''' MSIL (IL ASM)
    ''' </summary>
    IL = 256
    ''' <summary>
    ''' Para colorear los ficheros con formato XML
    ''' </summary>
    XML = 512
End Enum

''' <summary>
''' Clase para almacenar los tipos de comentarios m�ltiples
''' </summary>
Public Class Comentarios
    ''' <summary>
    ''' Los caracteres que indican el inicio del comentario m�ltiple
    ''' </summary>
    Public Inicio As String
    ''' <summary>
    ''' Los caracteres que indican el final del comentario m�ltiple
    ''' </summary>
    Public Final As String
    ''' <summary>
    ''' Propiedad predeterminada (indizador) que devuelve
    ''' el comentario de inicio o final dependiendo del valor
    ''' indicado en el par�metro.
    ''' 0 para la cadena de inicio.
    ''' 1 para la cadena final.
    ''' </summary>
    ''' <param name="index">El valor a recuperar
    ''' 0 para la cadena de inicio.
    ''' 1 para la cadena final.
    ''' </param>
    ''' <value>El valor a asignar</value>
    ''' <returns>Una cadena con el valor</returns>
    ''' <remarks></remarks>
    Default Public Property Item(ByVal index As Integer) As String
        Get
            If index = 0 Then
                Return Inicio
            Else
                Return Final
            End If
        End Get
        Set(ByVal value As String)
            If index = 0 Then
                Inicio = value
            Else
                Final = value
            End If
        End Set
    End Property
    ''' <summary>
    ''' Constructor en el que indicamos los valores
    ''' a asignar.
    ''' </summary>
    ''' <param name="inicio">
    ''' El texto de la cadena de inicio del comentario m�ltiple.
    ''' Ver <seealso cref="Inicio">Inicio</seealso>
    ''' </param>
    ''' <param name="final">
    ''' El texto de la cadena final del comentario m�ltiple.
    ''' Ver <seealso cref="Inicio">Final</seealso>
    ''' </param>
    Sub New(ByVal inicio As String, ByVal final As String)
        Me.Inicio = inicio
        Me.Final = final
    End Sub
End Class

''' <summary>
''' Colecci�n de palabras clave
''' </summary>
''' <remarks>
''' Se mantiene una colecci�n con todas las palabras clave.
''' Tambi�n se mantiene una para cada lenguaje especificado.
''' La instrucci�n se debe pasar en el mismo estado que la original.
''' 
''' Todos los miembros de la colecci�n se pueden usar de dos formas:
''' Indicando el lenguaje o sin indicarlo.
''' </remarks>
Public Class PalabrasClave
    Private palabras As New System.Collections.Generic.SortedDictionary(Of Lenguajes, SortedDictionary(Of String, PalabraClave))
    Private Shared ficheros As New System.Collections.Generic.Dictionary(Of Lenguajes, String)
    Private Shared extensiones As New System.Collections.Generic.Dictionary(Of Lenguajes, String)
    Private Shared caseSensibles As New System.Collections.Generic.Dictionary(Of Lenguajes, Boolean)
    '
    ' Para los comentarios
    Private Shared remSimple1 As New System.Collections.Generic.Dictionary(Of Lenguajes, String)
    'Private Shared remSimple2 As New System.Collections.Generic.Dictionary(Of Lenguajes, String)
    Private Shared remMultip1 As New System.Collections.Generic.Dictionary(Of Lenguajes, Comentarios)
    Private Shared remMultip2 As New System.Collections.Generic.Dictionary(Of Lenguajes, Comentarios)

    '----------------------------------------------------------------------
    ' M�todos y propiedades compartidas
    '----------------------------------------------------------------------

    ' M�todos separados para poder restablecer                  (18/Dic/05) 0.40842
    Public Shared Sub RestablecerValores()
        AsignarExtensiones()
        AsignarCaseSensitive()
        AsignarComentarios()
    End Sub

    Public Shared Sub AsignarExtensiones()
        ' A�adir las extensiones de los ficheros
        Extension(Lenguajes.CS) = "*.cs"
        Extension(Lenguajes.VB) = "*.vb"
        Extension(Lenguajes.VB6) = "*.bas; *.cls; *.frm; *.ctl"
        Extension(Lenguajes.dotNet) = "" ' "*.aspx; *.asmx; *.ascx; *.master"
        ' Instrucciones de Java aportadas por Emilio P. Egido (miliuco)
        Extension(Lenguajes.Java) = "*.java"
        Extension(Lenguajes.SQL) = "*.sql"
        Extension(Lenguajes.FSharp) = "*.fs; *.fsi; *.ml"
        Extension(Lenguajes.Pascal) = "*.pas; *.cp; *.pp; *.dpr"
        Extension(Lenguajes.CPP) = "*.c; *.cpp; *.cc; *.h; *.hpp"
        Extension(Lenguajes.IL) = "*.il"
        ' Para colorear XML (y HTML)                            (28/Ago/06)
        ' Los proyectos de VS2005 tambi�n son XML               (17/Nov/06)
        Extension(Lenguajes.XML) = "*.xml; *.xaml; *.xsl; *.rss;" &
                                       " *.cfg; *.config; *.manifest;" &
                                       " *.snippet; *.htm; *.html; *.asp;" &
                                       " *.aspx; *.asmx; *.ascx; *.master;" &
                                       " *.vbproj; *.csproj; *.user; *.xbap"
    End Sub

    Public Shared Sub AsignarCaseSensitive()
        CaseSensitive(Lenguajes.CPP) = True
        CaseSensitive(Lenguajes.CS) = True
        CaseSensitive(Lenguajes.dotNet) = False
        CaseSensitive(Lenguajes.FSharp) = True
        CaseSensitive(Lenguajes.IL) = True
        CaseSensitive(Lenguajes.Java) = True
        CaseSensitive(Lenguajes.Pascal) = True
        CaseSensitive(Lenguajes.SQL) = False
        CaseSensitive(Lenguajes.VB) = False
        CaseSensitive(Lenguajes.VB6) = False
        CaseSensitive(Lenguajes.XML) = False
    End Sub

    Public Shared Sub AsignarComentarios()
        ComentarioSimple1(Lenguajes.CPP) = "//"
        ComentarioSimple1(Lenguajes.CS) = "//"
        ComentarioSimple1(Lenguajes.dotNet) = "//"
        ComentarioSimple1(Lenguajes.FSharp) = "//"
        ComentarioSimple1(Lenguajes.IL) = "//"
        ComentarioSimple1(Lenguajes.Java) = "//"
        ComentarioSimple1(Lenguajes.Pascal) = "//"
        ComentarioSimple1(Lenguajes.SQL) = "--"
        ComentarioSimple1(Lenguajes.VB) = "'"
        ComentarioSimple1(Lenguajes.VB6) = "'"
        ComentarioSimple1(Lenguajes.XML) = ""
        '
        CommentMult1(Lenguajes.CPP, 0) = "/*"
        CommentMult1(Lenguajes.CPP, 1) = "*/"
        CommentMult1(Lenguajes.CS, 0) = "/*"
        CommentMult1(Lenguajes.CS, 1) = "*/"
        CommentMult1(Lenguajes.dotNet, 0) = "/*"
        CommentMult1(Lenguajes.dotNet, 1) = "*/"
        CommentMult1(Lenguajes.FSharp, 0) = "(*"
        CommentMult1(Lenguajes.FSharp, 1) = "*)"
        CommentMult2(Lenguajes.FSharp, 0) = "/*"
        CommentMult2(Lenguajes.FSharp, 1) = "*/"
        CommentMult1(Lenguajes.Java, 0) = "/*"
        CommentMult1(Lenguajes.Java, 0) = "*/"
        CommentMult1(Lenguajes.Pascal, 0) = "(*"
        CommentMult1(Lenguajes.Pascal, 1) = "*)"
        CommentMult2(Lenguajes.Pascal, 0) = "{"
        CommentMult2(Lenguajes.Pascal, 1) = "}"
        CommentMult1(Lenguajes.SQL, 0) = "/*"
        CommentMult1(Lenguajes.SQL, 1) = "*/"
        CommentMult1(Lenguajes.XML, 0) = "<!--"
        CommentMult1(Lenguajes.XML, 1) = "-->"
    End Sub

    Shared Sub New()
        RestablecerValores()
    End Sub

    '----------------------------------------------------------------------
    ' Para los comentarios usados por cada lenguaje
    '----------------------------------------------------------------------
    Public Const RemMIni As Byte = 0
    Public Const RemMFin As Byte = 1

    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y los comentarios m�ltiples.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve los comentarios de l�nea simple</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, String)
    ''' ....
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared ReadOnly Property CommentsMult1() As Dictionary(Of Lenguajes, Comentarios)
        Get
            Return remMultip1
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar el valor.</param>
    ''' <value>
    ''' </value>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property CommentMult1(ByVal lenguaje As Lenguajes, ByVal index As Integer) As String
        Get
            If remMultip1.ContainsKey(lenguaje) Then
                Return remMultip1(lenguaje)(index)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If remMultip1.ContainsKey(lenguaje) Then
                remMultip1(lenguaje)(index) = value
            Else
                If index = RemMIni Then
                    remMultip1.Add(lenguaje, New Comentarios(value, ""))
                Else
                    remMultip1.Add(lenguaje, New Comentarios("", value))
                End If
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' </param>
    ''' <value>
    ''' </value>
    ''' <returns>Las extensiones de ficheros o una cadena vac�a.</returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property CommentMult1(ByVal lenguaje As String, ByVal index As Integer) As String
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return CommentMult1(le, index)
        End Get
        Set(ByVal value As String)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            CommentMult1(le, index) = value
        End Set
    End Property

    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y los comentarios m�ltiples.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve los comentarios de l�nea simple</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, String)
    ''' ....
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared ReadOnly Property CommentsMult2() As Dictionary(Of Lenguajes, Comentarios)
        Get
            Return remMultip2
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar el valor.</param>
    ''' <value>
    ''' </value>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property CommentMult2(ByVal lenguaje As Lenguajes, ByVal index As Integer) As String
        Get
            If remMultip2.ContainsKey(lenguaje) Then
                Return remMultip2(lenguaje)(index)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If remMultip2.ContainsKey(lenguaje) Then
                remMultip2(lenguaje)(index) = value
            Else
                If index = RemMIni Then
                    remMultip2.Add(lenguaje, New Comentarios(value, ""))
                Else
                    remMultip2.Add(lenguaje, New Comentarios("", value))
                End If
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' </param>
    ''' <value>
    ''' </value>
    ''' <returns>Las extensiones de ficheros o una cadena vac�a.</returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property CommentMult2(ByVal lenguaje As String, ByVal index As Integer) As String
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return CommentMult2(le, index)
        End Get
        Set(ByVal value As String)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            CommentMult2(le, index) = value
        End Set
    End Property


    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y los comentarios simples.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve los comentarios de l�nea simple</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, String)
    ''' ....
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared ReadOnly Property ComentariosSimples() As Dictionary(Of Lenguajes, String)
        Get
            Return remSimple1
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar el valor.</param>
    ''' <value>
    ''' </value>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property ComentarioSimple1(ByVal lenguaje As Lenguajes) As String
        Get
            If remSimple1.ContainsKey(lenguaje) Then
                Return remSimple1(lenguaje)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If remSimple1.ContainsKey(lenguaje) Then
                remSimple1(lenguaje) = value
            Else
                remSimple1.Add(lenguaje, value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna los caracteres usados para el comentario simple.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' </param>
    ''' <value>
    ''' </value>
    ''' <returns>Las extensiones de ficheros o una cadena vac�a.</returns>
    ''' <remarks>
    ''' </remarks>
    Public Shared Property ComentariosSimple(ByVal lenguaje As String) As String
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return ComentarioSimple1(le)
        End Get
        Set(ByVal value As String)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            ComentarioSimple1(le) = value
        End Set
    End Property

    '----------------------------------------------------------------------
    ' Para indicar si el lenguaje es sensible a may�sculas y min�sculas
    '----------------------------------------------------------------------
    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y si son case sensitive.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve si el lenguaje es sensible a may�sculas / min�sculas</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, Boolean)
    ''' indicando si el lenguajes es case sensitive.
    ''' </returns>
    ''' <remarks>
    ''' Esta colecci�n se usa para saber si los lenguajes son case sensitive (sensibles a may�sculas / min�sculas).
    ''' </remarks>
    Public Shared ReadOnly Property CaseSensitives() As Dictionary(Of Lenguajes, Boolean)
        Get
            Return caseSensibles
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna si el lenguajes es case sensitive.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar el valor de case sensitive.</param>
    ''' <value>
    ''' Usar esta propiedad compartida para indicar si el lenguaje es sensible a may�sculas/min�sculas.
    ''' </value>
    ''' <returns>Un valor verdadero (True) si es sensible a may�sculas / min�sculas.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar si el lenguaje es case sensitive.
    ''' </remarks>
    Public Shared Property CaseSensitive(ByVal lenguaje As Lenguajes) As Boolean
        Get
            If caseSensibles.ContainsKey(lenguaje) Then
                Return caseSensibles(lenguaje)
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            If caseSensibles.ContainsKey(lenguaje) Then
                caseSensibles(lenguaje) = value
            Else
                caseSensibles.Add(lenguaje, value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna si el lenguajes es case sensitive.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' al que queremos asignar o del que queremos recuperar las extensiones.
    ''' </param>
    ''' <value>
    ''' Usar esta propiedad compartida para indicar si el lenguaje es sensible a may�sculas/min�sculas.
    ''' </value>
    ''' <returns>Un valor verdadero (True) si es sensible a may�sculas / min�sculas.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar si el lenguaje es case sensitive.
    ''' </remarks>
    Public Shared Property CaseSensitive(ByVal lenguaje As String) As Boolean
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return CaseSensitive(le)
        End Get
        Set(ByVal value As Boolean)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            CaseSensitive(le) = value
        End Set
    End Property


    '----------------------------------------------------------------------
    ' Para las extensiones asociadas con cada lenguaje
    '----------------------------------------------------------------------
    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y extensiones asociadas.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve las extensiones asociadas a cada lenguaje</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, String)
    ''' con las extensiones asociadas a cada lenguaje.
    ''' </returns>
    ''' <remarks>
    ''' Esta colecci�n s�lo se usa como una forma de asociar extensiones de ficheros 
    ''' con lenguajes, pero no tiene ninguna relaci�n con las colecciones de instrucciones
    ''' que esta clase pueda contener.
    ''' </remarks>
    Public Shared ReadOnly Property Extensions() As Dictionary(Of Lenguajes, String)
        Get
            Return extensiones
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna las extensiones a un lenguaje.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar o del que queremos recuperar las extensiones.</param>
    ''' <value>
    ''' Usar esta propiedad compartida para almacenar las extensiones de ficheros relacionados con los lenguajes
    ''' soportados por esta clase.
    ''' </value>
    ''' <returns>Las extensiones o una cadena vac�a.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar las extensiones
    ''' de los ficheros asociados con un lenguaje, pero no tiene ninguna
    ''' relaci�n con los idiomas utilizados para almacenar las instrucciones.
    ''' </remarks>
    Public Shared Property Extension(ByVal lenguaje As Lenguajes) As String
        Get
            If extensiones.ContainsKey(lenguaje) Then
                Return extensiones(lenguaje)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If extensiones.ContainsKey(lenguaje) Then
                extensiones(lenguaje) = value
            Else
                extensiones.Add(lenguaje, value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna las extensiones a un lenguaje.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' al que queremos asignar o del que queremos recuperar las extensiones.
    ''' </param>
    ''' <value>
    ''' Usar esta propiedad compartida para almacenar las extensiones de ficheros relacionados con los lenguajes
    ''' soportados por esta clase. El lenguaje se indica como cadena.
    ''' </value>
    ''' <returns>Las extensiones de ficheros o una cadena vac�a.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar las extensiones
    ''' de los ficheros asociados con un lenguaje, pero no tiene ninguna
    ''' relaci�n con los idiomas utilizados para almacenar las instrucciones.
    ''' </remarks>
    Public Shared Property Extension(ByVal lenguaje As String) As String
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return Extension(le)
        End Get
        Set(ByVal value As String)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Extension(le) = value
        End Set
    End Property


    '----------------------------------------------------------------------
    ' Para los nombres de ficheros asociados con las palabras clave
    '----------------------------------------------------------------------

    ''' <summary>
    ''' Devuelve una colecci�n con los lenguajes y ficheros asociados.
    ''' </summary>
    ''' <value>Propiedad compartida de solo lectura que devuelve los ficheros y lenguajes asociados.</value>
    ''' <returns>
    ''' Una colecci�n del tipo Dictionary(Of Lenguajes, String)
    ''' con los ficheros y lenguajes que hemos asociado.
    ''' </returns>
    ''' <remarks>
    ''' Esta colecci�n s�lo se usa como una forma de asociar nombres de ficheros 
    ''' con lenguajes, pero no tiene ninguna relaci�n con las colecciones de instrucciones
    ''' que esta clase pueda contener.
    ''' </remarks>
    Public Shared ReadOnly Property Filenames() As Dictionary(Of Lenguajes, String)
        Get
            Return ficheros
        End Get
    End Property

    ''' <summary>
    ''' Devuelve o asigna un fichero a un lenguaje.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje al que queremos asignar o del que queremos recuperar el fichero.</param>
    ''' <value>
    ''' Usar esta propiedad compartida para almacenar unos nombres de ficheros relacionados con los lenguajes
    ''' soportados por esta clase.
    ''' </value>
    ''' <returns>El nombre del fichero o una cadena vac�a.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar los ficheros
    ''' de palabras que queramos asociar con un lenguaje, pero no tiene ninguna
    ''' relaci�n con los idiomas utilizados para almacenar las instrucciones.
    ''' Debido a que C# no soporta las propiedades con par�metros,
    ''' existen dos m�todos para asignar y recuperar los ficheros:
    ''' SetFilename y GetFilename.
    ''' </remarks>
    Public Shared Property Filename(ByVal lenguaje As Lenguajes) As String
        Get
            If ficheros.ContainsKey(lenguaje) Then
                Return ficheros(lenguaje)
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If ficheros.ContainsKey(lenguaje) Then
                ficheros(lenguaje) = value
            Else
                ficheros.Add(lenguaje, value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Devuelve o asigna un fichero a un lenguaje.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' El lenguaje en formato texto 
    ''' al que queremos asignar o del que queremos recuperar el fichero.
    ''' </param>
    ''' <value>
    ''' Usar esta propiedad compartida para almacenar unos nombres de ficheros relacionados con los lenguajes
    ''' soportados por esta clase. El lenguaje se indica como cadena.
    ''' </value>
    ''' <returns>El nombre del fichero o una cadena vac�a.</returns>
    ''' <remarks>
    ''' Esta propiedad solo se usar� como repositorio para almacenar los ficheros
    ''' de palabras que queramos asociar con un lenguaje, pero no tiene ninguna
    ''' relaci�n con los idiomas utilizados para almacenar las instrucciones.
    ''' Debido a que C# no soporta las propiedades con par�metros,
    ''' existen dos m�todos para asignar y recuperar los ficheros:
    ''' SetFilename y GetFilename.
    ''' </remarks>
    Public Shared Property Filename(ByVal lenguaje As String) As String
        Get
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Return Filename(le)
        End Get
        Set(ByVal value As String)
            Dim le As Lenguajes = CType(System.Enum.Parse(GetType(Lenguajes), lenguaje), Lenguajes)
            Filename(le) = value
        End Set
    End Property

    ''' <summary>
    ''' A�ade un nombre de fichero a la colecci�n de nombres de ficheros.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje con el que se asociar� el nombre del fichero.</param>
    ''' <param name="filename">Nombre del fichero a asociar con el lenguaje indicado.</param>
    ''' <remarks>
    ''' Este m�todo compartido es una alternativa a la propiedad Filenames, ya que C# no soporta
    ''' propiedades con par�metros, y as� se podr� usar este m�todo tanto en VB como en C#.
    ''' </remarks>
    Public Shared Sub SetFilename(ByVal lenguaje As Lenguajes, ByVal filename As String)
        If ficheros.ContainsKey(lenguaje) Then
            ficheros(lenguaje) = filename
        Else
            ficheros.Add(lenguaje, filename)
        End If
    End Sub

    ''' <summary>
    ''' Recupera el nombre del fichero asociado con el lenguaje indicado.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje con el que est� asociado el fichero.</param>
    ''' <returns>
    ''' Devuelve el nombre del fichero asociado con el lenguaje
    ''' o una cadena vac�a si no hay ninguno asociado.
    ''' </returns>
    ''' <remarks>
    ''' Este m�todo compartido es una alternativa a la propiedad Filenames, ya que C# no soporta
    ''' propiedades con par�metros, y as� se podr� usar este m�todo tanto en VB como en C#.
    ''' </remarks>
    Public Shared Function GetFilename(ByVal lenguaje As Lenguajes) As String
        Return Filename(lenguaje)
    End Function

    '----------------------------------------------------------------------
    ' M�todos y propiedades de instancia
    '----------------------------------------------------------------------

    '''' <summary>
    '''' M�todo para a�adir nuevas instrucciones a la colecci�n
    '''' de palabras clave a la colecci�n general.
    '''' </summary>
    '''' <param name="instruccion">
    '''' Instrucci�n escrita en el mismo estado de may�sculas/min�sculas del lenguaje.
    '''' </param>
    '''' <remarks>
    '''' Las instrucciones se guardar�n como clave en min�sculas,
    '''' de forma que se puedan buscar sin importar c�mo se indiquen.
    '''' </remarks>
    'Public Sub Add(ByVal instruccion As String)
    '    Dim palabra As New PalabraClave
    '    palabra.Instruccion = instruccion
    '    Me.Add(palabra)
    'End Sub
    '''' <summary>
    '''' M�todo para a�adir nuevas instrucciones a la colecci�n
    '''' de palabras clave a la colecci�n general.
    '''' </summary>
    '''' <param name="palabra">
    '''' Objeto con la instrucci�n a a�adir a la colecci�n.
    '''' </param>
    '''' <remarks>
    '''' Las instrucciones se guardar�n como clave en min�sculas,
    '''' de forma que se puedan buscar sin importar c�mo se indiquen.
    '''' </remarks>
    'Public Sub Add(ByVal palabra As PalabraClave)
    '    If todas.ContainsKey(palabra.Instruccion) = False Then
    '        todas.Add(palabra.Instruccion, palabra)
    '    End If
    '    If todasLower.ContainsKey(palabra.InstruccionToLower) = False Then
    '        todasLower.Add(palabra.InstruccionToLower, palabra)
    '    End If
    'End Sub
    ''' <summary>
    ''' M�todo para a�adir nuevas instrucciones a la colecci�n
    ''' de palabras clave a la colecci�n general y a la del lenguaje
    ''' indicado.
    ''' </summary>
    ''' <param name="lenguaje">
    ''' Lenguaje en el que se guardar� la instrucci�n indicada.
    ''' </param>
    ''' <param name="instruccion">
    ''' Instrucci�n escrita en el mismo estado de may�sculas/min�sculas del lenguaje.
    ''' </param>
    ''' <remarks>
    ''' El lenguaje ser� uno de los valores de la enumeraci�n Lenguajes,
    ''' si se indica Lenguajes.Ninguno s�lo se almacenar� en la colecci�n general.
    ''' Las instrucciones se guardar�n como clave en min�sculas,
    ''' de forma que se puedan buscar sin importar c�mo se indiquen.
    ''' </remarks>
    Public Sub Add(ByVal lenguaje As Lenguajes, ByVal instruccion As String)
        Dim palabra As New PalabraClave
        palabra.Instruccion = instruccion
        'If lenguaje = Lenguajes.Ninguno Then
        '    Me.Add(palabra)
        'Else
        '    Me.Add(lenguaje, palabra)
        'End If
        Me.Add(lenguaje, palabra)
    End Sub
    ''' <summary>
    ''' M�todo para a�adir nuevas instrucciones a la colecci�n
    ''' de palabras clave a la colecci�n general y a la del lenguaje
    ''' indicado.
    ''' </summary>
    ''' <param name="lenguaje">Lenguaje en el que se guardar� la instrucci�n indicada.</param>
    ''' <param name="palabra">Objeto con la instrucci�n a a�adir a la colecci�n.</param>
    ''' <remarks>
    ''' El lenguaje ser� uno de los valores de la enumeraci�n Lenguajes,
    ''' si se indica Lenguajes.Ninguno s�lo se almacenar� en la colecci�n general.
    ''' Las instrucciones se guardar�n como clave en min�sculas,
    ''' de forma que se puedan buscar sin importar c�mo se indiquen.
    ''' </remarks>
    Public Sub Add(ByVal lenguaje As Lenguajes, ByVal palabra As PalabraClave)
        ' La colecci�n base tendr� todas las palabras
        ' y distingue entre may�sculas y min�sculas
        'If lenguaje = Lenguajes.Ninguno Then
        '    Me.Add(palabra)
        '    Return
        'End If
        'If todas.ContainsKey(palabra.Instruccion) = False Then
        '    todas.Add(palabra.Instruccion, palabra)
        'End If
        'If todasLower.ContainsKey(palabra.InstruccionToLower) = False Then
        '    todasLower.Add(palabra.InstruccionToLower, palabra)
        'End If

        ' Existir� una colecci�n para cada lenguaje
        If palabras.ContainsKey(lenguaje) = False Then
            Dim col As New SortedDictionary(Of String, PalabraClave)
            palabras.Add(lenguaje, col)
        End If
        ' En la colecci�n de cada lenguaje, la clave se guarda en min�scula
        If palabras(lenguaje).ContainsKey(palabra.InstruccionToLower) = False Then
            palabras(lenguaje).Add(palabra.InstruccionToLower, palabra)
        End If
    End Sub

    '''' <summary>
    '''' Elimina la instrucci�n de la colecci�n general.
    '''' </summary>
    '''' <param name="instruccion">Instrucci�n a eliminar.</param>
    '''' <remarks></remarks>
    'Public Sub Remove(ByVal instruccion As String)
    '    If todas.ContainsKey(instruccion) Then
    '        todas.Remove(instruccion)
    '    End If
    'End Sub
    ''' <summary>
    ''' Elimina la instrucci�n de la colecci�n general y del lenguaje indicado.
    ''' </summary>
    ''' <param name="lenguaje">Lenguaje del que se eliminar� la instrucci�n.</param>
    ''' <param name="instruccion">Instrucci�n a eliminar.</param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal lenguaje As Lenguajes, ByVal instruccion As String)
        'If lenguaje = Lenguajes.Ninguno Then
        '    Remove(instruccion)
        '    Return
        'End If
        'If todas.ContainsKey(instruccion) Then
        '    todas.Remove(instruccion)
        'End If
        If palabras.ContainsKey(lenguaje) Then
            If palabras(lenguaje).ContainsKey(instruccion) Then
                palabras(lenguaje).Remove(instruccion)
            End If
        End If
    End Sub

    '''' <summary>
    '''' Comprueba si la instrucci�n est� en la colecci�n general.
    '''' No se hace distinci�n de may�sculas y min�sculas.
    '''' </summary>
    '''' <param name="instruccion">Instrucci�n a comprobar.</param>
    '''' <returns>True si est� en la colecci�n.</returns>
    '''' <remarks></remarks>
    'Public Function ContainsLower(ByVal instruccion As String) As Boolean
    '    Return todasLower.ContainsKey(instruccion)
    'End Function
    '''' <summary>
    '''' Comprueba si la instrucci�n est� en la colecci�n general.
    '''' Se hace distinci�n entre may�sculas y min�sculas.
    '''' </summary>
    '''' <param name="instruccion">Instrucci�n a comprobar.</param>
    '''' <returns>True si est� en la colecci�n.</returns>
    '''' <remarks></remarks>
    'Public Function Contains(ByVal instruccion As String) As Boolean
    '    Return todas.ContainsKey(instruccion)
    'End Function
    ''' <summary>
    ''' Comprueba si la instrucci�n est� en la colecci�n del idioma indicado.
    ''' </summary>
    ''' <param name="lenguaje">Lenguaje en el que se comprobar� la existencia de la instrucci�n indicada.</param>
    ''' <param name="instruccion">Instrucci�n a comprobar.</param>
    ''' <returns>True si est� en la colecci�n.</returns>
    ''' <remarks>Si el lenguaje no est� creado, devolver� False.</remarks>
    Public Function Contains(ByVal lenguaje As Lenguajes, ByVal instruccion As String) As Boolean
        'If lenguaje = Lenguajes.Ninguno Then
        '    Return Contains(instruccion)
        'End If
        If palabras.ContainsKey(lenguaje) Then
            Return palabras(lenguaje).ContainsKey(instruccion)
        End If
        Return False
    End Function

    '''' <summary>
    '''' Elimina las instrucciones de la colecci�n general.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub Clear()
    '    todas.Clear()
    '    todasLower.Clear()
    'End Sub
    '''' <summary>
    '''' Elimina las instrucciones de la colecci�n general y
    '''' opcionalmente de todas las de los lenguajes.
    '''' </summary>
    '''' <param name="todos">True para eliminar todas las instrucciones de todos los lenguajes.</param>
    '''' <remarks></remarks>
    'Public Sub Clear(ByVal todos As Boolean)
    '    Me.Clear()
    '    If todos Then
    '        For Each le As Lenguajes In palabras.Keys
    '            If palabras.ContainsKey(le) Then
    '                palabras(le).Clear()
    '            End If
    '        Next
    '    End If
    'End Sub

    ''' <summary>
    ''' Elimina todas las instrucciones.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        palabras.Clear()
    End Sub

    ''' <summary>
    ''' Elimina todas las instrucciones de la colecci�n general y
    ''' de la del lenguaje indicado.
    ''' </summary>
    ''' <param name="lenguaje">Lenguaje en el que se comprobar� la existencia de la instrucci�n indicada.</param>
    ''' <remarks>
    ''' Si hay definido m�s de un lenguaje, las instrucciones de los dem�s no se eliminar�n,
    ''' pero si de la colecci�n general, por tanto es preferible eliminar todos los lenguajes,
    ''' salvo que siempre usemos el lenguaje al a�adir, buscar, etc.
    ''' </remarks>
    Public Sub Clear(ByVal lenguaje As Lenguajes)
        'Me.Clear()
        If lenguaje <> Lenguajes.Ninguno Then
            If palabras.ContainsKey(lenguaje) Then
                palabras(lenguaje).Clear()
            End If
        End If
    End Sub

    '''' <summary>
    '''' El n�mero de instrucciones de la colecci�n general.
    '''' </summary>
    '''' <value>Un valor entero con el n�mero de instrucciones de la colecci�n general.</value>
    '''' <returns>Devuelve el n�mero de instrucciones de la colecci�n general.</returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property Count() As Integer
    '    Get
    '        Return todas.Count
    '    End Get
    'End Property
    ''' <summary>
    ''' El n�mero de instrucciones del lenguaje indicado.
    ''' </summary>
    ''' <param name="lenguaje">El lenguaje del que se quiere consultar el n�mero de instrucciones.</param>
    ''' <value>Un valor entero con el n�mero de instrucciones de la colecci�n general.</value>
    ''' <returns>Devuelve el n�mero de instrucciones de la colecci�n del lenguaje indicado.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count(ByVal lenguaje As Lenguajes) As Integer
        Get
            'If lenguaje = Lenguajes.Ninguno Then
            '    Return todas.Count
            'End If
            If palabras.ContainsKey(lenguaje) Then
                Return palabras(lenguaje).Count
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Devuelve el total de instrucciones de todos los lenguajes.
    ''' </summary>
    ''' <returns>Un valor entero con todas las instrucciones de todos los lenguajes.</returns>
    ''' <remarks>
    ''' No se tendr� en cuenta las instrucciones de la colecci�n general,
    ''' s�lo las definidas espec�ficamente en cada lenguaje.
    ''' Si s�lo se han a�adido instrucciones de forma gen�rica o usando el valor Lenguajes.Ninguno,
    ''' este m�todo devolver� cero a pesar de que existan instrucciones en la colecci�n general.
    ''' </remarks>
    Public Function CountAll() As Integer
        Dim i As Integer = 0
        For Each le As Lenguajes In palabras.Keys
            i += palabras(le).Count
        Next
        Return i
    End Function

    'Default Public ReadOnly Property Item(ByVal instruccion As String) As String
    '    Get
    '        If todas.ContainsKey(instruccion) Then
    '            Return todas(instruccion).Instruccion
    '        End If
    '        Return ""
    '    End Get
    'End Property
    Default Public ReadOnly Property Item(ByVal lenguaje As Lenguajes, ByVal instruccion As String) As String
        Get
            'If lenguaje = Lenguajes.Ninguno Then
            '    Return Me.Item(instruccion)
            'End If
            If palabras.ContainsKey(lenguaje) Then
                If palabras(lenguaje).ContainsKey(instruccion) Then
                    Return palabras(lenguaje)(instruccion).Instruccion
                End If
            End If
            Return ""
        End Get
    End Property

    'Public Function PalabraClaveLower(ByVal instruccion As String) As PalabraClave
    '    Dim p As PalabraClave = Nothing
    '    todasLower.TryGetValue(instruccion, p)
    '    Return p

    '    'If todas.ContainsKey(instruccion) Then
    '    '    Return todas(instruccion)
    '    'End If
    '    'Return Nothing
    'End Function
    'Public Function PalabraClave(ByVal instruccion As String) As PalabraClave
    '    Dim p As PalabraClave = Nothing
    '    todas.TryGetValue(instruccion, p)
    '    Return p

    '    'If todas.ContainsKey(instruccion) Then
    '    '    Return todas(instruccion)
    '    'End If
    '    'Return Nothing
    'End Function

    Public Function PalabraClave(ByVal lenguaje As Lenguajes, ByVal instruccion As String) As PalabraClave
        'If lenguaje = Lenguajes.Ninguno Then
        '    Return PalabraClave(instruccion)
        'End If
        Dim p As PalabraClave = Nothing
        If palabras.ContainsKey(lenguaje) Then
            palabras(lenguaje).TryGetValue(instruccion, p)
        End If
        Return p

        'If palabras.ContainsKey(lenguaje) Then
        '    If palabras(lenguaje).ContainsKey(instruccion) Then
        '        Return palabras(lenguaje)(instruccion)
        '    End If
        'End If
        'Return Nothing
    End Function

    'Public Function ToArray() As String()
    '    Dim ar(0 To todas.Count - 1) As String
    '    Dim i As Integer = 0
    '    For Each p As PalabraClave In todas.Values
    '        ar(i) = p.Instruccion
    '        i += 1
    '    Next
    '    Return ar
    'End Function
    Public Function ToArray(ByVal lenguaje As Lenguajes) As String()
        'If lenguaje = Lenguajes.Ninguno Then
        '    Return Me.ToArray()
        'End If
        If palabras.ContainsKey(lenguaje) Then
            Dim ar(0 To palabras(lenguaje).Count - 1) As String
            Dim i As Integer = 0
            For Each p As PalabraClave In palabras(lenguaje).Values
                ar(i) = p.Instruccion
                i += 1
            Next
            Return ar
        End If
        Return Nothing
    End Function

    'Public Sub CargarPalabras(ByVal palabras() As String)
    '    ' Cargar las palabras del fichero indicado
    '    todas.Clear()
    '    For Each s As String In palabras
    '        Me.Add(s)
    '    Next
    'End Sub
    Public Sub CargarPalabras(ByVal lenguaje As Lenguajes, ByVal palabras() As String)
        'If lenguaje = Lenguajes.Ninguno Then
        '    CargarPalabras(palabras)
        '    Return
        'End If

        Me.Clear(lenguaje)
        For Each s As String In palabras
            Me.Add(lenguaje, s)
        Next
    End Sub
    'Public Sub CargarPalabras(ByVal filename As String)
    '    ' Cargar las palabras del fichero indicado
    '    todas.Clear()

    '    Dim sr As System.IO.StreamReader = Nothing
    '    Try
    '        sr = New System.IO.StreamReader(filename, System.Text.Encoding.Default, True)
    '        Dim s As String
    '        While sr.Peek <> -1
    '            s = sr.ReadLine
    '            Me.Add(s)
    '        End While
    '        'sr.Close()
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        If sr IsNot Nothing Then
    '            sr.Close()
    '        End If
    '    End Try
    'End Sub
    Public Sub CargarPalabras(ByVal lenguaje As Lenguajes, ByVal filename As String)
        '' Si el lenguaje indicado es Ninguno, asignar solo la colecci�n normal
        'If lenguaje = Lenguajes.Ninguno Then
        '    CargarPalabras(filename)
        '    Return
        'End If

        Me.Clear(lenguaje)

        ' Cargar las palabras del fichero indicado
        Dim sr As System.IO.StreamReader = Nothing
        Try
            sr = New System.IO.StreamReader(filename, System.Text.Encoding.Default, True)
            Dim s As String
            While sr.Peek <> -1
                s = sr.ReadLine
                Me.Add(lenguaje, s)
            End While
            'sr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            If sr IsNot Nothing Then
                sr.Close()
            End If
        End Try
    End Sub
End Class

''' <summary>
''' Clase base para cada instrucci�n
''' </summary>
''' <remarks>
''' Cada palabra clave se almacenar� usando la forma de may�sculas y min�sculas del lenguaje
''' La propiedad InstruccionToLower devuelve la instrucci�n en min�sculas
''' </remarks>
Public Class PalabraClave
    Implements IComparable

    ' La palabra clave en el estado a mostrar
    Private _Instruccion As String
    ''' <summary>
    ''' La instrucci�n de esta palabra clave
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    Public Property Instruccion() As String
        Get
            Return _Instruccion
        End Get
        Set(ByVal value As String)
            _Instruccion = value
        End Set
    End Property
    ''' <summary>
    ''' Propiedad de solo lectura con la instrucci�n en min�sculas
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property InstruccionToLower() As String
        Get
            Return Instruccion.ToLower()
        End Get
    End Property

    ''' <summary>
    ''' M�todo para permitir la clasificaci�n de cada palabra
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
        Return String.Compare(Me.InstruccionToLower, obj.ToString, True)
    End Function

    ''' <summary>
    ''' Devuelve la palabra en min�sculas
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String
        Return InstruccionToLower
    End Function
End Class

'End Namespace
