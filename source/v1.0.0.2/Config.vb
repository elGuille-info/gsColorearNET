'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' Clase para manejar ficheros de configuración                      (15/Nov/05)
'
'
' Las secciones siempre estarán dentro de <configuration>
' al menos así lo guardará esta clase, aunque permite leer pares key / value.
' Para que se sepa que se lee de configuration,
' en el código se indica explícitamente.
'
' Basada en mi código publicado el 27/Feb/05 en:
' http://www.elguille.info/NET/dotnet/appSettings2.htm
' Pero para usarla de forma independiente de ConfigurationSettings
'
' Revisado para poder guardar automáticamente                       (21/Feb/06)
' Poder leer todas las secciones y las claves de una sección        (21/Feb/06)
'
' Nuevas sobrecargas de GetValue y SetValue para el tipo Double     (05/Sep/20)
' Nuevas sobrecargas de GetValue, SetValue y SetKeyValue            (10/Sep/20)
'   para el tipo System.Windows.Forms.CheckState
' .NET standard 2.0 no tiene la definción de CheckState             (11/Sep/20)
'   comento estas definiciones.
' Nueva sobrecarga de SetKeyValue para el tipo Double               (11/Sep/20)
'
'
' ©Guillermo 'guille' Som, 2005-2006, 2020
'------------------------------------------------------------------------------
Option Explicit On 
Option Strict On

Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic
Imports System

Imports System.Collections
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Xml
Imports System.IO

'Namespace elGuille.Util.Developer

''' <summary>
''' Manejar ficheros de configuración
''' </summary>
Public Class Config

    '----------------------------------------------------------------------
    ' Los campos y métodos privados
    '----------------------------------------------------------------------
    Private mGuardarAlAsignar As Boolean = True
    Private Const configuration As String = "configuration/"
    Private ficConfig As String = ""
    Private configXml As New XmlDocument
    '
    ''' <summary>
    ''' Si se debe guardar automáticamente después de asignar un valor.
    ''' </summary>
    ''' <value>Un valor de tipo Boolean</value>
    ''' <returns>Devuelve el valor actualmente asignado</returns>
    ''' <remarks>
    ''' Si no se guarda automáticamente hay que llamar
    ''' al método Save para que se guarden en el fichero.
    ''' </remarks>
    Public Property GuardarAlAsignar() As Boolean
        Get
            Return mGuardarAlAsignar
        End Get
        Set(ByVal value As Boolean)
            mGuardarAlAsignar = value
        End Set
    End Property
    '
    ''' <summary>
    ''' Recupera el valor de la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <returns>
    ''' El contenido de la clave y sección.
    ''' </returns>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean
    ''' </remarks>
    Public Function GetValue(ByVal seccion As String, ByVal clave As String) As String
        Return GetValue(seccion, clave, "")
    End Function
    ''' <summary>
    ''' Recupera el valor de la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="predeterminado">
    ''' El valor predeterminado a devolver si no existe la clave.
    ''' </param>
    ''' <returns>
    ''' El contenido de la clave y sección.
    ''' El tipo devuelto depende del tipo del prámetro predeterminado.
    ''' </returns>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean
    ''' </remarks>
    Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As String) As String
        Return cfgGetValue(seccion, clave, predeterminado)
    End Function
    ''' <summary>
    ''' Recupera el valor de la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="predeterminado">
    ''' El valor predeterminado a devolver si no existe la clave.
    ''' </param>
    ''' <returns>
    ''' El contenido de la clave y sección.
    ''' El tipo devuelto depende del tipo del prámetro predeterminado.
    ''' </returns>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean
    ''' </remarks>
    Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As Integer) As Integer
        Return CInt(cfgGetValue(seccion, clave, predeterminado.ToString))
    End Function

    ''' <summary>
    ''' Recupera un valor de la calve y sección indicados, el valor devuelto es Double.
    ''' </summary>
    ''' <remarks>05/Sep/2020</remarks>
    Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As Double) As Double
        Return CDbl(cfgGetValue(seccion, clave, predeterminado.ToString))
    End Function

    ''' <summary>
    ''' Recupera el valor de la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="predeterminado">
    ''' El valor predeterminado a devolver si no existe la clave.
    ''' </param>
    ''' <returns>
    ''' El contenido de la clave y sección.
    ''' El tipo devuelto depende del tipo del prámetro predeterminado.
    ''' </returns>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' Cuando es Boolean, se guarda 1 ó 0 según sea True o False.
    ''' </remarks>
    Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As Boolean) As Boolean
        Dim def As String = "0"
        If predeterminado Then def = "1"
        def = cfgGetValue(seccion, clave, def)
        If def = "1" Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' </remarks>
    Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
        cfgSetValue(seccion, clave, valor)
    End Sub
    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' </remarks>
    Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Integer)
        cfgSetValue(seccion, clave, valor.ToString)
    End Sub

    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados usando el valor Double.
    ''' </summary>
    ''' <remarks>05/Sep/20</remarks>
    Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Double)
        cfgSetValue(seccion, clave, valor.ToString)
    End Sub

    ''' <summary>
    ''' Asigna el valor Double en la clave y la sección indicados.
    ''' Usando atributos dentro de la sección.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">El valor a asignar</param>
    ''' <remarks>
    ''' 11/Sep/2020
    ''' </remarks>
    Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Double)
        cfgSetKeyValue(seccion, clave, valor.ToString)
    End Sub

    '''' <summary>
    '''' Devuelve el valor del tipo System.Windows.Forms.CheckState
    '''' Los valores pueden ser: Checked, Unchecked, Indeterminate
    '''' </summary>
    '''' <remarks>10/Sep/2020</remarks>
    'Public Function GetValue(seccion As String, clave As String, predeterminado As System.Windows.Forms.CheckState) As System.Windows.Forms.CheckState
    '    Dim def As String = predeterminado.ToString
    '    def = cfgGetValue(seccion, clave, def)
    '    If def = "Checked" Then
    '        Return CheckState.Checked
    '    ElseIf def = "Unchecked" Then
    '        Return System.Windows.Forms.CheckState.Unchecked
    '    Else
    '        Return System.Windows.Forms.CheckState.Indeterminate
    '    End If
    'End Function

    '''' <summary>
    '''' Asigna el valor en la clave y sección indicados.
    '''' El tipo es System.Windows.Forms.CheckState
    '''' </summary>
    '''' <remarks>10/Sep/2020</remarks>
    'Public Sub SetValue(seccion As String, clave As String, valor As System.Windows.Forms.CheckState)
    '    cfgSetValue(seccion, clave, valor.ToString)
    'End Sub

    '''' <summary>
    '''' Asigna el valor en la clave y la sección indicados.
    '''' Usando atributos dentro de la sección.
    '''' El tipo es System.Windows.Forms.CheckState
    '''' </summary>
    '''' <remarks>10/Sep/2020</remarks>
    'Public Sub SetKeyValue(seccion As String, clave As String, valor As System.Windows.Forms.CheckState)
    '    cfgSetKeyValue(seccion, clave, valor.ToString)
    'End Sub

    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' Cuando es Boolean, se guarda 1 ó 0 según sea True o False.
    ''' </remarks>
    Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Boolean)
        If valor Then
            cfgSetValue(seccion, clave, "1")
        Else
            cfgSetValue(seccion, clave, "0")
        End If
    End Sub

    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' Usando atributos dentro de la sección.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' </remarks>
    Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
        cfgSetKeyValue(seccion, clave, valor)
    End Sub
    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' Usando atributos dentro de la sección.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' </remarks>
    Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Integer)
        cfgSetKeyValue(seccion, clave, valor.ToString)
    End Sub
    ''' <summary>
    ''' Asigna el valor en la clave y la sección indicados.
    ''' Usando atributos dentro de la sección.
    ''' </summary>
    ''' <param name="seccion">La sección</param>
    ''' <param name="clave">La clave dentro de la sección</param>
    ''' <param name="valor">
    ''' El valor a asignar.
    ''' </param>
    ''' <remarks>
    ''' Existen sobrecargas para los tipos String, Integer y Boolean.
    ''' Cuando es Boolean, se guarda 1 ó 0 según sea True o False.
    ''' </remarks>
    Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Boolean)
        If valor Then
            cfgSetKeyValue(seccion, clave, "1")
        Else
            cfgSetKeyValue(seccion, clave, "0")
        End If
    End Sub

    ''' <summary>
    ''' Elimina la sección, en realidad la deja vacía,
    ''' </summary>
    ''' <param name="seccion">La sección a eliminar</param>
    Public Sub RemoveSection(ByVal seccion As String)
        Dim n As XmlNode
        n = configXml.SelectSingleNode(configuration & seccion)
        If Not n Is Nothing Then
            n.RemoveAll()
            If mGuardarAlAsignar Then
                Me.Save()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Devuelve una colección con los nombres de las secciones.
    ''' </summary>
    ''' <returns>
    ''' Devuelve una colección de tipo List(Of String)
    ''' con las secciones del fichero de configuración.
    ''' </returns>
    Public Function Secciones() As List(Of String)
        Dim d As New List(Of String)
        Dim root As XmlNode
        Dim s As String = "configuration"
        root = configXml.SelectSingleNode(s)
        If root IsNot Nothing Then
            For Each n As XmlNode In root.ChildNodes
                d.Add(n.Name)
            Next
        End If
        Return d
    End Function

    ''' <summary>
    ''' Devuelve una colección con las claves de la sección indicada.
    ''' </summary>
    ''' <param name="seccion">
    ''' Sección de la que se quieren las claves
    ''' </param>
    ''' <returns>
    ''' Devuelve una colección de tipo Dictionary(Of String, String)
    ''' con las claves de la sección indicada.
    ''' </returns>
    Public Function Claves(ByVal seccion As String) As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)
        Dim root As XmlNode
        seccion = seccion.Replace(" ", "_")
        root = configXml.SelectSingleNode(configuration & seccion)
        If root IsNot Nothing Then
            For Each n As XmlNode In root.ChildNodes
                If d.ContainsKey(n.Name) = False Then
                    d.Add(n.Name, n.InnerText)
                End If
            Next
        End If
        Return d
    End Function
    '
    ''' <summary>
    ''' Guardar los datos en el fichero de configuración.
    ''' </summary>
    ''' <remarks>
    ''' Si no se llama a este método, no se guardará de forma permanente.
    ''' </remarks>
    Public Sub Save()
        configXml.Save(ficConfig)
    End Sub

    ''' <summary>
    ''' Lee el contenido del fichero de configuración.
    ''' Si no existe, se crea uno nuevo
    ''' </summary>
    Public Sub Read()
        Dim fic As String = ficConfig
        Const revDate As String = "Sun, 27 Aug 2006 18:12:04 GMT"
        If File.Exists(fic) Then
            configXml.Load(fic)
            ' Actualizar los datos de la información de esta clase
            Dim b As Boolean = mGuardarAlAsignar
            mGuardarAlAsignar = False
            Me.SetValue("configXml_Info", "info", "Generado con Config para Visual Basic 2005")
            Me.SetValue("configXml_Info", "revision", revDate)
            Me.SetValue("configXml_Info", "formatoUTF8", "El formato de este fichero debe ser UTF-8")
            mGuardarAlAsignar = b
            Me.Save()
        Else
            ' Crear el XML de configuración con la sección General
            Dim sb As New System.Text.StringBuilder
            sb.Append("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            sb.Append("<configuration>")
            ' Por si es un fichero appSetting
            sb.Append("<configSections>")
            sb.Append("<section name=""General"" type=""System.Configuration.DictionarySectionHandler"" />")
            sb.Append("</configSections>")
            sb.Append("<General>")
            sb.Append("<!-- Los valores irán dentro del elemento indicado por la clave -->")
            sb.Append("<!-- Aunque también se podrán indicar como pares key / value -->")
            sb.AppendFormat("<add key=""Revisión"" value=""{0}"" />", revDate)
            sb.Append("<!-- La clase siempre los añade como un elemento -->")
            sb.Append("<Copyright>©Guillermo 'guille' Som, 2005-2006</Copyright>")
            sb.Append("</General>")
            '
            sb.AppendFormat("<configXml_Info>{0}", vbCrLf)
            sb.AppendFormat("<info>Generado con Config para Visual Basic 2005</info>{0}", vbCrLf)
            sb.AppendFormat("<copyright>©Guillermo 'guille' Som, 2005-2006</copyright>{0}", vbCrLf)
            sb.AppendFormat("<revision>{0}</revision>{1}", revDate, vbCrLf)
            sb.AppendFormat("<formatoUTF8>El formato de este fichero debe ser UTF-8</formatoUTF8>{0}", vbCrLf)
            sb.AppendFormat("</configXml_Info>{0}", vbCrLf)
            '
            sb.Append("</configuration>")
            ' Asignamos la cadena al objeto
            configXml.LoadXml(sb.ToString)
            '
            ' Guardamos el contenido de configXml y creamos el fichero
            configXml.Save(ficConfig)
        End If
    End Sub

    ''' <summary>
    ''' El nombre del fichero de configuración.
    ''' </summary>
    ''' <value>El nuevo nombre a asignar.</value>
    ''' <returns>
    ''' Una cadena con el nombre del fichero de configuración
    ''' </returns>
    ''' <remarks>
    ''' Al asignarlo, NO se lee el contenido del fichero,
    ''' habrá que llamar al método <seealso cref="Read">Read</seealso>
    ''' </remarks>
    Public Property FileName() As String
        Get
            Return ficConfig
        End Get
        Set(ByVal value As String)
            ' Al asignarlo, NO leemos el contenido del fichero
            ficConfig = value
            'LeerFile()
        End Set
    End Property

    'Public Sub New()
    '    ' Asignamos automáticamente el nombre del fichero, y lo leemos
    '    ' Este constructor no deberíamos usarlo si esta clase está en una DLL
    '    ficConfig = System.Reflection.Assembly.GetExecutingAssembly.Location & ".cfg"
    '    Read()
    'End Sub
    ''' <summary>
    ''' Constructor indicando el nombre del fichero a usar.
    ''' </summary>
    ''' <param name="fic">Nombre del fichero a usar</param>
    Public Sub New(ByVal fic As String)
        ficConfig = fic
        ' Por defecto se guarda al asignar los valores
        mGuardarAlAsignar = True
        Read()
    End Sub

    ''' <summary>
    ''' Constructor con el nombre del fichero y si se guarda al asignar
    ''' </summary>
    ''' <param name="fic">Nombre del fichero ausar</param>
    ''' <param name="guardarAlAsignar">
    ''' True si se guarda automáticamente al asignar,
    ''' ver <seealso cref="GuardarAlAsignar">GuardarAlAsignar</seealso>.
    ''' </param>
    Public Sub New(ByVal fic As String, ByVal guardarAlAsignar As Boolean)
        ficConfig = fic
        mGuardarAlAsignar = guardarAlAsignar
        Read()
    End Sub
    '
    '----------------------------------------------------------------------
    ' Los métodos privados
    '----------------------------------------------------------------------
    '
    ' El método interno para guardar los valores
    ' Este método siempre guardará en el formato <seccion><clave>valor</clave></seccion>
    Private Sub cfgSetValue(
                        ByVal seccion As String,
                        ByVal clave As String,
                        ByVal valor As String)
        '
        Dim n As XmlNode
        '
        ' Filtrar los caracteres no válidos
        ' en principio solo comprobamos el espacio
        seccion = seccion.Replace(" ", "_")
        clave = clave.Replace(" ", "_")

        ' Se comrpueba si es un elemento de la sección:
        '   <seccion><clave>valor</clave></seccion>
        n = configXml.SelectSingleNode(configuration & seccion & "/" & clave)
        If Not n Is Nothing Then
            n.InnerText = valor
        Else
            Dim root As XmlNode
            Dim elem As XmlElement
            root = configXml.SelectSingleNode(configuration & seccion)
            If root Is Nothing Then
                ' Si no existe el elemento principal,
                ' lo añadimos a <configuration>
                elem = configXml.CreateElement(seccion)
                configXml.DocumentElement.AppendChild(elem)
                root = configXml.SelectSingleNode(configuration & seccion)
            End If
            If Not root Is Nothing Then
                ' Crear el elemento
                elem = configXml.CreateElement(clave)
                elem.InnerText = valor
                ' Añadirlo al nodo indicado
                root.AppendChild(elem)
            End If
        End If
        '
        If mGuardarAlAsignar Then
            Me.Save()
        End If
    End Sub

    ' Asigna un atributo a una sección
    ' Por ejemplo: <Seccion clave=valor>...</Seccion>
    ' También se usará para el formato de appSettings: <add key=clave value=valor />
    '   Aunque en este caso, debe existir el elemento a asignar.
    Private Sub cfgSetKeyValue(
                        ByVal seccion As String,
                        ByVal clave As String,
                        ByVal valor As String)
        '
        Dim n As XmlNode
        '
        ' Filtrar los caracteres no válidos
        ' en principio solo comprobamos el espacio
        seccion = seccion.Replace(" ", "_")
        clave = clave.Replace(" ", "_")

        n = configXml.SelectSingleNode(configuration & seccion & "/add[@key=""" & clave & """]")
        If Not n Is Nothing Then
            n.Attributes("value").InnerText = valor
        Else
            Dim root As XmlNode
            Dim elem As XmlElement
            root = configXml.SelectSingleNode(configuration & seccion)
            If root Is Nothing Then
                ' Si no existe el elemento principal,
                ' lo añadimos a <configuration>
                elem = configXml.CreateElement(seccion)
                configXml.DocumentElement.AppendChild(elem)
                root = configXml.SelectSingleNode(configuration & seccion)
            End If
            If Not root Is Nothing Then
                Dim a As XmlAttribute = CType(configXml.CreateNode(XmlNodeType.Attribute, clave, Nothing), XmlAttribute)
                a.InnerText = valor
                root.Attributes.Append(a)
            End If
        End If
        '
        If mGuardarAlAsignar Then
            Me.Save()
        End If
    End Sub

    ' Devolver el valor de la clave indicada
    Private Function cfgGetValue(
                        ByVal seccion As String,
                        ByVal clave As String,
                        ByVal valor As String
                        ) As String
        '
        Dim n As XmlNode
        '
        ' Filtrar los caracteres no válidos
        ' en principio solo comprobamos el espacio
        seccion = seccion.Replace(" ", "_")
        clave = clave.Replace(" ", "_")

        ' Primero comprobar si están el formato de appSettings: <add key = clave value = valor />
        n = configXml.SelectSingleNode(configuration & seccion & "/add[@key=""" & clave & """]")
        If Not n Is Nothing Then
            Return n.Attributes("value").InnerText
        End If
        '
        ' Después se comprueba si está en el formato <Seccion clave = valor>
        n = configXml.SelectSingleNode(configuration & seccion)
        If Not n Is Nothing Then
            Dim a As XmlAttribute = n.Attributes(clave)
            If Not a Is Nothing Then
                Return a.InnerText
            End If
        End If
        '
        ' Por último se comprueba si es un elemento de seccion:
        '   <seccion><clave>valor</clave></seccion>
        n = configXml.SelectSingleNode(configuration & seccion & "/" & clave)
        If Not n Is Nothing Then
            Return n.InnerText
        End If
        '
        ' Si no existe, se devuelve el valor predeterminado
        Return valor
    End Function
End Class

'End Namespace
