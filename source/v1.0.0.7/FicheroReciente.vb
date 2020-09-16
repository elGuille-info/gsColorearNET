'------------------------------------------------------------------------------
' Clase definida en la biblioteca para .NET Standard 2.0            (10/Sep/20)
' Basada en gsColorear y gsColorearCore
'
' Clase para la información de los ficheros recientes               (12/Mar/06)
' Se guardará el nombre, cadena de buscar, reemplazar y nombre de guardar como
' Añadido en la revisión 1.0.0.40852
'
' ©Guillermo 'guille' Som, 2006m 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

''' <summary>
''' Clase para usar con la lista de ficheros recientes.
''' Pensado para usar con mi editor de textos gsEditor.
''' </summary>
Public Class FicheroReciente
    ''' <summary>
    ''' Nombre del fichero
    ''' </summary>
    Public Nombre As String
    ''' <summary>
    ''' Cadena usada en la última búsqueda
    ''' </summary>
    Public Buscar As String
    ''' <summary>
    ''' Cadena usada en Buscar cuando se usa Buscar y Reemplazar
    ''' </summary>
    Public Reemp_Buscar As String
    ''' <summary>
    ''' Cadena usada en Reemplazar
    ''' </summary>
    ''' <remarks></remarks>
    Public Reemp_Poner As String

    Private m_Guardar_Como As String
    ''' <summary>
    ''' El nombre usado a la hora de guardar como.
    ''' </summary>
    ''' <value>El valor a asignar</value>
    ''' <returns>
    ''' Devuelve el nombre anteriormente asignado o
    ''' el nombre del fichero si aún no se ha guardado como.
    ''' </returns>
    Public Property Guardar_Como() As String
        Get
            If String.IsNullOrEmpty(m_Guardar_Como) = True _
                    AndAlso String.IsNullOrEmpty(Nombre) = False Then
                m_Guardar_Como = System.IO.Path.GetFileName(Nombre)
            End If
            Return m_Guardar_Como
        End Get
        Set(value As String)
            ' Guardar solo el nombre del fichero
            ' No quitarle el path!!!                                (23/Ago/06)
            m_Guardar_Como = value ' System.IO.Path.GetFileName(value)
        End Set
    End Property

    Private Shared ficheros As New Dictionary(Of String, FicheroReciente)

    ''' <summary>
    ''' Método compartido que devuelve un FicheroReciente
    ''' con el nombre indicado o uno nuevo si no existe.
    ''' </summary>
    ''' <param name="nombre">
    ''' El nombre del fichero
    ''' </param>
    ''' <returns>
    ''' Una referencia al objeto asociado a ese nombre.
    ''' </returns>
    Public Shared Function Fichero(nombre As String) As FicheroReciente
        If ficheros.ContainsKey(nombre.ToLower()) = False Then
            ' Añadirlo
            Dim ficRec As New FicheroReciente
            ficRec.Nombre = nombre
            ficheros.Add(nombre.ToLower(), ficRec)
            Return ficRec
        Else
            Return ficheros(nombre.ToLower())
        End If
    End Function
End Class
