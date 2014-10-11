Imports SAPbouiCOM.Framework
Imports System.IO
Imports SBO.SboAddOnBase

Imports System
Imports System.Xml
Imports System.Xml.XPath
Imports System.Windows.Forms
Imports System.Activator
Imports System.Runtime.Remoting
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine


Public Class DataLogAddon
    Inherits SboAddon
    Private RunApp As Boolean = True
    Public Sub Main()


    End Sub


    Public Sub New(ByVal StartUpPath As String, ByVal AddonName As String, ByRef pbo_RunApplication As Boolean)

        MyBase.New(StartUpPath, AddonName)
        m_Namespace = "DataLog"
        m_AssemblyName = "DataLog"
        TablePrefix = "OWA_INV"
        PermissionPrefix = "OWA_INV"
        MenuXMLFileName = "Menus.xml"
        'm_MenuImageFile = "truck.jpg"
        If IsNothing(m_SboApplication) Then
            pbo_RunApplication = False
            Exit Sub
        Else

            If Not initialise() Then
                pbo_RunApplication = False
                Exit Sub
            End If

        End If
        oApp.Run()
        pbo_RunApplication = True
        'Me.setFilters(oFilters)

    End Sub

    Private _dataSet As DataSet
    Private _Report As ReportDocument
    Private _ReportCaption As String
    Private _DisplayGroupTree As Boolean
    Private _ParamFields As CrystalDecisions.Shared.ParameterFields
    Private _PaperOrientation As Integer
    Private _PaperSize As Integer

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class


End Class