Imports System.Reflection

Module VersionData
    Public ReadOnly Property Version As String
        Get
            Return ApplicationVersion
        End Get
    End Property

    Public ReadOnly Property ApplicationVersion As String
        Get
            Return Assembly.GetExecutingAssembly().GetName().Version.ToString()
        End Get
    End Property
End Module
