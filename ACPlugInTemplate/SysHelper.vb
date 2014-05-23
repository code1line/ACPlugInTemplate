Public Class SysHelper

#Region "System/Intern"

    Public Enum VersionType
        Version
        FileVersion
        Major
        Minor
        Build
        Revision
    End Enum

    Public Shared Function GetTitle(Optional assembly As Reflection.Assembly = Nothing) As String

        If assembly Is Nothing Then assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim assemblyAttribute As Reflection.AssemblyProductAttribute
        assemblyAttribute = CType(assembly.GetCustomAttributes(GetType(Reflection.AssemblyProductAttribute), False)(0), Reflection.AssemblyProductAttribute)
        Return assemblyAttribute.Product.ToString

    End Function

    Public Shared Function GetVersion(Optional assembly As Reflection.Assembly = Nothing,
                                      Optional type As VersionType = VersionType.Version) As String

        If assembly Is Nothing Then assembly = Reflection.Assembly.GetExecutingAssembly()

        Select Case type
            Case VersionType.Version
                Return assembly.GetName.Version.ToString
            Case VersionType.FileVersion
                Dim fi = FileVersionInfo.GetVersionInfo(assembly.Location)
                Return fi.ProductVersion
            Case VersionType.Major
                Return assembly.GetName.Version.Major.ToString
            Case VersionType.Minor
                Return assembly.GetName.Version.Minor.ToString
            Case VersionType.Build
                Return assembly.GetName.Version.Build.ToString
            Case VersionType.Revision
                Return assembly.GetName.Version.Revision.ToString
            Case Else
                Return String.Empty
        End Select

    End Function

#End Region

End Class
