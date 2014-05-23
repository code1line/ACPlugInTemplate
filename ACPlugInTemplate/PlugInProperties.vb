Public Class PlugInProperties
    Implements ILSPlugInProperties

    Public Function goPluginProperties(nFunktion As Integer, sExpression As String) As clsPlugInProperties Implements ILSPlugInProperties.goPluginProperties

        Dim oProperties = New clsPlugInProperties

        Select Case nFunktion

            Case 1, 2  ' Test
                oProperties.Add("Name", "Angezeigter Name", "Tooltip", "Defaultwert", "Gruppe")
                oProperties.AddBoolean("Name", "Angezeigter Name", "Tooltip", "Defaultwert", "Gruppe")

            Case Else ' Nicht definiert
                oProperties = Nothing

        End Select

        Return oProperties

    End Function

    Public Shared Function goGetProperties(ByVal nFunktion As Integer, ByVal sArgs As String) As clsPlugInProperties

        Dim oProperties = New PlugInProperties
        Dim oParameter As clsPlugInProperties

        oParameter = oProperties.goPluginProperties(nFunktion, sArgs)
        oParameter.SetExpression(sArgs)

        Return oParameter

    End Function

End Class

