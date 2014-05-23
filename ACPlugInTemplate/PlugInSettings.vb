Public Class PlugInSettings
    Implements ILSPlugInSettings

    Public Sub gExecuteFunction(nFunction As Integer) Implements ILSPlugInSettings.gExecuteFunction

        Try
            Select Case nFunction
                Case 1
                    goApplication.ExecuteMenuEintrag("MenuSchlüssel:=MENUXYZ12345", True)

                Case 2
                    MsgBox("Test-Funktion Nr. 2")

            End Select

        Catch ex As Exception
            gDebugException(ex, Reflection.MethodBase.GetCurrentMethod)

        End Try

    End Sub

    Public Function goPlugInSettings() As clsPlugInSettings Implements ILSPlugInSettings.goPlugInSettings

        Try
            Dim oSettings = New clsPlugInSettings

            oSettings.Add(1, "Caption", "Tooltip", GetImage("xyz.png"))
            oSettings.Add(2, "Caption2", "Tooltip2", GetImage("xyz2.png"))
            '...

            Return oSettings

        Catch ex As Exception
            gDebugException(ex, Reflection.MethodBase.GetCurrentMethod)
            Return Nothing

        End Try

    End Function

    Public Shared Function GetImage(sImage As String) As Drawing.Image

        Try
            Return Drawing.Image.FromStream(Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(sImage))

        Catch ex As Exception
            gDebugException(ex, Reflection.MethodBase.GetCurrentMethod)
            Return Nothing
        End Try

    End Function

End Class
