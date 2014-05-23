Public Module SysGlobals

    Private _oInfo As clsInfoDaten

    Public ReadOnly Property goInfo As clsInfoDaten

        Get
            Return _oInfo
        End Get

    End Property

    Public ReadOnly Property goMandant As Sagede.OfficeLine.Engine.Mandant

        Get
            Return DirectCast(goApplication.oSageMandant, Sagede.OfficeLine.Interop62.Mandant).GetRealObject
        End Get

    End Property

    Public ReadOnly Property AppTitle() As String

        Get
            Return SysHelper.GetTitle
        End Get

    End Property

    Public Sub SetInfo(oInfo As clsInfoDaten)

        _oInfo = oInfo

    End Sub

End Module