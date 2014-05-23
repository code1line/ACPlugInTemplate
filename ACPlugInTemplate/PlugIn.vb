Public Class PlugIn
    Implements ILSPlugIn

    '----------------------------------------------------------------------------------------------------------------------------
    'Funktion........:  gsPlugInName
    'Beschreibung....:  Gibt die Bezeichnung des PlugIns zurück
    '----------------------------------------------------------------------------------------------------------------------------
    Public Function gsPlugInName() As String Implements ILSPlugIn.gsPlugInName

        Return String.Format("{0} {1}", AppTitle, SysHelper.GetVersion(, SysHelper.VersionType.FileVersion))

    End Function

    '----------------------------------------------------------------------------------------------------------------------------
    'Funktion........:  gbEntry
    'Beschreibung....:  Eintrittspunkt für Funktionsaufrufe in das PlugIn.
    'Parameter.......:  nFunktion - Auswahl für Funktionen
    '                   oInfo     - Aufgaben-Center Datenobjekt
    '                   sArgs     - Zusätzliche Argumente
    'Ergebnis........:  Liefert einen Boolean der bei erfolgreicher Aktion True, ansonsten False ist, zurück          
    '----------------------------------------------------------------------------------------------------------------------------
    Public Function gbEntry(nFunktion As Integer, oInfo As clsInfoDaten, sArgs As String) As Boolean Implements ILSPlugIn.gbEntry

        Dim valid = False

        'Globales goInfo setzen
        SetInfo(oInfo)

        Try
            Select Case nFunktion

                Case 10
                    '---------------------------------------------------------------------------------------
                    ' FUNKTION 10: Eigene Funktionen hier einbauen...
                    '---------------------------------------------------------------------------------------
                    '...
                    'bValid = True

                Case 9001
                    '---------------------------------------------------------------------------------------
                    ' FUNKTION 9001: Testfunktion für Properties
                    '
                    ' Properties aus sArgs für eine selektierte Zeile mit [$Spalte: Platzhalter auswerten
                    '---------------------------------------------------------------------------------------
                    Dim oProperties = PlugInProperties.goGetProperties(1, sArgs)
                    MsgBox(oProperties.GetString("Test"))
                    valid = True

                Case 9002
                    '---------------------------------------------------------------------------------------
                    ' FUNKTION 9002: Testfunktion für Properties
                    '
                    ' Properties aus sArgs für mehrere selektierte Zeilen mit [$Feld: Platzhalter auswerten
                    '---------------------------------------------------------------------------------------
                    Dim oZeile As clsInfoDatenZeile
                    Dim oProperties = PlugInProperties.goGetProperties(2, sArgs)

                    For Each oZeile In CType(oInfo.SelektierteZeilen, IEnumerable)
                        MsgBox(oProperties.GetString("Test", oZeile))
                    Next
                    valid = True

                Case 9999
                    '---------------------------------------------------------------------------------------
                    ' FUNKTION 9999: Allgemeine Testfunktion
                    '---------------------------------------------------------------------------------------
                    goFunction.nMessageBox("Funktion 9999: PlugIn funktioniert!", frmInfoHinweis.lsMsgButton.lsOk, frmInfoHinweis.lsMsgImage.lsOk, AppTitle,
                                           "Die Funktion 9999 ist eine Testfunktion, um die Funktionsweise des PlugIns zu testen.")
                    valid = True

                Case Else
                    goDebugger.LogError(String.Format("Funktion {0} nicht definiert.", nFunktion), AppTitle() & ":PlugIn:gbEntry:" & nFunktion)
                    valid = False

            End Select

            Return valid

        Catch ex As Exception
            goDebugger.LogError(ex.Message, AppTitle() & ":PlugIn:gEntry:" & nFunktion, ex.StackTrace)
            Return False

        End Try

    End Function

#Region "Implementierung Menü Externer Client"

    '----------------------------------------------------------------------------------------------------------------------------
    'Prozedur........:  gbMenuUpdate
    'Beschreibung....:  Liefert für die neue Gruppe der Navigationsleiste eine Menüstruktur
    '----------------------------------------------------------------------------------------------------------------------------
    Public Function gbMenuUpdate(ByVal oParentNode As Windows.Forms.TreeNode) As Boolean Implements ILSPlugIn.gbMenuUpdate

        'Dim oFolder As Windows.Forms.TreeNode
        'Dim oNode As Windows.Forms.TreeNode
        'oFolder = oParentNode.Nodes.Add("PlugIn:=" & AppTitle, "PlugIn Ordner", 1, 1)
        'oNode = oFolder.Nodes.Add("PlugIn:=" & AppTitle & ";Entry:=1", "PlugIn Funktion Nr.1", 2, 2)
        'oNode = oFolder.Nodes.Add("PlugIn:=" & AppTitle & ";Entry:=2", "PlugIn Funktion Nr.2", 2, 2)

        'oder

        'Dim oNode As TreeNode
        'Dim oNodeSub As TreeNode

        ''Event für Aufruf über Doppelklick
        '_tree = oParentNode.TreeView
        'AddHandler _tree.DoubleClick, AddressOf MenuClickOrEnter
        'AddHandler _tree.KeyPress, AddressOf MenuClickOrEnter

        ''Zusätzliche Icons
        '_tree.ImageList.Images.Add(My.Resources.Icon)

        'oNode = New TreeNode("Testmenü")
        'oNodeSub = New TreeNode("Eintrag1", 1, 1)
        'oNode.Nodes.Add(oNodeSub)
        'oNodeSub = New TreeNode("Eintrag2", 2, 2)
        'oNode.Nodes.Add(oNodeSub)
        'oNodeSub = New TreeNode("Eintrag3", 3, 3)
        'oNode.Nodes.Add(oNodeSub)
        'oParentNode.Nodes.Add(oNode)

        'oNode = New TreeNode("Erfassung")
        'oNodeSub = New TreeNode("Test", 4, 4)
        'oNodeSub.Tag = 1
        'oNode.Nodes.Add(oNodeSub)
        'oParentNode.Nodes.Add(oNode)

        Return True

    End Function

    '----------------------------------------------------------------------------------------------------------------------------
    'Funktion........:  goMenuImage
    'Beschreibung....:  Liefert eine Grafik für die neue Gruppe der Navigationsleiste
    '----------------------------------------------------------------------------------------------------------------------------
    Public Function goMenuImage() As Drawing.Image Implements ILSPlugIn.goMenuImage

        Return Nothing

        'Return My.Resources.Icon
        'oder
        'Return Drawing.Image.FromStream(Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(AppTitle & ".PlugIn.png"))

    End Function

    '----------------------------------------------------------------------------------------------------------------------------
    'Funktion........:  gsMenuName
    'Beschreibung....:  Erzeugt in der Navigationsleiste des externen Aufgaben-Center-Clients
    '                   eine neue Gruppe mit der angegebenen Bezeichung.
    '                   Liefert die Funktion einen Leerstring so wird keine Gruppe erzeugt.
    '----------------------------------------------------------------------------------------------------------------------------
    Public Function gsMenuName() As String Implements ILSPlugIn.gsMenuName

        Return String.Empty
        'Return "Test-MenuName"

    End Function


    'Public Sub MenuClickOrEnter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim oNode As TreeNode = CType(sender, TreeView).SelectedNode
    '    Dim nodeID As Integer = CType(oNode.Tag, Integer)

    '    If nodeID = 0 Then Exit Sub

    '    If e.GetType Is GetType(KeyPressEventArgs) Then
    '        Dim key As Char = CType(e, KeyPressEventArgs).KeyChar
    '        If CType(e, KeyPressEventArgs).KeyChar <> Chr(Keys.Return) Then Exit Sub
    '    End If

    '    Select Case nodeID
    '        Case 1
    '            MessageBox.Show(nodeID.ToString)

    '    End Select

    'End Sub

#End Region

End Class

