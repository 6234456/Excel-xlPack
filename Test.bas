Sub fetch()
    Dim pack As New xlPack
    With pack
        .addClassModules "FormatUtil", repoName:="Excel-FormatUtil"
        .addClassModules "Dicts"
        .addClassModules "Nodes"
        .addClassModules "Lists"
        .fetchModuleFiles
    End With
  End Sub
