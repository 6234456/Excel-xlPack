Sub fetch()
    Dim pack As New xlPack
    With pack
        .addClassModules "FormatUtil", repoName:="Excel-FormatUtil"
        .addClassModules "xlMiner", repoName:="xlMiner"
        .addClassModules "Dicts"
        .addClassModules "Nodes"
        .addClassModules "Lists"
        .addClassModules "TreeSets"
        .fetchModuleFiles
    End With
  End Sub
