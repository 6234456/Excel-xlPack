Sub main()
    
    Dim pack As New xlPack

    With pack
        .addClassModules "Dicts", "v2"
        .addClassModules "Nodes"
        .addClassModules "Lists"
        .fetchModuleFiles
    End With

End Sub
