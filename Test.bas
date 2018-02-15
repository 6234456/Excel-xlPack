Sub main()
    
    Dim pack As New xlPack
    
    With pack
        .ini
        .addClassModules Split("Dicts_Nodes_Lists", "_")
        .fetchModuleFiles
    End With

End Sub
