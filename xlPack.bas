'    Dim pack As New xlPack
'    With pack
'        .addClassModules "Dicts", "v2"
'        .addClassModules "Nodes"
'        .addClassModules "Lists"
'        .fetchModuleFiles
'    End With

Option Explicit

Private pURL As String
Private pURLFallBack As String
Private pModuleNames As Object
Private pNonModule As Object

' #TODO pack all the modules in a single file
' #TODO unpack the file above to the modules
' #TODO load cfg from json/text-file??

Public Function fetchModuleFiles()

    Dim i
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    For Each i In pModuleNames.Keys
        With http
            .Open "GET", pURL & pModuleNames(i), False
            .send
            addVBEModules .responseText, i, True
        End With
    Next i
    
     For Each i In pNonModule.Keys
        With http
            .Open "GET", pURL & pNonModule(i), False
            .send
            addVBEModules .responseText, i, False
        End With
    Next i

    Set http = Nothing
End Function

Private Sub Class_Initialize()
    Me.ini
End Sub

Public Function ini()
    
    pURL = "https://raw.githubusercontent.com/6234456/Excel-VBA-Dicts/master/"
    Set pModuleNames = CreateObject("scripting.dictionary")
    Set pNonModule = CreateObject("scripting.dictionary")

End Function

Public Function addModule(ByVal elem As String, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas", Optional ByVal isClass As Boolean = True)
    
    If Len(version) > 0 Then
        If Left(version, 1) <> "." Then
            version = "." & version
        End If
    End If
    
    If isClass Then
        pModuleNames(Trim(elem)) = Trim(elem) & version & extension
    Else
        pNonModule(Trim(elem)) = Trim(elem) & version & extension
    End If
    
End Function

Public Function addClassModules(elem, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas")
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i, version, extension
        Next i
    Else
        addModule elem, version, extension
    End If
End Function

Public Function addNonClassModules(elem, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas")
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i, version, extension, False
        Next i
    Else
        addModule elem, version, extension, False
    End If
End Function

Private Function addVBEModules(ByRef codeText As String, ByVal nm As String, ByVal isClass As Boolean)
    
    ' #TODO if the module duplicated
    If isClass Then
        With ThisWorkbook.VBProject.VBComponents
            With .add(vbext_ct_ClassModule)
               .Name = nm
               .CodeModule.AddFromString codeText
            End With
        End With
    Else
    
    End If

End Function
