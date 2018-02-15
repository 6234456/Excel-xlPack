Option Explicit

Private pURL As String
Private pURLFallBack As String
Private pModuleNames As Object

' #TODO pack all the modules in a single file
' #TODO unpack the file above to the modules
' #TODO load cfg from json/text-file??

Public Function fetchModuleFiles()

    Dim i
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    For Each i In pModuleNames.Keys
        With http
            .Open "GET", pURL & i, False
            .send
            addVBEModules .responseText, i, pModuleNames(i)
        End With
    Next i

    Set http = Nothing
End Function

Public Function ini()
    
    pURL = "http://www.qiou.eu/vba/"
    Set pModuleNames = CreateObject("scripting.dictionary")

End Function

Public Function addModule(ByVal elem As String, Optional ByVal isClassModule As Boolean = True)
   
    pModuleNames(Trim(elem)) = isClassModule
    
End Function

Public Function addClassModules(elem)
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i
        Next i
    Else
        addModule elem
    End If
End Function

Public Function addNonClassModules(elem)
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i, False
        Next i
    Else
        addModule elem, False
    End If
End Function

Private Function addVBEModules(ByRef codeText As String, ByVal nm As String, ByVal isClass As Boolean)
    
    ' #TODO if the module duplicated
    If isClass Then
        With ThisWorkbook.VBProject.VBComponents.add(vbext_ct_ClassModule)
            .Name = nm
            .CodeModule.AddFromString codeText
        End With
    Else
    
    End If

End Function
