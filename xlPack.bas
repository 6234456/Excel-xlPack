' Sub fetch()
'    Dim pack As New xlPack
'    With pack
'        .addClassModules "FormatUtil", repoName:="Excel-FormatUtil"
'        .addClassModules "Dicts"
'        .addClassModules "Nodes"
'        .addClassModules "Lists"
'        .fetchModuleFiles
'    End With
'  End Sub

Option Explicit

Private pURL As String
Private isURLReg As Object
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
    
    For Each i In pModuleNames.keys
        With http
            .Open "GET", pModuleNames(i)(1) & pModuleNames(i)(0), False
            .send
            addVBEModules .responseText, i, True
        End With
    Next i
    
     For Each i In pNonModule.keys
        With http
            .Open "GET", pModuleNames(i)(1) & pModuleNames(i)(0), False
            .send
            addVBEModules .responseText, i, False
        End With
    Next i

    Set http = Nothing
End Function

Private Sub Class_Initialize()
    Me.ini
End Sub

Private Sub Class_Terminate()
    Set pModuleNames = Nothing
    Set pNonModule = Nothing
End Sub

Public Function ini()
    
    pURL = "https://raw.githubusercontent.com/6234456/Excel-VBA-Dicts/master/"
    Set pModuleNames = CreateObject("scripting.dictionary")
    Set pNonModule = CreateObject("scripting.dictionary")
    Set isURLReg = CreateObject("vbscript.regexp")
    isURLReg.pattern = "^https?\:\/\/"

End Function


' if repoName starts with http:// or https://,  the url will be applied
' if repoName is the name of certain repository buildURL will be evoked
Public Function addModule(ByVal elem As String, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas", Optional ByVal isClass As Boolean = True, Optional ByVal repoName As String)
    
    If Len(version) > 0 Then
        If Left(version, 1) <> "." Then
            version = "." & version
        End If
    End If
    
    If IsMissing(repoName) Or repoName = "" Then
        repoName = pURL
    ElseIf Not isURLReg.Test(repoName) Then
        repoName = buildURL(repo:=repoName)
    End If
    
    If isClass Then
        pModuleNames(Trim(elem)) = Array(Trim(elem) & version & extension, repoName)
    Else
        pNonModule(Trim(elem)) = Array(Trim(elem) & version & extension, repoName)
    End If
    
End Function

Public Function buildURL(Optional ByVal host As String = "https://raw.githubusercontent.com/6234456/", Optional ByVal repo As String = "Excel-VBA-Dicts/", Optional ByVal branch As String = "master/") As String
    buildURL = addEndSlashIfNotExists(host) & addEndSlashIfNotExists(repo) & addEndSlashIfNotExists(branch)
End Function

Private Function addEndSlashIfNotExists(ByVal s As String) As String
    If Right(s, 1) <> "/" Then
        s = s & "/"
    End If
    
    addEndSlashIfNotExists = s
End Function

Public Function addClassModules(elem, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas", Optional ByVal repoName As String)
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i, version, extension, True, repoName
        Next i
    Else
        addModule elem, version, extension, True, repoName
    End If
End Function

Public Function addNonClassModules(elem, Optional ByVal version As String = "", Optional ByVal extension As String = ".bas", Optional ByVal repoName As String)
    If IsArray(elem) Then
        Dim i
        For Each i In elem
            addModule i, version, extension, False, repoName
        Next i
    Else
        addModule elem, version, extension, False, repoName
    End If
End Function

Private Function addVBEModules(ByRef codeText As String, ByVal nm As String, ByVal isClass As Boolean)
    
    ' #TODO if the module duplicated
    If isClass Then
        With ThisWorkbook.VBProject.VBComponents
            ' vbext_ct_ClassModule replaced by 2,  in case of reference error
            ' vbext_ct_StdModule  1
            With .add(2)
               .Name = nm
               .CodeModule.AddFromString codeText
            End With
        End With
    Else
    
    End If

End Function
