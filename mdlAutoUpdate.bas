#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
    Private Const INTERNET_CONNECTION_MODEM As LongPtr = &H1
    Private Const INTERNET_CONNECTION_LAN As LongPtr = &H2
    Private Const INTERNET_CONNECTION_PROXY As LongPtr = &H4
    Private Const INTERNET_CONNECTION_OFFLINE As LongPtr = &H20
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
    Private Const INTERNET_CONNECTION_MODEM As Long = &H1
    Private Const INTERNET_CONNECTION_LAN As Long = &H2
    Private Const INTERNET_CONNECTION_PROXY As Long = &H4
    Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
#End If

Const gistOwner = "paulsteigel"

Function ActiveInternet() As Boolean
    Dim L As Long, R As Long
    R = InternetGetConnectedState(L, 0&)
    If R <= 4 And R <> 0 Then ActiveInternet = True
End Function

Private Function gtCreateReferences(dom As Object) 'DOMDocument)
    ' adds all current references to an xml
    Dim R As Object ' Reference
    
    With ActiveWorkbook.VBProject
        For Each R In .References
            gtAddRefToManifest dom, R
        Next R
    End With
End Function

Private Function gtUpdateAll()
    ' do all updates for known manifests in this project
    Dim modle As Object 'VBComponent
    Dim stampLine As Long, co As New Collection, manifest As String, s As String, v As Variant
    For Each modle In ThisWorkbook.VBProject.VBComponents
        ' do we know this module?
        stampLine = gtManageable(modle)
        If stampLine > 0 Then
            ' yes we do - get the manifest
            manifest = gtStampManifest(modle, stampLine)
            If (manifest = vbNullString) Then
                MsgBox ("gistThat stamp in module " & modle.Name & " fiddled with.Run again as greenField")
            Else
            ' add it to the collection of manifests we need to process
              If gtCoIndex(manifest, co) = 0 Then
                co.Add manifest, manifest
              End If
            End If
        End If
    Next modle
    ' todo check if versions are up to date
    If co.Count > 0 Then
        s = vbNullString
        For Each v In co
            gtDoit CStr(v)
            If s <> vbNullString Then s = s & ","
            s = s & CStr(v)
        Next v
        MsgBox ("updated " & co.Count & " manifests(" & s & ")")
    End If
End Function

Private Function gtCoIndex(sid As Variant, co As Collection) As Long
    ' find index in a collection
    Dim i As Long
    gtCoIndex = 0
    For i = 1 To co.Count
        If co(i) = sid Then
          gtCoIndex = i
          Exit Function
        End If
    Next i
End Function

Private Function gtPreventCaching(url As String) As String
    ' this will tweak the url with an extra random parameter to prevent any accidental caching
    Dim p As String
    If (InStr(1, url, "?") > 0) Then
        p = "&"
    Else
        p = "?"
    End If
    Randomize
    gtPreventCaching = url & p & "gtPreventCaching=" & CStr(Int(10000 * Rnd()))
End Function

Public Function UpdateAvaiable() As Boolean
    ' First check for active internet connection
    If Not ActiveInternet Then Exit Function
    
    Dim myDom As Object, t As String, xNode As Object, lRet As Boolean
    ' get the requested manifest
    Set myDom = gtRecreateManifest("manifest.xml")
    
    Set xNode = myDom.selectSingleNode("//gists").FirstChild
    
    t = xNode.Attributes.getNamedItem("version").Text
    ' Try to get version of this app
    lRet = IIf(Val(t) <= GetAppVersion(), False, True)
    
    ' Now forced update if needed
    If lRet Then
        ShowStatus "The application is updating... please wait!"
        Call gtDoit("", True, myDom)
        WriteLog "Get to Update Avaiable message..."
    End If
    UpdateAvaiable = lRet
End Function

Public Function gtDoit(gtDoitmanifestID As String, Optional greenField As Boolean = False, Optional dom As Object) As Boolean
    Dim rawUrl As String, t As String, n As String, g As String
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim vbCom As Object 'VBComponent
    
    ' now we know which gists are needed here
    If (gtWillItWork(dom, greenField)) Then
        ' theres a good chance it will work
        ' for each module
        For Each xNode In dom.selectSingleNode("//gists").childNodes
            t = xNode.Attributes.getNamedItem("type").Text
            Select Case t
                Case "class", "module"
                    ' get the gist
                    rawUrl = gtConstructRawUrl(, xNode.Attributes.getNamedItem("filename").Text)
                    ' prevent caching will make it look like a different request
                    g = gtHttpGet(gtPreventCaching(rawUrl))
                    ' module name
                    n = xNode.Attributes.getNamedItem("module").Text
                    ' does it exist - if so then delete it
                    Set vbCom = gtModuleExists(n, ThisWorkbook)
                    
                    If (Not vbCom Is Nothing) Then
                        ' delete everything in it
                        WriteLog "Deleting code in module..." & vbTab & vbCom.Name
                        With vbCom.codeModule
                            .DeleteLines 1, .CountOfLines
                        End With
                    Else
                        ' And now add back...
                        Set vbCom = gtAddModule(n, ThisWorkbook, xNode.Attributes.getNamedItem("type").Text)
                    End If
        
                    ' add in the new code
                    WriteLog "Start inserting code in module..." & vbTab & vbCom.Name
                    With vbCom.codeModule
                        .AddFromString g
                    End With
        
                    ' stamp it
                    WriteLog "Writing stam for code in module..." & vbTab & vbCom.Name
                    gtInsertStamp vbCom, gtDoitmanifestID, rawUrl
                
                Case "reference"
                    gtAddReference xNode.Attributes.getNamedItem("name").Text, _
                                   xNode.Attributes.getNamedItem("guid").Text, _
                                   xNode.Attributes.getNamedItem("major").Text, _
                                   xNode.Attributes.getNamedItem("minor").Text
                Case Else
                    'Debug.Assert False
            
            End Select
        Next xNode
        WriteLog "End of generating code..."
        gtDoit = True
    Else
 
    End If
    WriteLog "Exit gtDoit..."
End Function
 
Private Function gtAddReference(Name As String, guid As String, major As String, minor As String) As Object ' Reference
    ' add a reference (if its not already there)
    Dim R As Object ' Reference
    On Error GoTo handle
    With ActiveWorkbook.VBProject
        For Each R In .References
            If (R.Name = Name) Then
                If (R.major < major Or R.major = major And R.minor < minor And Not R.BuiltIn) Then
                    .References.AddFromGuid guid, major, minor
                    .References.Remove (R)
                End If
                Exit Function
            End If
        Next R
    ' if we get here then we need to add it
      Set gtAddReference = .References.AddFromGuid(guid, major, minor)
      Exit Function
    End With
    
handle:
    MsgBox ("warning - tried and failed to add reference to " & Name & "v" & major & "." & minor)
    Exit Function
    
End Function

Private Function gtStampManifest(vbCom As Object, line As Long) As String 'VBComponent
    ' the manifest should be on the given line
    Dim s As String, n As Long, p As Long, marker As String
    marker = "manifest:"
    s = vbNullString
    With vbCom.codeModule
       n = InStr(1, LCase(.Lines(line, 1)), marker)
       If (n > 0) Then
        s = Mid(.Lines(line, 1), n + Len(marker))
        p = InStr(1, s, " ")
        s = Left(s, p - 1)
       End If
    End With
    gtStampManifest = s
End Function

Private Function gtInsertStamp(vbCom As Object, manifest As String, rawUrl As String) As Long 'VBComponent
    Dim stampLine As Long
    stampLine = gtManageable(vbCom)
    ' if it wasnt found then insert at line 1
    With vbCom.codeModule
        If stampLine <> 0 Then
            .DeleteLines stampLine, 1
        Else
            stampLine = 1
        End If
        .InsertLines stampLine, gtStampLog(manifest, rawUrl)
    End With
    gtInsertStamp = stampLine
    
End Function

Private Function gtWillItWork(dom As Object, _
                Optional greenField As Boolean = False) As Boolean 'DOMDocument
    
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim n As String, s As String, t As String
    Dim modle As Object 'VBComponent
    
    ' check we have something to do
    gtWillItWork = Not dom Is Nothing
    If Not gtWillItWork Then
        Exit Function
    End If
    ' first we check if these are new modules
    s = vbNullString
    For Each xNode In dom.selectSingleNode("//gists").childNodes
        ' the target module
        t = xNode.Attributes.getNamedItem("type").Text
        Select Case t
            Case "class", "module"
                n = xNode.Attributes.getNamedItem("module").Text
                Set modle = gtModuleExists(n, ThisWorkbook)
                
                If (Not modle Is Nothing) Then
                    ' it exists - validate its not somethig else with the same name
                    If (gtManageable(modle) = 0 And Not greenField) Then
                        s = gtAddStr(s, n)
                    End If
                End If
            Case "reference"
            Case "version"
            Case Else
                s = gtAddStr(s, "unknown type " & t)
        End Select
    Next xNode
    
    If (s <> vbNullString) Then
        MsgBox ("there may be a conflict with these modules names (" & s & _
            ") and some others in your project. " & _
            "If this is the first time you have run this - run with greenfield set to true to override this check")
        gtWillItWork = False
        Exit Function
    End If
   
   ' now check all gists are getable
   ' todo
   
End Function

Private Function gtAddStr(t As String, n As String) As String
    Dim s As String
    s = t
    If (s <> vbNullString) Then s = s & ","
    gtAddStr = s & n
End Function
 
Private Function gtRecreateManifest(Optional manifestID As String = "") As Object
    Dim dom As Object
    Dim manifest As String
    
    ' get the xml string
    manifest = gtHttpGet(gtPreventCaching(gtConstructRawUrl(manifestID)))
    
    If manifest <> vbNullString Then
    ' parse the xml
        Set dom = CreateObject("MSXML.DOMDocument")
        dom.LoadXML (manifest)
        Set gtRecreateManifest = dom
    Else
        MsgBox ("Could not get manifest for " & manifestID)
    End If
End Function
 
Private Function gtModuleExists(Name As String, wb As Workbook) As Object
    ' determine whether this module exists in the given workbook
    Dim modle As Object 'VBComponent
    
    For Each modle In wb.VBProject.VBComponents
       If Trim(LCase(modle.Name)) = Trim(LCase(Name)) Then
            Set gtModuleExists = modle
            Exit Function
       End If
    Next modle
End Function
 
Private Function gtAddModule(Name As String, wb As Workbook, modType As String) As Object ' VBComponent
    ' determine whether this module exists in the given workbook
    Dim modle As Object, t As String ' VBComponent, t As Long
 
    Select Case LCase(modType)
        Case "class"
            t = 2
        Case "module"
            t = 1
        Case Else
            MsgBox ("unknown module type " & modType)
    End Select
        
    Set modle = wb.VBProject.VBComponents.Add(t)
    modle.Name = Name
    
    ' added by andypope.info
    If modle.codeModule.CountOfLines > 1 Then
        ' remove Option Explict lines if it was added automatically
        modle.codeModule.DeleteLines 1, modle.codeModule.CountOfLines
    End If
    
    Set gtAddModule = modle
End Function
 
Private Function gtConstructRawUrl(Optional gistID As String = "", Optional gistFileName As String = vbNullString) As String
    ' given a gist, where is it?
    Dim s As String
    ' raw URL
    s = "https://raw.githubusercontent.com/" & gistOwner & "/les/master/" & gistID
    ' a gist can have multiple files in it
    If gistFileName <> vbNullString Then s = s & IIf(gistID = "", "", "/") & gistFileName
    ' TODO - specific versions
    gtConstructRawUrl = s
    Debug.Print s
End Function
 
Private Function gtAddToManifest(dom As Object, _
                                 gistID As String, _
                                 modType As String, _
                                 modle As String, _
                                 Optional FileName As String = vbNullString, _
                                 Optional version As String = vbNullString _
                        ) As Object ' DOMDocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.childNodes.Length + 1))
    head.appendChild Element
    
    Set attrib = dom.createAttribute("gistid")
    attrib.NodeValue = gistID
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("version")
    attrib.NodeValue = version
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("filename")
    attrib.NodeValue = FileName
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("module")
    attrib.NodeValue = modle
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = modType
    Element.setAttributeNode attrib
    
    Set gtAddToManifest = dom
End Function

Private Function gtAddRefToManifest(dom As Object, R As Object) As Object   ' DOMDocument, reference, domdocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.childNodes.Length + 1))
    head.appendChild Element
    'r.GUID, r.name, r.Major, r.Minor, r.description
    Set attrib = dom.createAttribute("guid")
    attrib.NodeValue = R.guid
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("name")
    attrib.NodeValue = R.Name
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("major")
    attrib.NodeValue = R.major
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("minor")
    attrib.NodeValue = R.minor
    Element.setAttributeNode attrib
 
    Set attrib = dom.createAttribute("description")
    attrib.NodeValue = R.description
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = "reference"
    Element.setAttributeNode attrib
    
    Set gtAddRefToManifest = dom
End Function

Private Function gtInitManifest(Optional description As String = vbNullString, _
                                 Optional contact As String = vbNullString) As Object ' DOMDocument
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim dom As Object ' DOMDocument
    Dim e2 As Object ' IXMLDOMElement
   
 ' creates an xml manifest of required gists
   Set dom = CreateObject("MSXML.DOMDocument")
 
    Set Element = dom.createElement("gistThat")
    Set attrib = dom.createAttribute("info")
    attrib.NodeValue = _
            "this is a manifest for gistThat VBA code distribution " & _
            " - see ramblings.mcpher.com for details"
    Element.setAttributeNode attrib
 
    
    With dom.appendChild(Element)
        Set Element = dom.createElement("manifest")
        .appendChild Element
        
 
        
        Set attrib = dom.createAttribute("description")
        attrib.NodeValue = description
        Element.setAttributeNode attrib
 
        Set attrib = dom.createAttribute("contact")
        attrib.NodeValue = contact
        Element.setAttributeNode attrib
 
        Element.appendChild dom.createElement("gists")
 
    End With
    Set gtInitManifest = dom
   
End Function
 
Private Function gtHttpGet(url As String) As String
    ' TODO oAuth
    Dim ohttp As Object
    Set ohttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    Call ohttp.Open("GET", url, False)
    Call ohttp.Send("")
    gtHttpGet = ohttp.ResponseText
    Set ohttp = Nothing
End Function

Private Function gtStampLog(manifest As String, rawUrl As String) As String
    ' create a comment to identify this as manageable
    gtStampLog = gtStamp & _
        " updated on " & Now() & " : from manifest:" & _
        manifest & _
        " gist " & rawUrl
End Function

Private Function gtStamp() As String
' this marks a module as manageable
    gtStamp = "'gistThat@mcpher.com :do not modify this line" & _
    " - see ramblings.mcpher.com for details:"
End Function

Private Function gtManageable(vbCom As Object) As Long  ' VBComponent
    ' return the line number of the gtStamp
    ' parameters as passed by ref in .find method
    Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
    startLine = 1: endLine = vbCom.codeModule.CountOfLines: startColumn = 1: endColumn = 255
    
    If (vbCom.codeModule.Find(gtStamp(), startLine, startColumn, endLine, endColumn)) Then
        gtManageable = startLine
    End If
    
End Function

Sub ShowStatus(msgStatus As String)
    Application.StatusBar = msgStatus
End Sub
