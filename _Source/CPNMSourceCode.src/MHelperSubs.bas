Attribute VB_Name = "MHelperSubs"
' Chemtech - A Siemens Business ========================================================
'
'=======================================================================================
'Option Explicit
' Desenvolvimento ======================================================================
' <iniciais>            Renan       <email>
' <iniciais>            Paulo       <email>
'=======================================================================================
' Versões ==============================================================================
' v0.5 - Implementação da
' v0.4 - Melhorias gerais e consolidação do CPNM para Excel e Autocad.
'=======================================================================================

Public gCnn                                       As ADODB.Connection
Public gUnitDef                                   As unitDefinition
Public Const unitOnlyChar                         As String = "U"
Public Const valueOnlyChar                        As String = "V"
Public Const unitAndValueChar                     As String = "D"
Public Const trackingChar                         As String = "Z"
Public Const calcChar                             As String = "C"
Public Const breakString                          As String = "dummy"
Public Const calculatedPropClassKey               As Integer = 5
Public Const nonExistantValueString               As String = "-"

Public Function isInCollection(col As collection, key As String) As Boolean
    '=======================================================================================
    ' checa se o argumento "key" aparece na coleção "col", retornando True
    '---------------------------------------------------------------------------------------
    ' [key] chave a ser buscada dentro da coleção
    ' [col] coleção que deve conter a chave
    '---------------------------------------------------------------------------------------
    ' < Observações >
    '---------------------------------------------------------------------------------------
    ' < Histórico de revisões>
    '=======================================================================================
    '
    Dim var                                       As Variant
    Dim errNumber                                 As Long
    Dim InCollection                              As Boolean

    InCollection = False
    Set var = Nothing

    Err.Clear
    On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(Err.Number)
    On Error GoTo 0

    '5 is not in, 0 and 438 represent incollection
    If errNumber = 5 Then                                            ' it is 5 if not in collection
        isInCollection = False
    Else
        isInCollection = True
    End If
End Function

Public Function createTrackingDictionary()
    ' This function creates a dictionary between field names and the names to display to individual users
    Dim colDictionary                             As collection

    Set colDictionary = New collection

    colDictionary.Add "NOME_ITEM", "Nome do item"
    colDictionary.Add "NOME_SUB_ARE", "Sub-area do item"
    colDictionary.Add "NOME_ARE", "Area do item"
    colDictionary.Add "NOME_PLA", "Planta do item"
    colDictionary.Add "NOME_IND", "Industrial do item"
    colDictionary.Add "NOME_UNI", "Unidade de negócio do item"

    colDictionary.Add "Nome do item", "NOME_ITEM"
    colDictionary.Add "Sub-area do item", "NOME_SUB_ARE"
    colDictionary.Add "Area do item", "NOME_ARE"
    colDictionary.Add "Planta do item", "NOME_PLA"
    colDictionary.Add "Industrial do item", "NOME_IND"
    colDictionary.Add "Unidade de negócio do item", "NOME_UNI"

    Set createTrackingDictionary = colDictionary
End Function

Public Sub createGlobalUnitDefinitionObject()
    Set gUnitDef = New unitDefinition
End Sub

Public Function defineDatabaseConnection(Optional promptOption As Integer = 1) As ADODB.Connection
    Dim strDbPath                                 As String
    Dim strConnect                                As String
    Dim cnn                                       As ADODB.Connection

    Set cnn = New ADODB.Connection

    strConnect = "Provider=sqloledb;" _
               & "Database=CHT-CPNM;" _
               & "Server=WSP-I02-V; " _
               & "DataTypeCompatibility=80;" _
               & "Integrated Security = SSPI;" _
               & "MARS Connection=True;"

    

    cnn.ConnectionString = strConnect
    Set defineDatabaseConnection = cnn
End Function

' #############
' Configuration Paths

Public Sub storeConnectionString(storeConnection)
    'This sub stores the connection string into Windows registry
    SaveSetting "CPNM", "ConnectionConfig", "ConnectionString", storeConnection
End Sub

Public Function getStoredConnectionString() As String
    ' This sub stores the connection string into the windows registry
    getStoredConnectionString = GetSetting("CPNM", "ConnectionConfig", "ConnectionString")
End Function

Public Sub storeUnitsXmlPath(xmlPath As String)
    'This sub stores the connection string into Windows registry
    SaveSetting "CPNM", "UnitsConfig", "unitsXmlPath", xmlPath
End Sub

Public Function getUnitsXmlPath() As String
    ' This sub stores the connection string into the windows registry
    getUnitsXmlPath = GetSetting("CPNM", "UnitsConfig", "unitsXmlPath")
End Function

Public Sub storeSharesXmlPath(sharesXmlPath As String)
    'This sub stores the connection string into Windows registry
    SaveSetting "CPNM", "SharesConfig", "sharesXmlPath", sharesXmlPath
End Sub

Public Function getSharesXmlPath() As String
    ' This sub stores the connection string into the windows registry
    getSharesXmlPath = GetSetting("CPNM", "SharesConfig", "sharesXmlPath")
End Function

Public Sub appendToShareXML(itemTypeKey1 As Long, itemTypeKey2 As Long, propList1 As collection, _
                            propList2 As collection, strShareName As String)

    '=======================================================================================
    ' Esta rotina salva um compartilhamento definido na interface para o XML.
    '---------------------------------------------------------------------------------------
    ' [itemTypeKey1] - chave primaria do tipo do primeiro item.
    ' [itemTypeKey2] - chave primaria do tipo do segundo item.
    ' [propList1]    - colecao de propriedades do primeiro item para compartilhamento.
    ' [propList2]    - colecao de propriedades do segundo item para compartilhamento.
    ' [strShareName] - nome a ser dado para o compartilhamento.
    '---------------------------------------------------------------------------------------
    ' a rotina assume que todos os argumentos foram validados antes da chamada.
    '---------------------------------------------------------------------------------------
    ' < Histórico de revisões>
    '=======================================================================================
    '
    Dim objXml                                    As MSXML2.DOMDocument
    Dim objThisNode                               As MSXML2.IXMLDOMNode
    Dim objShareNode                              As MSXML2.IXMLDOMNode
    Dim objPropPairNode                           As MSXML2.IXMLDOMNode

    ' iniciando o documento xml e o nó pai.
    Set objXml = New MSXML2.DOMDocument
    If False = objXml.Load(getSharesXmlPath()) Then                  'opening the shared xml path.
        MsgBox "Redefina a localização do xml de configuração dos compartilhamentos"
        storeSharesXmlPath (getFileDialog())
        objXml.Load (getSharesXmlPath())
    End If

    Set objShareNode = objXml.createNode(MSXML2.NODE_ELEMENT, "share", "")

    ' incluindo no nó pai o nome fornecido.
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "name", "")
    objThisNode.Text = strShareName
    Call objShareNode.appendChild(objThisNode)
    '#######
    ' including into the parent node the user name who created the share.
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "user", "")
    objThisNode.Text = Environ$("username")
    Call objShareNode.appendChild(objThisNode)

    ' including into the parent node the date which the share was created
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "data", "")
    objThisNode.Text = DateTime.Now
    Call objShareNode.appendChild(objThisNode)
    '########
    ' including the item types nde into the xml
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "typePair", "")
    Call objShareNode.appendChild(objThisNode)

    ' including the first item type
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "type1", "")
    objThisNode.Text = itemTypeKey1
    Call objShareNode.SelectSingleNode("typePair").appendChild(objThisNode)

    ' including the second item type
    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "type2", "")
    objThisNode.Text = itemTypeKey2
    Call objShareNode.SelectSingleNode("typePair").appendChild(objThisNode)

    Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "propPairSet", "")
    Call objShareNode.appendChild(objThisNode)

    For i = 1 To propList1.Count
        ' creating the propPair node.
        Set objPropPairNode = objXml.createNode(MSXML2.NODE_ELEMENT, "propPair", "")

        ' adding the first prop to the pair node.
        Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "prop1", "")
        objThisNode.Text = propList1(i)
        Call objPropPairNode.appendChild(objThisNode)

        ' adding the second prop to the pair node
        Set objThisNode = objXml.createNode(MSXML2.NODE_ELEMENT, "prop2", "")
        objThisNode.Text = propList2(i)
        Call objPropPairNode.appendChild(objThisNode)

        ' appending the whole propPair to the propPairSet node.
        Call objShareNode.SelectSingleNode("propPairSet").appendChild(objPropPairNode)
    Next i

    ' Now, the share node is ready to be appended to the whole XML document.
    Call objXml.SelectSingleNode("storedShares").appendChild(objShareNode)
    Call objXml.Save(getSharesXmlPath())
End Sub

Public Sub checkShareCollection(arrCollection As collection)
    Dim propList1                                 As collection
    Dim propList2                                 As collection
    Dim strShareName                              As String
    Dim itemTypeKey1                              As Long
    Dim itemTypeKey2                              As Long
    Dim itemName1                                 As String
    Dim itemName2                                 As String
    Dim propName1                                 As String
    Dim propName2                                 As String
    Dim j                                         As Integer
    Dim varArray                                  As Variant

    ' initiating propLists
    On Error GoTo checkShareCollection_Error

    Set propList1 = New collection
    Set propList2 = New collection

    varArray = arrCollection(1)

    itemName1 = varArray(0)
    itemName2 = varArray(2)

    itemTypeKey1 = getItemTypeFromItemKey(getItemKey(itemName1))
    itemTypeKey2 = getItemTypeFromItemKey(getItemKey(itemName2))

    strShareName = InputBox("Selectione um nome para o compartilhamento")

    For j = 1 To arrCollection.Count
        varArray = arrCollection(j)

        propName1 = varArray(1)
        propName2 = varArray(3)

        propList1.Add getPropKey(propName1)
        propList2.Add getPropKey(propName2)

        If itemName1 <> varArray(0) Or itemName2 <> varArray(2) Then
            Err.Raise vbObjectError + 554565, Description:="Para salvar um compartilhamento deve-se adicionar apenas referencias a um par de items."
        End If
    Next j

    Call appendToShareXML(itemTypeKey1, itemTypeKey2, propList1, propList2, strShareName)

    On Error GoTo 0
    GoTo checkShareCollection_ExitSub

checkShareCollection_ExitSub:
    Exit Sub

checkShareCollection_Error:
    errorAction = handleMyError(True)
    Select Case errorAction
        Case -1
Stop
        Case 1
            Resume Next
        Case 2
            GoTo checkShareCollection_ExitSub
        Case Else
            GoTo checkShareCollection_ExitSub
    End Select

End Sub

Public Function getFileDialog()
    Dim strDbPath                                 As String

    With Application.FileDialog(msoFileDialogOpen)
        .Show
        If .SelectedItems.Count = 1 Then
            strDbPath = .SelectedItems(1)
        End If
    End With
    getFileDialog = strDbPath
End Function

Public Sub resetEnvironment()
    ' This sub creates the connection to the database
    Call initializeCPNM
    
    If gCnn.State = adStateOpen Then
        gCnn.Close
    End If

    If gCnn.State = adStateClosed Then
        MsgBox "selecione o .mdb do banco de dados"
        Set gCnn = defineDatabaseConnection(1)
        gCnn.Open
    Else
        MsgBox "there is a problem with the database connection object."
        Exit Sub
    End If

    ' redefining the unitDefinition storage
    Call gUnitDef.redefPath

    ' redefining the shares storage
    MsgBox "selecione o xml de compartilhamentos"

    Dim sharesPath                                As String
    sharesPath = getFileDialog
    If sharesPath <> "" Then
        Call storeSharesXmlPath(sharesPath)
    End If
End Sub
