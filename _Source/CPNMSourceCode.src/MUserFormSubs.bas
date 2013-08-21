Attribute VB_Name = "MUserFormSubs"
' Chemtech - A Siemens Business ========================================================
'
'=======================================================================================
'Option Explicit
' Desenvolvimento ======================================================================
' <iniciais>            Renan       <email>
'=======================================================================================
' Versões ==============================================================================
'
'
'
'=======================================================================================

Sub runCPNM()
    ' This sub runs the information manager
    Call initializeCPNM
    frmRunManager.Show vbModeless
End Sub

Public Sub initializeCPNM()
    Call createGlobalUnitDefinitionObject
    Set gCnn = defineDatabaseConnection(-1)
    gCnn.Open
End Sub

Public Sub populateItems(strItemTypeName As String, cmbItemName, Optional strSubAreaName As String)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim itemTypeKey                               As Long
    Dim strSQL                                    As String
    Dim subAreaKey                                As Long

    On Error GoTo populateItems_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' POPULATING THE ITEM COMBOBOX
    itemTypeKey = getItemTypeKey(strItemTypeName)

    ' Now, I query over all item's where item.itemtype = itemtypekey
    If strSubAreaName = Empty Then
        strSQL = "Select NOME_ITEM from ITEM where ID_TIPO_ITEM = " & itemTypeKey & _
               " and ATIVO = -1 ORDER BY NOME_ITEM"
    Else
        subAreaKey = getSubAreaKey(strSubAreaName)
        strSQL = "Select NOME_ITEM from ITEM where ID_TIPO_ITEM = " & itemTypeKey & _
                 "and ID_SUB_ARE = " & subAreaKey & " and ATIVO = true ORDER BY NOME_ITEM"
    End If
    rs.Open strSQL, gCnn

    ' Adding the Item names into the combo
    If rs.EOF = True Then
        MsgBox "You have no item of the selected type!"
        Exit Sub
    End If

    Do While rs.EOF <> True
        cmbItemName.AddItem rs.Fields("NOME_ITEM").value
        rs.MoveNext
    Loop
    ' Voilá
    On Error GoTo 0

    Exit Sub

populateItems_Error:
    Call handleMyError
End Sub

Public Sub populateItemType(cmbItemType, Optional strItemClassName As String)
    ' The following code populates the item type combobox with item_type's name

    ' Declarations
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set colDictionary = createTrackingDictionary()

    ' Here i put the command for the query, filtering for the specifications
    If strItemClassName <> Empty Then
        itemClassKey = getItemClassKey(strItemClassName)
        strSQL = "select NOME_TIPO_ITEM from TIPO_ITEM where ID_CLASSE_TIPO_ITEM = " & itemClassKey & " order by NOME_TIPO_ITEM"
    Else
        strSQL = "Select NOME_TIPO_ITEM from TIPO_ITEM ORDER BY NOME_TIPO_ITEM "
    End If

    ' Here i open the query in the recordset rs
    rs.Open strSQL, gCnn

    'Now, I iterate over the query and fill the value of each item type in the combo
    Do While rs.EOF <> True
        cmbItemType.AddItem (rs.Fields("NOME_TIPO_ITEM").value)
        rs.MoveNext
    Loop
End Sub

Public Sub populateItemClass(cmbItemClass)
    ' The following code populates the item type combobox with item_type's name

    ' Declarations
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    Set colDictionary = createTrackingDictionary()

    ' Here i put the command for the query,
    strSQL = "Select NOME_CLASSE_TIPO_ITEM from CLASSE_TIPO_ITEM ORDER BY NOME_CLASSE_TIPO_ITEM"

    ' Here i open the query in the recordset rs
    rs.Open strSQL, gCnn

    'Now, I iterate over the query and fill the value of each item type in the combo
    Do While rs.EOF <> True
        cmbItemClass.AddItem (rs.Fields("NOME_CLASSE_TIPO_ITEM").value)
        rs.MoveNext
    Loop
End Sub

Public Sub populateShares(strItemName As String, cmbShares, Optional strPropNameShared As String)
    ' The following code populates the item type combobox with item_type's name

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim strInsideQuery                            As String
    Dim itemKey                                   As Long
    Dim propKey                                   As Long

    On Error GoTo populateShares_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' getting itemKey from name
    itemKey = getItemKey(strItemName)

    ' Here i put the command for the query
    If strPropNameShared <> "" Then
        propKey = getPropKey(strPropNameShared)

        strInsideQuery = "(select ID_VALOR from LINK_VALORES inner join ITEM on ITEM.ID_ITEM = LINK_VALORES.ID_ITEM" & _
                       " where LINK_VALORES.ID_ITEM = " & itemKey & " and LINK_VALORES.ID_TIPO_PROP = " & propKey & ")"

        strSQL = "select ITEM.NOME_ITEM,ITEM.ID_ITEM,LINK_VALORES.ID_TIPO_PROP " & _
                 "from (LINK_VALORES inner join ITEM on ITEM.ID_ITEM = LINK_VALORES.ID_ITEM) " & _
                 "where ID_VALOR in " & strInsideQuery
    Else
        strInsideQuery = "(select ID_VALOR from LINK_VALORES inner join ITEM on ITEM.ID_ITEM = LINK_VALORES.ID_ITEM" & _
                       " where LINK_VALORES.ID_ITEM = " & itemKey & ")"

        strSQL = "select ITEM.NOME_ITEM,ITEM.ID_ITEM,LINK_VALORES.ID_TIPO_PROP " & _
                 "from (LINK_VALORES inner join ITEM on ITEM.ID_ITEM = LINK_VALORES.ID_ITEM) " & _
                 "where ID_VALOR in " & strInsideQuery & " and ITEM.ID_ITEM <> " & itemKey
    End If

    ' Here i open the query in the recordset rs
    rs.Open strSQL, gCnn

    'Now, I iterate over the query and fill the value of each item type in the combo
    Do While rs.EOF <> True
        If Not (rs.Fields("ID_ITEM").value = itemKey And _
                rs.Fields("ID_TIPO_PROP").value = propKey) Then
            cmbShares.AddItem (rs.Fields("NOME_ITEM").value)
        End If
        rs.MoveNext
    Loop
    On Error GoTo 0

    Exit Sub

populateShares_Error:
    Call handleMyError
End Sub

Public Sub populateProps(strItemTypeName As String, cmbPropName, Optional strPropClass As String)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim itemTypeKey                               As Long
    Dim propClassKey                              As Long

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Getting the keys
    itemTypeKey = getItemTypeKey(strItemTypeName)

    ' POPULATING THE PROPERTY COMBOBOX

    ' Now, I query over all item's properties where item.itemtype = itemtypekey
    ' Getting the item type primary key
    If strPropClass = Empty Then
        strSQL = " SELECT PROPRIEDADES_ITEMS.ID_TIPO_ITEM, TIPO_PROPRIEDADES.NOME_TIPO_PROP, PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " FROM TIPO_PROPRIEDADES INNER JOIN PROPRIEDADES_ITEMS ON TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " WHERE (((PROPRIEDADES_ITEMS.ID_TIPO_ITEM)= " & itemTypeKey & "))" & _
               " ORDER BY TIPO_PROPRIEDADES.NOME_TIPO_PROP "
    Else
        propClassKey = getPropClassKey(strPropClass)
        strSQL = " SELECT PROPRIEDADES_ITEMS.ID_TIPO_ITEM, TIPO_PROPRIEDADES.NOME_TIPO_PROP, PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " FROM TIPO_PROPRIEDADES INNER JOIN PROPRIEDADES_ITEMS ON TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " WHERE PROPRIEDADES_ITEMS.ID_TIPO_ITEM= " & itemTypeKey & " AND TIPO_PROPRIEDADES.ID_CLASSE_TIPO_PROP = " & propClassKey & _
               " ORDER BY TIPO_PROPRIEDADES.NOME_TIPO_PROP "
    End If
    rs.Open strSQL, gCnn

    ' Adding the fields from the table into the combo
    On Error Resume Next
    rs.MoveFirst
    Do While rs.EOF <> True
        cmbPropName.AddItem rs.Fields("NOME_TIPO_PROP").value
        rs.MoveNext
    Loop
    On Error GoTo 0
End Sub

Public Sub populateExistantProps(strItemName As String, cmbPropName, Optional strPropClass As String)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim itemKey                                   As Long
    Dim propClassKey                              As Long

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Getting itemTypeKey
    itemKey = getItemKey(strItemName)
    itemTypeKey = getItemTypeFromItemKey(itemKey)

    ' POPULATING THE PROPERTY COMBOBOX

    ' Now, I query over all item's properties where item.itemtype = itemtypekey
    ' Getting the item type primary key
    If strPropClass = Empty Then
        strSQL = " select TIPO_PROPRIEDADES.NOME_TIPO_PROP" & _
               " from TIPO_PROPRIEDADES inner join LINK_VALORES on LINK_VALORES.ID_TIPO_PROP = TIPO_PROPRIEDADES.ID_TIPO_PROP" & _
               " where LINK_VALORES.ID_ITEM = " & itemKey & ";" & _
               " union" & _
               " select NOME_TIPO_PROP FROM (TIPO_PROPRIEDADES inner join PROPRIEDADES_ITEMS on PROPRIEDADES_ITEMS.ID_TIPO_PROP = TIPO_PROPRIEDADES.ID_TIPO_PROP) " & _
               " where TIPO_PROPRIEDADES.ID_CLASSE_TIPO_PROP = " & calculatedPropClassKey & " and PROPRIEDADES_ITEMS.ID_TIPO_ITEM = " & itemTypeKey & ";"
    Else
        propClassKey = getPropClassKey(strPropClass)
        If propClassKey <> calculatedPropClassKey Then
            strSQL = "Select NOME_TIPO_PROP from (TIPO_PROPRIEDADES inner join PROPRIEDADES_ITEMS on TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP) " & _
                   " left join LINK_VALORES on PROPRIEDADES_ITEMS.ID_TIPO_PROP = LINK_VALORES.ID_TIPO_PROP" & _
                   " where TIPO_PROPRIEDADES.ID_CLASSE_TIPO_PROP = " & propClassKey & _
                   " and LINK_VALORES.ID_ITEM = " & itemKey
        Else
            strSQL = "Select NOME_TIPO_PROP from (TIPO_PROPRIEDADES inner join PROPRIEDADES_ITEMS on TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP) " & _
                   " left join LINK_VALORES on PROPRIEDADES_ITEMS.ID_TIPO_PROP = LINK_VALORES.ID_TIPO_PROP" & _
                   " where TIPO_PROPRIEDADES.ID_CLASSE_TIPO_PROP = " & propClassKey
        End If
    End If

    rs.Open strSQL, gCnn

    ' Adding the fields from the table into the combo
    On Error Resume Next
    rs.MoveFirst
    Do While rs.EOF <> True
        cmbPropName.AddItem rs.Fields("NOME_TIPO_PROP").value
        rs.MoveNext
    Loop
    On Error GoTo 0
End Sub

Public Sub populateSharedProps(strItemName As String, cmbPropName, Optional strPropClass As String)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim itemTypeKey                               As Long
    Dim propClassKey                              As Long
    Dim propKey                                   As Long
    Dim itemKey                                   As Long

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' Getting the keys
    itemKey = getItemKey(strItemName)
    itemTypeKey = getItemTypeFromItemKey(itemKey)

    ' POPULATING THE PROPERTY COMBOBOX

    ' Now, I query over all item's properties where item.itemtype = itemtypekey
    ' Getting the item type primary key
    If strPropClass = Empty Then
        strSQL = " SELECT PROPRIEDADES_ITEMS.ID_TIPO_ITEM, TIPO_PROPRIEDADES.NOME_TIPO_PROP, PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " FROM TIPO_PROPRIEDADES INNER JOIN PROPRIEDADES_ITEMS ON TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " WHERE (((PROPRIEDADES_ITEMS.ID_TIPO_ITEM)= " & itemTypeKey & "))" & _
               " ORDER BY TIPO_PROPRIEDADES.NOME_TIPO_PROP "
    Else
        propClassKey = getPropClassKey(strPropClass)
        strSQL = " SELECT PROPRIEDADES_ITEMS.ID_TIPO_ITEM, TIPO_PROPRIEDADES.NOME_TIPO_PROP, PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " FROM TIPO_PROPRIEDADES INNER JOIN PROPRIEDADES_ITEMS ON TIPO_PROPRIEDADES.ID_TIPO_PROP = PROPRIEDADES_ITEMS.ID_TIPO_PROP" & _
               " WHERE PROPRIEDADES_ITEMS.ID_TIPO_ITEM= " & itemTypeKey & " AND TIPO_PROPRIEDADES.ID_CLASSE_TIPO_PROP = " & propClassKey & _
               " ORDER BY TIPO_PROPRIEDADES.NOME_TIPO_PROP "
    End If
    rs.Open strSQL, gCnn

    ' Adding the fields from the table into the combo
    On Error Resume Next
    rs.MoveFirst
    Do While rs.EOF <> True
        propKey = rs.Fields("ID_TIPO_PROP").value
        If checkIfIsShared(itemKey, propKey) Then
            cmbPropName.AddItem rs.Fields("NOME_TIPO_PROP").value
        End If
        rs.MoveNext
    Loop
    On Error GoTo 0
End Sub

Public Sub populateTrackingProps(strItemTypeName As String, cmbTrackingPropName)
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim itemTypeKey                               As Long
    Dim colDictionary                             As collection
    Dim field                                     As Variant
    Dim propToAdd                                 As String

    On Error GoTo populateTrackingProps_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Getting itemTypeKey
    itemTypeKey = getItemTypeKey(strItemTypeName)

    ' For visual porpuses only, I create this collection to translate the field names
    Set colDictionary = New collection
    Set colDictionary = createTrackingDictionary()

    ' Now, I get the whole table of all the tracking properties of this item (The parent's general properties)
    strSQL = "select ITEM.NOME_ITEM, SUB_AREA.NOME_SUB_ARE, AREA.NOME_ARE, PLANTA.NOME_PLA, INDUSTRIAL.NOME_IND, UNIDADE_NEGOCIO.NOME_UNI " & _
             "from UNIDADE_NEGOCIO INNER JOIN ((((INDUSTRIAL INNER JOIN PLANTA ON INDUSTRIAL.ID_IND = PLANTA.ID_IND) INNER JOIN AREA ON " & _
             "PLANTA.ID_PLA = AREA.ID_PLA) INNER JOIN SUB_AREA ON AREA.ID_ARE = SUB_AREA.ID_ARE) INNER JOIN ITEM ON SUB_AREA.ID_SUB_ARE = ITEM.ID_SUB_ARE) ON UNIDADE_NEGOCIO.ID_UNI = INDUSTRIAL.ID_UNI " & _
             "where ITEM.ID_TIPO_ITEM = " & itemTypeKey

    ' Now, I open this query
    rs.Open strSQL, gCnn

    For Each field In rs.Fields
        propToAdd = colDictionary(field.Name)
        cmbTrackingPropName.AddItem (propToAdd)
    Next field
    On Error GoTo 0

    Exit Sub

populateTrackingProps_Error:
    Call handleMyError

End Sub

Public Sub populateSubArea(cmbSubArea)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    On Error GoTo populateSubArea_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Now, I query over all item's where item.itemtype = itemtypekey
    strSQL = "Select NOME_SUB_ARE from SUB_AREA"
    rs.Open strSQL, gCnn

    ' checking keyness
    Call checkKeyness(rs)

    Do While rs.EOF <> True
        cmbSubArea.AddItem rs.Fields("NOME_SUB_ARE").value
        rs.MoveNext
    Loop
    ' Voilá
    On Error GoTo 0

    Exit Sub

populateSubArea_Error:
    Call handleMyError
End Sub

Public Sub populatePropClass(cmbPropClass)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    On Error GoTo populatePropClass_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Now, I query over all item's where item.itemtype = itemtypekey
    strSQL = "Select NOME_CLASSE_TIPO_PROP from CLASSE_TIPO_PROP"
    rs.Open strSQL, gCnn
    ' Adding the Item names into the combo
    If rs.EOF = True Then
        MsgBox "You have no Sub Area's in the database!"
        Exit Sub
    End If

    Do While rs.EOF <> True
        cmbPropClass.AddItem rs.Fields("NOME_CLASSE_TIPO_PROP").value
        rs.MoveNext
    Loop
    ' Voilá
    On Error GoTo 0

    Exit Sub

populatePropClass_Error:
    Call handleMyError
End Sub

Public Sub populateUnits(strPropName As String, cmbUnits)
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim dimKey                                    As Long
    Dim propKey                                   As Long
    Dim tempUnit                                  As unit

    propKey = getPropKey(strPropName)
    dimKey = getDimKey(propKey)

    If dimKey <> 0 Then
        On Error GoTo populateUnits_Error

        For Each tempUnit In gUnitDef.pUnitCollection
            If tempUnit.dimension = dimKey Then
                cmbUnits.AddItem tempUnit.symbol
            End If
        Next tempUnit
    End If

    ' Voilá
    On Error GoTo 0

    Exit Sub

populateUnits_Error:
    Call handleMyError
End Sub

Public Sub populateSavedShares(cmbSavedShares, strItemType1 As String, strItemType2 As String)
    Dim itemTypeKey1                              As Long
    Dim itemTypeKey2                              As Long

    ' getting the type keys
    itemTypeKey1 = getItemTypeKey(strItemType1)
    itemTypeKey2 = getItemTypeKey(strItemType2)

    Call populateSavedSharesFromKeys(cmbSavedShares, itemTypeKey1, itemTypeKey2)
End Sub

Public Sub populateSavedSharesFromKeys(cmbSavedShares, itemTypeKey1 As Long, itemTypeKey2 As Long)
    ' getting the xml document
    Set objXml = New MSXML2.DOMDocument
    If False = objXml.Load(getSharesXmlPath()) Then
        MsgBox "Redefina a localização do xml de configuração dos compartilhamentos"
        storeSharesXmlPath (getFileDialog())
        objXml.Load (getSharesXmlPath())
    End If

    Set objNodeSet = objXml.getElementsByTagName("share")

    For j = 0 To objNodeSet.Length - 1
        If objNodeSet.Item(j).SelectSingleNode("typePair").SelectSingleNode("type1").Text = itemTypeKey1 And _
           objNodeSet.Item(j).SelectSingleNode("typePair").SelectSingleNode("type2").Text = itemTypeKey2 Then
            cmbSavedShares.AddItem objNodeSet.Item(j).SelectSingleNode("name").Text
        End If
    Next j
End Sub

Public Sub addToShareList(shareList As ListBox, strItemName1 As String, strItemName2 As String, _
                          strPropName1 As String, strPropName2 As String)
    ' Adding to the list.
    shareList.AddItem strItemName1
    shareList.List(shareList.ListCount - 1) = strItemName2
    shareList.AddItem strPropName1
    shareList.List(shareList.ListCount - 1) = strPropName2
    shareList.AddItem "----------------"
    shareList.List(shareList.ListCount - 1) = "---------------"
End Sub

Public Sub removeFromShareList(shareList As ListBox, indexClicked As Integer)
    For j = 1 To 3
        shareList.RemoveItem (3 * indexClicked)
    Next j
End Sub

Public Function loadShareCollection(strShareName As String, strItemName1 As String, strItemName2 As String) As shareQueue
    Dim objXml                                    As MSXML2.DOMDocument
    Dim objXmlShareNode                           As MSXML2.IXMLDOMNode
    Dim objXmlNodes                               As MSXML2.IXMLDOMNodeList
    Dim propKey1                                  As Long
    Dim propKey2                                  As Long
    Dim strPropName1                              As String
    Dim strPropName2                              As String
    Dim j                                         As Integer
    Dim tempQueue                                 As shareQueue

    Set tempQueue = New shareQueue

    Set objXml = New MSXML2.DOMDocument
    If False = objXml.Load(getSharesXmlPath) Then
        MsgBox "Redefina a localização do xml de configuração dos compartilhamentos"
        storeSharesXmlPath (getFileDialog())
        objXml.Load (getSharesXmlPath())
    End If

    Set objXmlShareNode = objXml.SelectSingleNode("//share[name=" & Chr(34) & strShareName & Chr(34) & "]/propPairSet")
    Set objXmlNodes = objXmlShareNode.SelectNodes("propPair")

    For j = 0 To objXmlNodes.Length - 1
        propKey1 = objXmlNodes.Item(j).SelectSingleNode("prop1").Text
        propKey2 = objXmlNodes.Item(j).SelectSingleNode("prop2").Text

        strPropName1 = getPropName(propKey1)
        strPropName2 = getPropName(propKey2)

        Call tempQueue.enqueue(strItemName1, strPropName1, strItemName2, strPropName2)
    Next j

    Set loadShareCollection = tempQueue

End Function

Public Sub populateAllProps(ltbPropType)
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strSQL = "select NOME_TIPO_PROP from TIPO_PROPRIEDADES order by NOME_TIPO_PROP"

    rs.Open strSQL, gCnn

    Do While rs.EOF <> True
        ltbPropType.AddItem rs.Fields("NOME_TIPO_PROP").value
        rs.MoveNext
    Loop
End Sub

Public Sub populateDimension(cmbDimension As Variant)
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strSQL = "select NOME_DIMENSAO from DIMENSOES"

    rs.Open strSQL, gCnn

    Do While rs.EOF <> True
        cmbDimension.AddItem rs.Fields("NOME_DIMENSAO").value
        rs.MoveNext
    Loop
End Sub

Public Sub deleteXrefs(itemTypeKey As Long)
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset
    strSQL = "delete from PROPRIEDADES_ITEMS where ID_TIPO_ITEM = " & itemTypeKey & ";"
    gCnn.Execute strSQL
End Sub

Public Sub createXref(itemTypeKey As Long, propKey As Long)
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset
    strSQL = "insert into [CHT-CPNM].[dbo].[PROPRIEDADES_ITEMS](ID_TIPO_ITEM, ID_TIPO_PROP) values(" & itemTypeKey & "," & propKey & ");"
    gCnn.Execute strSQL
End Sub

Public Function GetAllProperties(itemKey As Long)
    Dim rs As ADODB.Recordset
    Dim strSQL As String
        Dim itemTypeKey As Long
        
    Set rs = New ADODB.Record
    itemTypeKey = getItemTypeFromItemKey(itemKey)
    strSQL = "select ID_TIPO_PROP from PROPRIEDADES_ITEMS where ID_TIPO_ITEM = " & itemTypeKey
    rs.Open strSQL, gCnn
    
    Set GetAllProperties = rs
End Function
