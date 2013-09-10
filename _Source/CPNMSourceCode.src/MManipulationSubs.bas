Attribute VB_Name = "MManipulationSubs"
' Chemtech - A Siemens Business ========================================================
'
'=======================================================================================
Option Explicit
' Desenvolvimento ======================================================================
' <iniciais>            Renan       <email>
'=======================================================================================
' Versões ==============================================================================
'
'
'
'=======================================================================================

Public Sub exportSingleData(strItemName As String, strPropName As String, _
                            strInsertValue As String, strInsertUnit As String, Optional bolAskForConf As Boolean = False)

' Declarations
    Dim itemKey                    As Long
    Dim propKey                    As Long
    Dim unitKey                    As Long

    itemKey = getItemKey(strItemName)
    propKey = getPropKey(strPropName)
    unitKey = getUnitKey(strPropName, strInsertUnit)

    Call exportSingleDataFromKeys(itemKey, propKey, strInsertValue, unitKey, bolAskForConf)
End Sub

Public Sub exportSingleDataFromKeys(itemKey As Long, propKey As Long, _
                                    strInsertValue As String, unitKey As Long, Optional bolAskForConf As Boolean = False)
' Declarations
    Dim rs                         As ADODB.Recordset
    Dim varResponse                As Variant
    Dim valueExistance             As Boolean
    Dim strCommand                 As String
    Dim idValueJustInserted        As Long
    Dim strConfirmString           As String
    Dim valueKey                   As Long

    On Error GoTo exportSingleData_Error

    ' And i create the recordset to receive some queries
    Set rs = New ADODB.Recordset

    ' Now I trim the value
    strInsertValue = Trim(strInsertValue)

    If unitKey <> 0 Then                                             'Will convert the value ONLY if there is a unit to convert to.
        strInsertValue = Format(CDbl(strInsertValue), "Scientific")
        strInsertValue = gUnitDef.convertValue(strInsertValue, propKey, unitKey, False)
    End If

    ' Now, we gotta check if the value is already in the database. If it is already there, we have to take a different course of action
    valueExistance = checkValueExistance(itemKey, propKey)

    If valueExistance = False Then
        'Now, we are free to insert the value into the VLUE table
        strCommand = "insert into [CHT-CPNM].[dbo].[VALOR_PROPRIEDADES](VALOR_PROP)" & _
                     " values('" & strInsertValue & "')"

        ' Executing the thing
        gCnn.Execute (strCommand)

        ' Getting the just inserted value ID
        rs.Open "SELECT @@Identity AS ID", gCnn
        idValueJustInserted = rs.Fields("ID").value

        ' Creating the sql command to INSERT the registry into the LINK table
        strCommand = "insert into [CHT-CPNM].[dbo].[LINK_VALORES](ID_ITEM,ID_TIPO_PROP,ID_VALOR) values('" & itemKey & "','" & propKey & "','" & idValueJustInserted & "')"
        ' Executing the thing =)
        Debug.Print strCommand
        gCnn.Execute (strCommand)
    Else

        strConfirmString = "Property " & getPropName(propKey) & " of the item " & getItemName(itemKey) & " already exists" & Chr(10) & _
                           "Are you sure you want to update the old value?"

        ' Checking if the user asked for update confirmation
        If bolAskForConf Then
            varResponse = MsgBox(strConfirmString, vbYesNo, "Confirmation")
        Else
            varResponse = vbYes
        End If

        ' If the update was confirmed, run the things down
        If varResponse = vbYes Then
            ' Getting the Value ID that corresponds to our property-item pair.
            valueKey = getValueKey(itemKey, propKey)

            ' Creating the sql command to UPDATE the registry
            strCommand = "update VALOR_PROPRIEDADES " & _
                         " set VALOR_PROP = '" & strInsertValue & "'" & _
                         " where ID_VALOR = " & valueKey
            ' Executing the thing =)
            gCnn.Execute (strCommand)
        End If
    End If
    On Error GoTo 0

    Exit Sub

exportSingleData_Error:
    Call handleMyError
End Sub

Public Sub changeTrackingData(strItemName, strPropNameTrack, strNewValue)
' Declarations
    Dim rs                         As ADODB.Recordset
    Dim strSQL                     As String
    Dim colDictionary              As collection
    Dim itemKey                    As Long
    Dim propTrackKey               As String
    Dim bigSelect                  As String


    On Error GoTo changeTrackingData_Error

    ' And i create the recordset to receive some queries
    Set rs = New ADODB.Recordset

    ' getting the dictionary
    Set colDictionary = createTrackingDictionary()

    ' getting the keys
    itemKey = getItemKey(strItemName)
    propTrackKey = colDictionary(strPropNameTrack)

    ' Now, this is the MONSTER QUERY
    bigSelect = "select ITEM.NOME_ITEM, SUB_AREA.NOME_SUB_ARE, AREA.NOME_ARE, PLANTA.NOME_PLA, INDUSTRIAL.NOME_IND, UNIDADE_NEGOCIO.NOME_UNI " & _
                " from UNIDADE_NEGOCIO INNER JOIN ((((INDUSTRIAL INNER JOIN PLANTA ON INDUSTRIAL.ID_IND = PLANTA.ID_IND) INNER JOIN AREA ON " & _
                " PLANTA.ID_PLA = AREA.ID_PLA) INNER JOIN SUB_AREA ON AREA.ID_ARE = SUB_AREA.ID_ARE) INNER JOIN ITEM ON SUB_AREA.ID_SUB_ARE = ITEM.ID_SUB_ARE) ON UNIDADE_NEGOCIO.ID_UNI = INDUSTRIAL.ID_UNI " & _
                " where ITEM.ID_ITEM = " & itemKey

    ' The update command
    strSQL = "update (" & bigSelect & ")" & _
             " set " & propTrackKey & "= '" & strNewValue & "'"

    ' Executing the update command my friend. We are done.
    gCnn.Execute strSQL

    'closing the connection
    On Error GoTo 0

    Exit Sub

changeTrackingData_Error:
    Call handleMyError
End Sub

Public Sub createItem(strItemName As String, strItemType As String, strSubArea As String)
' Declarations
    Dim subAreaKey                 As String
    Dim itemTypeKey                As Long
    Dim itemExistance              As Boolean
    Dim strCommand                 As String
    Dim strConfirmString           As String

    ' Checking if the item already exists
    On Error GoTo createItem_Error
    If checkItemExistance(strItemName) = True Then
        MsgBox "Já existe um Item com este nome cadastrado!"
        Exit Sub
    End If

    ' Here, i get the primary keys
    subAreaKey = getSubAreaKey(strSubArea)
    itemTypeKey = getItemTypeKey(strItemType)
    ' Here i check if the item already exists
    itemExistance = checkItemExistance(strItemName)
    If itemExistance = False Then
        ' Creating the sql command to INSERT the registry
        strCommand = "insert into ITEM(NOME_ITEM,ID_TIPO_ITEM,ID_SUB_ARE, ATIVO) values('" & strItemName & "'," & itemTypeKey & "," & subAreaKey & "," & -1 & ")"
        ' Executing the thing =)
        gCnn.Execute (strCommand)
    Else
        strConfirmString = "Item " & strItemName & "already exists"
    End If
    On Error GoTo 0

    Exit Sub

createItem_Error:
    Call handleMyError

End Sub

Sub createSharing(strItemName1 As String, strPropName1 As String, strItemName2 As String, strPropName2 As String)
' Essa rotina é um wrapper

    Dim itemKey1                   As Long
    Dim itemKey2                   As Long
    Dim propKey1                   As Long
    Dim propKey2                   As Long

    itemKey1 = getItemKey(strItemName1)
    propKey1 = getPropKey(strPropName1)
    itemKey2 = getItemKey(strItemName2)
    propKey2 = getPropKey(strPropName2)

    Call createSharingFromKeys(itemKey1, propKey1, itemKey2, propKey2)
End Sub

Sub createSharingFromKeys(itemKey1 As Long, propKey1 As Long, itemKey2 As Long, propKey2 As Long)
'=======================================================================================
' Essa rotina cria o compartilhamento a partir dos nomes dos items e propriedades
'---------------------------------------------------------------------------------------
' [strItemName1] - Nome do primeiro item que entra no compartilhamento
' [strPropName1] - Nome da propriedade do primeiro item para compartilhamento.
' [strItemName2] - Nome do segunda item que entra no compartilhamento
' [strPropName2] - Nome da propriedade do segundo item para compartilhamento.
'---------------------------------------------------------------------------------------
    Dim rs                         As New ADODB.Recordset
    Dim strSQL                     As String
    Dim value1Exists               As Boolean
    Dim value2Exists               As Boolean
    Dim valueKey                   As Long
    Dim valueKey4deletion          As Long
    Dim itemKey4creation           As Long
    Dim propKey4creation           As Long
    Dim itemKey4change             As Long
    Dim propKey4change             As Long
    Dim valueKey1                  As Long
    Dim valueKey2                  As Long
    Dim optionLeft                 As Variant

    ' And i create the recordset to receive the query
    On Error GoTo createSharingFromKeys_Error
    Set rs = New ADODB.Recordset

    ' Now, we need to get the values.
    value1Exists = checkValueExistance(itemKey1, propKey1)
    value2Exists = checkValueExistance(itemKey2, propKey2)

    ' setting the keys to -1
    valueKey4deletion = -1
    itemKey4creation = -1
    propKey4creation = -1
    itemKey4change = -1
    propKey4change = -1

    ' Asserting that both properties to be shared have the same dimension
    Call assertDimensionallity(propKey1, propKey2)

    ' Now the cases are going to start
    ' both 1 and 2 dont exist
    If Not value1Exists And Not value2Exists Then                    'Here I add a dummy value for the first pair.
        MsgBox "None of the chosen values exists." & vbCr & "I will now create the share with a dummy value"

        Call exportSingleDataFromKeys(itemKey1, propKey1, "Aguardando Upload.", 0, False)
        value1Exists = checkValueExistance(itemKey1, propKey1)       'With this, the system will progress accordingly.
        Exit Sub
    End If

    ' 1 exists, 2 dont
    If value1Exists And Not value2Exists Then
        'now i get the value1 key
        valueKey = getValueKey(itemKey1, propKey1)
        'this is the entry i will link to value1
        itemKey4creation = itemKey2
        propKey4creation = propKey2
    End If

    ' 2 exists, 2 dont
    If Not value1Exists And value2Exists Then
        'now i get the value2 key
        valueKey = getValueKey(itemKey2, propKey2)
        'this is the entry i will link to value2
        itemKey4creation = itemKey1
        propKey4creation = propKey1
    End If

    ' both values exist
    If value1Exists And value2Exists Then

        'now i get the value keys.
        valueKey1 = getValueKey(itemKey1, propKey1)
        valueKey2 = getValueKey(itemKey2, propKey2)

        If valueKey1 <> valueKey2 Then
            ' now, i will delete the second value
            optionLeft = MsgBox("Both values exist! Do you want to keep Value 1?", vbYesNo)

            If optionLeft = vbYes Then
                valueKey = valueKey1                                 'this defines the final value key
                valueKey4deletion = getValueKey(itemKey2, propKey2)

                itemKey4change = itemKey2                            ' and the link to be changed is the deleted one.
                propKey4change = propKey2
            Else
                valueKey = valueKey2                                 'this defines the final value key
                valueKey4deletion = getValueKey(itemKey1, propKey1)  'the other value will be deleted

                itemKey4change = itemKey1                            ' and the link to be changed is the deleted one.
                propKey4change = propKey1
            End If

        Else
            'if the keys are equal, it makes no difference which to choose.
            valueKey = valueKey1
            itemKey4change = itemKey2
            propKey4change = propKey2

        End If
    End If

    ' executing the UPDATE statement
    If itemKey4change <> -1 Then
        strSQL = " update LINK_VALORES " & _
                 " set ID_VALOR = " & valueKey & _
                 " where ID_ITEM = " & itemKey4change & " AND ID_TIPO_PROP = " & propKey4change

        gCnn.Execute (strSQL)
    End If

    ' executing the INSERTION statement
    If itemKey4creation <> -1 Then
        strSQL = " insert into LINK_VALORES(ID_ITEM,ID_TIPO_PROP,ID_VALOR) " & _
                 " values (" & itemKey4creation & "," & propKey4creation & "," & valueKey & ")"

        gCnn.Execute (strSQL)
    End If

    ' executing the DELETE statement for the old value
    If valueKey4deletion <> -1 Then
        strSQL = "delete from VALOR_PROPRIEDADES where ID_VALOR = " & valueKey4deletion
        gCnn.Execute (strSQL)
    End If

    On Error GoTo 0

    GoTo createSharingFromKeys_Finally

createSharingFromKeys_Finally:

    Exit Sub

    ' Procedure Error Handler
createSharingFromKeys_Error:
    Dim errorAction                As Integer
    'here goes your specific error handling code.

    ' here comes the generic global error handling code.
    errorAction = handleMyError()
    Select Case errorAction
    Case -1
        Stop
    Case 1
        Resume Next
    Case 2
        GoTo createSharingFromKeys_Finally
    Case Else
        Stop
    End Select
End Sub

Sub breakSharing(strItemName1 As String, strPropName1 As String, strItemName2 As String, strPropName2 As String)
' Declarations
    Dim rs                         As New ADODB.Recordset
    Dim strSQL                     As String
    Dim itemKey1                   As Long
    Dim itemKey2                   As Long
    Dim propKey1                   As Long
    Dim propKey2                   As Long
    Dim valueKey                   As Long
    Dim valueShared                As String
    Dim idValueJustInserted        As Long

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    itemKey1 = getItemKey(strItemName1)
    propKey1 = getPropKey(strPropName1)
    itemKey2 = getItemKey(strItemName2)
    propKey2 = getPropKey(strPropName2)
    valueKey = getValueKey(itemKey1, propKey1)
    valueShared = getValue(valueKey)

    strSQL = "insert into VALOR_PROPRIEDADES (VALOR_PROP)" & _
             " values ('" & valueShared & "')"

    ' executing the command
    gCnn.Execute (strSQL)

    ' Now I need to get the key of the new inserted values
    rs.Open "SELECT @@Identity AS ID", gCnn
    idValueJustInserted = rs.Fields("ID").value

    ' now, we have to change the break the old references.
    strSQL = "update LINK_VALORES " & _
             " set ID_VALOR = " & idValueJustInserted & _
             " where ID_ITEM = " & itemKey2 & " AND ID_TIPO_PROP = " & propKey2

    ' and now I execute the new value
    gCnn.Execute (strSQL)

    If Err Then
        MsgBox "Houve erro na quebra do compartilhamento"
    End If
End Sub

Public Sub assertDimensionallity(propKey1 As Long, propKey2 As Long)
    Dim dimKey1                    As Long
    Dim dimKey2                    As Long

    dimKey1 = getDimKey(propKey1)
    dimKey2 = getDimKey(propKey2)

    If dimKey1 <> dimKey2 Then
        Err.Raise 175645 + vbObjectError, Description:="You are trying to share values which do not agree dimensionally"
    End If
End Sub

Public Sub createProperty(strPropName As String, strPropClassName As String, strDimName As String, Optional isCalc As Integer = 0)
    Dim rs                         As ADODB.Recordset
    Dim strSQL                     As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    If strPropName <> "" And strPropClassName <> "" Then
        If checkPropertyExistance(strPropName) = False Then
            strSQL = "INSERT INTO [CHT-CPNM].[dbo].[TIPO_PROPRIEDADES](NOME_TIPO_PROP, ID_DIMENSAO, ID_CLASSE_TIPO_PROP, PROP_CALCULADA) VALUES ('" & strPropName & "'," & getDimKeyFromDimName(strDimName) & _
                     "," & getPropClassKey(strPropClassName) & "," & isCalc & ");"
        Else
            strSQL = "UPDATE [CHT-CPNM].[dbo].[TIPO_PROPRIEDADES] SET ID_DIMENSAO = " & getDimKeyFromDimName(strDimName) & _
                     ", ID_CLASSE_TIPO_PROP = " & getPropClassKey(strPropClassName) & ", PROP_CALCULADA = " & isCalc & _
                     " WHERE NOME_TIPO_PROP = '" & strPropName & "';"
        End If
    End If

    gCnn.Execute strSQL
End Sub

Public Sub createItemType(strItemTypeName As String, strItemTypeClassName As String)
    Dim rs                         As ADODB.Recordset
    Dim strSQL                     As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    If strItemTypeName <> "" Then
        strSQL = "INSERT INTO [CHT-CPNM].[dbo].[TIPO_ITEM](NOME_TIPO_ITEM, ID_CLASSE_TIPO_ITEM) values ('" & strItemTypeName & "'," & getItemClassKey(strItemTypeClassName) & ")"
        gCnn.Execute strSQL
    End If
End Sub

'##############
' Changing Subs
'##############

Public Sub changeItemTypeByKey(itemKey As Long, newItemTypeKey As String)
    Dim strSQL                     As String

    strSQL = "UPDATE ITEM SET ID_TIPO_ITEM = " & newItemTypeKey & " where ID_ITEM = " & itemKey & ";"

    gCnn.Execute strSQL
End Sub

Public Sub changeItemType(itemName As String, newItemTypeName As String)
    Dim newItemTypeKey             As Long
    Dim itemKey                    As Long

    If newItemTypeName <> "" And itemName <> "" Then
        itemKey = getItemKey(itemName)
        newItemTypeKey = getItemTypeKey(newItemTypeName)

        Call changeItemTypeByKey(itemKey, newItemTypeKey)
    End If
End Sub

Public Sub changeItemNameByKey(itemKey As Long, newItemName As String)
    Dim strSQL                     As String

    If newItemName <> "" Then
        strSQL = "UPDATE ITEM set NOME_ITEM = '" & newItemName & "' where ID_ITEM = " & itemKey & ";"

        gCnn.Execute strSQL
    End If
End Sub

Public Sub changeItemName(itemName As String, newItemName As String)
    Dim itemKey                    As Long
    itemKey = getItemKey(itemName)
    Call changeItemNameByKey(itemKey, newItemName)
End Sub

Public Sub changeItemActiveStatusByKey(itemKey As Long, IsActive As Integer)
    Dim activeStatus               As Integer
    Dim strSQL                     As String

    strSQL = "UPDATE ITEM SET ATIVO = " & IsActive & " where ID_ITEM = " & itemKey & ";"

    gCnn.Execute strSQL
End Sub

Public Sub changeItemActiveStatus(itemName As String, IsActive As Integer)
    Dim itemKey                    As Long
    itemKey = getItemKey(itemName)
    Call changeItemActiveStatusByKey(itemKey, IsActive)
End Sub
