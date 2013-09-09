Attribute VB_Name = "MQueries"
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

Public Function getItemTypeKey(strItemTypeName) As Long
    ' The following code populates the comboboxes with item's names and prop's names

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Getting the item type primary key from query
    strSQL = "Select ID_TIPO_ITEM from TIPO_ITEM where NOME_TIPO_ITEM = '" & strItemTypeName & "'"
    rs.Open strSQL, gCnn

    ' Checking if something weird just happenned
    Call checkKeyness(rs)

    ' Checking if there is a problem
    getItemTypeKey = rs.Fields("ID_TIPO_ITEM").value

End Function

Public Function getItemKey(strItemName) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_ITEM from ITEM where NOME_ITEM = '" & strItemName & "'"

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the keys.
    getItemKey = rs.Fields("ID_ITEM")

End Function

Public Function getItemTypeFromItemKey(itemKey As Long) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_TIPO_ITEM from ITEM where ID_ITEM = " & itemKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the keys.
    getItemTypeFromItemKey = rs.Fields("ID_TIPO_ITEM").value

End Function

Public Function getPropKey(strPropName As String) As Long
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_TIPO_PROP from TIPO_PROPRIEDADES where NOME_TIPO_PROP = '" & strPropName & "'"

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the key.
    getPropKey = rs.Fields("ID_TIPO_PROP")

End Function

Public Function getDimKey(propKey As Long) As Long
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset

    strSQL = "select ID_DIMENSAO, NOME_TIPO_PROP from TIPO_PROPRIEDADES where ID_TIPO_PROP = " & propKey

    rs.Open strSQL, gCnn

    If Not IsNull(rs.Fields("ID_DIMENSAO").value) Then
        getDimKey = rs.Fields("ID_DIMENSAO").value
    Else
        getDimKey = 0
    End If
End Function

Public Function getDimKeyFromDimName(dimName As String) As Long
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    Set rs = New ADODB.Recordset

    strSQL = "select ID_DIMENSAO from [CHT-CPNM].[dbo].[DIMENSOES] where NOME_DIMENSAO = " & "'" & dimName & "'"

    rs.Open strSQL, gCnn

    If Not rs.EOF Then
        If Not IsNull(rs.Fields("ID_DIMENSAO").value) Then
            getDimKeyFromDimName = rs.Fields("ID_DIMENSAO").value
        Else
            getDimKeyFromDimName = 0
        End If
    End If
End Function

Public Function getSubAreaKey(strSubAreaName As String) As Long
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_SUB_ARE from SUB_AREA where NOME_SUB_ARE = '" & strSubAreaName & "'"

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the key.
    getSubAreaKey = rs.Fields("ID_SUB_ARE").value

End Function

Public Function getSubAreaKeyFromItemKey(itemKey As Long) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_SUB_ARE from ITEM where ID_ITEM = " & itemKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the key.
    getSubAreaKeyFromItemKey = rs.Fields("ID_SUB_ARE").value

End Function

Public Function getPropClassKey(strPropClass As String) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_CLASSE_TIPO_PROP from [CHT-CPNM].[dbo].[CLASSE_TIPO_PROP] where NOME_CLASSE_TIPO_PROP = '" & strPropClass & "'"

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the key.
    getPropClassKey = rs.Fields("ID_CLASSE_TIPO_PROP").value

End Function

Public Function getItemName(itemKey As Long) As String
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select NOME_ITEM from ITEM where ID_ITEM = " & itemKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the name.
    getItemName = rs.Fields("NOME_ITEM")

End Function

Public Function getItemClassKeyFromItemKey(itemKey As Long) As String
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim itemTypeKey                               As Long

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    itemTypeKey = getItemTypeFromItemKey(itemKey)

    ' Query string for the value
    strSQL = "select ID_CLASSE_TIPO_ITEM from TIPO_ITEM where ID_TIPO_ITEM = " & itemTypeKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the name.
    getItemClassKeyFromItemKey = rs.Fields("ID_CLASSE_TIPO_ITEM")

End Function

Public Function getItemClassKey(strItemClassName As String) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select ID_CLASSE_TIPO_ITEM from [CHT-CPNM].[dbo].[CLASSE_TIPO_ITEM] where NOME_CLASSE_TIPO_ITEM = '" & strItemClassName & "'"

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the key.
    getItemClassKey = rs.Fields("ID_CLASSE_TIPO_ITEM").value

End Function

Public Function getValueKey(itemKey, propKey) As Long
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' Aqui entra o comando em SQL para fazer a query
    strSQL = " SELECT ID_VALOR " & _
           " FROM LINK_VALORES WHERE ID_ITEM = " & itemKey & " And ID_TIPO_PROP = " & propKey

    ' Criando um recordset com o resultado da query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, gCnn

    ' Call checkKeyness(rs)

    ' Returning the value
    If rs.EOF <> True Then
        getValueKey = rs.Fields("ID_VALOR").value
    Else
        getValueKey = 0
    End If
End Function

Public Function getPropName(propKey) As String
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select NOME_TIPO_PROP from TIPO_PROPRIEDADES where ID_TIPO_PROP = " & propKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the name.
    getPropName = rs.Fields("NOME_TIPO_PROP")

End Function

Public Function getShareName(valueKey As Long, itemKeyComp As Long, _
                             itemKeyOrig As Long, propKeyOrig As Long) As String
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' Query string for the value
    strSQL = "select NOME_TIPO_PROP from TIPO_PROPRIEDADES inner join LINK_VALORES on TIPO_PROPRIEDADES.ID_TIPO_PROP = LINK_VALORES.ID_TIPO_PROP" & _
           " where LINK_VALORES.ID_VALOR = " & valueKey & " and LINK_VALORES.ID_ITEM = " & itemKeyComp & _
           " and NOT(LINK_VALORES.ID_TIPO_PROP = " & propKeyOrig & " and LINK_VALORES.ID_ITEM = " & itemKeyOrig & ")"
    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' getting to know if the query returned stuff
    Call checkKeyness(rs)

    ' If register is found, get the name.
    getShareName = rs.Fields("NOME_TIPO_PROP")
End Function

Public Function getUnitKey(strPropName As String, strUnitSymbol As String) As Long
    '=======================================================================================
    ' [strPropName] nome da propriedade, conforme definido no database
    ' [strUnitName] nome/simbolo da unidade. E.g. Pa, kgf/cm², °C.
    '---------------------------------------------------------------------------------------
    ' Essa rotina confia na implementação das classes "unit" e "unitDefinition".
    ' Deve existir uma instancia global gUnitDef de unitDefinition que define todas as units.
    ' A classe "unitDefinition" depende de um arquivo XML para a definição dos objetos unit.
    '=======================================================================================
    '
    getUnitKey = gUnitDef.getUnitKey(strPropName, strUnitSymbol)
End Function

Public Function getUnitSymbol(propKey As Long, unitKey As Long) As Long
    '=======================================================================================
    ' Retorna o simbolo da unidade a partir da chave da propriedade e chave da unidade
    ' Esta função é um wrapper.
    '---------------------------------------------------------------------------------------
    ' [propKey] chave da propriedade conforme definida no database
    ' [unitKey] chave da unidade conforme definida no XML
    '---------------------------------------------------------------------------------------
    ' Essa rotina confia na implementação das classes "unit" e "unitDefinition".
    ' Deve existir uma instancia global gUnitDef de unitDefinition que define todas as units.
    ' A classe "unitDefinition" depende de um arquivo XML para a definição dos objetos unit.
    '=======================================================================================
    '
    getUnitSymbol = gUnitDef.getUnitSymbol(propKey, unitKey)
End Function

Public Function getValue(valueKey) As String
    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Query string for the value
    strSQL = "select VALOR_PROP from VALOR_PROPRIEDADES where ID_VALOR = " & valueKey

    ' getting the query into rs
    rs.Open strSQL, gCnn

    ' If register is found, get the name.
    If rs.EOF <> True Then
        getValue = rs.Fields("VALOR_PROP")
    Else
        getValue = nonExistantValueString
    End If
End Function

Public Function checkItemExistance(strItemName As String) As Boolean
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String

    ' now I check if the value already exists
    strSQL = "select ID_ITEM from ITEM " & _
           " where NOME_ITEM = '" & strItemName & "'"

    ' now I open this check query;
    Set rs = New ADODB.Recordset
    rs.Open strSQL, gCnn

    ' If the query is empty, means the value is not yet there.
    If rs.EOF Then
        checkItemExistance = False
    Else
        checkItemExistance = True
    End If

End Function

Public Function checkValueExistance(itemKey As Long, propKey As Long) As Boolean
    Dim valueKey                                  As Long

    valueKey = getValueKey(itemKey, propKey)

    If valueKey <> 0 Then
        checkValueExistance = True
    Else
        checkValueExistance = False
    End If
End Function

Public Function checkIfIsCalc(propKey As Long) As Boolean

    Dim strSQL                                    As String
    Dim rs                                        As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' agora eu preciso verificar se a propriedade é comum, ou é calculada.
    strSQL = "select ID_TIPO_PROP from TIPO_PROPRIEDADES where PROP_CALCULADA = -1 AND ID_TIPO_PROP = " & CLng(propKey)
    rs.Open strSQL, gCnn
    Select Case rs.RecordCount
        Case 1
            checkIfIsCalc = True
        Case 0
            checkIfIsCalc = False
        Case Else
            Err.Raise vbObjectError + 112525, Description:="OMFG YOU GOT A KEYNESS PROBLEM!!!"
    End Select
End Function

Public Function checkIfIsShared(itemKey As Long, propKey As Long)
    ' The following code populates the item type combobox with item_type's name

    ' Declarations
    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim valueKey                                  As Long

    On Error GoTo checkIfIsShared_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' getting the valueKey of the pair
    valueKey = getValueKey(itemKey, propKey)

    ' Here i put the command for the query
    strSQL = "select ID_ITEM from LINK_VALORES where ID_VALOR = " & valueKey

    ' Here i open the query in the recordset rs
    rs.Open strSQL, gCnn

    ' Returning the result
    If rs.RecordCount > 1 Then
        checkIfIsShared = True
    Else
        checkIfIsShared = False
    End If

    On Error GoTo 0

    Exit Function

checkIfIsShared_Error:
    'getValueKey will raise this error if no value is found. Which means the value does not exist, so, can't be shared.
    If Err.Number = 100000 + vbObjectError Then
        Resume Next
    Else
        Call handleMyError
    End If
End Function


Public Sub populateAllItems(cmbItemName)
    Dim rs                                        As ADODB.Recordset
    Dim strSQL As String

    Set rs = New ADODB.Recordset

    strSQL = "select NOME_ITEM from ITEM"

    rs.Open strSQL, gCnn

    Do While rs.EOF <> True
        cmbItemName.AddItem rs.Fields("NOME_ITEM").value
        rs.MoveNext
    Loop
End Sub

Public Function checkPropertyExistance(strPropName) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    strSQL = "select * from TIPO_PROPRIEDADES where NOME_TIPO_PROP = '" & strPropName & "'"

    rs.Open strSQL, gCnn
    
    If rs.EOF Then
        checkPropertyExistance = False
    Else
        checkPropertyExistance = True
    End If
End Function
