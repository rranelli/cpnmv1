Attribute VB_Name = "MReadData"
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

Public Function getData(strAddress As String, Optional showMsgBoxes As Boolean = False) As String
    '=======================================================================================
    ' Esta é a rotina que recebe a string de endereço e aciona a rotina correta para o carre
    ' gamento dos dados.
    '---------------------------------------------------------------------------------------
    ' [showMsgBoxes] booleano que indica o desejo de apresentar msgboxes com informações
    ' [strAddress] string na forma "itemKey|propKey"
    ' Se o propKey começa com Z, trata-se de um endereço de rastreamento.
    ' Se o propKey começa com D, trata-se de uma referência com valor e unidade
    ' Se o propKey começa com U, trata-se de uma referência de unidade apenas
    ' Se o propKey começa com V, trata-se de uma referência de valor apenas
    '---------------------------------------------------------------------------------------
    ' As propriedades que são calculadas aparecem na classe de propriedades 5.
    ' Referências de rastreamento tem a propKey começando com o caractere "Z" e neste caso,
    ' a propKey é uma string com o nome do campo de rastreamento (e.g. Nome_item)
    '---------------------------------------------------------------------------------------
    ' < Histórico de revisões>
    '=======================================================================================
    '
    ' This function decides if your data is a normal one or a tracking one
    Dim Namez                                     As Variant
    Dim propKey                                   As Variant         'here the property primary key may assume the value of a trackingPropKey, which is a string
    Dim itemKey                                   As Long
    Dim unitKey                                   As Long
    Dim rs                                        As ADODB.Recordset

    ' Aqui vc pega as chaves de equipamento e valor da propriedade
    Namez = split(strAddress, breakString)
    itemKey = CLng(Namez(1))
    propKey = Namez(2)
    On Error Resume Next
    unitKey = Namez(3)                                               'ignoring the possible error for a tracking value
    On Error GoTo 0

    ' Transformando propKey em string
    Namez = CStr(propKey)
    propKey = Right(Namez, Len(Namez) - 1)                           'Preparando a chave de propriedade

    ' aqui iniciamos o recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' Agora, você tem de checar o tipo de referência, e puxar o resultado com a rotina/opção correta;
    Select Case Left(Namez, 1)
        Case trackingChar
            getData = getTrackingDataFromDB(itemKey, CStr(propKey), showMsgBoxes)

        Case unitOnlyChar
            getData = getDataFromDB(itemKey, CLng(propKey), unitKey, 0, showMsgBoxes)

        Case valueOnlyChar
            getData = getDataFromDB(itemKey, CLng(propKey), unitKey, 1, showMsgBoxes)

        Case unitAndValueChar
            getData = getDataFromDB(itemKey, CLng(propKey), unitKey, 2, showMsgBoxes)

        Case calcChar
            getData = getCalculatedDataFromDB(itemKey, CLng(propKey))

        Case Else
            Err.Raise vbObjectError + 440040, Description:="Não há especificação correta do endereço :" & strAddress
    End Select
End Function

Public Function getDataFromDB(itemKey As Long, propKey As Long, unitKey As Long, refOption As Integer _
     , Optional showMsgBoxes As Boolean = True)
    ' Esta função vai extrair um valor do banco e retornar no formato apresentável.
    Dim strSQL                                    As String
    Dim rs                                        As ADODB.Recordset
    Dim varValueFromDatabase                      As String
    Dim errorAction                               As Integer
    Dim tempValue                                 As String
    Dim tempUnit                                  As String
    Dim valueKey                                  As Long
    Dim placeHolderString                         As String

    ' Criando um recordset para receber a query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    ' Aqui eu vou rodar uma query que mostra se o valor que eu quero buscar é compartilhado com algum valor calculado
    ' Caso seja, vou atualizar este valor calculado.
    If checkValueExistance(itemKey, propKey) Then
        valueKey = getValueKey(itemKey, propKey)
        strSQL = "select * " & _
                 "from " & _
                 "[CHT-CPNM].[dbo].[LINK_VALORES] inner join [CHT-CPNM].[dbo].[TIPO_PROPRIEDADES]" & _
                 "on [TIPO_PROPRIEDADES].[ID_TIPO_PROP] = [LINK_VALORES].[ID_TIPO_PROP]" & _
                 "where" & _
                 "[TIPO_PROPRIEDADES].[PROP_CALCULADA] = -1 and [LINK_VALORES].[ID_VALOR] = " & valueKey

        rs.Open strSQL, gCnn

        Do While rs.EOF <> True
            placeHolderString = getCalculatedDataFromDB(rs.Fields("ID_ITEM"), rs.Fields("ID_TIPO_PROP"))
            rs.MoveNext
        Loop
    End If

    ' Aqui entra o comando em SQL para fazer a query
    On Error GoTo getDataFromDB_Error
    tempValue = getValue(valueKey)

    If unitKey <> 0 And tempValue <> nonExistantValueString Then
        tempValue = gUnitDef.convertValue(tempValue, propKey, unitKey, True)
        tempUnit = gUnitDef.getUnitSymbol(propKey, unitKey)
    End If

    ' The code here only happens if the thing ran without errors.
    Select Case refOption
        Case 0
            varValueFromDatabase = tempUnit
        Case 1
            varValueFromDatabase = tempValue
        Case 2
            varValueFromDatabase = tempValue & " " & tempUnit

    End Select

    getDataFromDB = varValueFromDatabase

    Exit Function

getDataFromDB_ExitSub:
    Exit Function

getDataFromDB_Error:
    errorAction = handleMyError(showMsgBoxes)
    If Err.Number = vbObjectError + 100211 Or Err.Number = vbObjectError + _
       100212 Or Err.Number = vbObjectError + 100213 Then
        getDataFromDB = Err.Description
    Else
        Select Case errorAction
            Case -1
                Stop
            Case 1
                Resume Next
            Case 2
                GoTo getDataFromDB_ExitSub
            Case Else
                GoTo getDataFromDB_ExitSub
        End Select
    End If

End Function

Public Function getTrackingDataFromDB(itemKey As Long, propKey As Variant, Optional showMsgBoxes As Boolean = True)
    ' Esta função vai extrair um valor de propriedade de rastreamento (nomes, etc.)
    Dim strSQL                                    As String
    Dim rs                                        As ADODB.Recordset
    Dim varValueFromDatabase                      As String
    Dim errorAction                               As Integer

    ' Aqui entra o comando em SQL para fazer a query
    ' Query string for the value
    strSQL = "select ITEM.NOME_ITEM, SUB_AREA.NOME_SUB_ARE, AREA.NOME_ARE, PLANTA.NOME_PLA, INDUSTRIAL.NOME_IND, UNIDADE_NEGOCIO.NOME_UNI " & _
             "from UNIDADE_NEGOCIO INNER JOIN ((((INDUSTRIAL INNER JOIN PLANTA ON INDUSTRIAL.ID_IND = PLANTA.ID_IND) INNER JOIN AREA ON " & _
             "PLANTA.ID_PLA = AREA.ID_PLA) INNER JOIN SUB_AREA ON AREA.ID_ARE = SUB_AREA.ID_ARE) INNER JOIN ITEM ON SUB_AREA.ID_SUB_ARE = ITEM.ID_SUB_ARE) ON UNIDADE_NEGOCIO.ID_UNI = INDUSTRIAL.ID_UNI " & _
             "where ITEM.ID_ITEM = " & itemKey

    ' Criando um recordset com o resultado da query
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, gCnn

    On Error GoTo getTrackingDataFromDB_Error
    Call assertGetTrackingData(rs, itemKey, propKey)

    varValueFromDatabase = rs.Fields(propKey).value
    getTrackingDataFromDB = varValueFromDatabase

    Exit Function

getTrackingDataFromDB_ExitSub:
    Exit Function

getTrackingDataFromDB_Error:
    errorAction = handleMyError(showMsgBoxes)
    If Err.Number = vbObjectError + 100211 Or Err.Number = vbObjectError + _
       100212 Or Err.Number = vbObjectError + 100213 Then
        getTrackingDataFromDB = Err.Description
    Else
        Select Case errorAction
            Case -1
                Stop
            Case 1
                Resume Next
            Case 2
                GoTo getTrackingDataFromDB_ExitSub
            Case Else
                GoTo getTrackingDataFromDB_ExitSub
        End Select
    End If
End Function

Private Function getCalculatedDataFromDB(itemKey As Long, propKey As Long) As String
    '
    Dim strPartial                                As String

    strPartial = ""

    Select Case propKey                                              'Aqui eu faço a verificação dos casos especiais de propKeys
        Case 519                                                      'Tag da linha
            'strPartial = getDataFromDB(itemKey, 46, 0, 1)
            'strPartial = strPartial & "-" & getDataFromDB(itemKey, 61, 0, 1)
            strPartial = getTrackingDataFromDB(itemKey, "NOME_ITEM")
            'strPartial = strPartial & "-" & getDataFromDB(itemKey, 38, 0, 1)

    End Select

    Call exportSingleDataFromKeys(itemKey, propKey, strPartial, 0) 'Reexportando dados para o banco
    getCalculatedDataFromDB = getValue(getValueKey(itemKey, propKey))
End Function


Public Sub assertGetData(rs As ADODB.Recordset, itemKey As Long, propKey As Variant)
    '=======================================================================================
    ' A rotina faz a asserção que o dado obtido do banco de dados é correto.
    ' a rotina checa se a query retornou um único valor. Se não, gera um erro com a descrição
    ' apropriada
    '---------------------------------------------------------------------------------------
    ' < descrição dos argumentos>
    '---------------------------------------------------------------------------------------
    ' < Observações >
    '---------------------------------------------------------------------------------------
    ' < Histórico de revisões>
    '=======================================================================================
    '
    If rs.RecordCount = 1 Then
        If IsNull(rs.Fields("VALOR_PROP").value) Then
            Err.Raise vbObjectError + 100211, Description:="Erro! Valor da propriedade" & itemKey & "|" & propKey & " não está cadastrado (Null)"
        End If
    Else
        If rs.RecordCount > 1 Then
            Err.Raise vbObjectError + 100212, Description:="Erro! Valor da propriedade" & itemKey & "|" & propKey & " está dubplicado"
        Else
            Err.Raise vbObjectError + 100213, Description:="Erro! O par " & itemKey & "|" & propKey & " não está cadastrado"
        End If
    End If
End Sub


Public Sub assertGetTrackingData(rs As ADODB.Recordset, itemKey As Long, propKey As Variant)
    '=======================================================================================
    ' A rotina faz a asserção que o dado obtido do banco de dados é correto.
    ' a rotina checa se a query retornou um único valor. Se não, gera um erro com a descrição
    ' apropriada
    '---------------------------------------------------------------------------------------
    ' < descrição dos argumentos>
    '---------------------------------------------------------------------------------------
    ' < Observações >
    '---------------------------------------------------------------------------------------
    ' < Histórico de revisões>
    '=======================================================================================
    '
    If rs.RecordCount = 1 Then
        If IsNull(rs.Fields(propKey).value) Then
            Err.Raise vbObjectError + 100211, Description:="Erro! Valor da propriedade" & itemKey & "|" & propKey & " não está cadastrado (Null)"
        End If
    Else
        If rs.RecordCount > 1 Then
            Err.Raise vbObjectError + 100212, Description:="Erro! Valor da propriedade" & itemKey & "|" & propKey & " está dubplicado"
        Else
            Err.Raise vbObjectError + 100213, Description:="Erro! O par " & itemKey & "|" & propKey & " não está cadastrado"
        End If
    End If
End Sub

Public Function createAddress(itemKey As Long, propKey As Long, unitKey As Long, refOption As Integer) As String
    Dim strNewAddress                             As String
    ' checking which kind of reference to create
    If checkIfIsCalc(propKey) Then                                   'checking if it is a calculated reference
        strNewAddress = breakString & itemKey & breakString & calcChar & propKey & breakString & 0
    Else
        Select Case refOption                                        'choosing the non-calculated reference type.
            Case 0
                strNewAddress = breakString & itemKey & breakString & unitOnlyChar & propKey & breakString & unitKey
            Case 1
                strNewAddress = breakString & itemKey & breakString & valueOnlyChar & propKey & breakString & unitKey
            Case 2
                strNewAddress = breakString & itemKey & breakString & unitAndValueChar & propKey & breakString & unitKey
        End Select
    End If

    createAddress = strNewAddress
End Function

Public Function createTrackingAddress(itemKey As Long, propKey As Variant, unitKey As Long) As String
    createTrackingAddress = breakString & itemKey & breakString & trackingChar & propKey & breakString & unitKey
End Function
