Attribute VB_Name = "MTestSuite"
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

Public Sub runTestSuite()

    Call initializeCPNM

    MsgBox "iniciando teste de tratamento de erro controlado"
    Call testCustomErrorCatch
    MsgBox "iniciando teste de tratamento de erro não tratado"
    Call testErrorCatch
    MsgBox "iniciando teste do getData"
    Call testGetDataError
    MsgBox "iniciando teste de existência de valor"
    Call testValueExistance
    MsgBox "iniciando teste de upload de dado simples"
    Call upload_dado
    MsgBox "iniciando teste de compartilhamento e quebra"
    Call testeShare
    MsgBox "iniciando teste de update de propriedade de rastreamento"
    Call test_track_change
End Sub

Private Sub testErrorCatch()

    Dim a                As Single
    On Error GoTo testErrorCatch_Error
    a = 1 / 0

    Exit Sub

testErrorCatch_Error:
    Call handleMyError

End Sub

Private Sub testCustomErrorCatch()

    On Error GoTo testCustomErrorCatch_Error
    Err.Raise Number:=vbObjectError + 22000, Description:="Caught your test error"

    Exit Sub

testCustomErrorCatch_Error:
    Call handleMyError
End Sub

Private Sub testGetDataError()

    Dim strAddress       As String
    Dim resultado        As String

    'item que não existe
    strAddress = "dummy1345dummyD3867567"
    resultado = getData(strAddress, True)
    MsgBox "Resultado = " & resultado

    'item que existe
    strAddress = "dummy3dummyD7"
    resultado = getData(strAddress, True)
    MsgBox "Resultado = " & resultado

End Sub

Private Sub testValueExistance()

    Dim itemKey          As Long
    Dim propKey          As Long

    itemKey = -115
    propKey = -19

    If Not checkValueExistance(itemKey, propKey) Then
        MsgBox "teste concluido com sucesso"
    Else
        MsgBox "vishhhhh, ta errado", vbCritical
    End If
End Sub


Private Sub upload_dado()
    Dim itemName         As String
    Dim itemType         As String
    Dim propName         As String
    Dim value            As String
    Dim unit             As String
    Dim itemExists       As Boolean
    Dim itemKey          As Long
    Dim propKey          As Long
    Dim valueKey         As Long

    'start
    itemName = "BA-050234"
    itemType = "BOMBA"
    propName = "Calor específico"
    value = Rnd()
    unit = 0

    itemExists = checkItemExistance(itemName)

    If Not itemExists Then
        Call createItem(itemName, itemType, "1000 - Água Industrial e Potável")
        If Err Then MsgBox "erro na criação do item"
    End If
    

    itemKey = getItemKey(itemName)
    propKey = getPropKey(propName)
    valueKey = getValueKey(itemKey, propKey)

    'Teste de inserção de um dado
    Call exportSingleData(itemName, propName, value, unit, True)

    Debug.Assert value = getValue(valueKey)

    ' report
    If Not Err Then
        MsgBox "teste executado com sucesso"
    Else
        MsgBox "teste não foi concluido com sucesso", vbCritical
    End If

End Sub

Private Sub testeShare()


    Dim name1            As String
    Dim name2            As String
    Dim prop1            As String
    Dim prop2            As String
    Dim value1           As String
    Dim value2           As String
    Dim unit1            As String
    Dim unit2            As String
    Dim itemKey1         As Long
    Dim itemKey2         As Long
    Dim propKey1         As Long
    Dim propKey2         As Long
    Dim valKey1          As Long
    Dim valKey2          As Long
    Dim result1          As String
    Dim result2          As String


    name1 = "10-HT-107"
    name2 = "10-HT-108"
    prop1 = "Área"
    prop2 = "Área"
    value1 = "12"
    value2 = "24"
    unit1 = "m²"
    unit2 = "m¹"

    ' getting keys to check if the thing was created
    itemKey1 = getItemKey(name1)
    itemKey2 = getItemKey(name2)
    propKey1 = getPropKey(prop1)
    propKey2 = getPropKey(prop2)

    ' getting the values after sharing
    valKey1 = getValueKey(itemKey1, propKey1)
    valKey2 = getValueKey(itemKey2, propKey2)

    ' if they are already shared, break the sharing
    If valKey1 = valKey2 Then
        Call breakSharing(name1, prop1, name2, prop2)
        ' getting the values after sharing
        valKey1 = getValueKey(itemKey1, propKey1)
        valKey2 = getValueKey(itemKey2, propKey2)
    End If

    ' putting two different values there
    Call exportSingleData(name1, prop1, value1, unit1)
    Call exportSingleData(name2, prop2, value2, unit2)

    Debug.Print getValue(valKey2)
    Debug.Print getValue(valKey1)
    Debug.Assert (getValue(valKey1) <> getValue(valKey2))

    ' creating the sharing
    Call createSharing(name1, prop1, name2, prop2)

    ' getting the values after sharing
    valKey1 = getValueKey(itemKey1, propKey1)
    valKey2 = getValueKey(itemKey2, propKey2)

    ' asserting that the sharing did happen
    result1 = getValue(valKey1)
    result2 = getValue(valKey2)
    Debug.Assert (valKey1 = valKey2)
    Debug.Assert (result1 = result2)

    ' breaking the sharing
    Call breakSharing(name1, prop1, name2, prop2)

    ' getting the values after the break
    valKey1 = getValueKey(itemKey1, propKey1)
    valKey2 = getValueKey(itemKey2, propKey2)

    ' asserting that the sharing was broke
    Debug.Assert (valKey1 <> valKey2)
    Debug.Assert (getValue(valKey1) = getValue(valKey2))

    Debug.Assert (getValue(valKey1) = value1)                        ' this now will only work if you choose yes
    Debug.Assert (getValue(valKey2) = value1)

    If Not Err Then
        MsgBox "teste executado com sucesso"
    Else
        MsgBox "teste não foi concluido com sucesso", vbCritical
    End If
End Sub

Private Sub test_track_change()
    Dim itemName         As String
    Dim itemKey          As Long
    Dim newName          As String
    Dim propName         As String


    ' itemName
    itemName = "10-HT-108"

    ' start of the testing
    itemKey = getItemKey(itemName)
    newName = "NOMENOVO"
    propName = "Nome do item"

    ' changing
    Call changeTrackingData(itemName, propName, newName)

    ' checking change
    MsgBox getItemName(itemKey) & "  o certo aqui é: " & newName

    ' undoing the change
    Call changeTrackingData(newName, propName, itemName)

    ' checking change
    MsgBox getItemName(itemKey) & "  o certo aqui é: " & itemName
End Sub



Private Sub test_query()
    Dim rs               As ADODB.Recordset
    Set rs = getQuery("viewItemProp")
    Call query2msgbox(rs)

End Sub

Private Function getQuery(strQueryName As String) As ADODB.Recordset
    Dim strConnect       As String
    Dim rs               As ADODB.Recordset
    Dim cnn              As ADODB.Connection

    gCnn.Open strConnect

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    rs.Open strQueryName, cnn

    Set getQuery = rs
End Function

Private Sub query2msgbox(rs As ADODB.Recordset)

    Dim strReport        As String
    Dim field            As Variant

    strReport = ""
    For Each field In rs.Fields
        strReport = strReport & field.Name & "   "
    Next field

    strReport = strReport & vbCr

    Do While rs.EOF <> True
        For Each field In rs.Fields
            strReport = strReport & " | " & field.value
        Next field

        strReport = strReport & " | " & vbCr

        rs.MoveNext
    Loop
    MsgBox strReport
End Sub
