Attribute VB_Name = "MSpecificExcelExport"
' Chemtech - A Siemens Business ========================================================
'
'=======================================================================================
Option Explicit
' Desenvolvimento ======================================================================
' RR            Renan Ranelli                    renan.ranelli@chemtech.com.br
' PRP           Paulo Roberto Polastri           paulo.polastri@chemtech.com.br
'=======================================================================================
' Versões ==============================================================================
'12/13/2012  - Emissão inicial
'=======================================================================================
'
Private Const propRow                             As Long = 3
Private Const itemColumn                          As Long = 2
Private Const keyColumn                           As Long = 1
Private Const unitRow                             As Long = 5
Private Const startingColumn                      As Long = 3
Private Const startingRow                         As Long = 6

Public Sub uploadDataWithTimer()
    Dim myTimer                                   As obTimer
    Set myTimer = New obTimer

    ' Abrindo a conexão com o banco de dados.
    Set gCnn = defineDatabaseConnection(-1)
    Call createGlobalUnitDefinitionObject

    On Error Resume Next
    gCnn.Open
    If Err Then                                                      'se houver problema na abertura da conexão, ela precisa ser redefinida.
        MsgBox "Você precisa configurar uma conexão com o banco de dados"
        Exit Sub
    End If
    On Error GoTo 0

    ' Running the upload task
    myTimer.StartTimer                                               'starting the timer
    Call uploadDataToDatabase                                        'running the task
    myTimer.StopTimer                                                'stopping the timer

    ' reporting the time
    MsgBox "Exportação realizada com sucesso! " & vbCr & vbCr & "Tempo decorrido para exportação:" & myTimer.Elapsed & " segundos"

    ' Closing the connection if it stills open
    If gCnn.State = adStateOpen Then gCnn.Close
End Sub

Private Sub uploadDataToDatabase()
    ' This sub exports data from the active worksheet into the database
    Dim c                                         As Long
    Dim r                                         As Long
    Dim itemKey                                   As Long
    Dim propKey                                   As Long
    Dim valueKey                                  As Long
    Dim strItemName                               As String
    Dim strSubArea                                As String
    Dim strItemType                               As String
    Dim strPropName                               As String
    Dim strInsertValue                            As String
    Dim strInsertUnit                             As String
    Dim strSharedAddress                          As String
    Dim anomalySentinel                           As Boolean
    Dim sharedCollection                          As collection

    ' initial definitions
    On Error GoTo uploadDataToDatabase_Error
    c = startingColumn                                               ' the start column in the worksheet
    r = startingRow                                                  ' the start row in the worksheet
    Set sharedCollection = New collection

    ' Getting the tracking information for this worksheet
    strItemType = Cells(1, 5).Text
    strSubArea = Cells(1, 3).Text

    ' Validanting the items
    Call validateItems(True)

    'This is the exporting part of the code

    c = startingColumn                                               ' the start column in the worksheet
    r = startingRow                                                  ' the start row in the worksheet

    Do While Cells(r, keyColumn).Formula <> Empty                    ' repeat until first empty cell in column c
        Do While Cells(propRow, c).Formula <> Empty                  ' repeat until first empty cell in row r
            If Cells(propRow + 1, c) <> 1 Then
                ' Getting Information for data export
                strItemName = Cells(r, itemColumn).Text
                strPropName = Cells(propRow, c).Text                 'property name

                itemKey = getItemKey(strItemName)
                On Error Resume Next
                propKey = getPropKey(strPropName)
                If Err = 0 Then                                      ' If the prop exists...
                    On Error GoTo 0

                    strInsertValue = Cells(r, c).Text                'Text to be inserted
                    strInsertUnit = Cells(unitRow, c).Text           'unit to be inserted

                    ' Data export
                    If strInsertValue <> "" Then
                        Call exportSingleData(strItemName, strPropName, strInsertValue, strInsertUnit, False)    'The "False" argument tells the sub to update existing items without prompt
                    End If

                    ' Testing if the current cell is in the sharedCollection.
                    If checkIfIsShared(itemKey, propKey) Then

                        valueKey = getValueKey(itemKey, propKey)
                        If isInCollection(sharedCollection, CStr(valueKey)) Then
                            strSharedAddress = sharedCollection(CStr(valueKey))
                            Range(strSharedAddress).Interior.ColorIndex = 36

                            Cells(r, c).Interior.ColorIndex = 36

                            anomalySentinel = True
                        Else
                            sharedCollection.Add Cells(r, c).Address, CStr(valueKey)
                        End If

                    End If
                End If

                On Error GoTo 0
            End If
            c = c + 1                                                ' next column
        Loop

        r = r + 1                                                    ' next row
        c = startingColumn                                           ' go back to starting column
    Loop

    If anomalySentinel Then MsgBox "Há uma anomalia na exportação. Você está fazendo upload de valores compartilhados mais de uma vez."

    On Error GoTo 0

    GoTo uploadDataToDatabase_Finally

uploadDataToDatabase_Finally:

    Exit Sub

    ' Procedure Error Handler
uploadDataToDatabase_Error:
    Dim errorAction                               As Integer
    'here goes your specific error handling code.

    ' here comes the generic global error handling code.
    errorAction = handleMyError()
    Select Case errorAction
        Case -1
            Stop
        Case 1
            Resume Next
        Case 2
            GoTo uploadDataToDatabase_Finally
        Case Else
            Stop
    End Select
End Sub

Public Sub importDataWithTimer()

    Dim myTimer                                   As obTimer
    Set myTimer = New obTimer

    ' Opening the connection
    Set gCnn = defineDatabaseConnection(-1)
    On Error Resume Next
    gCnn.Open
    If Err Then
        MsgBox "Você precisa configurar uma conexão com o banco de dados"
        Exit Sub
    End If

    Call createGlobalUnitDefinitionObject

    On Error GoTo 0
    ' Running the upload task
    myTimer.StartTimer
    Call importDataFromDatabase
    myTimer.StopTimer

    MsgBox " Importação realizada com sucesso!Tempo decorrido para exportação:" & myTimer.Elapsed

    ' Closing the connection
    If gCnn.State = adStateClosed Then gCnn.Close
End Sub

Private Sub importDataFromDatabase()

    Dim c                                         As Long
    Dim r                                         As Long
    Dim itemKey                                   As Long
    Dim unitKey                                   As Long
    Dim propKey                                   As Long
    Dim strItemName                               As String
    Dim strPropName                               As String
    Dim strAddress                                As String
    Dim strInsertValue                            As String
    Dim strUnitName                               As String

    ' validando as itemKeys e criando os items que precisam ser criados.
    Call validateItems(False)

    c = startingColumn                                               ' the start column in the worksheet
    r = startingRow                                                  ' the start row in the worksheet

    Do While Cells(r, itemColumn).Formula <> Empty                   ' repeat until first empty cell in column c
        Do While Cells(propRow, c).Formula <> Empty                  ' repeat until first empty cell in row r
            ' Getting Information for data export

            ' getting the names
            strInsertValue = ""

            strItemName = Cells(r, itemColumn).Text
            strPropName = Cells(propRow, c).Text
            strUnitName = Cells(unitRow, c).Text

            ' getting the keys from the names;
            On Error Resume Next
            itemKey = getItemKey(strItemName)
            propKey = getPropKey(strPropName)
            unitKey = getUnitKey(strPropName, strUnitName)

            If Err = 0 Then
                If itemKey <> Cells(r, 1).value Then
                    MsgBox "O nome do seu item " & strItemName & " não corresponde à chave primaria guardada"
                    Exit Sub
                End If

                On Error GoTo 0
                strAddress = createAddress(itemKey, propKey, unitKey, 1)
                strInsertValue = getData(strAddress, False)
                Cells(r, c).value = strInsertValue
            End If

            c = c + 1                                                ' next column
        Loop

        r = r + 1                                                    ' next row
        c = startingColumn                                           ' go back to starting column
    Loop
    Exit Sub
End Sub

Public Sub deleteRow()
    ' This sub deletes the row of the item

    Dim confirm                                   As Variant

    ' checking if he can actually delete this line
    If ActiveCell.row < 6 Then
        MsgBox "Você não pode deletar essa linha!"
        Exit Sub
    End If

    ' asking the user if he is really sure
    confirm = MsgBox("Tem certeza que deseja deletar a linha " & ActiveCell.row & "?", vbYesNo)
    If confirm = vbYes Then
        ActiveCell.EntireRow.Delete
    End If
End Sub

Private Sub validateItems(Optional createItems As Boolean = True)
    Dim c                                         As Long
    Dim r                                         As Long
    Dim itemKey                                   As Long
    Dim strItemType                               As String
    Dim strSubArea                                As String
    Dim strItemName                               As String

    Dim check                                     As Variant

    c = startingColumn                                               ' the start column in the worksheet
    r = startingRow                                                  ' the start row in the worksheet

    ' Getting the tracking information for this worksheet
    strItemType = Cells(1, 5).Text
    strSubArea = Cells(1, 3).Text

    ' This part will create every item in the worksheet which is not yet present in the database!
    ' Also, this will update the item name according the the primary Key!
    Do While Cells(r, itemColumn).Formula <> Empty


        strItemName = Cells(r, itemColumn).Text                      ' Getting the item name for test

        ' Checking if I have to create the item
        If Cells(r, keyColumn) = Empty And createItems Then
            ' First, we get the needed info to create the item
            If checkItemExistance(strItemName) = False Then
                ' Now, we create the item
                Call createItem(strItemName, strItemType, strSubArea)
                'MsgBox "Creating the item: " & strItemName & " of the type " & strItemType
            End If

            ' Now, I get the item key from the created Item
            itemKey = getItemKey(strItemName)
            ' And I put it into the first column
            Cells(r, keyColumn).value = itemKey
        End If

        ' Checking if there is already an intem in the row
        If Cells(r, keyColumn) <> Empty Then
            ' Get the name of the existing item
            itemKey = CLng(Cells(r, keyColumn).value)
            strItemName = getItemName(itemKey)

            'Checking if the item name of an entry was changed
            If strItemName <> Cells(r, itemColumn).Text Then
                ' Asking the user if he wants to revert to the original name
                check = MsgBox("O nome do item " & Cells(r, 2).value & " não corresponde à chave primária nesta planilha!" & Chr(10) & _
                               "O sistema acredita que este item deveria se chamar: " & strItemName & Chr(10) & Chr(10) & _
                               "O sistema está certo? Em caso de dúvida, selecione NÃO", vbYesNo)
                If check = vbYes Then
                    Cells(r, itemColumn).value = strItemName
                    'Else, does not let the user change the entry name. He needs to create a new entry and delete the old one.
                ElseIf check = vbNo Then
                    MsgBox "A importação será interrompida"
                    Exit Sub
                End If
            End If
        End If

        r = r + 1                                                    'Now, we go to the next item
    Loop

End Sub

Public Sub wrapSetUpConnection()
    ' This sub creates the connection to the database
    Call initializeCPNM

    If Not gCnn Is Nothing Then
        If gCnn.State = adStateOpen Then
            gCnn.Close
        End If
    End If

    Set gCnn = defineDatabaseConnection(0)
    gCnn.Open

    Call createGlobalUnitDefinitionObject
End Sub

Public Sub resetConnection()
    If Not gCnn Is Nothing Then
        If gCnn.State = adStateOpen Then
            gCnn.Close
        End If
    End If

    Set gCnn = defineDatabaseConnection(1)
    gCnn.Open

    Call createGlobalUnitDefinitionObject
End Sub

Public Sub dumpItemsByType()
    Dim rs                                        As Recordset
    Dim strSQL                                    As String
    Dim row                                       As Integer
    Dim itemTypeName                              As String
    Dim itemTypeKey                               As Long

    itemTypeName = Cells(1, 5).Text
    itemTypeKey = getItemTypeKey(itemTypeName)

    Set rs = New Recordset
    strSQL = "select ID_ITEM, NOME_ITEM from ITEM where ID_TIPO_ITEM = " & itemTypeKey
    rs.Open strSQL, gCnn

    row = startingRow
    Do While Not rs.EOF
        Cells(row, itemColumn).value = rs.Fields("NOME_ITEM").value
        Cells(row, keyColumn).value = rs.Fields("ID_ITEM").value
        rs.MoveNext
        row = row + 1
    Loop
End Sub

Public Sub dumpInstrumentos()
    Dim rs                                        As Recordset
    Dim strSQL                                    As String
    Dim row                                       As Integer
    Dim itemTypeName                              As String
    Dim itemTypeKey                               As Long

    Set rs = New Recordset
    strSQL = "select ID_ITEM, NOME_ITEM from ITEM inner join TIPO_ITEM on ITEM.ID_TIPO_ITEM = TIPO_ITEM.ID_TIPO_ITEM where ID_CLASSE_TIPO_ITEM = 3"    '3 é a chave da classe instrumentos
    rs.Open strSQL, gCnn
    row = startingRow

    Do While Not rs.EOF
        Cells(row, itemColumn).value = rs.Fields("NOME_ITEM").value
        Cells(row, keyColumn).value = rs.Fields("ID_ITEM").value
        rs.MoveNext
        row = row + 1
    Loop
End Sub

Public Sub runDiagnose()
    Dim folderPath                                As String
    Dim strWbName                                 As String
    Dim strConcat                                 As String
    Dim propRow                                   As Integer
    Dim startCol                                  As Integer
    Dim col                                       As Integer
    Dim row                                       As Integer
    Dim thisKey                                   As Variant
    Dim itemTypeCol                               As Integer
    Dim itemTypeRow                               As Integer
    Dim reportRowStart                            As Integer
    Dim thisWorksheet                             As Worksheet
    Dim thisWorkbook                              As Workbook
    Dim reportWorksheet                           As Worksheet
    Dim thisPropName                              As String
    Dim thisItemTypeName                          As String
    Dim strLog                                    As String
    Dim uploadDic                                 As dictionary
    Dim varWorksheet                              As Variant
    Dim splitz                                    As Variant
    Dim startRow                                  As Integer

    propRow = 3                                                      'estes dois parametros são de configuração para a leitura das planilhas
    startCol = 3
    startRow = 6
    itemTypeRow = 1
    itemTypeCol = 5
    reportRowStart = 7

    Set uploadDic = New dictionary
    Call initializeCPNM

    folderPath = displayFolderOpen("Selecione a pasta com as planilhas de origem dos dados", ActiveWorkbook.Path)
    Set reportWorksheet = ActiveSheet

    strWbName = Dir(folderPath & "\*.xls*")
    Do While strWbName <> ""
        If strWbName Like reportWorksheet.Cells(2, 2).Text Then
            Set thisWorkbook = Workbooks.Open(folderPath & "\" & strWbName)
            For Each varWorksheet In thisWorkbook.Worksheets
                Set thisWorksheet = varWorksheet

                Call validateItems

                col = startCol
                Do While thisWorksheet.Cells(propRow, col).value <> ""
                    row = startRow
                    Do While thisWorksheet.Cells(row, keyColumn).value <> ""
                        thisPropName = thisWorksheet.Cells(propRow, col).Text
                        thisItemTypeName = getItemTypeFromItemKey(CLng(thisWorksheet.Cells(row, keyColumn).value))
                        strConcat = thisItemTypeName & "|" & thisPropName

                        If Not uploadDic.Exists(strConcat) Then
                            uploadDic.Add strConcat, thisWorkbook.Name
                        Else
                            uploadDic(strConcat) = uploadDic(strConcat) & "/" & thisWorkbook.Name
                            strLog = strLog & vbCr & "O par " & strConcat & " aparece em workbooks repetidos"
                        End If
                        row = row + 1
                    Loop
                    col = col + 1
                Loop
            Next varWorksheet
            thisWorkbook.Close
        End If

        strWbName = Dir()
    Loop

    row = reportRowStart
    For Each thisKey In uploadDic.Keys
        strConcat = thisKey
        splitz = split(strConcat, "|")
        thisItemTypeName = splitz(0)
        thisPropName = splitz(1)

        With reportWorksheet
            .Cells(row, 1).value = thisItemTypeName
            .Cells(row, 2).value = thisPropName
            .Cells(row, 3).value = uploadDic(strConcat)
            If uploadDic(strConcat) Like "*/*" Then .Cells(row, 4).value = "Sim" Else .Cells(row, 4).value = "Não"
        End With
        row = row + 1
    Next thisKey

    Debug.Print strLog
End Sub
