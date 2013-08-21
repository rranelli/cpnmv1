Attribute VB_Name = "MSpecificExcel"
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

Public Sub getDataFromDatabase(Optional showMsgBoxes As Boolean = True)
    Call getRangeVarsFromDatabase(showMsgBoxes)
End Sub

Sub cleanUpWholeDocument(Optional showMsgBoxes As Boolean = True)
'=======================================================================================
' < descrição da rotina >
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
'Sub cleanUpWholeDocument()
' This sub cleans up the document
    On Error GoTo cleanUpWholeDocument_Error
    Debug.Assert Application.Name = "Microsoft Excel"

    Dim rangeVarz                            As Variant
    Dim namedRange                           As Name
    Dim strReport                            As String
    Dim strAddress                           As String
    Dim rangeVarsInDocument                  As collection

    ' This part cleans up the document of docvariables which do not appear in the text
    Set rangeVarsInDocument = New collection
    strReport = "I Deleted the following docvariables, with the following value: "

    For Each namedRange In ActiveWorkbook.Names
        If InStr(namedRange.Name, breakString) Then
            strAddress = namedRange.Name

            If checkIfRangeVarIsInSheet(strAddress) Then
                rangeVarsInDocument.Add "found", strAddress
            End If

            On Error GoTo 0

        End If
    Next namedRange

    For Each rangeVarz In ActiveWorkbook.Names
        If Not isInCollection(rangeVarsInDocument, rangeVarz.Name) And InStr(rangeVarz.Name, breakString) Then
            strReport = strReport & vbCr & rangeVarz.Name & "   : " & rangeVarz.RefersTo
            rangeVarz.Delete
        End If
    Next rangeVarz

    If showMsgBoxes Then MsgBox strReport
    On Error GoTo 0

    Exit Sub

cleanUpWholeDocument_Error:
    Call handleMyError
End Sub

Private Sub getRangeVarsFromDatabase(Optional showMsgBoxes As Boolean = True)
'=======================================================================================
' This sub gets data from the database
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' [*] É necessário fazer uma formatação especial na hora de atribuir para o range nomeado.
' quando mandamos atribuir ao range nomeado o valor "-1 °C" o excel traduz este valor para
' "=-1 °C", o que gera um erro. Precisamos atribuir o valor " ="-1°C" " para evitar isso.
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Dim problemCount                         As Integer
    Dim valueToRangeVar                      As String
    Dim rangeVarz                            As Name

    Debug.Assert Application.Name = "Microsoft Excel"

    On Error GoTo getDocVarsFromDatabase_Error

    ' Starting the number of blown up variables
    problemCount = 0

    ' This is the main loop, all magic goes here
    For Each rangeVarz In ActiveWorkbook.Names
        If rangeVarz.Name Like "dummy*dummy*" Then
            ' sending the address string (which is the docvar name) to the getData function
            valueToRangeVar = getData(rangeVarz.Name)

            ' Getting the value into the docvar
            rangeVarz.RefersTo = "=" & Chr(34) & valueToRangeVar & Chr(34)    ' veja [*] na observação

            ' Checking if the value returned was null
            If InStr(valueToRangeVar, "Erro!") Then
                problemCount = problemCount + 1
            End If
        End If
    Next rangeVarz

    ' Now, i get to tell you if everything went as expected
    If problemCount = 0 Then
    Else
        If showMsgBoxes Then MsgBox "Warning!" & Chr(10) & Chr(10) & "There are " & problemCount & " fields with no value in the database!!         "
    End If

    ' And we clean the thing up
    Call cleanUpWholeDocument(False)
    On Error GoTo 0

    Exit Sub

getDocVarsFromDatabase_Error:
    Call handleMyError

End Sub

Public Sub createReference(strItemName As String, strPropName As String, _
                           strUnitName As String, isTrack As Boolean, Optional refOption As Integer = 2)
'=======================================================================================
' < descrição da rotina >
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Dim itemKey                              As Long
    Dim propKey                              As Long
    Dim unitKey                              As Long
    Dim dummy                                As Variant
    Dim strNewAddress                        As String
    Dim itemTrackingPropKey                  As String
    Dim colDictionary                        As collection

    Debug.Assert Application.Name = "Microsoft Excel"
    On Error GoTo createReference_Error
    ' If register is found, get the keys.

    If isTrack Then
        itemKey = getItemKey(strItemName)

        Set colDictionary = createTrackingDictionary()
        itemTrackingPropKey = colDictionary(strPropName)

        strNewAddress = createTrackingAddress(itemKey, itemTrackingPropKey, 0)
    Else

        itemKey = getItemKey(strItemName)
        propKey = getPropKey(strPropName)
        unitKey = getUnitKey(strPropName, strUnitName)

        strNewAddress = createAddress(itemKey, CVar(propKey), unitKey, refOption)
    End If


    'Checking if the variable already exists, and creating the thing nice.
    On Error Resume Next
    dummy = ActiveWorkbook.Names(strNewAddress).value

    If Err.Number = 0 Then                            ' docvar exists!
        MsgBox "I will now create the field in the document, but the rangeVar " & strNewAddress & " already exists"

        'Creating the field in the document
        Selection.value = "=" & strNewAddress
    Else
        'Adding the named range into excel named ranges
        ActiveWorkbook.Names.Add Name:=strNewAddress, RefersTo:="waiting for database update"
        ' Creating the field in the document
        Selection.value = "=" & strNewAddress
    End If

    'Updating all the docvars
    Call getDataFromDatabase

    On Error GoTo 0

    Exit Sub

createReference_Error:
    Call handleMyError

End Sub

Sub changeTheReferences(strOriginalItemName As String, strNewItemName As String, Optional bolColor As Boolean = False, Optional dummy1 As Boolean = True, Optional dummy2 As Boolean = True)
'=======================================================================================
' This sub will change the name of the DocVars
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
' Declarations
    Dim originalItemKey                      As Long
    Dim newItemKey                           As Long
    Dim changeCount                          As Integer
    Dim problemCount                         As Integer
    Dim rangeVarz                            As Variant
    Dim splitz                               As Variant
    Dim itemKey                              As Long
    Dim unitKey                              As Long
    Dim propKey                              As Variant         'the propKey here can refer to the tracking property primary key! which is a string.
    Dim newRangeVarName                      As String

    Debug.Assert Application.Name = "Microsoft Excel"

    ' Here, I put the connect string
    On Error GoTo changeTheReferences_Error

    ' getting the keys
    originalItemKey = getItemKey(strOriginalItemName)
    newItemKey = getItemKey(strNewItemName)

    ' Now, for each docVar that points into the originalItemKey, i will change it to the newItemKey
    changeCount = 0                                   'counting the number of changes
    problemCount = 0                                  'counting the number of null new fields

    For Each rangeVarz In ActiveWorkbook.Names
        ' here i break the rangeVarz names
        If rangeVarz.Name Like "*" & breakString & "*" Then
            splitz = split(rangeVarz.Name, breakString)
            itemKey = CLng(splitz(1))                 ' Now, Here is some crazy and important catch!!
            propKey = splitz(2)
            unitKey = CLng(splitz(3))

            ' here I check if the active docvar points to the original key
            If itemKey = originalItemKey Then
                'here I get the hold of the new and the old docvar names
                newRangeVarName = breakString & newItemKey & breakString & propKey & breakString & unitKey
                rangeVarz.Name = newRangeVarName
                
                ' here I change the color of the changed stuff.
                If bolColor Then
                    findAndHighlight (rangeVarz.Name)
                End If
            End If
        End If
    Next rangeVarz

    ' Voilá my good friend =)
    ' And we clean the thing up
    Call cleanUpWholeDocument(False)
    ' And for the glory of CopyPasteNoMore: Update what is there
    Call getRangeVarsFromDatabase
    On Error GoTo 0

    Exit Sub

changeTheReferences_Error:
    Call handleMyError
End Sub

Sub populateOriginalItem(cmbOriginalItemName, Optional strItemClass As String, Optional strSubArea As String)
'=======================================================================================
' This sub populate the "original item" combobox in the "change reference form"
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Debug.Assert Application.Name = "Microsoft Excel"

    ' Declarations
    Dim itemName                             As String
    Dim splited                              As Variant
    Dim itemKey                              As Long
    Dim subAreaKey                           As Long
    Dim subAreaKeyDesired                    As Long
    Dim itemClassKey                         As Long
    Dim itemClassKeyDesired                  As Long
    Dim rangeVarz                            As Variant
    Dim element                              As Variant
    Dim bolSentinel                          As Boolean
    Dim itemNameCol                          As collection

    ' Here, I put the connect string
    On Error GoTo populateOriginalItem_Error

    ' initial attributions
    Set itemNameCol = New collection
    If strItemClass <> "" Then itemClassKeyDesired = getItemClassKey(strItemClass)
    If strSubArea <> "" Then subAreaKeyDesired = getSubAreaKey(strSubArea)

    If ActiveWorkbook.Names.Count > 0 Then
        For Each rangeVarz In ActiveWorkbook.Names
            If InStr(rangeVarz.Name, breakString) Then
                ' parsing the address to the value
                splited = split(rangeVarz.Name, breakString)

                ' getting the item name
                itemKey = CLng(splited(1))
                itemName = getItemName(itemKey)

                ' getting the auxiliary keys
                subAreaKey = getSubAreaKeyFromItemKey(itemKey)
                itemClassKey = getItemClassKeyFromItemKey(itemKey)

                ' Checking the aditional filters
                bolSentinel = True                    ' the bolSentinel is a watch for constraint violation.
                If strItemClass <> "" Then
                    If itemClassKey <> itemClassKeyDesired Then bolSentinel = False
                End If
                If strSubArea <> "" Then
                    If subAreaKey <> subAreaKeyDesired Then bolSentinel = False
                End If

                ' Adding to the collection, but only DISTINCT values.
                If Not isInCollection(itemNameCol, itemName) And bolSentinel Then
                    itemNameCol.Add Item:=itemName, key:=itemName
                End If

            End If
        Next rangeVarz
    Else
        MsgBox "You have no references in the document."
        Exit Sub
    End If

    'Now, I iterate over the query and fill the value of each item name in the combo
    For Each element In itemNameCol
        cmbOriginalItemName.AddItem element
    Next element
    On Error GoTo 0

    Exit Sub

populateOriginalItem_Error:
    Call handleMyError

End Sub

Sub populateNewItem(strOriginalItemName As String, cmbNewItemName, Optional strSubArea As String)
'=======================================================================================
' < descrição da rotina >
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
' This sub populate the newItem combobox given the type of the OriginalItem

' Declarations
    Debug.Assert Application.Name = "Microsoft Excel"

    Dim rs                                   As ADODB.Recordset
    Dim strSQL                               As String
    Dim originalItemTypeKey                  As Long
    Dim originalItemKey                      As Long
    Dim subAreaKey                           As Long

    On Error GoTo populateNewItem_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Now, I will get the keys
    originalItemKey = getItemKey(strOriginalItemName)
    originalItemTypeKey = getItemTypeFromItemKey(originalItemKey)

    'Now, I will query over all the items with the itemtype I got above
    strSQL = "Select NOME_ITEM from ITEM where ID_TIPO_ITEM = " & originalItemTypeKey

    ' filtering by subArea
    If strSubArea <> "" Then
        subAreaKey = getSubAreaKey(strSubArea)
        strSQL = strSQL & " AND ID_SUB_ARE = " & subAreaKey
    End If

    ' Now, I open this new query
    rs.Open strSQL, gCnn

    'Now, I iterate over the query and fill the value of each item name in the combo
    Do While rs.EOF <> True
        cmbNewItemName.AddItem (rs.Fields("NOME_ITEM").value)
        rs.MoveNext
    Loop
    ' Voilá =)
    On Error GoTo 0

    Exit Sub

populateNewItem_Error:
    Call handleMyError
End Sub

Public Function checkIfRangeVarIsInSheet(strAddress As String) As Boolean
'=======================================================================================
' This sub checks if the rangeVar appears in some worksheet
'---------------------------------------------------------------------------------------
' < descrição dos argumentos>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Dim found                                As Range
    Dim worksheetz                           As Worksheet

    For Each worksheetz In ActiveWorkbook.Sheets
        Set found = Nothing
        worksheetz.Activate
        'found = Cells.Find(What:=strAddress, LookIn:=xlFormulas _
         , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
         MatchCase:=False, SearchFormat:=False)
        Set found = Cells.Find(what:=strAddress)

        If Not (found Is Nothing) Then
            checkIfRangeVarIsInSheet = True
            Exit Function
        End If
    Next worksheetz

    checkIfRangeVarIsInSheet = False
End Function

Private Sub findAndHighlight(strAddress As String)
'=======================================================================================
' Esta rotina busca na planilha por endereços especificados em strAddress e troca a cor
' para amarelo.
'---------------------------------------------------------------------------------------
' [strAddress] endereço  no formato dummy<itemKey>dummy<refType><propKey>
'---------------------------------------------------------------------------------------
' < Observações >
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Dim c                                    As Variant
    Dim firstAddress                         As String

    Set c = Cells.Find(what:=strAddress)
    firstAddress = c.Address
    If Not c Is Nothing Then
        Do
            c.Select

            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

            Set c = Cells.FindNext(c)
        Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
End Sub
