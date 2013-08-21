Attribute VB_Name = "MSpecificWord"
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
    Call getDocVarsFromDatabase(showMsgBoxes)
End Sub

Sub cleanUpWholeDocument(Optional showMsgBoxes As Boolean = True)
    ' This sub cleans up the document
    On Error GoTo cleanUpWholeDocument_Error
    Debug.Assert Application.Name = "Microsoft Word"

    Dim docVarz                                   As Variant
    Dim fieldz                                    As field
    Dim strReport                                 As String
    Dim strAddress                                As String
    Dim splitz                                    As Variant
    Dim dummy                                     As Variant
    Dim element                                   As Variant
    Dim docVarsInDocument                         As Collection

    ' This part cleans up the document of docvariables which do not appear in the text
    Set docVarsInDocument = New Collection
    strReport = "I Deleted the following docvariables, with the following value: "

    For Each fieldz In ActiveDocument.Fields
        If InStr(fieldz.Code.Text, "DOCVARIABLE") Then
            splitz = split(fieldz.Code.Text, " ")

            For Each element In splitz
                If InStr(element, breakString) Then                  'Found a CPNM address docvar
                    strAddress = element                             'Found the strAddress which is the docvar name
                    Exit For
                End If
            Next element

            On Error Resume Next
            dummy = ActiveDocument.Variables(strAddress).value

            If Not Err Then
                docVarsInDocument.Add "found", strAddress
            End If

            On Error GoTo 0

        End If
    Next fieldz

    For Each docVarz In ActiveDocument.Variables
        If Not isInCollection(docVarsInDocument, docVarz.Name) And InStr(docVarz.Name, breakString) Then
            strReport = strReport & vbCr & docVarz.Name & "   : " & docVarz.value
            docVarz.Delete
        End If
    Next docVarz

    If showMsgBoxes Then MsgBox strReport
    On Error GoTo 0

    Exit Sub

cleanUpWholeDocument_Error:
    Call handleMyError
End Sub

Private Sub createTheNonExistanteDocVars()
    ' This sub creates the docvars in the documents fields

    Dim thisField                                 As field
    Dim thisText                                  As String
    Dim thisDocVarName                            As String
    Dim thisVariant                               As Variant

    For Each thisField In ActiveDocument.Fields
        thisText = thisField.Code.Text
        If UCase(thisText) Like "*DOCVARIABLE*" Then
            thisVariant = split(thisText, " ")
            thisDocVarName = thisVariant(2)

            On Error Resume Next
            ActiveDocument.Variables.Add (thisDocVarName)            'Ignore error if docvar exists
            On Error GoTo 0
        End If
    Next thisField

End Sub

Private Sub getDocVarsFromDatabase(Optional showMsgBoxes As Boolean = True)
    ' This sub gets data from the database
    Dim colDictionary                             As Collection
    Dim problemCount                              As Integer
    Dim docVarz                                   As Variant
    Dim valueToDocVar                             As String

    Debug.Assert Application.Name = "Microsoft Word"

    On Error GoTo getDocVarsFromDatabase_Error

    Set colDictionary = New Collection

    ' First, I will add all docVariables which are not created
    Call createTheNonExistanteDocVars

    ' Starting the number of blown up variables
    problemCount = 0

    ' This is the main loop, all magic goes here
    For Each docVarz In ActiveDocument.Variables
        If docVarz.Name Like "*" & breakString & "*" Then
            ' sending the address string (which is the docvar name) to the getData function
            valueToDocVar = getData(docVarz.Name)

            ' Getting the value into the docvar
            If valueToDocVar Like "*Erro!*" And docVarz.value <> " " And Not docVarz.value Like "*Erro!*" Then
                Select Case MsgBox("O valor da propriedade do item será substituida por um valor de Erro de importação." _
                                 & vbCrLf & "Você deseja realizar a substituição ?" _
                                 & vbCrLf & "" _
                                 & vbCrLf & "Se você deseja preservar o valor atual, selecione Não." _
                     , vbYesNo Or vbQuestion Or vbDefaultButton2, Application.Name)
                    Case vbYes
                        docVarz.value = valueToDocVar
                End Select
            Else
                docVarz.value = valueToDocVar
            End If

            ' Checking if the value returned was null
            If InStr(valueToDocVar, "Erro!") Then
                problemCount = problemCount + 1
            End If
        End If
    Next docVarz

    ' Now, I update the whole document.
    ActiveDocument.Fields.Update

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
    Dim itemKey                                   As Long
    Dim propKey                                   As Long
    Dim unitKey                                   As Long
    Dim dummy                                     As Variant
    Dim strNewAddress                             As String
    Dim colDictionary                             As Collection
    Dim itemTrackingPropKey                       As String

    Debug.Assert Application.Name = "Microsoft Word"
    On Error GoTo createReference_Error

    ' checking if the thing is a tracking reference
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
    dummy = ActiveDocument.Variables(strNewAddress)

    If Err.Number = 0 Then                                           ' docvar exists!
        MsgBox "I will now create the field in the document, but the docvar " & strNewAddress & " already exists"

        'Creating the field in the document
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                             "DOCVARIABLE " & strNewAddress, PreserveFormatting:=True
    Else
        'Adding the variable into the word variables
        'MsgBox "I will now create the docvar " & strNewDocVarName & " and the field in the document"
        ActiveDocument.Variables.Add strNewAddress

        ' Creating the field in the document
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                             "DOCVARIABLE " & strNewAddress, PreserveFormatting:=True
    End If

    'Updating all the docvars
    Call getDataFromDatabase
    'Updating the fields
    ActiveDocument.Fields.Update
    On Error GoTo 0

    Exit Sub

createReference_Error:
    Call handleMyError

End Sub

Sub changeTheReferences(strOriginalItemName, strNewItemName, _
                        Optional bolColor As Boolean = False, _
                        Optional onlySelectedArea As Boolean = False, _
                        Optional updateAll As Boolean = True)
    ' This sub will change the name of the DocVars

    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim originalItemKey                           As Long
    Dim newItemKey                                As Long
    Dim changeCount                               As Integer
    Dim problemCount                              As Integer
    Dim docVarz                                   As Variant
    Dim splitz                                    As Variant
    Dim itemKey                                   As Long
    Dim propKeyExtended                           As Variant         'the propKey here can refer to the tracking property primary key! which is a string.
    Dim unitKey                                   As Variant
    Dim oldDocVarName                             As String
    Dim newDocVarName                             As String
    Dim dummy                                     As Variant
    Dim field                                     As Variant
    Dim fieldIndex                                As Integer
    Dim newCode                                   As String
    Dim thisSelectionBkp                          As Range

    Debug.Assert Application.Name = "Microsoft Word"

    ' Here, I put the connect string
    On Error GoTo changeTheReferences_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Now, i write the query to give me the primaryKey of the original item
    strSQL = "select ID_ITEM from ITEM where NOME_ITEM = " & Chr(34) & strOriginalItemName & Chr(34)

    ' Now, I open this query
    rs.Open strSQL, gCnn

    ' Checking if something weird just happenned
    Call checkKeyness(rs)

    ' Now, I will get the item key from this query
    originalItemKey = rs.Fields("ID_ITEM").value

    ' Closing the recordset so it can receive another query
    rs.Close

    ' Now, I will start another query to give me the new item key
    strSQL = "select ID_ITEM from ITEM where NOME_ITEM = " & Chr(34) & strNewItemName & Chr(34)

    ' Now, I open this query
    rs.Open strSQL, gCnn

    ' Checking if something weird just happenned
    Call checkKeyness(rs)

    ' Now, I will get the item key from this query
    newItemKey = rs.Fields("ID_ITEM").value

    ' Now, for each docVar that points into the originalItemKey, i will change it to the newItemKey
    changeCount = 0                                                  'counting the number of changes
    problemCount = 0                                                 'counting the number of null new fields
    Set thisSelectionBkp = Selection.Range.Duplicate                 'getting current selection

    For Each docVarz In ActiveDocument.Variables
        ' here i break the docvarz names
        If InStr(docVarz.Name, breakString) Then
            splitz = split(docVarz.Name, breakString)
            itemKey = CLng(splitz(1))
            propKeyExtended = splitz(2)                              ' Importante! Aqui o propKey carrega o caractere de especificação! (se é só unidade, tracking, etc.)
            unitKey = splitz(3)

            ' here I check if the active docvar points to the original key
            If itemKey = originalItemKey Then
                'here I get the hold of the new and the old docvar names
                oldDocVarName = docVarz.Name
                newDocVarName = breakString & newItemKey & breakString & propKeyExtended & breakString & unitKey

                ' Now, i need to create the new docVar. They wont create themselves.
                On Error Resume Next
                dummy = ActiveDocument.Variables.Add(newDocVarName, docVarz.value)
                On Error GoTo 0

                ' Now, I need to change the fields inside the document! The simples change of docvars do not change the fields.
                If Not onlySelectedArea Then
                    For Each field In ActiveDocument.Fields          ' Loop through whole document
                        fieldIndex = field.index

                        If InStr(field.Code, oldDocVarName) Then     'here i check if the old doc var name appears into the field code
                            'Replacing the old docVar name in the field code by the new one
                            newCode = Replace(field.Code, oldDocVarName, newDocVarName)

                            'Getting the code with the replaced docvar name
                            ActiveDocument.Fields.Item(fieldIndex).Code.Text = newCode

                            ' coloring the field if we want to
                            If bolColor Then
                                ActiveDocument.Fields.Item(fieldIndex).Select
                                Selection.Range.HighlightColorIndex = wdYellow
                            End If

                            changeCount = changeCount + 1
                        End If
                    Next field
                Else

                    For Each field In thisSelectionBkp.Fields        ' Loop through selection only
                        If InStr(field.Code, oldDocVarName) Then     'here i check if the old doc var name appears into the field code
                            'Replacing the old docVar name in the field code by the new one
                            newCode = Replace(field.Code, oldDocVarName, newDocVarName)

                            'Getting the code with the replaced docvar name
                            field.Code.Text = newCode

                            ' coloring the field if we want to
                            If bolColor Then
                                field.Select
                                Selection.Range.HighlightColorIndex = wdYellow
                            End If

                            changeCount = changeCount + 1
                        End If
                    Next
                End If
            End If
        End If
    Next docVarz

    ' Voilá my good friend =)
    ' And we clean the thing up
    Call cleanUpWholeDocument(False)
    ' And for the glory of CopyPasteNoMore: Update what is there
    If updateAll Then Call getDocVarsFromDatabase
    On Error GoTo 0

    Exit Sub

changeTheReferences_Error:
    Call handleMyError
End Sub

Sub populateOriginalItem(cmbOriginalItemName, Optional strItemClass As String, _
                         Optional strSubArea As String)
    ' This sub populate the "original item" combobox in the "change reference form"

    Debug.Assert Application.Name = "Microsoft Word"

    ' Declarations
    Dim itemName                                  As String
    Dim splited                                   As Variant
    Dim itemKey                                   As Long
    Dim subAreaKey                                As Long
    Dim itemClassKey                              As Long
    Dim subAreaKeyDesired                         As Long
    Dim itemClassKeyDesired                       As Long
    Dim bolSentinel                               As Boolean
    Dim docVarz                                   As Variant
    Dim element                                   As Variant
    Dim itemNameCol                               As Collection

    ' Here, I put the connect string
    On Error GoTo populateOriginalItem_Error

    ' Initiating the itemName collection, which will be used to populate the combo.
    Set itemNameCol = New Collection

    ' getting the filtered keys
    If strItemClass <> "" Then itemClassKeyDesired = getItemClassKey(strItemClass)
    If strSubArea <> "" Then subAreaKeyDesired = getSubAreaKey(strSubArea)

    If ActiveDocument.Variables.Count > 0 Then
        For Each docVarz In ActiveDocument.Variables
            If InStr(docVarz.Name, breakString) Then
                ' parsing the address to the value
                splited = split(docVarz.Name, breakString)

                ' getting the item name
                itemKey = CLng(splited(1))
                itemName = getItemName(itemKey)

                ' getting the auxiliary keys
                subAreaKey = getSubAreaKeyFromItemKey(itemKey)
                itemClassKey = getItemClassKeyFromItemKey(itemKey)

                ' Checking the aditional filters
                bolSentinel = True                                   ' the bolSentinel is a watch for constraint violation.
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
        Next docVarz
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

Sub populateNewItem(originalItemName, cmbNewItemName, Optional strSubArea As String)
    ' This sub populate the newItem combobox given the type of the OriginalItem

    ' Declarations
    Debug.Assert Application.Name = "Microsoft Word"

    Dim rs                                        As ADODB.Recordset
    Dim strSQL                                    As String
    Dim originalItemTypeKey                       As Long
    Dim originalItemKey                           As Long
    Dim subAreaKey                                As Long

    On Error GoTo populateNewItem_Error

    ' And i create the recordset to receive the query
    Set rs = New ADODB.Recordset

    ' Getting the keys
    originalItemKey = getItemKey(originalItemName)
    originalItemTypeKey = getItemTypeFromItemKey(originalItemKey)

    'Now, I will query over all the items with the itemtype I got above
    strSQL = "Select NOME_ITEM from ITEM where ID_TIPO_ITEM = " & originalItemTypeKey

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
