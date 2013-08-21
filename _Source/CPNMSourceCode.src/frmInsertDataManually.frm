VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertDataManually 
   Caption         =   "Inserir Dado Simples no Banco de Dados"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   OleObjectBlob   =   "frmInsertDataManually.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsertDataManually"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_initialize()
' populando para a aba de referência a valor
    Call populateItemType(ltbItemType)
    Call populateItemClass(cmbItemClass)
    Call populateSubArea(cmbSubArea)
    Call populatePropClass(cmbPropClass)
    ' populando para a aba de referência a rastreamento
    Call populateItemClass(cmbItemClassTrack)
    Call populateSubArea(cmbSubAreaTrack)
    Call populateItemType(ltbItemTypeTrack)
    ' populando a sub para criação do item
    Call populateSubArea(ltbSubArea)
    Call populateItemClass(cmbItemClassCreate)
    Call populateItemType(ltbItemTypeCreate)
End Sub

' ############
' BUTTONS

Private Sub btnExportData_Click()
' This sub is used to export data to the database, given the names especification
    If Not IsNull(ltbUnitName.value) Then
        Call exportSingleData(ltbItemName.value, ltbPropName.value, txbValue.value, ltbUnitName.value, True)
    Else
        Call exportSingleData(ltbItemName.value, ltbPropName.value, txbValue.value, "", True)
    End If
    MsgBox "Dado exportado com sucesso!"
End Sub

Private Sub btnChangeTrackingProperty_Click()
    If txbNewValue.value <> "" Then
        Call changeTrackingData(ltbItemNameTrack.value, ltbPropNameTrack.value, txbNewValue.value)
        If Not Err Then MsgBox "Referência de rastreamento alterada com sucesso"

        ltbItemNameTrack.Clear
        Call populateItems(ltbItemTypeTrack.value, ltbItemNameTrack, cmbSubArea.value)
    End If
End Sub

Private Sub btnCreateNewItem_Click()
' This sub creates a new item using the defined values in their comboboxes
    If checkItemExistance(txbNewItemName.value) = False Then
        Call createItem(txbNewItemName.value, ltbItemTypeCreate.value, ltbSubArea.value)
        MsgBox "Item criado com sucesso!"
    Else
        MsgBox "Já existe um Item com este nome cadastrado!"
    End If
    'Unload Me
End Sub

' End of buttons
' ##############

' ##############
' create
Private Sub cmbItemClassCreate_Change()
    ltbItemTypeCreate.Clear
    Call populateItemType(ltbItemTypeCreate, cmbItemClassCreate.value)
End Sub

' End of create
' ##############

' ####################################################
' Start of the property location lump of code

Private Sub ckbExistantProps_Click()
    If ckbExistantProps.value = True Then                            'O click resolve antes da mudança.
        If ltbItemName.value <> Empty Then
            ltbPropName.Clear
            Call populateExistantProps(ltbItemName.value, ltbPropName, cmbPropClass.value)
        End If
    Else
        If ltbItemType.value <> Empty Then
            ltbPropName.Clear
            Call populateProps(ltbItemType.value, ltbPropName, cmbPropClass.value)
        End If
    End If
End Sub

Private Sub cmbItemClass_Change()
    ltbItemType.Clear
    Call populateItemType(ltbItemType, cmbItemClass.value)
End Sub

Private Sub ltbItemType_Change()
' This sub populates item names and property names on their comboboxes
    ltbItemName.Clear
    ltbPropName.Clear
    txbCurrentValue.Text = ""
    If Not txbNewValue Is Nothing Then txbNewValue.Text = ""
    If ltbItemType.value <> Empty Then
        If ckbExistantProps.value = False Then
            Call populateItems(ltbItemType.value, ltbItemName, cmbSubArea.value)
            Call populateProps(ltbItemType.value, ltbPropName, cmbPropClass.value)
        Else
            Call populateItems(ltbItemType.value, ltbItemName, cmbSubArea.value)
        End If
    End If
End Sub

Private Sub ltbPropName_Change()
    ltbUnitName.Clear
    Call showCurrentData
    If ltbPropName.value <> Empty Then
        Call populateUnits(ltbPropName.value, ltbUnitName)
    End If
End Sub

Private Sub ltbUnitName_Change()
    Call showCurrentData
End Sub

Private Sub ltbItemName_Change()
' This sub populates item names and property names on their comboboxes
    Call showCurrentData
    If ckbExistantProps.value And ltbItemName.value <> Empty Then
        ltbPropName.Clear
        Call populateExistantProps(ltbItemName.value, ltbPropName, cmbPropClass.value)
    End If
End Sub

Private Sub cmbSubArea_Change()
    If ltbItemType.value <> Empty Then
        ltbItemName.Clear
        Call populateItems(ltbItemType.value, ltbItemName, cmbSubArea.value)
    End If
End Sub

Private Sub cmbPropClass_Change()
    ltbPropName.Clear
    If ltbItemType.value <> Empty Then
        If ckbExistantProps.value And ltbItemName.value <> Empty Then
            Call populateExistantProps(ltbItemName.value, ltbPropName, cmbPropClass.value)
        Else
            Call populateProps(ltbItemType.value, ltbPropName, cmbPropClass.value)
        End If
    End If
End Sub

' End of the normal property lump of code
' ####################################################

' ####################################################
' Start of the Tracking property location lump of code

Private Sub cmbItemClassTrack_Change()
    ltbItemTypeTrack.Clear
    Call populateItemType(ltbItemTypeTrack, cmbItemClassTrack.value)
End Sub

Private Sub ltbItemTypeTrack_Change()
' This sub populates item names and tracking property names on their comboboxes
    ltbItemNameTrack.Clear
    ltbPropNameTrack.Clear
    txbCurrentValue.Text = ""
    If Not txbNewValue Is Nothing Then txbNewValue.Text = ""
    If ltbItemTypeTrack.value <> Empty Then
        Call populateItems(ltbItemTypeTrack.value, ltbItemNameTrack, cmbSubArea.value)
        Call populateTrackingProps(ltbItemTypeTrack.value, ltbPropNameTrack)
    End If
End Sub

Private Sub cmbSubAreaTrack_Change()
    If ltbItemTypeTrack.value <> Empty Then
        ltbItemNameTrack.Clear
        Call populateItems(ltbItemTypeTrack.value, ltbItemNameTrack, cmbSubAreaTrack.value)
    End If
End Sub

Private Sub ltbItemNameTrack_Click()
    Call showCurrentData
End Sub

Private Sub ltbPropNameTrack_Click()
    Call showCurrentData
End Sub

' End of the tracking property location lump of code
' ####################################################

'######
' Auxiliary

Private Sub showCurrentData()
    Dim propKey          As Variant                                  'propKey may receive track or normal type of reference.
    Dim itemKey          As Long
    Dim unitKey          As Long
    Dim colDictionary    As collection

    txbCurrentValue.value = ""
    txbCurrentValueTrack.value = ""

    If (ltbPropName.value <> Empty) And (ltbItemName.value <> Empty) Then
        propKey = getPropKey(ltbPropName.value)
        If (ltbUnitName.value <> Empty) Or getDimKey(CLng(propKey)) = 0 Then
            Set colDictionary = createTrackingDictionary

            itemKey = getItemKey(ltbItemName.value)
            propKey = getPropKey(ltbPropName.value)
            If Not IsNull(ltbUnitName.value) Then _
               unitKey = getUnitKey(ltbPropName.value, ltbUnitName.value)

            txbCurrentValue.value = getDataFromDB(itemKey, CLng(propKey), unitKey, 2, False)    'This "2" makes the unit show up in the value
        End If
    End If

    If ltbPropNameTrack.value <> Empty And ltbItemNameTrack.value <> Empty Then
        Set colDictionary = createTrackingDictionary

        itemKey = getItemKey(ltbItemNameTrack.value)
        propKey = colDictionary(ltbPropNameTrack.value)
        txbCurrentValueTrack.value = getTrackingDataFromDB(itemKey, CStr(propKey), False)
    End If
End Sub

