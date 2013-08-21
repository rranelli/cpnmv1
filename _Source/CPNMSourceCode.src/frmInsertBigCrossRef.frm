VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertBigCrossRef 
   Caption         =   "Gerador de Referências do CopyPasteNuncaMais"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   OleObjectBlob   =   "frmInsertBigCrossRef.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsertBigCrossRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private txbNewValue      As TextBox                                  'This is just a fix. Without it, I cant use the lump of code.
Private refOption        As Integer

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

    ' iniciando a opção padrão de unidade & valor.
    opbValueAndUnit.value = True
End Sub

'#################
' Buttons

Private Sub btnCreateReference_Click()
' This sub creates the reference by using the values defined in the comboboxes
    If Application.Name = "AutoCAD" Then Me.Hide
    If (ltbItemName.value <> Empty And ltbPropName.value <> Empty _
        And (ltbUnitName.ListCount = 0 Or (ltbUnitName.value <> Empty And ltbUnitName <> Empty))) Then
        If Not IsNull(ltbUnitName.value) Then
            Call createReference(ltbItemName.value, ltbPropName.value, ltbUnitName.value, False, refOption)
        Else
            Call createReference(ltbItemName.value, ltbPropName.value, "", False, refOption)
        End If
    Else
        MsgBox "Faltam items para especificar!"
    End If
    If Application.Name = "AutoCAD" Then Me.Show
End Sub

Private Sub btnCreateTrackingReference_Click()
' this sub creates a tracking reference
    If Application.Name = "AutoCAD" Then Me.Hide
    Call createReference(ltbItemNameTrack.value, ltbPropNameTrack.value, "0", True)
    If Application.Name = "AutoCAD" Then Me.Show
End Sub

' Subs for the selection of the reference style to be inserted
Private Sub opbValue_Click()
    refOption = 1
End Sub

Private Sub opbUnit_Click()
    refOption = 0
End Sub

Private Sub opbValueAndUnit_Click()
    refOption = 2
End Sub

' End of buttons
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
