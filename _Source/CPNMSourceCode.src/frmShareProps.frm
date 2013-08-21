VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShareProps 
   Caption         =   "Criar Compartilhamento"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14670
   OleObjectBlob   =   "frmShareProps.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShareProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private txbNewValue1     As TextBox
Private txbNewValue2     As TextBox
Dim objShareQueue        As shareQueue
Dim objBreakQueue        As shareQueue

'###########
' INIT
Private Sub userform_initialize()
' This sub populates the item type combobox
    Call populateItemType(ltbItemType1)
    Call populateItemType(ltbItemType2)
    Call populateItemType(ltbItemTypeBreak)

    Call populateItemClass(cmbItemClass1)
    Call populateItemClass(cmbItemClass2)
    Call populateItemClass(cmbItemClassBreak)

    Call populateSubArea(cmbSubArea1)
    Call populateSubArea(cmbSubArea2)
    Call populateSubArea(cmbSubAreaBreak)

    Call populatePropClass(cmbPropClass1)
    Call populatePropClass(cmbPropClass2)
    Call populatePropClass(cmbPropClassBreak)

    Set objShareQueue = New shareQueue
    Set objBreakQueue = New shareQueue
End Sub

'##################################
' BUTTONS

'#####
' Buttons at share

Private Sub btnCreateShare_Click()
' this sub creates the sharing given the options at the comboboxes
    Dim j                As Integer
    Dim strItemName1     As String
    Dim strItemName2     As String
    Dim strPropName1     As String
    Dim strPropName2     As String
    Dim arrCollection    As collection
    Dim varArray()       As Variant

    Set arrCollection = objShareQueue.sharesCollection

    For j = 1 To objShareQueue.countShares

        varArray = arrCollection(j)

        strItemName1 = varArray(0)
        strPropName1 = varArray(1)
        strItemName2 = varArray(2)
        strPropName2 = varArray(3)

        Call createSharing(strItemName1, strPropName1, strItemName2, strPropName2)
    Next j

    MsgBox "Execução Completada"
End Sub

Private Sub btnAddToList_Click()

    Dim dimKey1          As Long
    Dim dimKey2          As Long
    ' checking if the adding is valid
    If ltbItemName1.value <> Empty And ltbItemName2.value <> Empty _
       And ltbPropName1.value <> Empty And ltbPropName2.value <> Empty Then

        dimKey1 = getDimKey(getPropKey(ltbPropName1.value))
        dimKey2 = getDimKey(getPropKey(ltbPropName2.value))

        If dimKey1 = dimKey2 Then
            Call objShareQueue.enqueue(ltbItemName1.value, ltbPropName1.value, ltbItemName2.value, ltbPropName2.value)
            Call addToShareList(ltbCreationQueue, ltbItemName1.value, ltbItemName2.value, ltbPropName1.value, ltbPropName2.value)
        Else
            MsgBox "Você está tentando adicionar à lista propriedades cujas dimensões não concordam."
        End If
    Else
        MsgBox "especificação incompleta"
    End If
End Sub

Private Sub btnCleanList_Click()
    objShareQueue.Clear
    ltbCreationQueue.Clear
End Sub

Private Sub btnRemoveFromList_Click()
    Dim indexClicked     As Integer

    If objShareQueue.countShares > 0 And ltbCreationQueue.ListIndex + 1 <> Empty Then

        indexClicked = Int(ltbCreationQueue.ListIndex / 3)

        ' deleting from the listbox
        Call removeFromShareList(ltbCreationQueue, indexClicked)
        ' dequeuing the clicked value
        objShareQueue.dequeue (indexClicked)
    End If
End Sub

Private Sub btnSaveList_Click()
    Call checkShareCollection(objShareQueue.sharesCollection)
    MsgBox "execução completa"
End Sub

Private Sub btnLoadShare_Click()

    Dim varArray         As Variant
    Dim strPropName1     As String
    Dim strPropName2     As String
    Dim strItemName1     As String
    Dim strItemName2     As String
    Dim shareCollection  As collection
    Dim j                As Integer


    If ltbSavedShares.value <> Empty And ltbItemName1.value <> Empty And ltbItemName2.value <> Empty Then
        Set objShareQueue = loadShareCollection(ltbSavedShares.value, ltbItemName1.value, ltbItemName2.value)
        Set shareCollection = objShareQueue.sharesCollection

        For j = 1 To shareCollection.Count
            varArray = shareCollection(j)

            strItemName1 = varArray(0)
            strPropName1 = varArray(1)
            strItemName2 = varArray(2)
            strPropName2 = varArray(3)

            Call addToShareList(ltbCreationQueue, strItemName1, strItemName2, strPropName1, strPropName2)
        Next j

    Else
        MsgBox "especificação incompleta"
    End If
End Sub

'#####
'Buttons at break
Private Sub btnBreakShare_Click()
' this sub creates the sharing given the options at the comboboxes
    Dim j                As Integer
    Dim strItemName1     As String
    Dim strItemName2     As String
    Dim strPropName1     As String
    Dim strPropName2     As String
    Dim arrCollection    As collection
    Dim varArray()       As Variant

    Set arrCollection = objBreakQueue.sharesCollection

    For j = 1 To objBreakQueue.countShares

        varArray = arrCollection(j)

        strItemName1 = varArray(0)
        strPropName1 = varArray(1)
        strItemName2 = varArray(2)
        strPropName2 = varArray(3)

        Call breakSharing(strItemName1, strPropName1, strItemName2, strPropName2)
    Next j

    MsgBox "Execução Completada"
End Sub

Private Sub btnAddToListBreak_Click()

    If ltbItemNameOrig.value <> Empty And ltbPropNameOrig.value <> Empty _
       And ltbItemNameComp.value <> Empty And txbPropNameComp.value <> Empty Then
        Call objBreakQueue.enqueue(ltbItemNameOrig.value, ltbPropNameOrig.value, ltbItemNameComp.value, txbPropNameComp.value)
        Call addToShareList(ltbBreakQueue, ltbItemNameOrig.value, ltbPropNameOrig.value, ltbItemNameComp.value, txbPropNameComp.value)
    Else
        MsgBox "especificação incompleta"
    End If
End Sub

Private Sub btnCleanListBreak_Click()
    objBreakQueue.Clear
    ltbBreakQueue.Clear
End Sub

Private Sub btnRemoveFromListBreak_Click()
    Dim indexClicked     As Integer

    If objBreakQueue.countShares > 0 And ltbBreakQueue.ListIndex + 1 <> Empty Then

        indexClicked = Int(ltbBreakQueue.ListIndex / 3)

        ' deleting from the listbox
        Call removeFromShareList(ltbBreakQueue, indexClicked)
        ' dequeuing the clicked value
        objBreakQueue.dequeue (indexClicked)
    End If
End Sub

Private Sub btnLoadBreak_Click()

    Dim varArray         As Variant
    Dim strPropName1     As String
    Dim strPropName2     As String
    Dim strItemName1     As String
    Dim strItemName2     As String
    Dim shareCollection  As collection
    Dim j                As Integer


    If ltbSavedSharesBreak.value <> Empty And ltbItemNameComp.value <> Empty And ltbItemNameOrig.value <> Empty Then
        Set objBreakQueue = loadShareCollection(ltbSavedSharesBreak.value, ltbItemNameOrig.value, ltbItemNameComp.value)
        Set shareCollection = objBreakQueue.sharesCollection

        For j = 1 To shareCollection.Count
            varArray = shareCollection(j)

            strItemName1 = varArray(0)
            strPropName1 = varArray(1)
            strItemName2 = varArray(2)
            strPropName2 = varArray(3)

            Call addToShareList(ltbBreakQueue, strItemName1, strItemName2, strPropName1, strPropName2)
        Next j

    Else
        MsgBox "especificação incompleta"
    End If
End Sub

' ####################################################
' Start of the property location lump of code <frame 1>

Private Sub ckbExistantProps1_Click()
    If ckbExistantProps1.value = True Then                           'O click resolve antes da mudança.
        If ltbItemName1.value <> Empty Then
            ltbPropName1.Clear
            Call populateExistantProps(ltbItemName1.value, ltbPropName1, cmbPropClass1.value)
        End If
    Else
        If ltbItemType1.value <> Empty Then
            ltbPropName1.Clear
            Call populateProps(ltbItemType1.value, ltbPropName1, cmbPropClass1.value)
        End If
    End If
End Sub

Private Sub cmbItemClass1_Change()
    ltbItemType1.Clear
    Call populateItemType(ltbItemType1, cmbItemClass1.value)
End Sub

Private Sub ltbItemType1_Change()
' This sub populates item names and property names on their comboboxes
    ltbItemName1.Clear
    ltbPropName1.Clear
    txbCurrentValue1.Text = ""
    If Not txbNewValue1 Is Nothing Then txbNewValue1.Text = ""
    If ltbItemType1.value <> Empty Then
        If ckbExistantProps1.value = False Then
            Call populateItems(ltbItemType1.value, ltbItemName1, cmbSubArea1.value)
            Call populateProps(ltbItemType1.value, ltbPropName1, cmbPropClass1.value)
        Else
            Call populateItems(ltbItemType1.value, ltbItemName1, cmbSubArea1.value)
        End If
    End If

    ltbSavedShares.Clear
    If ltbItemType1.value <> Empty And ltbItemType2.value <> Empty Then
        Call populateSavedShares(ltbSavedShares, ltbItemType1.value, ltbItemType2.value)
    End If
End Sub

Private Sub ltbItemName1_Change()
' This sub populates item names and property names on their comboboxes
    Call showCurrentData1
    If ckbExistantProps1.value And ltbItemName1.value <> Empty Then
        ltbPropName1.Clear
        Call populateExistantProps(ltbItemName1.value, ltbPropName1, cmbPropClass1.value)
    End If
End Sub

Private Sub ltbPropName1_Change()
    Call showCurrentData1
End Sub

Private Sub cmbSubArea1_Change()
    If ltbItemType1.value <> Empty Then
        ltbItemName1.Clear
        Call populateItems(ltbItemType1.value, ltbItemName1, cmbSubArea1.value)
    End If
End Sub

Private Sub cmbPropClass1_Change()
    ltbPropName1.Clear
    If ltbItemType1.value <> Empty Then
        If ckbExistantProps1.value And ltbItemName1.value <> Empty Then
            Call populateExistantProps(ltbItemName1.value, ltbPropName1, cmbPropClass1.value)
        Else
            Call populateProps(ltbItemType1.value, ltbPropName1, cmbPropClass1.value)
        End If
    End If
End Sub

' End of the normal property lump of code <frame 1>
' ####################################################

' ####################################################
' Start of the property location lump of code <frame 2>

Private Sub ckbExistantProps2_Click()
    If ckbExistantProps2.value = True Then                           'O click resolve antes da mudança.
        If ltbItemName2.value <> Empty Then
            ltbPropName2.Clear
            Call populateExistantProps(ltbItemName2.value, ltbPropName2, cmbPropClass2.value)
        End If
    Else
        If ltbItemType2.value <> Empty Then
            ltbPropName2.Clear
            Call populateProps(ltbItemType2.value, ltbPropName2, cmbPropClass2.value)
        End If
    End If
End Sub

Private Sub cmbItemClass2_Change()
    ltbItemType2.Clear
    Call populateItemType(ltbItemType2, cmbItemClass2.value)
End Sub

Private Sub ltbItemType2_Change()
' This sub populates item names and property names on their comboboxes
    ltbItemName2.Clear
    ltbPropName2.Clear
    txbCurrentValue2.Text = ""
    If Not txbNewValue2 Is Nothing Then txbNewValue2.Text = ""
    If ltbItemType2.value <> Empty Then
        If ckbExistantProps2.value = False Then
            Call populateItems(ltbItemType2.value, ltbItemName2, cmbSubArea2.value)
            Call populateProps(ltbItemType2.value, ltbPropName2, cmbPropClass2.value)
        Else
            Call populateItems(ltbItemType2.value, ltbItemName2, cmbSubArea2.value)
        End If
    End If

    ltbSavedShares.Clear
    If ltbItemType1.value <> Empty And ltbItemType2.value <> Empty Then
        Call populateSavedShares(ltbSavedShares, ltbItemType1.value, ltbItemType2.value)
    End If
End Sub

Private Sub ltbItemName2_Change()
' This sub populates item names and property names on their comboboxes
    Call showCurrentData2
    If ckbExistantProps2.value And ltbItemName2.value <> Empty Then
        ltbPropName2.Clear
        Call populateExistantProps(ltbItemName2.value, ltbPropName2, cmbPropClass2.value)
    End If
End Sub

Private Sub ltbPropName2_Change()
    Call showCurrentData2
End Sub

Private Sub cmbSubArea2_Change()
    If ltbItemType2.value <> Empty Then
        ltbItemName2.Clear
        Call populateItems(ltbItemType2.value, ltbItemName2, cmbSubArea2.value)
    End If
End Sub

Private Sub cmbPropClass2_Change()
    ltbPropName2.Clear
    If ltbItemType2.value <> Empty Then
        If ckbExistantProps2.value And ltbItemName2.value <> Empty Then
            Call populateExistantProps(ltbItemName2.value, ltbPropName2, cmbPropClass2.value)
        Else
            Call populateProps(ltbItemType2.value, ltbPropName2, cmbPropClass2.value)
        End If
    End If
End Sub

' End of the normal property lump of code <Frame 2>
' ####################################################

'#####################################
' Start of the Break Sharing From Code

Private Sub ltbItemTypeBreak_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' This sub populates item names and property names on their comboboxes
    If KeyCode = 13 Then                                             '13 = Enter / Return
        ltbItemName.Clear
        ltbPropName.Clear
        Call populateItems(ltbItemTypeBreak.value, ltbItemNameBreak)
        Call populateProps(ltbItemTypeBreak.value, ltbPropNameBreak)
    End If
End Sub

Private Sub ltbItemTypeBreak_Click()
' This sub populates item names and property names on their comboboxes
    ltbItemNameOrig.Clear
    ltbPropNameOrig.Clear
    Call populateItems(ltbItemTypeBreak.value, ltbItemNameOrig)
End Sub

Private Sub ltbItemNameOrig_Change()
' This sub populates item names and property names on their comboboxes
    ltbItemNameComp.Clear
    ltbPropNameOrig.Clear
    If ltbItemNameOrig.value <> Empty Then
        Call populateShares(ltbItemNameOrig.value, ltbItemNameComp)
    End If

    If ltbItemNameOrig.value <> Empty Then
        Call populateSharedProps(ltbItemNameOrig.value, ltbPropNameOrig, cmbPropClassBreak.value)
    End If
End Sub

Private Sub ltbPropNameOrig_Change()
' This sub populates item names and property names on their comboboxes
    txbPropNameComp.value = ""
    Call showCurrentShare

    If ltbItemNameOrig.value <> Empty Then
        ltbItemNameComp.Clear
        If ltbPropNameOrig.value <> Empty Then
            Call populateShares(ltbItemNameOrig.value, ltbItemNameComp, ltbPropNameOrig.value)
        Else
            Call populateShares(ltbItemNameOrig.value, ltbItemNameComp)
        End If
    End If

    ltbSavedSharesBreak.Clear
    If ltbItemNameComp.value <> Empty And ltbItemNameOrig.value <> Empty Then
        Call populateSavedShares(ltbSavedSharesBreak, getItemTypeFromItemKey(getItemKey(ltbItemNameComp.value)) _
                                                      , getItemTypeFromItemKey(getItemKey(ltbItemNameOrig.value)))
    End If
End Sub

Private Sub ltbItemNameComp_Change()
    txbPropNameComp.value = ""
    Call showCurrentShare

    ltbSavedSharesBreak.Clear
    If ltbItemNameComp.value <> Empty And ltbItemNameOrig.value <> Empty Then
        Call populateSavedSharesFromKeys(ltbSavedSharesBreak, getItemTypeFromItemKey(getItemKey(ltbItemNameComp.value)) _
                                                              , getItemTypeFromItemKey(getItemKey(ltbItemNameOrig.value)))
    End If
End Sub

' End of the Break Sharing From Code
'#####################################

'#########
' Auxiliary
Private Sub showCurrentShare()
    Dim itemKey          As Long
    Dim propKey          As Long
    Dim valueKey         As Long
    Dim itemKeyComp      As Long

    If ltbPropNameOrig.value <> Empty And ltbItemNameComp.value <> Empty Then
        itemKey = getItemKey(ltbItemNameOrig.value)                  'itemkey/propkey is the original item/prop pair.
        propKey = getPropKey(ltbPropNameOrig.value)
        itemKeyComp = getItemKey(ltbItemNameComp.value)

        If checkValueExistance(itemKey, propKey) Then
            valueKey = getValueKey(itemKey, propKey)
            On Error Resume Next                                     'se não existir o compartilhamento, não quero erro.
            txbPropNameComp.value = getShareName(valueKey, itemKeyComp, itemKey, propKey)
        End If
    Else
        txbPropNameComp.value = ""
    End If
End Sub

Private Sub showCurrentData1()
    Dim propKey          As Variant                                  'propKey may receive track or normal type of reference.
    Dim itemKey          As Long
    Dim unitKey          As Long
    Dim colDictionary    As collection

    txbCurrentValue1.value = ""

    If (ltbPropName1.value <> Empty) And (ltbItemName1.value <> Empty) Then
        propKey = getPropKey(ltbPropName1.value)
        Set colDictionary = createTrackingDictionary

        itemKey = getItemKey(ltbItemName1.value)
        propKey = getPropKey(ltbPropName1.value)

        unitKey = 0
        If getDimKey(CLng(propKey)) <> 0 Then unitKey = 1

        txbCurrentValue1.value = getDataFromDB(itemKey, CLng(propKey), unitKey, 2, False)    'This "2" makes the unit show up in the value
    End If
End Sub

Private Sub showCurrentData2()
    Dim propKey          As Variant                                  'propKey may receive track or normal type of reference.
    Dim itemKey          As Long
    Dim unitKey          As Long
    Dim colDictionary    As collection

    txbCurrentValue2.value = ""

    If (ltbPropName2.value <> Empty) And (ltbItemName2.value <> Empty) Then
        propKey = getPropKey(ltbPropName2.value)
        Set colDictionary = createTrackingDictionary

        itemKey = getItemKey(ltbItemName2.value)
        propKey = getPropKey(ltbPropName2.value)

        unitKey = 0
        If getDimKey(CLng(propKey)) <> 0 Then unitKey = 1

        txbCurrentValue2.value = getDataFromDB(itemKey, CLng(propKey), unitKey, 2, False)    'This "2" makes the unit show up in the value
    End If
End Sub

