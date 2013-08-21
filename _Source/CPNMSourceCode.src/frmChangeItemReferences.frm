VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeItemReferences 
   Caption         =   "Reuso de Referências - MC a partir de template em 10 cliques no máximo."
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   OleObjectBlob   =   "frmChangeItemReferences.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangeItemReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBatchIt_Click()
    Me.Hide
    frmChangeBatch.Show
    Unload Me
End Sub

Private Sub userform_initialize()
' This sub populates the combobox with the names of the referenced items in the documents
    Call populateOriginalItem(ltbOriginalItemName)

    Call populateItemClass(cmbItemClassOrig)

    Call populateSubArea(cmbSubAreaOrig)
    Call populateSubArea(cmbSubAreaNew)
End Sub

' ##############
' Buttons

Private Sub btnChangeReferences_Click()
' This sub will then issue the command to change the references
    If ltbOriginalItemName.value <> Empty And ltbNewItemName.value <> Empty Then

        Call changeTheReferences(ltbOriginalItemName.value, ltbNewItemName.value, _
                                 ckbColorChanges.value, ckbSelectedAreaOnly.value, False)
        ' #Region# Bad Idea
        'Call ltbOriginalItemName.Clear
        'Call populateOriginalItem(ltbOriginalItemName) 're-populating the list of existant items
        ' #End Region# Bad Idea
    Else
        MsgBox "Especificação incompleta"
    End If
End Sub

' End of buttons
' ##############

' ##############
' Populating

Private Sub cmbItemClassOrig_Change()
    ltbOriginalItemName.Clear
    Call populateOriginalItem(ltbOriginalItemName, cmbItemClassOrig.value, cmbSubAreaOrig.value)
End Sub

Private Sub cmbSubAreaOrig_Change()
    ltbOriginalItemName.Clear
    Call populateOriginalItem(ltbOriginalItemName, cmbItemClassOrig.value, cmbSubAreaOrig.value)
End Sub

Private Sub cmbSubAreaNew_Change()
    ltbNewItemName.Clear
    If ltbOriginalItemName.value <> Empty Then
        Call populateNewItem(ltbOriginalItemName.value, ltbNewItemName, cmbSubAreaNew.value)
    End If
End Sub

Private Sub ltbOriginalItemName_Change()
' This sub populates the newItem combobox in accordance with the items type in the originalItem combobox
    ltbNewItemName.Clear
    If ltbOriginalItemName.value <> Empty Then
        Call populateNewItem(ltbOriginalItemName.value, ltbNewItemName, cmbSubAreaNew.value)
    End If
End Sub

' End of Populating
' ##############
