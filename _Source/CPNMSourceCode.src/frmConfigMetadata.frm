VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfigMetadata 
   Caption         =   "Config ItemType "
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   OleObjectBlob   =   "frmConfigMetadata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfigMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Initializing the userform
Private Sub userform_initialize()
    Call populatePropClass(cmbPropTypeClass)
    Call populateItemClass(cmbItemTypeClass)
    Call populateDimension(cmbDimension)

    Call populateItemType(ltbItemType)
    Call populateAllProps(ltbPropType)
End Sub

' ###############
' Buttons

Private Sub btnAddXref_Click()
    If ltbPropType.value <> "" Then
        ltbXrefs.AddItem (ltbPropType.value)
    End If
End Sub

Private Sub btnCreateItemType_Click()
    Call createItemType(txbItemTypeName.value, cmbItemTypeClass.value)
End Sub



Private Sub btnCreatePropType_Click()
    Call createProperty(txbPropTypeName, cmbPropTypeClass.value, cmbDimension.value)
End Sub

Private Sub btnRemoveXrefs_Click()
    If ltbXrefs.value <> "" Then
        ltbXrefs.RemoveItem (ltbXrefs.ListIndex)
    End If
End Sub


Private Sub btnUpdateItemType_Click()
    ltbItemType.Clear
    Call populateItemType(ltbItemType)
End Sub

Private Sub btnUpdatePropType_Click()
    ltbPropType.Clear
    Call populateAllProps(ltbPropType)
End Sub

Private Sub ltbItemType_Change()
    If ltbItemType.value <> "" Then
        ltbXrefs.Clear
        Call populateProps(ltbItemType.value, ltbXrefs)
    End If
End Sub

Private Sub btnCommitChanges_Click()
    Dim thisPropKey                               As Long
    Dim itemTypeKey                               As Long
    Dim thisPropName                              As String
    Dim thisIndex                                 As Integer

    If (ltbItemType.value) <> "" Then
        itemTypeKey = getItemTypeKey(ltbItemType.value)

        'This loop clears all Xrefs
        Call deleteXrefs(itemTypeKey)

        'This loop adds all the Xrefs into the listbox
        For thisIndex = 0 To ltbXrefs.ListCount - 1
            thisPropName = ltbXrefs.List(thisIndex)
            thisPropKey = getPropKey(thisPropName)

            On Error Resume Next                                     'Here I expect to face non-keyness errors. I just ignore then.
            Call createXref(itemTypeKey, thisPropKey)
            On Error GoTo 0
        Next

        'reloading the Xrefs listbox
        ltbXrefs.Clear
        Call populateProps(ltbItemType.value, ltbXrefs)
    Else
        MsgBox "Incomplete Specification"
    End If
End Sub
