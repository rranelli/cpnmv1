VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeBatch 
   Caption         =   "Savior Batch Gen"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9225
   OleObjectBlob   =   "frmChangeBatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangeBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################
'##### This is what is different relative to the simple one #####
'################################################################

Private Sub btnAddToList_Click()
    ltbTargetList.AddItem (ltbNewItemName.value)
End Sub

Private Sub btnRemoveFromList_Click()
    If ltbTargetList.value <> Empty Then
        ltbTargetList.RemoveItem (ltbTargetList.ListIndex)
    End If
End Sub

Private Sub btnChangeReferences_Click()
' This sub will then issue the command to change the references as a batch process
    If ltbOriginalItemName.value <> Empty And ltbTargetList.ListCount > 0 Then
        Dim thisDocx As Word.Document
        Dim objFileSys As Object

        ' Making the file system object so I can copy without using the bad vba's FileCopy method.
        Set objFileSys = CreateObject("Scripting.FileSystemObject")

        ' Preparing the environment.
        originalItemName = ltbOriginalItemName.value
        originalDocPath = ActiveDocument.Path
        originalDocName = ActiveDocument.Name
        originalDocFullPath = ActiveDocument.FullName
        varz = split(originalDocFullPath, ".")
        originalDocExtension = varz(1)
        
        For selectedIndex = 1 To ltbTargetList.ListCount
            newItemName = ltbTargetList.List(selectedIndex - 1)
            newDocFullPath = originalDocPath & "\" & originalDocName & " - " & newItemName & "." & originalDocExtension
            
            ' Here, I will copy this file into a new one.
            objFileSys.CopyFile originalDocFullPath, newDocFullPath, True
        
            ' ORSC - Open, Run, Save, Close
            Set thisDocx = Word.Documents.Open(newDocFullPath)
            
            Call thisDocx.Application.Run("changeTheReferences", originalItemName, newItemName, ckbColorChanges.value)
            Call thisDocx.SaveAs
            Call thisDocx.Close
        Next
    Else
        MsgBox "Especificação incompleta"
    End If
    
    ' Finishing the execution
    If Err = 0 Then MsgBox "Batelada de documentos realizada com sucesso!"
    Unload Me
End Sub

'########################
'########################
'########################

Private Sub userform_initialize()
' This sub populates the combobox with the names of the referenced items in the documents
    Call populateOriginalItem(ltbOriginalItemName)

    Call populateItemClass(cmbItemClassOrig)

    Call populateSubArea(cmbSubAreaOrig)
    Call populateSubArea(cmbSubAreaNew)

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

