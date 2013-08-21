VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRunManager 
   Caption         =   "Chemtech's CopyPasteNuncaMais"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   OleObjectBlob   =   "frmRunManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRunManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConfigMetadata_Click()
    Me.Hide
    frmConfigMetadata.Show
    Me.Show
End Sub

Private Sub userform_Terminate()
    '    MsgBox "Terminando a seção" & vbCr & vbCr & "Obrigado por usar o Copy Paste Nunca Mais"
    'If gCnn.State = 1 Then gCnn.Close
End Sub

Private Sub btnChangeReference_Click()
    ' This sub activates the form to create references
    Me.Hide
    frmChangeItemReferences.Show vbModeless
    Unload Me
End Sub

Private Sub btnCreateConnection_Click()
    Call resetEnvironment
End Sub

Private Sub btnInsertSingleValue_Click()
    ' This sub activates the form to create references
    Me.Hide
    frmInsertDataManually.Show
    Me.Show
End Sub

Private Sub btnNewReference_Click()
    ' This sub activates the form to create references
    Me.Hide
    frmInsertBigCrossRef.Show vbModeless
    Unload Me
End Sub

Private Sub btnDownloadData_Click()
    Call getDataFromDatabase
    If Not Err Then MsgBox "Download dos dados do Database: Sucesso!"
    Unload Me
End Sub

Private Sub cmbCriaCompart_Click()
    Me.Hide
    frmShareProps.Show
    Me.Show
End Sub
