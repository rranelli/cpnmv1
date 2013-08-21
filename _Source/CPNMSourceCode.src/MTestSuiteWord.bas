Attribute VB_Name = "MTestSuiteWord"
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

Private Sub deleteAllDocVars()
    ' This sub is for test porpuses only. It deletes all the docvariables from the document
    Dim strReport                                 As String
    Dim docVarz                                   As Variant

    For Each docVarz In ActiveDocument.Variables
        strReport = strReport & "I WILL DELETE :" & docVarz.Name & " Whose value is: " & docVarz.value & vbCr
        docVarz.Delete
    Next docVarz

    ' sending the report
    MsgBox strReport
End Sub

Private Sub testAdd()
    ' This sub adds some variables for test porpuses only.

    Dim error1                                    As Integer

    On Error Resume Next
    ActiveDocument.Variables.Add createAddress(2, 5, 1, 2)
    ActiveDocument.Variables.Add createAddress(2, 6, 1, 2)
    ActiveDocument.Variables.Add createAddress(2, 7, 1, 2)
    ActiveDocument.Variables.Add createAddress(3, 5, 1, 2)
    ActiveDocument.Variables.Add createAddress(3, 6, 1, 2)
    ActiveDocument.Variables.Add createAddress(3, 7, 1, 2)
    ActiveDocument.Variables.Add createAddress(235435, 3453453, 0, 2)

    ActiveDocument.Variables.Add createTrackingAddress(2, CVar("NOME_ITEM"), 0)
    ActiveDocument.Variables.Add createTrackingAddress(3, CVar("NOME_ITEM"), 0)
    ActiveDocument.Variables.Add ("dummy242432dummyD6554646dummy0")

    If Err Then MsgBox "erro ocorreu na criação"

    ' clearing the error
    error1 = Err.Number
    Err.Clear

    'finish
    If (error1 = 0) And (Not Err) Then
        MsgBox "teste concluido com sucesso"
    Else
        MsgBox "teste não foi concluido com sucesso", vbExclamation
    End If

End Sub

Public Sub cleanDocument()
    Call cleanUpWholeDocument(True)
End Sub

Public Sub benchmarkImport()
    ' getting the data from the database
    Dim timer                                     As obTimer
    Dim docvarname                                As String
    Dim i                                         As Integer
    Set timer = New obTimer

    timer.StartTimer

    For i = 1 To 500
        docvarname = breakString & CLng(1000 + Rnd() * 18000) & breakString & CLng(62 + Rnd() * 110)
        ActiveDocument.Variables.Add (docvarname)
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                             "DOCVARIABLE " & docvarname, PreserveFormatting:=True
    Next i

    Call getDataFromDatabase(False)

    timer.StopTimer

    If Err Then
        MsgBox "erro ocorreu na importação"
    Else
        MsgBox " teste de importacao concluido com sucesso" & vbCr & _
               "I took " & timer.Elapsed & " seconds to run the document update"
    End If

End Sub

Private Sub listDocvars()
    '
    Dim strReport                                 As String
    Dim docVarz                                   As Variant

    strReport = ""
    For Each docVarz In ActiveDocument.Variables
        strReport = strReport & docVarz.Name & " :  " & docVarz.value & vbCr
    Next docVarz
    MsgBox strReport
End Sub

Private Sub fieldBombing()
    ' Esta sub vai substituir todos os campos do documento por valores "hard coded". Deve ser usada apenas no envio ao cliente.
    Dim i                                         As Integer
    Dim field                                     As Variant

    For Each field In ActiveDocument.Fields
        i = field.index
        ' Moving selection to the field.
        Selection.GoTo what:=wdGoToField, which:=wdGoToAbsolute, Count:=i
        ' Inserting field result at selection.
        Selection.InsertAfter ActiveDocument.Fields(i).result.Text
        ' Deleting the field.
        ActiveDocument.Fields(i).Delete
    Next field

End Sub

Private Sub testClean()
    Call cleanUpWholeDocument
End Sub
