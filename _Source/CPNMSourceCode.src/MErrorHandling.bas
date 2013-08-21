Attribute VB_Name = "MErrorHandling"
' Chemtech - A Siemens Business ========================================================
'
'=======================================================================================
Option Explicit
' Desenvolvimento ======================================================================
' RR            Renan       renan.ranelli@chemtech.com.br
'=======================================================================================
' Versões ==============================================================================
'
'
'
'=======================================================================================

Public Function handleMyError(Optional showMsgBoxes As Boolean = True) As Integer
'=======================================================================================
'
'---------------------------------------------------------------------------------------
' [showMsgBoxes] - booleano que indica se serão apresentadas msgboxes com resultados do erro
'---------------------------------------------------------------------------------------
' Retorna: -1 para stop. (para a execução)
'           1 para resume next. (ignorar o erro e continuar a execução do código principal)
'           2 para goto ExitSub. (o equivalente ao "finally" do try/catch/except
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
    Dim intMyErrorNumber As Long

    'I extract my user defined error number.
    intMyErrorNumber = Err.Number - vbObjectError

    'If Err.Number = vbObjectError + 22000 Then
    Select Case intMyErrorNumber

        Case 100000                                                  'No value in checkKeyness
            If showMsgBoxes Then MsgBox "An error has occurred and was caught" & _
               vbCr & vbCr & "Error number = " & Err.Number & _
               vbCr & vbCr & "Error description = " & Err.Description & _
               vbCr & vbCr & "problem at line: " & Erl, vbCritical
        
        Case 100001                                                  'Duplicated values in checkKeyness
            MsgBox "An error has occurred and was caught" & _
                   vbCr & vbCr & "Error number = " & Err.Number & _
                   vbCr & vbCr & "Error description = " & Err.Description & _
                   vbCr & vbCr & "problem at line: " & Erl, vbCritical
            handleMyError = 1

        Case 22000                                                   'This case corresponds to the test case
            MsgBox "Your test error was caught, and properly treated!" & vbCr & "YAY!!"
            handleMyError = 1

            ' ########
            ' GetData error handling
        Case 100211
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 2
        Case 100212
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 2
        Case 100213
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 2

            ' collection creation error.
        Case 1121221
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 1

            ' dimensionallity error.
        Case 175645
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 2

        Case 554565
            If showMsgBoxes Then MsgBox Err.Description
            handleMyError = 2

        Case Else                                                    'This case corresponds to an untreated error.
            MsgBox "An error has occurred and was caught, but not treated" & _
                   vbCr & vbCr & "Error number = " & Err.Number & _
                   vbCr & vbCr & "Error description = " & Err.Description & _
                   vbCr & vbCr & "problem at line: " & Erl, vbCritical
            handleMyError = -1
    End Select
End Function

Public Sub checkKeyness(rs As ADODB.Recordset)
' This sub checks if there is more than one, or more than one, records in the recordset, where it should have a single record in it.
    If rs.EOF = True Then
        Err.Raise vbObjectError + 100000, Description:="The query returned no value!"
    Else
        If rs.RecordCount > 1 Then
            Err.Raise vbObjectError + 100001, "The query returned no value!"
        End If
    End If
End Sub
