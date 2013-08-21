VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "unitDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public pUnitCollection   As collection

Private Sub class_initialize()
    Set Me.pUnitCollection = getUnitsCollectionFromXML()
End Sub

Public Sub redefPath()
    Set Me.pUnitCollection = getUnitsCollectionFromXML(True)
End Sub

Public Function convertValue(strValue As String, propKey As Long, _
                             unitKey As Long, bolImporting As Boolean) As String
'=======================================================================================
' conversão de um valor de ou para o SI.
'---------------------------------------------------------------------------------------
' [valueSI] - Se TRUE, converte strValue do SI para a unidade dada em unitKey (import)
'           - Se FALSE, converte strValue da unidade dada em unitKey para o SI (export)
'---------------------------------------------------------------------------------------
' A conversão funciona da seguinte forma.
'           offsetFactor + (convFactor * <nãoSI>) = <SI>
'---------------------------------------------------------------------------------------
' < Histórico de revisões>
'=======================================================================================
'
'#############################
' Há um problema conhecido.
' Esse método Só vai funcionar corretamente se o separador de decimal do seu sistema for a virgula. (padrão br)
'#############################

    Dim convFactor       As Double
    Dim offsetFactor     As Double
    Dim dimKey           As Long
    Dim convertedValue   As Variant
    Dim dblValue         As Double
    Dim unitSymbol       As String

    On Error GoTo convertValue_Error
    ' getting the dimension key
    dimKey = getDimKey(propKey)

    If unitKey <> 0 And dimKey <> 0 Then
        ' with the dimension key, i get the dictionary with the conversion factors.

        convFactor = getConvFactor(propKey, unitKey)
        offsetFactor = getOffsetFactor(propKey, unitKey)

        dblValue = CDbl(strValue)
        unitSymbol = getUnitSymbol(propKey, unitKey)

        If bolImporting Then                                         'So, dblValue is IN SI.
            convertedValue = (dblValue - offsetFactor) / convFactor
        Else                                                         'So, dblValue is NOT IN SI
            convertedValue = dblValue * convFactor + offsetFactor
        End If
    Else
        convertedValue = strValue
    End If

    convertValue = Format(CStr(convertedValue), "standard")

    On Error GoTo 0

    GoTo convertValue_Finally

convertValue_Finally:

    Exit Function

    ' Procedure Error Handler
convertValue_Error:
    Dim errorAction      As Integer
    'here goes your specific error handling code.

    ' here comes the generic global error handling code.
    errorAction = handleMyError()
    Select Case errorAction
        Case -1
            Resume Next
        Case 1
            Resume Next
        Case 2
            GoTo convertValue_Finally
        Case Else
            Stop
    End Select
End Function

Public Property Get getUnitKey(strPropName As String, strUnitSymbol As String) As Long
    Dim tempUnit         As unit
    Dim propKey          As Long
    Dim dimKey           As Long

    propKey = getPropKey(strPropName)
    dimKey = getDimKey(propKey)

    If dimKey <> 0 And strUnitSymbol <> "" Then
        For Each tempUnit In Me.pUnitCollection
            If tempUnit.symbol = strUnitSymbol And tempUnit.dimension = dimKey Then

                getUnitKey = tempUnit.key
                Exit Property

            End If
        Next tempUnit
    Else
        getUnitKey = 0
    End If
End Property

Private Function getConvFactor(propKey As Long, unitKey As Long) As Double
    Dim tempUnit         As unit

    If unitKey <> 0 Then
        Set tempUnit = getUnitObjectFromKeys(propKey, unitKey)
        If Not tempUnit Is Nothing Then getConvFactor = tempUnit.convFactor
    Else
        getConvFactor = 1
    End If
End Function

Private Function getOffsetFactor(propKey As Long, unitKey As Long) As Double
    Dim tempUnit         As unit

    If unitKey <> 0 Then
        Set tempUnit = getUnitObjectFromKeys(propKey, unitKey)
        If Not tempUnit Is Nothing Then getOffsetFactor = tempUnit.convOffset
    Else
        getOffsetFactor = 0
    End If
End Function

Public Property Get getUnitSymbol(propKey As Long, unitKey As Long) As String
    Dim tempUnit         As unit

    If unitKey <> 0 Then
        Set tempUnit = getUnitObjectFromKeys(propKey, unitKey)
        If Not tempUnit Is Nothing Then getUnitSymbol = tempUnit.symbol
    Else
        getUnitSymbol = ""
    End If
End Property

'##########
' Privates

Private Function getUnitObjectFromKeys(propKey As Long, unitKey As Long) As unit
    Dim dimKey           As Long
    Dim tempUnit         As unit
    Dim strUnitAddress   As String

    dimKey = getDimKey(propKey)

    strUnitAddress = dimKey & "|" & unitKey

    '    If isInCollection(Me.pUnitCollection, strUnitAddress) Then
    Set tempUnit = Me.pUnitCollection(strUnitAddress)
    Set getUnitObjectFromKeys = tempUnit
    '    End If
End Function

Private Function getUnitsCollectionFromXML(Optional redefPaths As Boolean = False) As collection

    Dim objXml           As MSXML2.DOMDocument
    Dim unitList         As MSXML2.IXMLDOMNodeList
    Dim unit             As MSXML2.IXMLDOMElement
    Dim itemIndex        As Integer
    Dim tempUnit         As unit
    Dim unitCol          As collection

    Set objXml = New MSXML2.DOMDocument

    If False = objXml.Load(getUnitsXmlPath()) Or redefPaths Then
        MsgBox "Redefina a localização do xml de configuração do sistema de unidades"
        storeUnitsXmlPath (getFileDialog())
        objXml.Load (getUnitsXmlPath())
    End If

    Set unitList = objXml.getElementsByTagName("unit")

    Set unitCol = New collection

    For itemIndex = 0 To unitList.Length - 1
        Set tempUnit = New unit

        tempUnit.symbol = CStr(unitList.Item(itemIndex).SelectSingleNode("symbol").Text)
        tempUnit.convFactor = CDbl(unitList.Item(itemIndex).SelectSingleNode("convFactor").Text)
        tempUnit.convOffset = CDbl(unitList.Item(itemIndex).SelectSingleNode("convOffset").Text)
        tempUnit.dimension = CInt(unitList.Item(itemIndex).SelectSingleNode("dimension").Text)
        tempUnit.key = CLng(unitList.Item(itemIndex).SelectSingleNode("key").Text)

        unitCol.Add Item:=tempUnit, key:=tempUnit.dimension & "|" & tempUnit.key

    Next itemIndex

    Set getUnitsCollectionFromXML = unitCol
End Function
