VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shareQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim shareCount           As Long
Dim queue(1000, 3)       As Variant

Private Sub class_initialize()
    shareCount = 0
End Sub

Public Property Get countShares()
    countShares = shareCount
End Property

Public Sub dequeue(index As Integer)
    Dim j                As Integer

    For j = 0 To 3
        queue(index, j) = queue(shareCount - 1, j)
    Next j

    shareCount = shareCount - 1
End Sub

Public Property Get sharesCollection() As collection
    Dim i                As Long
    Dim tempArray(3)     As Variant

    Set sharesCollection = New collection

    For i = 0 To shareCount - 1
        tempArray(0) = queue(i, 0)
        tempArray(1) = queue(i, 1)
        tempArray(2) = queue(i, 2)
        tempArray(3) = queue(i, 3)

        sharesCollection.Add tempArray
    Next i
End Property

Public Sub Clear()
    shareCount = 0
End Sub

Public Sub enqueue(strItemName1 As String, strPropName1 As String, strItemName2 As String, strPropName2 As String)
    queue(shareCount, 0) = strItemName1
    queue(shareCount, 1) = strPropName1
    queue(shareCount, 2) = strItemName2
    queue(shareCount, 3) = strPropName2

    shareCount = shareCount + 1
End Sub
