VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'How many times per second is the counter updated?
Private Declare Function QueryFrequency Lib "kernel32" _
                                        Alias "QueryPerformanceFrequency" ( _
                                        lpFrequency As Currency) As Long

'What is the counter's value
Private Declare Function QueryCounter Lib "kernel32" _
                                      Alias "QueryPerformanceCounter" ( _
                                      lpPerformanceCount As Currency) As Long

'Variables to store the counter information
Dim mcyFrequency         As Currency
Dim mcyOverhead          As Currency
Dim mcyStarted           As Currency
Dim mcyStopped           As Currency

Private Sub class_initialize()
    Dim cyCount1 As Currency, cyCount2 As Currency

    'Get the counter frequency
    QueryFrequency mcyFrequency

    'Call the hi-res counter twice, to check how long it takes
    QueryCounter cyCount1
    QueryCounter cyCount2

    'Store the call overhead
    mcyOverhead = cyCount2 - cyCount1
End Sub


Public Sub StartTimer()
'Get the time that you started
    QueryCounter mcyStarted
End Sub

Public Sub StopTimer()
'Get the time that you stopped
    QueryCounter mcyStopped
End Sub

Public Property Get Elapsed() As Double
    Dim cyTimer          As Currency

    'Have you stopped or not?
    If mcyStopped = 0 Then
        QueryCounter cyTimer
    Else
        cyTimer = mcyStopped
    End If

    'If you have a frequency, return the duration, in seconds
    If mcyFrequency > 0 Then
        Elapsed = (cyTimer - mcyStarted - mcyOverhead) / mcyFrequency
    End If
End Property

