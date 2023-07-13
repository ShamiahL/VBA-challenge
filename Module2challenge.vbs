VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TwentyEighteen()
    Dim Ticker As String
    Dim openyear As Double
    Dim Closeyear As Double
    Dim yearlychange As Double
    Dim PercentChange As Double
    Dim VolumeChange As LongLong
    Dim lastrow As Integer
    
    For i = 2 To lastrow

    Cells(i, 10) = Cells(i, 6) - Cells(i, 3)

    Cells(i, 11) = (1 - (Cells(i, 6) / Cells(i, 3))) * 100

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Next i
    
    
End Sub
    

