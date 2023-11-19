Attribute VB_Name = "Module1"
Option Explicit

' KEVIN BYROM
' VB SLOTS
' CIS 140 : ASSIGNMENT 4

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Global Const PIXEL = 3
Global Const SRCCOPY = &HCC0020
Global Const MODAL = 1

Global vSlot(0 To 2) As slotproperties
Global CurrentMoney As Integer

Type slotproperties
    value(-1 To 1) As Integer
    speed As Integer
    y As Integer
End Type

Function CalcPrize() As Integer

    'first, check if 3 slots are equal
    If vSlot(0).value(0) = vSlot(1).value(0) And vSlot(1).value(0) = vSlot(2).value(0) Then
        If vSlot(0).value(0) = 0 Then 'cherry
            CalcPrize = 4
        ElseIf vSlot(0).value(0) = 1 Then 'grape
            CalcPrize = 10
        ElseIf vSlot(0).value(0) = 2 Then 'lemon
            CalcPrize = 15
        ElseIf vSlot(0).value(0) = 3 Then 'lime
            CalcPrize = 25
        ElseIf vSlot(0).value(0) = 4 Then 'orange
            CalcPrize = 35
        ElseIf vSlot(0).value(0) = 5 Then 'seven
            CalcPrize = 50
        End If
        Exit Function
    ElseIf vSlot(0).value(0) = 0 Then 'one cherry
        CalcPrize = 1
        Exit Function
    End If

End Function


Sub DrawSlots()

    Dim i, ii As Integer
    Dim rc As Long
    
    For i = 0 To 2
        frmMain.Slot(i).ScaleMode = PIXEL
    Next
    
    For i = 0 To 2
        For ii = -1 To 1
            rc = BitBlt(frmMain.Slot(i).hDC, 0, vSlot(i).y + (100 * ii), 100, 100, frmBuffer.picImage(vSlot(i).value(ii)).hDC, 0, 0, SRCCOPY)
        Next
    Next
    
End Sub
Sub Timeout(count As Long)

    Dim Start As Long
    
    Start = Timer
    While Timer < Start + count
        DoEvents
    Wend
    
End Sub


