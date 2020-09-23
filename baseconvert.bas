Attribute VB_Name = "Module2"
Public Function ConvertBase(NumIn As String, BaseIn As Integer, BaseOut As Integer) As String
    Dim i As Integer, CurrentCharacter As String, CharacterValue As Integer, PlaceValue As Integer, RunningTotal As Double, Remainder As Double, BaseOutDouble As Double, NumInCaps As String
    ' Ensure input data is valid
    If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or BaseOut < 1 Or BaseOut > 36 Then
    ConvertBase = "Error"
    Exit Function
End If
' Ensure any letters in the input mumber
'     are capitals
NumInCaps = UCase$(NumIn)
' Convert NumInCaps into Decimal
PlaceValue = Len(NumInCaps)


For i = 1 To Len(NumInCaps)
    PlaceValue = PlaceValue - 1
    CurrentCharacter = Mid$(NumInCaps, i, 1)
    CharacterValue = 0
    If Asc(CurrentCharacter) > 64 And _
    Asc(CurrentCharacter) < 91 Then _
    CharacterValue = Asc(CurrentCharacter) - 55


    If CharacterValue = 0 Then
        ' Ensure NumIn is correct
        If Asc(CurrentCharacter) < 48 Or _
        Asc(CurrentCharacter) > 57 Then
        ConvertBase = "Error"
        Exit Function
    Else
        CharacterValue = Val(CurrentCharacter)
    End If
End If


If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
    ' Ensure NumIn is correct
    ConvertBase = "Error"
    Exit Function
End If
RunningTotal = RunningTotal + CharacterValue * (BaseIn ^ PlaceValue)
Next i
' Convert Decimal Number into the desire
'     d base using
    ' Repeated Division


Do
BaseOutDouble = CDbl(BaseOut)
Remainder = ModDouble(RunningTotal, BaseOutDouble)
RunningTotal = (RunningTotal - Remainder) / BaseOut


If Remainder >= 10 Then
    CurrentCharacter = Chr$(Remainder + 55)
Else
    CurrentCharacter = Right$(Str$(Remainder), Len(Str$(Remainder)) - 1)
End If
ConvertBase = CurrentCharacter & ConvertBase
Loop While RunningTotal > 0
End Function


Public Function ModDouble(NumIn As Double, DivNum As Double) As Double
    ' Returns the Remainder when a number is
    '     divided by another
    ' (Works for double data-type)
    ModDouble = NumIn - (Int(NumIn / DivNum) * DivNum)
End Function


