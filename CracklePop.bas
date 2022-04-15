Attribute VB_Name = "Crack"
Sub RC()

Dim i As Long, r As Range
Dim myArray() As Variant


ReDim myArray(1 To 101)
Check = True: Counter = 1: Total = 0 ' The homies

    Do ' Outer loop.
        Do While Counter < 101 ' Inner Loop

            If Counter Mod 5 = 0 And Counter Mod 3 = 0 Then
                'Debug.Print "CracklePop"
                myArray(Counter) = "CracklePop"
            ElseIf Counter Mod 3 = 0 Then
                'Debug.Print "Crackle"
                myArray(Counter) = "Crackle"
            ElseIf Counter Mod 5 = 0 Then
                'Debug.Print "Pop"
                myArray(Counter) = "Pop"
            Else
                'Debug.Print Counter
                myArray(Counter) = Counter
            End If

            Counter = Counter + 1 'Plus
            RangeToArray = arr

        Loop

        Total = Total + Counter ' Exit Do
        Counter = 0

        If Total = 101 Then
            Check = False
        End If

    Loop Until Check = False ' Exit outer loop

    MsgBox Join(myArray, vbCrLf)

End Sub
