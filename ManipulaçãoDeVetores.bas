Attribute VB_Name = "Módulo1"
Sub vetores()
Dim vetor(3) As Double
vetor(0) = 10
vetor(1) = 21
vetor(2) = 32
vetor(3) = 43
Debug.Print ("O elemento é: " & vetor(1))
End Sub

Sub vetores2()
Dim vetor(2 To 4) As Double
'vetor(1) = 10
vetor(2) = 21
vetor(3) = 32
vetor(4) = 43
'vetor(5) = 54
Debug.Print ("O elemento é: " & vetor(4))
End Sub

Sub vetor3()
   Dim vetor(4) As Double
   vetor(1) = 10
   vetor(2) = 21
   vetor(3) = 32
   vetor(4) = 43
   soma = vetor(1) + vetor(2) + vetor(3) + vetor(4)
   Debug.Print ("O elemento é: " & soma)
End Sub
 
Sub vetores4()
   Dim vetor1(3) As Double
   vetor1(1) = 10
   vetor1(2) = 20
   vetor1(3) = 30
   Dim vetor2(4 To 6) As Double
   vetor2(4) = 15
   vetor2(5) = 30
   vetor2(6) = 45
   Dim vetor_concat As Double
   vetor_concat = vetor1(1) & "," & vetor2(6)
   Debug.Print ("Os elementos são: " & vetor_concat)
   End Sub

Function SomaRecursiva(numeros() As Double, tamanho As Integer) As Double
   If tamanho = 0 Then
SomaRecursiva = 0
Else
   SomaRecursiva = numeros(tamanho - 1) + SomaRecursiva(numeros, tamanho - 1)
End If
End Function

Sub TesteRecursivo()
   Dim numeros(5) As Double
   numeros(0) = 10
   numeros(1) = 20
   numeros(2) = 30
   numeros(3) = 40
   numeros(4) = 50
   Dim soma As Double
   soma = SomaRecursiva(numeros, 5)
   MsgBox "A soma é: " & soma
End Sub

Sub dobro()
   Dim vetor(5) As Double
   Dim num As Double
   Dim i As Integer
   For i = 1 To 5
   num = InputBox("Informe a " & i & "o. número")
   vetor(i) = num
   Cells(i, "A") = vetor(i) * 2
Next i
End Sub

Sub limpar()
Range("A1:B5").ClearContents
End Sub




