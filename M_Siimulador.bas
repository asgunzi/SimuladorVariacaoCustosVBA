Attribute VB_Name = "M_Siimulador"
Option Explicit


Sub simulaTriangular()

Dim nl As Long, i As Long, t As Long
Dim arrItem As Variant
Dim nTrials As Long
Dim arrResults As Variant
Dim min As Double, med As Double, max As Double

'Configuracoes
nTrials = 200

Math.Randomize

ReDim arrResults(1 To nTrials, 1 To 2)

'Le linhas
nl = Application.WorksheetFunction.CountA(Range("a5:a50000"))

arrItem = Range("a5:d" & 4 + nl)


'Faz mil simulações
For t = 1 To nTrials
    'Para cada simulação, sorteia cada valor segundo a distribuicao triangular
    For i = 1 To nl
        min = arrItem(i, 2)
        med = arrItem(i, 3)
        max = arrItem(i, 4)
        arrResults(t, 1) = t
        arrResults(t, 2) = arrResults(t, 2) + evalTriangular(min, med, max) 'soma o sorteio da triangular
     Next i
Next t


Range("v5:w5000").ClearContents
Range("v5").Resize(nTrials, 2) = arrResults

End Sub


Private Function evalTriangular(min As Double, med As Double, max As Double)
'Recebe os parametros min med e max
'Faz o sorteio segundo a distribuição triangular
'Retorna o valor sorteado
    
Dim swap As Double
Dim FC As Double

swap = Math.Rnd() 'Sorteia um número aleatório

'Encontra o valor segundo a inversa da CDF da triangular
'https://en.wikipedia.org/wiki/Triangular_distribution



FC = (max - min) / (med - min)


If swap < FC Then
    evalTriangular = min + Math.Sqr(swap * (med - min) * (max - min))

Else
    evalTriangular = med - Math.Sqr((1 - swap) * (med - min) * (med - max))
    
End If



End Function

