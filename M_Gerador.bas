Attribute VB_Name = "M_Gerador"
Option Explicit



Sub gerarCaso()
'Gera um caso com números aleatórios, apenas para ilustrar o funcionamento do simulador
Dim ncasos As Long
Dim arrCasos As Variant
Dim i  As Long


'Sorteia o número de itens
ncasos = Application.WorksheetFunction.RandBetween(4, 20)


'Redimensiona o array para conter os itens
ReDim arrCasos(1 To ncasos, 1 To 4) 'Ncasos linhas e 4 colunas: item, min, médio e max


For i = 1 To ncasos
    arrCasos(i, 1) = "Item " & i
    
    'Media
    arrCasos(i, 3) = Math.Rnd() * 100
    
    'Mín
    arrCasos(i, 2) = arrCasos(i, 3) - Math.Rnd() * 10
    If arrCasos(i, 2) < 0 Then arrCasos(i, 2) = 0
    'Max
    arrCasos(i, 4) = arrCasos(i, 3) + Math.Rnd() * 10
    
Next i

P_Simulador.Activate
Range("a5:d5000").ClearContents
Range("a5").Resize(ncasos, UBound(arrCasos, 2)) = arrCasos

End Sub





