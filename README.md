VB-CEPEL
========

Ler e armazenar dados em TXT
Private Sub Command1_Click()
Dim linha, text, x, data1, data2 As String
Dim dia(500) As String
Dim cracha(500) As String
Dim hora(500) As String
Dim codigo(500) As String

Open "C:\ABONO09.txt" For Input As #1

Line Input #1, linha
Do While Not EOF(1)
'Line Input #1, linha ' ler linha
  If UCase(Mid(linha, 34, 11)) = "OCORRENCIAS" Then
  data1 = Mid(linha, 60, 8)
  data2 = Mid(linha, 71, 8)
 Debug.Print data1
 Debug.Print data2
  End If
      
    Line Input #1, linha
    If UCase(Mid(linha, 12, 1)) = ":" Then
    cracha(x) = Mid(linha, 14, 8)
    'cracha(0) = Mid(linha, 14, 8)
   ' cracha(1) = Mid(linha, 14, 8)
    Debug.Print cracha(x)
    End If
    
  ' Line Input #1, linha
   If UCase(Mid(linha, 72, 1)) = ":" Then
   dia(x) = Mid(linha, 1, 6)
   Debug.Print dia(x)
   End If
   
   
 'Line Input #1, linha
If UCase(Mid(linha, 72, 1)) = ":" Then
hora(x) = Mid(linha, 70, 5)
Debug.Print hora(x)
End If

 
'Line Input #1, linha
If UCase(Mid(linha, 72, 1)) = ":" Then
codigo(x) = Mid(linha, 77, 3)
Debug.Print codigo(x)
End If
'Form1.Text1 = "data: " & "  " & data1 & "  " & "data:  " & "  " & data2 & "  " & "cracha: " & "  " & cracha(x) & "  " & "dia: " & dia(x) & "  " & "hora: " & hora(x) & "  " & "codigo: " & codigo(x)
'Form1.List1.AddItem "data: " & "  " & data1 & "  " & "data:  " & "  " & data2 & "  " & "cracha: " & "  " & cracha(x) & "  " & "dia: " & dia(x) & "  " & "hora: " & hora(x) & "  " & "codigo: " & codigo(x)
'Form1.List1.AddItem (data1) & " " & " " & (data2) & " " & "cracha: " & (cracha(x)) & " " & " " & (dia(x))
'Form1.List1.AddItem (data2)
'x = x + 1 ' indice do proximo elemento

   Loop
   Close #1
   End Sub

