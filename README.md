# Exercicio-Elevador-VBA
Alunos: André Neri, Nathália dos Santos, Gabrielle Bellini, Juliana Angelo, Michelle Cavallaro

Sub elevador()
    Dim nviagem As Integer
    Dim pessoasEntrar As Integer
    Dim pessoasSair As Integer
    Dim viagens(29) As Integer
    Dim qtdPessoasElevador As Integer
    Dim pessoasEntraram As Integer
    Dim vaga As Integer
    Dim i As Integer
    Dim somaPessoas As Integer
    
    
    For i = 1 To 30
        nviagem = i
        pessoasEntrar = InputBox("Quantidade de pessoas disponiveis no andar para entrar")
        
        If nviagem = 1 Then
            If pessoasEntrar <= 6 Then
                qtdPessoasElevador = pessoasEntrar
            Else
                qtdPessoasElevador = 6
            End If
            viagens(i - 1) = qtdPessoasElevador
            
        End If
        
        
        If nviagem <> 1 Then
            pessoasSair = InputBox("Quantidade de pessoas que irão sair do elevador")
            
            If pessoasSair > qtdPessoasElevador Then
                MsgBox ("Impossível sair mais pessoas do que têm no elevador")
                
            Else
                qtdPessoasElevador = qtdPessoasElevador - pessoasSair
                
            End If
        
            vaga = 6 - qtdPessoasElevador
            
            If pessoasEntrar >= vaga Then
                pessoasEntraram = vaga
                qtdPessoasElevador = qtdPessoasElevador + vaga
            Else
                pessoasEntraram = pessoasEntrar
                qtdPessoasElevador = qtdPessoasElevador + pessoasEntrar
            End If
        
            viagens(i - 1) = pessoasEntraram
            
        End If
                
               
    Next
    For i = 0 To 29
        somaPessoas = somaPessoas + viagens(i)
    Next
        MsgBox ("Quantidade de Pessoas que entraram no elevador: " & somaPessoas)
          
End Sub
