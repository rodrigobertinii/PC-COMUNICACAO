# Descrição do Projeto

Atendendo às necessidades da PC Comunicação, desenvolvi um arquivo Excel que inclui um formulário para que os colaboradores possam registrar reuniões com clientes, sejam elas internas ou externas. O projeto também contempla uma seção para métricas e indicadores que permitem aos gestores realizar análises detalhadas.

## 1. Criação do Formulário

Utilizando VBA, desenvolvi o design do formulário e inseri as informações necessárias conforme as exigências da empresa.

![image](https://github.com/user-attachments/assets/4d3f14a4-0ea3-4598-a744-00d814582a81)

## 2. Criação do Banco de Dados

Após a aprovação do design pelo cliente, configurei o banco de dados necessário para preencher o formulário. Esses dados servem como a base para o preenchimento do formulário pelos colaboradores.

![image](https://github.com/user-attachments/assets/e5f7bd8c-6600-404e-91ac-44cb19a5171d)

## 3. Transferência de Dados

Integrei a transferência dos dados inseridos no formulário para uma aba específica da planilha Excel. O código VBA a seguir foi utilizado para popular os campos do formulário e para transferir os dados preenchidos:

### Código para Carregar Dados no Formulário

```vba
Private Sub UserForm_Initialize()
    ' Adiciona dados ao formulário para o campo "nome"
    Dim Lin As Integer
    Lin = 2
    Do Until Planilha3.Cells(Lin, 1) = ""
        ComboBoxNome.AddItem Planilha3.Cells(Lin, 1)
        Lin = Lin + 1
    Loop
    
    ' Adiciona dados ao formulário para o campo "data"
    Lin = 2
    While Planilha3.Cells(Lin, 2) <> ""
        ComboBoxData.AddItem Planilha3.Cells(Lin, 2)
        Lin = Lin + 1
    Wend
    
    ' Adiciona dados ao formulário para o campo "cliente"
    Lin = 2
    While Planilha3.Cells(Lin, 3) <> ""
        ComboBoxCliente.AddItem Planilha3.Cells(Lin, 3)
        Lin = Lin + 1
    Wend
    
    ' Adiciona dados ao formulário para o campo "Tempo trabalhado"
    Lin = 2
    Do Until Planilha3.Cells(Lin, 4) = ""
        ComboBoxTempoTrabalhado.AddItem Planilha3.Cells(Lin, 4)
        Lin = Lin + 1
    Loop
    
    ' Adiciona dados ao formulário para o campo "Local"
    Lin = 2
    Do Until Planilha3.Cells(Lin, 5) = ""
        ComboBoxLocal.AddItem Planilha3.Cells(Lin, 5)
        Lin = Lin + 1
    Loop
End Sub

```

### Código para Transferir Dados do Formulário para a Planilha

```vba
Private Sub CommandButtonenviardados_Click()
    ' Valida se os campos estão preenchidos corretamente
    If ComboBoxNome = "" Then
        MsgBox "Digite um valor para o campo 'nome'"
        ComboBoxNome.SetFocus
        Exit Sub
    ElseIf ComboBoxNome.ListIndex = -1 Then
        MsgBox "Valor inválido, escolha um valor da lista na aba 'nome'"
        ComboBoxNome.SetFocus
        ComboBoxNome = ""
        Exit Sub
    End If
    
    If ComboBoxData = "" Then
        MsgBox "Digite um valor para o campo 'Data'"
        ComboBoxData.SetFocus
        Exit Sub
    ElseIf ComboBoxData.ListIndex = -1 Then
        MsgBox "Valor inválido, escolha um valor da lista na aba 'Data'"
        ComboBoxData.SetFocus
        ComboBoxData = ""
        Exit Sub
    End If
    
    If ComboBoxCliente = "" Then
        MsgBox "Digite um valor para o campo 'Cliente'"
        ComboBoxCliente.SetFocus
        Exit Sub
    ElseIf ComboBoxCliente.ListIndex = -1 Then
        MsgBox "Valor inválido, escolha um valor da lista na aba 'Cliente'"
        ComboBoxCliente.SetFocus
        ComboBoxCliente = ""
        Exit Sub
    End If
    
    If ComboBoxTempoTrabalhado = "" Then
        MsgBox "Digite um valor para o campo 'Tempo trabalhado (em min)'"
        ComboBoxTempoTrabalhado.SetFocus
        Exit Sub
    ElseIf ComboBoxTempoTrabalhado.ListIndex = -1 Then
        MsgBox "Valor inválido, escolha um valor da lista na aba 'Tempo Trabalhando (em min)'"
        ComboBoxTempoTrabalhado.SetFocus
        ComboBoxTempoTrabalhado = ""
        Exit Sub
    End If
    
    If ComboBoxLocal = "" Then
        MsgBox "Digite um valor para o campo 'Local'"
        ComboBoxLocal.SetFocus
        Exit Sub
    ElseIf ComboBoxLocal.ListIndex = -1 Then
        MsgBox "Valor inválido, escolha um valor da lista na aba 'Local'"
        ComboBoxLocal.SetFocus
        ComboBoxLocal = ""
        Exit Sub
    End If
    
    ' Transfere os dados do formulário para a planilha "Resultados"
    Planilha4.Select
    Range("A5000").Select
    ActiveCell.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = ComboBoxNome
    ActiveCell.Offset(0, 1).Value = ComboBoxData
    ActiveCell.Offset(0, 2).Value = ComboBoxCliente
    ActiveCell.Offset(0, 3).Value = Val(ComboBoxTempoTrabalhado)
    ActiveCell.Offset(0, 4).Value = ComboBoxLocal
    ActiveCell.Offset(0, 5).Value = Date
    
    MsgBox "Material cadastrado com sucesso", vbExclamation
    Planilha1.Select
    Unload Me
End Sub

```
![image](https://github.com/user-attachments/assets/71327f08-8181-4515-b688-6bb748a8a4b3)

![image](https://github.com/user-attachments/assets/2a74bffd-db21-4304-87ab-5e2c475cc668)

![image](https://github.com/user-attachments/assets/7a02f8a0-638d-4add-844a-0d4fc339c3e0)


## 4. Análise de Resultados
Criei uma aba dedicada à análise dos resultados, onde os dados do formulário são armazenados. Esta aba inclui cálculos e funções para gerar métricas e indicadores.

![image](https://github.com/user-attachments/assets/0908c9f8-4ab9-4124-81d0-858347de4e63)

## 5. Indicadores
Com base na aba de análise de resultados, desenvolvi três indicadores principais:

- Tempo total gasto por cada funcionário

![image](https://github.com/user-attachments/assets/2dc97cb7-e2f3-4a45-aa57-567abcb347da)

- Custo operacional por cliente
- 
![image](https://github.com/user-attachments/assets/1d89dc85-6ab8-436b-b192-bf954d80759b)

- Relação entre colaborador e tempo dedicado ao cliente
- 
![image](https://github.com/user-attachments/assets/1922f137-8dd8-44c9-96b4-914b07ff2ba3)

## 6. Interface de Usuário
Finalmente, criei uma aba principal que limita o acesso dos colaboradores apenas ao formulário (em cor laranja) e fornece aos gestores acesso aos dados e métricas (em cor verde).

![image](https://github.com/user-attachments/assets/752fa243-6581-4ae7-a731-8d66d03479fd)

Observação: É necessário utilizar uma senha para acessar a aba de dados e métricas pelos gestores.
