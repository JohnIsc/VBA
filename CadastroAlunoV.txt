VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroAluno 
   Caption         =   "Cadastro de Alunos"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "CadastroAluno.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastroAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
txtDatadeMatricula = Date
End Sub

Private Sub bntPesquisa_Click()

If txtCep = Empty Then Exit Sub
If Len(txtCep) < 9 Then Exit Sub
' lean conta os caracters. Se ela tiver vazia ou com menos de 9 caracteres, ela n vai executar meu codigo

txtEndereço = UCase(ConsultaCep(txtCep, "logradouro"))
txtBairro = UCase(ConsultaCep(txtCep, "bairro"))
txtCidade = UCase(ConsultaCep(txtCep, "localidade"))
txtUF = UCase(ConsultaCep(txtCep, "uf"))
'Ucase é pra colocar o resultado em maiusculo
End Sub


' Botão Cancelar
Private Sub btnCancelar_Click()

Unload CadastroAluno
Msgbox "Cadastro não foi salvo!", vbInformation, "Cadastro Cancelado"

End Sub

Private Sub btnSalvar_Click()

If lblTipoCadastro.Caption = "NOVO CADASTRO" Then
SalvarNovoCadastro
ElseIf lblTipoCadastro.Caption = "EDIÇÃO DE CADASTRO" Then
SalvarEdicaoCadastro
End If

'Para ele executar a nossa sub personalizada abaixo
End Sub
'Sub personalizada, criada abaixo, para ele conseguir detectar se estou criando novo cadastro ou editando um já existente
Private Sub SalvarNovoCadastro()

Dim Campoembranco As Boolean
Dim Linha         As Integer

Campoembranco = False
'Iniciar a verificação se os campos obrigatórios estão preenchidos.

If txtNomeAluno = Empty Then Campoembranco = True
If txtCodigo = Empty Then Campoembranco = True
If txtDataNasc = Empty Then Campoembranco = True
If txtPai = Empty Then Campoembranco = True
If txtMae = Empty Then Campoembranco = True
If txtResponsavel = Empty Then Campoembranco = True
If txtNaturalidade = Empty Then Campoembranco = True
If txtEndereço = Empty Then Campoembranco = True
If txtN = Empty Then Campoembranco = True
If txtBairro = Empty Then Campoembranco = True
If txtCidade = Empty Then Campoembranco = True
If txtUF = Empty Then Campoembranco = True
If txtCelular1 = Empty Then Campoembranco = True
If cboPeriodo = Empty Then Campoembranco = True
If cboCiclo = Empty Then Campoembranco = True
If cboProfessores = Empty Then Campoembranco = True
If cboTurno = Empty Then Campoembranco = True
If txtDatadaMatricula = Empty Then Campoembranco = True
If cboSexo = Empty Then Campoembranco = True
If cboNE = Empty Then Campoembranco = True
If txtCep = Empty Then Campoembranco = True

If Campoembranco = True Then
Msgbox "Campo Obrigatório sem preenchimento.", vbInformation, "Campo em branco"
Exit Sub
'Pra parar o código caso tenha campo em branco

End If


'Vamos descobrir a ultima linha preenchida na tabela pra salvar na planilha

Linha = 391

Do While Planilha1.Cells(Linha, 2) <> Empty
Linha = Linha + 1
Loop
'Se a minha Linha 3, na coluna 2 for diferente de vazia, pegar a minha variavel linha e somar 1, colando um loop, DEVIDO A DO WHILE
'Inicia armazenamento dos dados de cadastro do Aluno. No exemplo abaixo, o meu ID sempre vai valer o valor que esta no meu contador de registro + 1

Planilha1.Cells(Linha, 2) = Planilha1.Range("B388").Value + 1
Planilha1.Cells(Linha, 3) = txtNomeAluno.Value
Planilha1.Cells(Linha, 4) = txtDataNasc.Value
Planilha1.Cells(Linha, 5) = txtPai.Value
Planilha1.Cells(Linha, 6) = txtMae.Value
Planilha1.Cells(Linha, 7) = txtResponsavel.Value
Planilha1.Cells(Linha, 8) = txtRestricao.Value
Planilha1.Cells(Linha, 9) = txtNaturalidade.Value
Planilha1.Cells(Linha, 10) = cboSexo.Value
Planilha1.Cells(Linha, 11) = txtCep.Value
Planilha1.Cells(Linha, 12) = txtLogradouro
Planilha1.Cells(Linha, 13) = txtEndereço
Planilha1.Cells(Linha, 14) = txtN.Value
Planilha1.Cells(Linha, 15) = txtComplemento.Value
Planilha1.Cells(Linha, 16) = txtBairro.Value
Planilha1.Cells(Linha, 17) = txtCidade.Value
Planilha1.Cells(Linha, 18) = txtUF.Value
Planilha1.Cells(Linha, 19) = txtCelular1.Value
Planilha1.Cells(Linha, 20) = txtCelular2.Value
Planilha1.Cells(Linha, 21) = cboPeriodo.Value
Planilha1.Cells(Linha, 22) = cboCiclo.Value
Planilha1.Cells(Linha, 23) = cboProfessores.Value
Planilha1.Cells(Linha, 24) = cboTurno.Value
Planilha1.Cells(Linha, 25) = cboSituacaodoAluno
Planilha1.Cells(Linha, 26) = Date 'Para ele colocar a data automaticamente
Planilha1.Cells(Linha, 27) = cboNE.Value
Planilha1.Cells(Linha, 28) = txtDataTransferencia.Value

'Inserir botão de Editar e Excluir Registro
InserirBotoesAcao Linha, Planilha1.Cells(Linha, 2)
'tinha errado uma virgula aqui


Unload CadastroAluno
'Para fechar o formulário
Msgbox "Cadastro salvo com sucesso!", vbInformation, "Cadastro salvo"


End Sub

Private Sub SalvarEdicaoCadastro() 'escrever sem acento pq é codigo
Dim Campoembranco As Boolean
Dim Linha         As Integer
Campoembranco = False

If txtNomeAluno = Empty Then Campoembranco = True
'If txtCodigo = Empty Then Campoembranco = True
'If txtDataNasc = Empty Then Campoembranco = True
'If txtPai = Empty Then Campoembranco = True
'If txtMae = Empty Then Campoembranco = True
'If txtResponsavel = Empty Then Campoembranco = True
'If txtNaturalidade = Empty Then Campoembranco = True
'If txtEndereço = Empty Then Campoembranco = True
'If txtN = Empty Then Campoembranco = True
'If txtBairro = Empty Then Campoembranco = True
'If txtCidade = Empty Then Campoembranco = True
'If txtUF = Empty Then Campoembranco = True
'If txtCelular1 = Empty Then Campoembranco = True
'If cboPeriodo = Empty Then Campoembranco = True
'If cboCiclo = Empty Then Campoembranco = True
'If cboProfessores = Empty Then Campoembranco = True
'If cboTurno = Empty Then Campoembranco = True
'If txtDatadaMatricula = Empty Then Campoembranco = True
'If cboSexo = Empty Then Campoembranco = True
'If cboNE = Empty Then Campoembranco = True
'If txtCep = Empty Then Campoembranco = True

If Campoembranco = True Then
Msgbox "Campo Obrigatório sem preenchimento.", vbInformation, "Campo em branco"
Exit Sub
'Pra parar o código caso tenha campo em branco

End If

Linha = 391

Do While Planilha1.Cells(Linha, 2) <> CDbl(lblTipoCadastro.Tag) 'cdbl trasnforma a tag texto em número
Linha = Linha + 1
Loop

'Iniciar armazenamento dos dados de cadastro do Aluno. No exemplo abaixo, o meu ID sempre vai valer o valor que esta no meu contador de registro + 1

Planilha1.Cells(Linha, 2) = Planilha1.Range("X2").Value + 1
Planilha1.Cells(Linha, 3) = txtNomeAluno.Value
Planilha1.Cells(Linha, 4) = txtDataNasc.Value
Planilha1.Cells(Linha, 5) = txtPai.Value
Planilha1.Cells(Linha, 6) = txtMae.Value
Planilha1.Cells(Linha, 7) = txtResponsavel.Value
Planilha1.Cells(Linha, 8) = txtRestricao.Value
Planilha1.Cells(Linha, 9) = txtNaturalidade.Value
Planilha1.Cells(Linha, 10) = cboSexo.Value
Planilha1.Cells(Linha, 11) = txtCep.Value
Planilha1.Cells(Linha, 12) = txtLogradouro
Planilha1.Cells(Linha, 13) = txtEndereço
Planilha1.Cells(Linha, 14) = txtN.Value
Planilha1.Cells(Linha, 15) = txtComplemento.Value
Planilha1.Cells(Linha, 16) = txtBairro.Value
Planilha1.Cells(Linha, 17) = txtCidade.Value
Planilha1.Cells(Linha, 18) = txtUF.Value
Planilha1.Cells(Linha, 19) = txtCelular1.Value
Planilha1.Cells(Linha, 20) = txtCelular2.Value
Planilha1.Cells(Linha, 21) = cboPeriodo.Value
Planilha1.Cells(Linha, 22) = cboCiclo.Value
Planilha1.Cells(Linha, 23) = cboProfessores.Value
Planilha1.Cells(Linha, 24) = cboTurno.Value
Planilha1.Cells(Linha, 25) = cboSituacaodoAluno
'Planilha1.Cells(Linha, 26) = Date 'Para ele colocar a data automaticamente
Planilha1.Cells(Linha, 27) = cboNE.Value
Planilha1.Cells(Linha, 28) = txtDataTransferencia.Value

Msgbox "Cadastro editado com sucesso.", vbInformation, "Cadastro Editado"



End Sub

Private Sub txtCep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

txtCep.MaxLength = 9
Select Case txtCep.SelStart
Case Is = 5
txtCep.SelText = "-"
End Select

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub

' Data de Nasc
Private Sub txtDataNasc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

txtDataNasc.MaxLength = 10
Select Case txtDataNasc.SelStart
Case Is = 2, 5
txtDataNasc.SelText = "/"
End Select

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub
'Numero da casa
Private Sub txtN_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub
'NCelular 1
Private Sub txtCelular1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

txtCelular1.MaxLength = 13
Select Case txtCelular1.SelStart
Case Is = 0
txtCelular1.SelText = "("
Case Is = 3
txtCelular1.SelText = ")"

End Select

End Sub

'NCelular 2
Private Sub txtCelular2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

txtCelular2.MaxLength = 13
Select Case txtCelular2.SelStart
Case Is = 0
txtCelular2.SelText = "("
Case Is = 3
txtCelular2.SelText = ")"

End Select

End Sub

' Data da Matricula
Private Sub txtDatadaMatricula_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

txtDatadaMatricula.MaxLength = 10
Select Case txtDatadaMatricula.SelStart
Case Is = 2, 5
txtDatadaMatricula.SelText = "/"
End Select

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If


End Sub

'Codigo
Private Sub txtCodigo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0

End If

End Sub

'Lançar codigos ou itens automaticos

Private Sub txtNomeAluno_Change()
Dim I As Long ' Dim é um numero longo
Planilha1.Select

Planilha1.Range("B5").Select 'B5 é onde tá meu primeiro ID
Planilha1.Range("B5") = "001"

Range("B1000").End(xlUp).Offset(1, 0).Select 'Ele vai começar da cedula 1000 e vai subir até na A8 verificando se tem número
I = Range("B1000").End(xlUp).Offset(0, 0).Value 'i é o numero id preenchido que está acima
Me.txtCodigo = I + 1 'vai pegar o id que eu tenho e somar +1


End Sub


