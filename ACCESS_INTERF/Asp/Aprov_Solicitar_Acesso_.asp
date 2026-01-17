<!--#include file="../../inc/data_interfanon.asp"-->
<!--#include file="../../inc/xmlAcessos.asp"-->
<%
'	- Sistema			: CLA
'	- Arquivo			: Aprov_Solicitar_Acesso.asp
'	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
'						Valida e Disponibiliza os Dados na Camada de Interface do CLA
'						Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno

'Retira Chr(10) e Chr(13) do objeto xml para enviar p/ o JS/SQL
Function FormatarXml(objXml)
	
	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(objXml.xml,Chr(13),"")
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")
	
	FormatarXml = strXmlDadosAux
End Function

Function indexPropAcesso(strProp)
	Select Case strProp
		Case "TER"
			indexPropAcesso = 0
		Case "EBT"
			indexPropAcesso = 1
		Case "CLI"
			indexPropAcesso = 2
	End Select
End Function

'Adiciona um Node a um objeto XML se existir atualiza
Function AdicionarNode(strNomeNode,objXML,varValorNode)
	
	Dim objNodeFilho
	Dim objNodeList
	
    If objXML.xml = "" Then
	   objXML.loadXML "<xmlDados></xmlDados>"
	End If
	
	'Verifica se já existe
	Set objNodeList = objXml.selectNodes("*/" & strNomeNode)
	
	if objNodeList.Length = 0 then
		'Cria
		Set objNodeFilho = objXML.createNode("element", strNomeNode, "")
		objNodeFilho.text = varValorNode
		objXML.documentElement.appendChild (objNodeFilho)
	Else
		'Atualiza
		objNodeList.Item(0).Text = varValorNode
	End If
	
	Set AdicionarNode = objXML
	
End Function

'Adiciona um Node a um objeto XML se existir atualiza
Function UpdNodeAcesso(strNomeNode,objXML,varValorNode)
	
	Dim objNodeFilho
	Dim objNodeList
	
	If objXML.xml = "" Then
		objXML.loadXML "<xmlDados></xmlDados>"
	End If

	'Verifica se já existe
	Set objNodeList = objXml.selectNodes("//Acesso/" & strNomeNode)

	if objNodeList.Length = 0 then
		'Cria
		Set objNodeList = objXml.selectNodes("//Acesso")
		Set objNodeFilho = objXML.createNode("element", strNomeNode, "")
		objNodeFilho.text = varValorNode
		objNodeList(0).appendChild (objNodeFilho)
	Else
		'Atualiza
		objNodeList.Item(0).Text = varValorNode
	End If
	
	Set UpdNodeAcesso = objXML
	
End Function

Dim objXmlDadosForm
Dim objXmlDadosDB
Dim objXmlRetorno
Dim strxmlResp

Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

'Recupera dados do xml da página anterior
If objXmlDadosForm.parseError.errorCode < 1 Then
	strxml = objXmlDadosForm.xml
	
	
'********** JCARTUS-Dez/2011 : Implementação da senha de interface para APROV-SOLICITAR-ACESSO	
	strsenha = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/senha").text)
	if UCASE(strsenha) <> "A3590023DF66AC92AE35E35E33160" then
	   strxmlResp = 				"<resposta-cla><codigo>999</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>Senha de interface incorreta</mensagem></resposta-cla>"
		Response.write strxmlResp
	  response.end
	end if	
'**********	
	
	objXmlDadosForm.Async = "false"
	objXmlDadosForm.preserveWhiteSpace = True
		
	'Acao = objXmlDadosForm.selectSingleNode("//requisicao-cla/acao").text


	'Response.Write "<B>readyState=</B> " & objXmlDadosForm.readyState &"<br>"
	
	If 	Len(Trim(objXmlDadosForm.parseError.errorCode)) > 1 Then
		
		strxmlResp = "<resposta-cla> "
		strxmlResp = strxmlResp  & 	"<codigo> " & objXmlDadosForm.parseError.errorCode & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem> Erro durante Parsing do xml. Motivo:<" & Trim(objXmlDadosForm.parseError.reason)
		strxmlResp = strxmlResp  & 	Trim(objXmlDadosForm.parseError.line) & "</mensagem>"
		strxmlResp = strxmlResp  & 	"</resposta-cla> "
		
		Response.Write strxmlResp
		
		'Checa se serviço é 0800 - E.
		Aprovisi_ID = 0		
		Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWChar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adWChar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWChar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,1" 'Solicitar Send
		Vetor_Campos(11)="adInteger,1,adParamInput,1"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)

		Response.end
		
	End IF
	
	Processo = ""
	Acao = ""
	ID_Tarefa_Acesso  = ""
	Conta_Corrente = ""
	SubConta = ""
	Razao_Social = ""
	Nome_Fantasia = ""
	OE_Ano = ""
	OE_Numero = ""
	OE_Item  = ""
	Servico = ""
	Designacao_Servico = ""
	Tipo_Contrato_Cliente = ""
	Numero_Contrato_Cliente = ""
	
	DT_Prevista_Ativacao_Serv = ""
	Contato_Tecnico_Cliente = ""
	Telefone_Tecnico_Cliente = ""
	Email_Tecnico_Cliente = ""
	
	Velocidade_Servico = ""
	Tipo_Interface_Cliente = ""
	Tipo_Interface_Embratel = ""
	DT_Inicio_Servico_Temp = ""
	DT_Fim_Servico_Temp = ""
	
	Username_Cadastrador = ""
	Telefone_Cadastrador = ""
	CNPJ = ""
	Inscricao_Estadual = ""
	Inscricao_Municipal = ""
	
	CNL = ""
	Proprietario_Endereco = ""
	Bairro = ""
	CEP = ""
	Cidade = ""
	
	Complemento = ""
	Nome_Logradouro = ""
	Numero_Predio = ""
	Tipo_Logradouro = ""
	UF = ""
	
	Identif_Centro_Cliente = ""
	Numero_SEV = ""
	Id_Logico = ""
	
	Dt_Desativacao = ""
	Reaproveitar_Fisico = ""
	
	Distribuicao			= ""
	Rota					= ""
	Observacao				= ""
	ID_Tarefa_Acesso_ATV	= ""
	Rollback				= ""
	Codigo_Sap				= ""
	Estacao_Cliente			= ""
	ServicoDesc				= ""
	
	Acesso_Migracao 		= ""
	Acesso_TipoRede			= ""
	
	'GPON
	Vel_Voz					= ""
	Dados					= ""
	
	'Chamar a Procedure passando os Parametros
	'Requisicao CLA
	
	Acao = objXmlDadosForm.selectSingleNode("//requisicao-cla/acao").text
	ID_Tarefa_Acesso = objXmlDadosForm.selectSingleNode("//requisicao-cla/id-tarefa").text
	
	if Acao = "CAN" then
		ID_Tarefa_Can = objXmlDadosForm.selectSingleNode("//requisicao-cla/id-tarefa-can").text
		Observacao = objXmlDadosForm.selectSingleNode("//requisicao-cla/observacao").text
	end if
	
	Id_Logico =  objXmlDadosForm.selectSingleNode("//requisicao-cla/id-logico").text
	
	if Acao <> "AltCadastral" then
	' Serviço/Order Entry
	OE_Ano 				= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/ano").text)
	OE_Numero 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/numero").text)
	OE_Item 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/item").text)
	OriSol_Descricao 	= trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/origem").text)
	end if 
	
	If ( Trim(Acao) = "ATV"  or Trim(Acao) = "ALT"  or Trim(Acao) = "DES" ) Then
		
		'Dados do Cliente
		 Conta_Corrente = objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/conta").text
		 SubConta = objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/subconta").text
		 Razao_Social = objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/razao-social").text
		 Nome_Fantasia = objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/nome-fantasia").text
		
		 Servico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico").text)
		 ServicoDesc = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico-desc").text)
		 Designacao_Servico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/designacao").text)
		 Numero_Contrato_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/contrato").text)
		 Tipo_Contrato_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/tipo-contrato").text)
		 DT_Prevista_Ativacao_Serv = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/data-prevista").text)
		
		'Serviço/Order Entry/Contato Tecnico
		 Contato_Tecnico_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/contato-tecnico/cliente").text)
		 Telefone_Tecnico_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/contato-tecnico/telefone").text)
		 Email_Tecnico_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/contato-tecnico/e-mail").text)
		
		'Serviço/Order Entry
		 Velocidade_Servico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/velocidade").text)
		 Tipo_Interface_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/interface-cliente").text)
		 Tipo_Interface_Embratel = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/interface-embratel").text)
		
		'Serviço/Order Entry/Serviço Temporário
		 DT_Inicio_Servico_Temp = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico-temporario/data-inicio").text)
		 DT_Fim_Servico_Temp = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico-temporario/data-fim").text)
		
		'Serviço/Order Entry/Cadastrador
		 Username_Cadastrador = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/cadastrador/username").text)
		 Telefone_Cadastrador = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/cadastrador/telefone").text)
		
		'Serviço/Order Entry/Endereço Instalação
		 CNPJ = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/cnpj").text)
		 Inscricao_Estadual = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/inscricao-estadual").text)
		 Inscricao_Municipal = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/inscricao-municipal").text)
		 CNL = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/cnl").text)
		 Proprietario_Endereco = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/proprietario").text)
		 Bairro =  trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/bairro").text)
		 CEP = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/cep").text)
		 Cidade = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/cidade").text)
		 Complemento = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/complemento").text)
		 Nome_Logradouro = objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/logradouro").text
		 Numero_Predio   = objXmlDadosForm.selectSingleNode("//requisicao-cla/servico/order-entry/endereco-instalacao/numero").text
		 UF = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/uf").text)
		 Identif_Centro_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/centro-cliente").text)
		 Tipo_Logradouro = objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/endereco-instalacao/tipo-logradouro").text
		 
		'Serviço/Order Entry
		 Numero_SEV 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/numero-sev").text)
		 
		 if isnumeric(Numero_SEV) = false then
  		  Numero_SEV = ""
		 end if
		 
		 ' Código SAP
		 Codigo_Sap 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/codDescargaSAP").text)
		 ' Estação Cliente Ex.: RJO AM
		 Estacao_Cliente 		= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/estCliente").text)
		 
		 'BSOD
		 if OriSol_Descricao = "SGAP" then
		 	Acesso_Migracao		= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acesso/migracao").text)
		 	Acesso_TipoRede		= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acesso/tipo-rede").text)
		 end if
		 
		 'GPON
		 Vel_Voz = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/vel_voz").text)
		 Dados = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/dados").text)
	End If
	
	if ( OriSol_Descricao = "SGAV" and ( Trim(Acao) = "ATV"  or Trim(Acao) = "ALT" or Trim(Acao) = "CAN" ) ) then
		'Reenvio
		Reenvio = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/reenvio").text)
		
	End if
	
	if ( OriSol_Descricao = "SGAV"  ) then
		'PABX_VIRTUAL - BROADSOFT / IMS		
		PABX_VIRTUAL = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/pabx_virtual_plataforma").text)
		
		'if PABX_VIRTUAL = "1" then	
		'	PABX_VIRTUAL = "SIM"
		'elseif PABX_VIRTUAL = "2" then
		'	PABX_VIRTUAL = "NÃO"
		'else 
		'	PABX_VIRTUAL = "NULL"
		'end if	
		
		if PABX_VIRTUAL = "S" then	
			PABX_VIRTUAL = "SIM"
		elseif PABX_VIRTUAL = "N" then
			PABX_VIRTUAL = "NÃO"
		else 
			PABX_VIRTUAL = "NULL"
		end if	
		
		
	End if
	
	if Acao = "AltCadastral" then
		 Servico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/servico").text)
		 ServicoDesc = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/servico-desc").text)
		 Numero_Contrato_Cliente = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/contrato").text)
		 Numero_SEV 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/numero-sev").text)
		 Velocidade_Servico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/velocidade").text)
		 
		 if isnumeric(Numero_SEV) = false then
  		  Numero_SEV = ""
		 end if
		 
	end if
	
	if ( OriSol_Descricao = "SGAV"  ) then
		'Reenvio
		DESIG_SITE = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/desig-site").text)

	End if	
	
	'Tratamento de campos.
	IF TRIM (CEP) <> "" THEN
		CEP = LEFT(CEP,5) & "-" & RIGHT(CEP,3)
	END IF
	
	'Preenche parâmetros para execução da Proc de Validação/Gravação da Interface
	
	Vetor_Campos(1)="adWChar,20,adParamInput,"& Acao
	Vetor_Campos(2)="adinteger,5,adParamInput,"& ID_Tarefa_Acesso
	Vetor_Campos(3)="adinteger,5,adParamInput,"& ID_Tarefa_Can
	
	Vetor_Campos(4)="adWChar,11,adParamInput,"& Conta_Corrente
	Vetor_Campos(5)="adWChar,4,adParamInput,"& SubConta
	Vetor_Campos(6)="adWChar,60,adParamInput,"& Razao_Social
	Vetor_Campos(7)="adWChar,60,adParamInput,"& Nome_Fantasia
	
	Vetor_Campos(8)="adWChar,4,adParamInput,"& left(OE_Ano,4)
	Vetor_Campos(9)="adWChar,7,adParamInput,"& OE_Numero
	Vetor_Campos(10)="adWChar,5,adParamInput,"& OE_Item
	
	Vetor_Campos(11)="adWChar,40,adParamInput,"& Servico
	Vetor_Campos(12)="adWChar,100,adParamInput,"& Designacao_Servico
	Vetor_Campos(13)="adWChar,1,adParamInput,"& Tipo_Contrato_Cliente
	Vetor_Campos(14)="adWChar,30,adParamInput,"& Numero_Contrato_Cliente
	
	Vetor_Campos(15)="adWChar,10,adParamInput,"& DT_Prevista_Ativacao_Serv
	Vetor_Campos(16)="adWChar,30,adParamInput,"& left(Contato_Tecnico_Cliente,30)
	Vetor_Campos(17)="adWChar,30,adParamInput,"& Telefone_Tecnico_Cliente
	Vetor_Campos(18)="adWChar,60,adParamInput,"& Email_Tecnico_Cliente
	
	Vetor_Campos(19)="adWChar,10,adParamInput,"& Velocidade_Servico
	Vetor_Campos(20)="adWChar,30,adParamInput,"& Tipo_Interface_Cliente
	Vetor_Campos(21)="adWChar,30,adParamInput,"& Tipo_Interface_Embratel
	Vetor_Campos(22)="adWChar,10,adParamInput,"& DT_Inicio_Servico_Temp
	Vetor_Campos(23)="adWChar,10,adParamInput,"& DT_Fim_Servico_Temp
	
	Vetor_Campos(24)="adWChar,30,adParamInput,"& Username_Cadastrador
	Vetor_Campos(25)="adWChar,30,adParamInput,"& Telefone_Cadastrador
	Vetor_Campos(26)="adWChar,20,adParamInput,"& CNPJ
	Vetor_Campos(27)="adWChar,20,adParamInput,"& Inscricao_Estadual
	Vetor_Campos(28)="adWChar,20,adParamInput,"& Inscricao_Municipal
	
	Vetor_Campos(29)="adWChar,4,adParamInput,"& CNL
	Vetor_Campos(30)="adWChar,60,adParamInput,"& Proprietario_Endereco
	Vetor_Campos(31)="adWChar,30,adParamInput,"& Bairro
	Vetor_Campos(32)="adWChar,10,adParamInput,"& CEP
	Vetor_Campos(33)="adWChar,60,adParamInput,"& Cidade
	
	'CH-48322GQG 
	Vetor_Campos(34)="adWChar,30,adParamInput,"& LEFT(Complemento,30)
	Vetor_Campos(35)="adWChar,60,adParamInput,"& LEFT(Nome_Logradouro,60)
	Vetor_Campos(36)="adWChar,10,adParamInput,"& Numero_Predio
	Vetor_Campos(37)="adWChar,20,adParamInput,"& Tipo_Logradouro
	Vetor_Campos(38)="adWChar,2,adParamInput,"& UF
	Vetor_Campos(39)="adWChar,10,adParamInput,"& Identif_Centro_Cliente
	Vetor_Campos(40)="addouble,18,adParamInput,"& Numero_SEV
	Vetor_Campos(41)="adWChar,10,adParamInput,"& Id_Logico
	Vetor_Campos(42)="adWChar,10,adParamInput,"& Dt_Desativacao
	Vetor_Campos(43)="adWChar,3,adParamInput,"& Reaproveitar_Fisico
	
	Vetor_Campos(44)="adinteger,12,adParamInput,"& Codigo_Sap
	Vetor_Campos(45)="adWChar,12,adParamInput,"& Estacao_Cliente
	Vetor_Campos(46)="adWChar,20,adParamInput,"& OriSol_Descricao
	Vetor_Campos(47)="adWChar,50,adParamInput,"& ServicoDesc
	
	'BSOD:
	Vetor_Campos(48)="adWChar,1,adParamInput,"& Acesso_Migracao
	Vetor_Campos(49)="adinteger,8,adParamInput,"& Acesso_TipoRede
	
	'Reenvio
	Vetor_Campos(50)="adWChar,1,adParamInput,"& Reenvio
	'GPON
	Vetor_Campos(51)="adWChar,30,adParamInput,"& Vel_Voz
	Vetor_Campos(52)="adWChar,1,adParamInput,"& Dados
	
	'ASMS:
	Vetor_Campos(53)="adWChar,10,adParamInput,"& Id_Logico_Temp
	Vetor_Campos(54)="adWChar,1,adParamInput,"& Indicador_Alterar
	Vetor_Campos(55)="adWChar,50,adParamInput,"& Id_Servico
	Vetor_Campos(56)="adWChar,10,adParamInput,"& Velocidade_Total
	
	Vetor_Campos(57)="adWChar,4000,adParamInput,"& Header_Original
	Vetor_Campos(58)="adWChar,3000,adParamInput,"& External_Reflist
	
	Vetor_Campos(59)="adWChar,50,adParamInput,"& Id_Acesso
	
	Vetor_Campos(60)="adInteger,2,adParamOutput,0"
	Vetor_Campos(61)="adWChar,100,adParamOutput,a "
	
	Vetor_Campos(62)="adWChar,200,adParamInput,"& Observacao
	
	Vetor_Campos(63)="adWChar,50,adParamInput,"& DESIG_SITE
	
	Vetor_Campos(64)="adWChar,10,adParamInput,"& PABX_VIRTUAL
	
	Call APENDA_PARAM("CLA_sp_ins_Aprovisionador",64,Vetor_Campos)
	
	'Vetor_Campos(53)="adInteger,2,adParamOutput,0"
	'Vetor_Campos(54)="adWChar,100,adParamOutput,a "
	
	'Call APENDA_PARAM("CLA_sp_ins_Aprovisionador",54,Vetor_Campos)
	'response.write APENDA_PARAMSTR("CLA_sp_ins_Aprovisionador",54,Vetor_Campos)
    ObjCmd.Execute'pega dbaction
	
	DBAction = ObjCmd.Parameters("ret").value
	DBDescricao = ObjCmd.Parameters("ret1").value
	
	''Verifica se Retornou erro
	If trim(DBAction) <> "0" and trim(DBAction) <> "" Then
		
		strxmlResp = 	"<resposta-cla><codigo>" & Trim(DBAction) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(DBDescricao) & "</mensagem></resposta-cla>"
		
		Response.write strxmlResp
		
		objXmlRetorno.loadXML(strxmlResp)
		
		strxmlResp = "Erro ao gravar na Interface Acesso => CLA - Acao: Solicitar Acesso Aprovisionador " & Trim(DBAction) & Trim(DBDescricao)
		
		'Checa se serviço é 0800 - E.
		Aprovisi_ID = 0		
		Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWChar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adWChar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWChar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,1" 'Solicitar Send
		Vetor_Campos(11)="adInteger,1,adParamInput,1"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & Id_Logico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)
		
		Response.End
		
	Else
		
		strxmlResp = "<resposta-cla> "
		strxmlResp = strxmlResp  &  "<codigo>0</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>Solicitacao Recebida com Sucesso. </mensagem>"
		strxmlResp = strxmlResp  & 	"</resposta-cla> "
		
		Response.Write strxmlResp
		
		'Checa se serviço é 0800.
		Aprovisi_ID = 0		
		Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWChar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adWChar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWChar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,1" 'Solicitar Send
		Vetor_Campos(11)="adInteger,1,adParamInput,0"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & Id_Logico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)
		
	End If
	
Else
	
	strxmlResp = "<resposta-cla> "
	strxmlResp = strxmlResp  & 	"<codigo>9999</codigo>"
	strxmlResp = strxmlResp  & 	"<mensagem>XML nao Enviado ao CLA</mensagem>"
	strxmlResp = strxmlResp  & 	"</resposta-cla> "
	
	Response.Write strxmlResp
	
	'Checa se serviço é 0800 - E.
	Aprovisi_ID = 0		
	Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
	Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
	Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
	Vetor_Campos(4)="adWChar,20,adParamInput," 	& ID_Tarefa_Acesso
	Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
	Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
	Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
	Vetor_Campos(8)="adWChar,200,adParamInput," 	& strxmlResp
	Vetor_Campos(9)="adWChar,8000,adParamInput," 	& strxml
	Vetor_Campos(10)="adInteger,1,adParamInput,1" 'Solicitar Send
	Vetor_Campos(11)="adInteger,1,adParamInput,1"
	Vetor_Campos(12)="adNumeric,10,adParamInput," & Id_Logico
	
	strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
	db.Execute(strSqlRet)
	
	Response.end
	
End If
%>
