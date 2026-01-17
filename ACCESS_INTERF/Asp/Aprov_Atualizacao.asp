<!--#include file="../../inc/data_interfanon.asp"-->
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
		
		'Response.Write strxmlResp
		
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
	
	rede_wireless           = ""
	
	'Chamar a Procedure passando os Parametros
	'Requisicao CLA
	
	Acao = objXmlDadosForm.selectSingleNode("//requisicao-cla/acao").text
	'ID_Tarefa_Acesso = objXmlDadosForm.selectSingleNode("//requisicao-cla/id-tarefa").text
	
	Id_Logico =  objXmlDadosForm.selectSingleNode("//requisicao-cla/id-logico").text
	
	PE 					= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/pe").text)
	porta_pe 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/porta_pe").text)
	cvlan 				= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/cvlan").text)
	svlan 				= trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/svlan").text)
	OriSol_Descricao 	= trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/origem").text)
	
	 
	
	Vetor_Campos(1)="adWChar,20,adParamInput,"& Acao
		
	Vetor_Campos(2)="adWChar,10,adParamInput,"& Id_Logico
	Vetor_Campos(3)="adWChar,20,adParamInput,"& OriSol_Descricao
	Vetor_Campos(4)="adWChar,50,adParamInput,"& PE
	
	Vetor_Campos(5)="adWChar,50,adParamInput,"& Porta_PE
	Vetor_Campos(6)="adWChar,10,adParamInput,"& cvlan
	Vetor_Campos(7)="adWChar,10,adParamInput,"& svlan
	
	Vetor_Campos(8)="adInteger,2,adParamOutput,0"
	Vetor_Campos(9)="adWChar,100,adParamOutput,a "
	
	Call APENDA_PARAM("CLA_sp_ins_AtualizacaoLote",9,Vetor_Campos)
	
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
