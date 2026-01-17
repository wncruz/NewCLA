<!--#include file="../inc/data_interfanon.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->
<%
''	- Sistema			: CLA
''	- Arquivo			: Aprov_Mudanca_Titularidade.asp
''	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
''						  Valida e Disponibiliza os Dados na Camada de Interface do CLA
''						  Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno
''	- Objetivo			: Mudança de Titularidade

strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

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
'objXmlDadosForm.save(Server.MapPath("../TesteMudancaTitularidade.xml"))

'Recupera dados do xml da página anterior
If objXmlDadosForm.parseError.errorCode < 1 Then
    strxml = objXmlDadosForm.xml

	OriSol_Descricao = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/origem").text)
	if OriSol_Descricao <> "ASMS" then
		strsenha = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/senha").text)
		if UCASE(strsenha) <> "A3590023DF66AC92AE35E35E33160" then
		   strxmlResp = 				"<resposta-cla><codigo>999</codigo>"
			strxmlResp = strxmlResp  & 	"<mensagem>Senha de interface incorreta</mensagem></resposta-cla>"
			response.write strxmlResp
		  	response.end
		end if
	end if
	
	Acao = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/acao").text)
	stridLogico = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-logico").text)
	OriSol_Descricao = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/origem").text)
	ID_Tarefa_Acesso = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-tarefa").text)
	
	Cli_CC = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/conta").text)
	Cli_SubCC = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/subconta").text)
	Cli_Nome = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/razao-social").text)
	Cli_NomeFantasia = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/cliente/nome-fantasia").text)
	
	if OriSol_Descricao <> "ASMS" then
		Oe_Ano = trim(objXmlDadosForm.selectSingleNode("//servico/order-entry/ano").text)
		Oe_Numero = trim(objXmlDadosForm.selectSingleNode("//servico/order-entry/numero").text)
		Oe_Item = trim(objXmlDadosForm.selectSingleNode("//servico/order-entry/item").text)
	else 
		OE_Numero 			=  trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/numero").text)	 
		OE_Item 			=  trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/item").text)
		Id_Servico 			=  trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-servico").text)
	end if	
	
	if OriSol_Descricao = "ASMS" then 
		
		Header_Original 	=  objXmlDadosForm.selectSingleNode("//requisicao-cla/header-original").xml 
		External_Reflist	=  objXmlDadosForm.selectSingleNode("//requisicao-cla/external-reflist").xml
		Servico 			= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico").text)
		ServicoDesc 		= trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/servico/order-entry/servico-desc").text)
		ServicoDesc 		= "REALIDADE IP/" + Servico
				
	end if 
	
	
	Vetor_Campos(1)="adWChar,20,adParamInput," & Acao
	Vetor_Campos(2)="adWChar,20,adParamInput,"& stridLogico
	Vetor_Campos(3)="adWChar,10,adParamInput,"& OriSol_Descricao
	Vetor_Campos(4)="adWChar,15,adParamInput,"& ID_Tarefa_Acesso
	
	Vetor_Campos(5)="adWChar,15,adParamInput,"& Cli_CC
	Vetor_Campos(6)="adWChar,4,adParamInput,"& Cli_SubCC
	Vetor_Campos(7)="adWChar,60,adParamInput,"& Cli_Nome
	Vetor_Campos(8)="adWChar,60,adParamInput,"& Cli_NomeFantasia
	
	Vetor_Campos(9)="adInteger,4,adParamInput,"& Oe_Ano
	Vetor_Campos(10)="adInteger,7,adParamInput,"& Oe_Numero
	Vetor_Campos(11)="adInteger,4,adParamInput,"& Oe_Item
	
	Vetor_Campos(12)="adWChar,30,adParamInput,"& strloginrede
	Vetor_Campos(13)="adWChar,15,adParamInput,"& Request.ServerVariables("REMOTE_ADDR")
	
	Vetor_Campos(14)="adInteger,2,adParamOutput,0"
	Vetor_Campos(15)="adWChar,200,adParamOutput,a "
		
	Vetor_Campos(16)="adWChar,50,adParamInput,"& Id_Servico
	
	Vetor_Campos(17)="adWChar,4000,adParamInput,"& Header_Original
	Vetor_Campos(18)="adWChar,3000,adParamInput,"& External_Reflist
	
	Vetor_Campos(19)="adWChar,40,adParamInput,"& Servico
	Vetor_Campos(20)="adWChar,50,adParamInput,"& ServicoDesc

	Call APENDA_PARAM("CLA_sp_upd_MudancaTitularidadeAprovASMS",20,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	DBDescricao = ObjCmd.Parameters("RET1").value

	If DBAction <> 0 Then

		strxmlResp = 				"<resposta-cla><codigo> " & Trim(DBAction) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(DBDescricao) & "</mensagem></resposta-cla>"
		'Response.write strxmlResp
		
		'Checa se serviço é 0800 - E.
		Vetor_Campos(1)="adWchar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWchar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWchar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWchar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWchar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWchar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput,0"
		Vetor_Campos(8)="adWchar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWchar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,8" 'Mud Tit
		Vetor_Campos(11)="adInteger,1,adParamInput,1"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & stridLogico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)
		
		Response.end
	Else
		strxmlResp = 				"<resposta-cla><codigo>0</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>Mudanca de Titularidade recebida com Sucesso. </mensagem></resposta-cla>"
		
		'Checa se serviço é 0800 - E.
		Vetor_Campos(1)="adWchar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWchar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWchar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWchar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWchar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWchar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput,0"
		Vetor_Campos(8)="adWchar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWchar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,8" 'Mud Tit
		Vetor_Campos(11)="adInteger,1,adParamInput,1"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & stridLogico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)

		Response.Write strxmlResp
		Response.end
	End If

Else

		strxmlResp = 				"<resposta-cla><codigo> 9999 </codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>XML nao Enviado ao CLA</mensagem></resposta-cla>"

		Response.Write strxmlResp
		Response.end

End If


on error resume next 

if err.number <> 0 then

		strxmlResp = 				"<resposta-cla><codigo> " & Trim(Str(err.number)) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(err.Description) & "</mensagem></resposta-cla>"

		Response.write strxmlResp
		Response.end

end if
%>