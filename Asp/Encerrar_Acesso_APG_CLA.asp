<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->


<%
''	- Sistema			: CLA
''	- Arquivo			: Encerra_Acesso_Apg_CLA.asp
''	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
''						  Valida e Disponibiliza os Dados na Camada de Interface do CLA
''						  Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno
''	- Objetivo			: Recebe notificação de Conclusão do processo de Aprovisionamento do Serviço
''						  (Ativação, Alteração, Cancelamento, Desativação)

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
Set objXmlDadosDB = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")

'Recupera dados do xml da página anterior
If Trim(Request.Form("hdnxml")) <> "" Then

	strxml = trim(Request.Form("hdnxml"))

	strxml = Replace(strxml,Chr(13),"")
	strxml = Replace(strxml,Chr(10),"")

	objXmlDadosForm.Async = "false"
	objXmlDadosForm.preserveWhiteSpace = True
	objXmlDadosForm.loadXML(strxml)

	'Response.Write "<B>readyState=</B> " & objXmlDadosForm.readyState &"<br>"

	If 	Len(Trim(objXmlDadosForm.parseError.errorCode)) > 1 Then

		strxmlResp = 				"<resposta-cla><codigo> Erro ao realizar parsing do XML enviado: " & objXmlDadosForm.parseError.errorCode & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(objXmlDadosForm.parseError.reason)
		strxmlResp = strxmlResp  & 	Trim(objXmlDadosForm.parseError.line) & "</mensagem></resposta-cla>"

		Response.Write strxmlResp
		Response.end

	End IF

	'*' Claudio - Atribuir campos do XML as Variaveis.
	'*' objXmlDadosForm
	strSolid = ""
	strprocesso = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/processo").text)
	stracao = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acao").text)
	stridTarefaApg = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/id-tarefa-apg").text)
	stridLogico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/id-logico").text)
	strdt_encerramento = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/data-encerramento").text)
	
	'Checa se serviço é 0800.
	Oe_numero = ""
	Oe_ano = ""
	Oe_item = ""
	Id_logico = stridLogico
	Processo = strprocesso
	Acao = stracao
	
	call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,7,strxml)
	
	StrSQL = "select max(sol_id) as sol_id from cla_solicitacao where Acl_IDAcessologico = '" & stridLogico & "'"
	Set objRSSolic = db.Execute(strSql)
	strSolid = objRSSolic("sol_id")

	'Finaliza = "0" 'Não Finaliza.
	'If 	Finaliza = 1 Then

	'	strxmlResp = 				"<resposta-cla><codigo>0</codigo>"
	'	strxmlResp = strxmlResp  & 	"<mensagem> Processo Realizado com Sucesso</mensagem></resposta-cla>"

	'	Response.Write strxmlResp
	'	Response.end

	'End IF

	'response.write "Processo: " & strprocesso & chr(13)
	'response.write "Ação: " & stracao & chr(13)
	'response.write "ID Logico: " & stridLogico & chr(13)
	'response.write "Data de Encerramento: " & strdt_encerramento & chr(13)
	'response.write chr(13)

	'Chama a Procedure passando os Parametros

	Vetor_Campos(1)="adWChar,20,adParamInput," & strprocesso
	Vetor_Campos(2)="adWChar,20,adParamInput,"& stracao
	Vetor_Campos(3)="adWChar,10,adParamInput,"& strSolid
	Vetor_Campos(4)="adWChar,15,adParamInput,"& stridLogico
	Vetor_Campos(5)="adWChar,20,adParamInput,"& stridTarefaApg
	Vetor_Campos(6)="adWChar,20,adParamInput,"& strOE_Solicitacao_OEU
	Vetor_Campos(7)="adWChar,10,adParamInput,"& strdt_encerramento

	Vetor_Campos(8)="adInteger,2,adParamOutput,0"
	Vetor_Campos(9)="adWChar,100,adParamOutput,a "

	Call APENDA_PARAM("CLA_sp_ins_Encerrar_Acesso",9,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	DBDescricao = ObjCmd.Parameters("RET1").value

	''Verifica se Retornou erro
	If DBAction <> 0 Then

		strxmlResp = 				"<resposta-cla><codigo> " & Trim(DBAction) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(DBDescricao) & "</mensagem></resposta-cla>"
		Response.write strxmlResp
		Response.end

	Else

		
		
		Vetor_Campos(1)="adWChar,10,adParamInput,"& strSolid
		Vetor_Campos(2)="adWChar,20,adParamInput, 7 " '7 - Termino
		Vetor_Campos(3)="adWChar,20,adParamInput," & trim(strprocesso)
		Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(stracao)

		StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",4,Vetor_Campos)

		db.execute(strSQL)

		strxmlResp = 				"<resposta-cla><codigo>0</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>Solicitacao Gravada com Sucesso. </mensagem></resposta-cla>"

		Response.Write strxmlResp
		Response.end

	End If

Else

		strxmlResp = 				"<resposta-cla><codigo> 9999 </codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>XML não Enviado ao CLA</mensagem></resposta-cla>"

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