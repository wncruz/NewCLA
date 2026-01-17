<!--#include file="../inc/data_interfanon.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->
<%
''	- Sistema			: CLA
''	- Arquivo			: Aprov_Notificar_Termino.asp
''	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
''						  Valida e Disponibiliza os Dados na Camada de Interface do CLA
''						  Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno
''	- Objetivo			: Recebe notificação de Conclusão do processo de Aprovisionamento do Serviço
''						  (Ativação, Alteração, Cancelamento, Desativação)

'strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
strLoginRede = "ASMS"

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
'objXmlDadosForm.save(Server.MapPath("../TesteConectividade-CLA-NOTIF.xml"))

'Recupera dados do xml da página anterior
If objXmlDadosForm.parseError.errorCode < 1 Then
    strxml = objXmlDadosForm.xml
	OriSol_Descricao = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/origem").text)
	strsenha = ""
	
	
	Acao = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/acao").text)
	stridLogico = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-logico").text)
	
	ID_Tarefa_Acesso = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-tarefa").text)
	strsolid = ""
	Id_Logico_Temp = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-logico-temp").text)
	Id_Servico = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/id-servico").text)
	Data_Fim = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/data-fim-aprov").text)
	
	PE = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/pe").text)
	Porta_Pe = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/porta_pe").text)
	Cvlan = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/cvlan").text)
	Svlan = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/svlan").text)
	elan = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/elan").text)
	
	if not isnull(elan) or elan <> "" then
		Svlan = elan
	end if 
	
	Switch = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/switch").text)
	Porta_Switch = trim(objXmlDadosForm.selectSingleNode("//requisicao-cla/porta_switch").text)
	
	

	Vetor_Campos(1)="adWChar,20,adParamInput," & Acao
	Vetor_Campos(2)="adWChar,20,adParamInput,"& stridLogico
	Vetor_Campos(3)="adWChar,20,adParamInput,"& OriSol_Descricao
	Vetor_Campos(4)="adWChar,20,adParamInput,"& ID_Tarefa_Acesso
	Vetor_Campos(5)="adInteger,4,adParamInput,"& strsolid
	Vetor_Campos(6)="adWChar,30,adParamInput,"& strloginrede
	Vetor_Campos(7)="adWChar,15,adParamInput,"& Request.ServerVariables("REMOTE_ADDR")

	Vetor_Campos(8)="adWChar,20,adParamInput,"& Id_Logico_Temp
	Vetor_Campos(9)="adWChar,40,adParamInput,"& Id_Servico
	Vetor_Campos(10)="adWChar,20,adParamInput,"& Data_Fim

	Vetor_Campos(11)="adInteger,2,adParamOutput,0"
	Vetor_Campos(12)="adWChar,200,adParamOutput,a "
	Vetor_Campos(13)="adInteger,8,adParamOutput,0"
	
	Vetor_Campos(14)="adWChar,13,adParamInput,"& PE
	Vetor_Campos(15)="adWChar,20,adParamInput,"& Porta_Pe
	Vetor_Campos(16)="adWChar,5,adParamInput,"& Cvlan
	Vetor_Campos(17)="adWChar,5,adParamInput,"& Svlan
	Vetor_Campos(18)="adWChar,20,adParamInput,"& Switch
	Vetor_Campos(19)="adWChar,20,adParamInput,"& Porta_Switch

	Call APENDA_PARAM("CLA_sp_ins_Encerrar_AcessoAprovASMS",19,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	DBDescricao = ObjCmd.Parameters("RET1").value
	Aprovisi_ID = ObjCmd.Parameters("RET2").value

	''Verifica se Retornou erro
	If DBAction <> 0 Then

		strxmlResp = 				"<resposta-cla><codigo> " & Trim(DBAction) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(DBDescricao) & "</mensagem></resposta-cla>"
		Response.write strxmlResp
		
		'Checa se serviço é 0800 - E.
		Vetor_Campos(1)="adWchar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWchar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWchar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWchar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWchar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWchar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adWchar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWchar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,7" 'Notif Término
		Vetor_Campos(11)="adInteger,1,adParamInput,1"
		Vetor_Campos(12)="adNumeric,10,adParamInput," & stridLogico
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
		db.Execute(strSqlRet)

		Response.end
	Else
		'Vetor_Campos(1)="adInteger,4,adParamInput," & strSolid
		'Vetor_Campos(2)="adInteger,4,adParamInput, 55"
		'Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
		'Vetor_Campos(4)="adWchar,1,adParamInput,"
		'Vetor_Campos(5)="adWchar,100,adParamInput,STATUS AUTOMATICO"  
		'Vetor_Campos(6)="adWchar,1,adParamInput,M"

		'strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
		
	'	db.Execute(strSqlRet)

		strxmlResp = 				"<resposta-cla><codigo>0</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>Notificacao de ativacao de servico recebida com Sucesso. </mensagem></resposta-cla>"
		
		'Checa se serviço é 0800.
		Vetor_Campos(1)="adWchar,4,adParamInput," 	& OE_Ano
		Vetor_Campos(2)="adWchar,7,adParamInput," 	& OE_Numero
		Vetor_Campos(3)="adWchar,3,adParamInput," 	& OE_Item
		Vetor_Campos(4)="adWchar,20,adParamInput," 	& ID_Tarefa_Acesso
		Vetor_Campos(5)="adWchar,20,adParamInput," 	& OriSol_Descricao
		Vetor_Campos(6)="adWchar,10,adParamInput," 	& Acao
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adWchar,200,adParamInput," 	& strxmlResp
		Vetor_Campos(9)="adWchar,8000,adParamInput," 	& strxml
		Vetor_Campos(10)="adInteger,1,adParamInput,7" 'Notif Término
		Vetor_Campos(11)="adInteger,1,adParamInput,0"
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


'on error resume next 

if err.number <> 0 then

		strxmlResp = 				"<resposta-cla><codigo> " & Trim(Str(err.number)) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(err.Description) & "</mensagem></resposta-cla>"

		Response.write strxmlResp
		Response.end

end if
%>