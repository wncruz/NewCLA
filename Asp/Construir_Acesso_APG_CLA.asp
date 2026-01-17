<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->
<!--#include file="../inc/EnviarRetornoConstr_Apg.asp"-->

<%
'	- Sistema			: CLA
'	- Arquivo			: Construir_Acesso_APG_CLA.asp
'	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
'						Valida e Disponibiliza os Dados na Camada de Interface do CLA
'						Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno

'Set db = server.createobject("ADODB.Connection")
'db.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
'db.ConnectionTimeout = 0
'db.CommandTimeout = 0
'db.open

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

Function EnviarRetornoConstr_Apg_CAN_Estocar(dblIdLogico, IdInterfaceAPG, dblSolid)

	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim strXmlEndereco
	Dim EncontrouDados
	Dim DescErro
	Dim CodErro

	Dim IDTarefa, DataEstocagem

	StrRetorno = ""
	IDTarefa = ""
	DataEstocagem = ""

    %><!--#include file="../inc/conexao_apg.asp"--><%
	strXmlEndereco = ""
	EncontrouDados = false

	StrClasse = "INTERFCONSTRUIRRETURN"
	
	if IdInterfaceAPG <> "" then

			''Response.Write "<script language=javascript>alert('Encontrou Logico:" & dblIdLogico &":"& dblSolid &":" & IdInterfaceAPG & " ')</script>"

			Vetor_Campos(1)="adWChar,50,adParamInput,null "
			Vetor_Campos(2)="adInteger,1,adParamInput,null "
			Vetor_Campos(3)="adWChar,50,adParamInput," & IdInterfaceAPG
			Vetor_Campos(4)="adInteger,1,adParamInput,null "
			Vetor_Campos(5)="adInteger,1,adParamInput,null "

			strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
			Set objRSDadosInterf = db.Execute(strSql)

			'response.write "Sql: " & strSql '@@DEBUG

			If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

				'Response.Write "<script language=javascript>alert('EncontrouDados True ')</script>"
				EncontrouDados = True

				IDTarefa = objRSDadosInterf("ID_Tarefa_Apg")

			End If

			If EncontrouDados = True Then

					'Response.Write "<script language=javascript>alert('Encontrou Dados')</script>" '@@DEBUG

					Strxml			=   "<soap:Envelope "
					Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance"""
					Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema"""
					Strxml = Strxml &   " xmlns:soap=" &"""http://schemas.xmlsoap.org/soap/envelope/"""

					Strxml = Strxml & 	"> <soap:Body> "

					'## <!-- Define a operação sendo realizada (executar classe) --> "
					Strxml = Strxml & 	"	<executeClass> "

					'## <!-- Ambiente do Apia a ser chamado --> "
					Strxml = Strxml & 	"		<envName>APG</envName> "

					'## <!-- Nome da classe de negócio, tal como configurada no Apia --> "
					Strxml = Strxml & 	"		<className>" & StrClasse & "</className> "

					'## <!-- Parâmetros configurados na classe --> "
					Strxml = Strxml & 	"		<parameters> "

					Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & objRSDadosInterf("Processo") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & objRSDadosInterf("Acao") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & dblIdLogico  & "</parameter> "


					Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & IdInterfaceAPG & "</parameter> " 'OBRIG
					Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" &  "</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoConstrucao""" &">" &"</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & "OK" &"</parameter> "   'OBRIG
					Strxml = Strxml & 	"		<parameter name=" & """dataEstocagem""" &">" & Date &"</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroCodigoProvedor""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroNomeProvedor""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroEstacao""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroSlot""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroTimeslot""" &">" & "</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeInstaladoraRecurso""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeEmpresaConstrutoraInfra""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtDataAceitacaoInfra""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroAcessoAde""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroOtsAcessoEmbratel""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtDesignacaoBandabasicaCriada""" &">" &  "</parameter> "

					Strxml = Strxml & 	"			<parameter name=" & """bloco""" &">" & "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """cabo"""  &">" & "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """par""" &">"  &  "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """pino""" &">" &  "</parameter> "

					Strxml = Strxml & 	"	</parameters> "

					'## <!-- Dados do usuário --> "
					Strxml = Strxml & 	"		<userData> "
					Strxml = Strxml & 	"			<usrLogin>" & StrLogin & "</usrLogin> "
					Strxml = Strxml & 	"			<password>" & StrSenha  & "</password> "
					Strxml = Strxml & 	"			</userData> "
					Strxml = Strxml & 	"	</executeClass> "
					Strxml = Strxml & 	"</soap:Body> "
					Strxml = Strxml & 	"</soap:Envelope> "
					
					Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico )
			
								
					Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
					Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse
					Vetor_Campos(3)="adVarchar,15,adParamInput," & dblIdLogico
					Vetor_Campos(4)="adVarchar,20,adParamInput," & IdInterfaceAPG
					Vetor_Campos(5)="adVarchar,20,adParamInput, " & objRSSol("sol_id")
					Vetor_Campos(6)="adVarchar,20,adParamInput, 4 " 'Construir Return
					Vetor_Campos(7)="adVarchar,20,adParamInput, " & objRSDadosInterf("Processo")
					Vetor_Campos(8)="adVarchar,20,adParamInput, " & objRSDadosInterf("Acao") 
					strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",8,Vetor_Campos)

					Call db.Execute(strSqlRet)

			Else

				'Response.Write "<script language=javascript>alert('Não Encontrou Dados')</script>"

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & ID_Interface_APG	 'Identificação do APG
				Vetor_Campos(3)="addouble,10,adParamInput," & Acl_idacessologico 'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"

				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value

				''Tratar erro
				'If DBAction <> "1" Then

				'	Response.Write "Erro na Inclusão do LOG"

				'End If

			End If

	Else

			'Response.Write "<script language=javascript>alert('ID Logico não informado.')</script>"

	End If
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

		strxmlResp = 				"<resposta-cla> "
		strxmlResp = strxmlResp  &	"<codigo> " & objXmlDadosForm.parseError.errorCode & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem> Não foi possivel realizar parsing do XML enviado. Motivo:" & Trim(objXmlDadosForm.parseError.reason)
		strxmlResp = strxmlResp  & 	Trim(objXmlDadosForm.parseError.line) & "</mensagem>"
		strxmlResp = strxmlResp  &	"</resposta-cla> "

		Response.Write strxmlResp
		Response.end
	End IF

	processo = ""
	acao = ""
	idLogico = ""
	idTarefaApg = ""

	OE_Solicitacao_OEU = ""
	numeroSolicitacao = ""
	numeroRecurso = ""
	txtRecursos = ""

	'*' Claudio - Atribuir campos do XML as Variaveis.
	'*' objXmlDadosForm
	'Chama a Procedure passando os Parametros
	
	
	

	processo = objXmlDadosForm.selectSingleNode("/requisicao-cla/processo").text
	acao = objXmlDadosForm.selectSingleNode("/requisicao-cla/acao").text
	idLogico = objXmlDadosForm.selectSingleNode("/requisicao-cla/id-logico").text
	idTarefaApg = objXmlDadosForm.selectSingleNode("/requisicao-cla/id-tarefa-apg").text
	'numeroSolicitacao = objXmlDadosForm.selectSingleNode("/requisicao-cla/numero-solicitacao").text
	
	

	Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & idLogico )
	If Not objRSSol.Eof and Not objRSSol.Bof Then
		numeroSolicitacao = objRSSol("sol_id")
	End If
	
	
	
	If acao = "ATV" Or acao = "ALT" Then

		OE_Solicitacao_OEU = ""
		numeroSolicitacao = objXmlDadosForm.selectSingleNode("/requisicao-cla/numero-solicitacao").text
		numeroRecurso = objXmlDadosForm.selectSingleNode("/requisicao-cla/numero-recurso").text

		Set recursos = objXmlDadosForm.selectNodes("/requisicao-cla/codigos-recurso/recurso")

		txtRecursos = ""

		for i = 1 to recursos.length
			if i = 1 then
				txtRecursos = recursos.nextNode().text
			else
				txtRecursos = txtRecursos & ":" & recursos.nextNode().text
			end if
		next
	End If

	'Checa se serviço é 0800.
	Oe_numero = ""
	Oe_ano = ""
	Oe_item = ""
	Id_logico = idLogico
	Processo = Processo
	Acao = Acao
	
	
	call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,3,strxml)
	
	Vetor_Campos(1)="adWChar,20,adParamInput," & Processo
	Vetor_Campos(2)="adWChar,20,adParamInput,"& Acao
	Vetor_Campos(3)="adWChar,15,adParamInput,"& idLogico
	Vetor_Campos(4)="adWChar,20,adParamInput,"& idTarefaApg
	Vetor_Campos(5)="adWChar,30,adParamInput,"& OE_Solicitacao_OEU
	Vetor_Campos(6)="adWChar,10,adParamInput,"& numeroSolicitacao
	Vetor_Campos(7)="adWChar,20,adParamInput,"& numeroRecurso

	Vetor_Campos(8)="adWChar,255,adParamInput,"& txtRecursos
	Vetor_Campos(9)="adWChar,20,adParamInput,null"
	Vetor_Campos(10)="adWChar,30,adParamInput,null"
	Vetor_Campos(11)="adWChar,30,adParamInput,null"
	Vetor_Campos(12)="adWChar,30,adParamInput,null"
	Vetor_Campos(13)="adWChar,30,adParamInput,null"
	Vetor_Campos(14)="adWChar,10,adParamInput,null"
	Vetor_Campos(15)="adWChar,20,adParamInput,null"
	Vetor_Campos(16)="adWChar,30,adParamInput,null"
	Vetor_Campos(17)="adWChar,30,adParamInput,null"
	Vetor_Campos(18)="adInteger,2,adParamOutput,0"
	Vetor_Campos(19)="adWChar,100,adParamOutput,a "

	Call APENDA_PARAM("CLA_sp_ins_Construir_Acesso",19,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	DBDescricao = ObjCmd.Parameters("RET1").value

	'Response.Write "<script language=javascript>alert('Resultado: " & DBAction & DBDescricao & "');</script>"

	''Verifica se Retornou erro
	''De Gravação ou Validação dos dados
	If Trim(DBAction) <> "0" Then

		strxmlResp = 				"<resposta-cla> "
		strxmlResp = strxmlResp  & 	"<codigo> " & Trim(DBAction) & "</codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(DBDescricao) & "</mensagem>"
		strxmlResp = strxmlResp  &	"</resposta-cla> "

		Response.write strxmlResp


		Vetor_Campos(1)="adInteger,6,adParamInput, " & numeroSolicitacao	'Solicitação
		Vetor_Campos(2)="adInteger,6,adParamInput, " & OE_Solicitacao_OEU	'Identificação do APG
		Vetor_Campos(3)="addouble,10,adParamInput, " & idLogico				'ID Logico
		Vetor_Campos(4)="adWChar,255,adParamInput, " & strxmlResp			'Descrição do Erro
		Vetor_Campos(5)="adInteger,4,adParamOutput,0"

		Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value

		''Tratar erro
		'If DBAction <> "1" Then

			Response.Write "Erro na Inclusão do LOG"

		'End If

		Response.end
	Else

	    Set objRSMisto = db.Execute("select top 1 cla_acessofisico.acf_id,Tec_ID from cla_acessofisico inner join cla_acessologicofisico on cla_acessofisico.Acf_Id = cla_acessologicofisico.Acf_ID where Acl_IDAcessoLogico = "&idLogico&" and Acf_IDAcessoFisico is null")
         If Not objRSMisto.eof and  not objRSMisto.Bof Then
	       dblacf_id = objRSMisto("acf_id")
	       dblTec_ID = objRSMisto("Tec_ID")
	     End if
	
		Set objRSPed = db.Execute("CLA_Sp_Sel_PedidoSolicitacao " & numeroSolicitacao)

		If Not objRSPed.Eof and Not objRSPed.Bof Then
			IntPed_ID = objRSPed("Ped_Id")
		Else
			IntPed_ID = "0"
		End If
		
		Vetor_Campos(1)="adWChar,10,adParamInput,"& numeroSolicitacao
		Vetor_Campos(2)="adWChar,20,adParamInput, 3" ' 3 - Construir Send 
		Vetor_Campos(3)="adWChar,20,adParamInput," & trim(processo)
		Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(acao)
		Vetor_Campos(5)="adInteger,4,adParamInput,"& dblacf_id
		Vetor_Campos(6)="adInteger,4,adParamInput,"& dblTec_ID

		StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",6,Vetor_Campos)

		db.execute(strSQL)
		
		Vetor_Campos(1)="adWChar,50,adParamInput,null "
		Vetor_Campos(2)="adInteger,1,adParamInput,null "
		Vetor_Campos(3)="adWChar,50,adParamInput,null "
		Vetor_Campos(4)="adInteger,1,adParamInput,null" 
		Vetor_Campos(5)="adInteger,10,adParamInput," & numeroSolicitacao
	
		strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)

		Set objRSDadosInterf = db.Execute(strSql)

		If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

			Alteracao_Cadastral = objRSDadosInterf("flag_altcadastral")
			dblIdLogico         = objRSDadosInterf("id_logico")
			IdInterfaceAPG      = ""

		End If
		
		if Alteracao_Cadastral = "S" then
		
			x = EnviarRetornoContr_Apg(dblIdLogico, IdInterfaceAPG, numeroSolicitacao)
			
		end if 

		strxmlResp = "<resposta-cla> "
		strxmlResp = strxmlResp  & "<codigo>0</codigo>"
		strxmlResp = strxmlResp  & "<mensagem>Solicitacao Gravada com Sucesso. </mensagem>"
		strxmlResp = strxmlResp  & "</resposta-cla> "

		Response.Write strxmlResp '& "<br>"
	End If

Else

		strxmlResp =	"<resposta-cla> "
		strxmlResp = strxmlResp  & 	"<codigo> 9999 </codigo>"
		strxmlResp = strxmlResp  & 	"<mensagem>XML não Enviado ao CLA</mensagem>"
		strxmlResp = strxmlResp  &	"</resposta-cla> "

		strxml = "XML não enviado ao CLA"

		Vetor_Campos(1)="adInteger,6,adParamInput,0" 	'Solicitação
		Vetor_Campos(2)="adInteger,6,adParamInput,0"	'Identificação do APG
		Vetor_Campos(3)="addouble,10,adParamInput,0" 	'ID Logico
		Vetor_Campos(4)="adWChar,255,adParamInput, " & strxml 'Descrição do Erro
		Vetor_Campos(5)="adInteger,4,adParamOutput,0"

		Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value

		''Tratar erro
		'If DBAction <> "1" Then

		'	Response.Write "Erro na Inclusão do LOG"

		'End If

		Response.Write strxmlResp '& "<br>"
		Response.end
End If


'on error resume next

'if err.number <> 0 then
'		strxmlResp = 				"<codigo> " & Trim(err.number) & "</codigo>"
'		strxmlResp = strxmlResp  & 	"<Descricao>" & Trim(err.Description) & "</Descricao>"
'
'		Response.write strxmlResp
'		Response.end
'
'end if

'Adicionado PRSS 21/04/2007
if acao = "CAN" or acao = "DES" Then
  'Enviar retorno automático do construir:
  'x = EnviarRetornoConstr_Apg_CAN_Estocar (idLogico,idTarefaApg,numeroSolicitacao)
  'response.write x
  EnviarRetornoConstr_Apg_CAN_Estocar idLogico,idTarefaApg,numeroSolicitacao
End if
%>
