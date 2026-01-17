<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->

<%
'	- Sistema			: CLA
'	- Arquivo			:Entregar_Acesso_APG_CLA.asp
'	- Descrição			: Recebe chamadas de Sistemas de Aprovisionamento Passando arquivo Xml,
'						Valida e Disponibiliza os Dados na Camada de Interface do CLA
'						Gera arquivo Xml de retorno contendo Codigo e Descricao do retorno

'Set db = server.createobject("ADODB.Connection")
'db.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
'db.ConnectionTimeout = 0
'db.CommandTimeout = 0
'db.open

strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarRetornoEntregar_Can_Des_Apg(dblIdLogico, IdInterfaceAPG, dblSolid, Interf_id)

	'Dim dblIdLogico, IdInterfaceAPG, dblSolid

	'dblIdLogico = 6787123722
	'IdInterfaceAPG = 1
	'dblSolid = 233993

	'On Error Resume Next
	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim strXmlEndereco
	Dim EncontrouDados
	Dim DescErro
	Dim CodErro
	Dim retorno

	%>
	<%'<!--#include file="../inc/data.asp"--> - removido PRSS 21/03/2007 - Função dentro de função.%>
	<%
	
	strXmlEndereco = ""
	EncontrouDados = false

	

	StrProcesso = ""
	StrAcao = ""
	StrIDLogico = ""
	StrIDTarefaApg = ""

	numeroSolicitacao = ""

	dt_entrega_acesso = ""
	dataRecebimentoRecursoAcesso = ""
	dataAceiteAcesso = ""
	numeroAcessoProvedor = ""
	dt_desinstalacao = Date()
	retorno = "OK"


	If dblIdLogico <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput,"
			Vetor_Campos(2)="adInteger,4,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

			Set objRSDadosCla = db.Execute(strSqlRet)

			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then

				Vetor_Campos(1)="adWChar,50,adParamInput,null "
				Vetor_Campos(2)="adInteger,1,adParamInput,null "
				Vetor_Campos(3)="adWChar,30,adParamInput, " & IdInterfaceAPG
				Vetor_Campos(4)="adInteger,1,adParamInput,null" '&IdInterfaceAPG

				strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",4,Vetor_Campos)

				Set objRSDadosInterf = db.Execute(strSql)

				If Not objRSDadosInterf.eof and  Not objRSDadosInterf.Bof Then

					EncontrouDados = True

					StrProcesso = Trim(objRSDadosInterf("Processo"))
					StrAcao = Trim(objRSDadosInterf("Acao"))
					StrIDLogico = Trim(dblIdLogico)
					StrIDTarefaApg = Trim(objRSDadosInterf("ID_Tarefa_Apg"))

				End If

			End If

			If Trim(StrAcao) = "DES" Then
				StrClasse = "INTERFDESINSTALARRETURN" 'INTERF_ENTREGAR_RETURN
			ELSE 
				StrClasse = "INTERFENTREGARRETURN" 'INTERF_ENTREGAR_RETURN
			END IF 

			If EncontrouDados = True Then

					Strxml			=   "<soap:Envelope "
					Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance"""
					Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema"""
					Strxml = Strxml &   " xmlns:soap="&"""http://schemas.xmlsoap.org/soap/envelope/"""
					Strxml = Strxml & 	"> <soap:Body> "

					Strxml = Strxml & 	"	<!-- Define a operacao sendo realizada (executar classe) --> "
					Strxml = Strxml & 	"	<executeClass> "
					Strxml = Strxml & 	"		<!-- Ambiente do Apia a ser chamado --> "
					Strxml = Strxml & 	"		<envName>APG</envName> "

					Strxml = Strxml & 	"		<!-- Nome da classe de negocio, tal como configurada no Apia --> "
					Strxml = Strxml & 	"		<className>"& StrClasse &"</className> "
					Strxml = Strxml & 	"		<!-- Parametros configurados na classe --> "
					Strxml = Strxml & 	"		<parameters> "

					Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & StrProcesso & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & StrAcao & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & mid(trim(StrIDLogico),1,2) & "8" & mid(trim(StrIDLogico),4,10)  & "</parameter> "
					Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "

					Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & StrIDTarefaApg & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" & ">" & "</parameter> "

				 	Strxml = Strxml & 	"		<parameter name=" & """dataEntregaAcesso""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataRecebimentoRecursoAcesso""" &">" & "</parameter> "
				 	Strxml = Strxml & 	"		<parameter name=" & """dataAceiteAcesso""" &">" & "</parameter> "
	 				Strxml = Strxml & 	"		<parameter name=" & """numeroAcessoProvedor""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataDesinstalacao""" &">" & Trim(dt_desinstalacao) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & Trim(retorno) & "</parameter> "
					Strxml = Strxml & 	"	</parameters> "

					Strxml = Strxml & 	"		<!-- Dados do usuario --> "
					Strxml = Strxml & 	"		<userData> "
					Strxml = Strxml & 	"		<!-- Usuario Apia executante --> "
					Strxml = Strxml & 	"			<usrLogin>"&StrLogin&"</usrLogin> "
					Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha ja encriptada) -->"
					Strxml = Strxml & 	"			<password>"&StrSenha&"</password> "
					Strxml = Strxml & 	"			</userData>"
					Strxml = Strxml & 	"	</executeClass> "
					Strxml = Strxml & 	"</soap:Body> "
					Strxml = Strxml & 	"</soap:Envelope> "
					
					Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico )
			
					Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
					Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse
					Vetor_Campos(3)="adVarchar,15,adParamInput," & StrIDLogico
					Vetor_Campos(4)="adVarchar,20,adParamInput," & StrIDTarefaApg
					Vetor_Campos(5)="adVarchar,20,adParamInput, " & objRSSol("sol_id")
					Vetor_Campos(6)="adVarchar,20,adParamInput, 6 " 'Entregar Return 
					Vetor_Campos(7)="adVarchar,20,adParamInput, " & StrProcesso
					Vetor_Campos(8)="adVarchar,20,adParamInput, " & StrAcao 
					strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",8,Vetor_Campos)
					
					Call db.Execute(strSqlRet)


					'Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

					'xmlhttp.Open "POST", AdresserPath, StrLogin, StrSenha
					'xmlhttp.setRequestHeader "SOAPAction", "executeClass"

					'xmlhttp.send(Strxml)
					'strRetorno = xmlhttp.ResponseText


					'Set doc = server.CreateObject("Microsoft.XMLDOM")
					'Set doc1 = server.CreateObject("Microsoft.XMLDOM")
					'doc.async= False
					'doc1.async= False

					'Set xmlhttp= Nothing
					'Set doc = Nothing
					'Set doc1= Nothing

'					Response.Write "<B>readyState=</B> "&doc.readyState &"<br>"

			Else

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid				'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & ID_Interface_APG		'Identificação do APG
				Vetor_Campos(3)="addouble,10,adParamInput," & Acl_idacessologico	'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp			'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"

				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value

				''Tratar erro
				'If DBAction <> "1" Then

				'	Response.Write "Erro na Inclusão do LOG"

				'End If


			End If

	End If

End Function


Function EnviarRetornoEntregar_Apg( dblIdLogico, IdInterfaceAPG, dblSolid, Interf_id)
	'response.write "<b>dblIdLogico: </b>" & dblIdLogico & "<br>"
	'response.write "<b>IdInterfaceAPG: </b>" & IdInterfaceAPG & "<br>"
	'response.write "<b>dblSolid: </b>" & dblSolid & "<br>"

	'Dim dblIdLogico, IdInterfaceAPG, dblSolid

	'dblIdLogico = 6787123862
	'IdInterfaceAPG = 1257
	'dblSolid = 234142

	%><!--#include file="../inc/conexao_apg.asp"--><%
	StrClasse = "INTERFENTREGARRETURN" 'INTERF_ENTREGAR_RETURN
	
	'On Error Resume Next
	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim strXmlEndereco
	Dim EncontrouDados
	Dim StrTipoAcessso
	Dim strTecnologia
	Dim strVelAcessoFis
	Dim DescErro
	Dim CodErro

	Dim dt_entrega_acesso, dt_recebimento_recurso_acesso, dt_aceite_acesso
	Dim numero_acesso_provedor, dt_desinstalacao, retorno

	strXmlEndereco = ""
	EncontrouDados = false

	'response.write "<b>AdresserPath: </b>" & AdresserPath & "<br>"
	'response.write "<b>StrLogin: </b>" & StrLogin & "<br>"
	'response.write "<b>StrSenha: </b>" & StrSenha & "<br>"
	'response.write "<b>StrClasse: </b>" & StrClasse & "<br>"
	
	dt_aceite_acesso = ""
	numero_acesso_provedor = ""
	dt_recebimento_recurso_acesso = "" 

	dt_entrega_acesso = date()
	dt_desinstalacao = "" 
	retorno = ""

	if dblIdLogico <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput,"
			Vetor_Campos(2)="adInteger,4,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

			Set objRSDadosCla = db.Execute(strSqlRet)

			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then

				'PRSS 23/03/2007 - Corrigido o número de parâmetros passados de 4 para 5.
				Vetor_Campos(1)="adWChar,30,adParamInput,null " 'Usuario
				Vetor_Campos(2)="adInteger,1,adParamInput,null " 'Tipo
				Vetor_Campos(3)="adWChar,30,adParamInput,null " 'Tarefa
				Vetor_Campos(4)="adInteger,1,adParamInput,null " 'Sol_Acesso_ID
				Vetor_Campos(5)="adInteger,1,adParamInput, " & dblSolid 'Sol_id
				
				strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)

				Set objRSDadosInterf = db.Execute(strSql)

				If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

					EncontrouDados = True
					
					Alteracao_Cadastral = objRSDadosInterf("flag_altcadastral")
					
				End If
			End If

			If EncontrouDados = True Then
					
					Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico )
			
					dblSolid = objRSSol("sol_id")
					
					Vetor_Campos(1)="adWChar,10,adParamInput,"& dblSolid
					
					if interf_id = 6 then
					  Vetor_Campos(2)="adWChar,20,adParamInput, 6 " '6 - Entregar Return
					else
					  Vetor_Campos(2)="adWChar,20,adParamInput, 5 " '5 - Entregar Send
					end if
					Vetor_Campos(3)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
					Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))

					StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",4,Vetor_Campos)

					db.execute(strSQL)

					dt_aceite_acesso = trim(objRSDadosCla("Acf_DtAceite"))

					If trim(objRSDadosCla("Acf_Proprietario")) = "EBT" Then
					
						numero_acesso_provedor = ""
					
						''@@ Davif Incluir validação da Tecnologia para atribuir valores.
						if objRSDadosCla("Tec_Id") <> "" then
						Set objRSAux = db.Execute("CLA_Sp_Sel_Tecnologia " & objRSDadosCla("Tec_Id"))
							if not objRSAux.Eof and Not objRSAux.Bof then
								strTecnologia	= objRSAux("Tec_Sigla")
							End if
						End if

						If Trim(strTecnologia) = "RADIO" Or Trim(strTecnologia) = "FIBRA" Or Trim(strTecnologia) = "SATELITE" Then

								'@@ Davif - Dados do Acesso Embratel
								Vetor_Campos(1)="adInteger,4,adParamInput," & dblIdLogico
								Vetor_Campos(2)="adWChar,25,adParamInput," & dblSolid
								
								strSqlRet = APENDA_PARAMSTRSQL("cla_sp_sel_crmsprocessoAcesso",2,Vetor_Campos)
								Set objRSDadosEbt = db.Execute(strSqlRet)

								If not objRSDadosEbt.eof and objRSDadosEbt.bof Then

									dt_recebimento_recurso_acesso = objRSDadosEbt("Entregaequipamento")

								End If

						End if
					Else 
					
						numero_acesso_provedor = trim(objRSDadosCla("Acf_NroAcessoPtaEbt"))
					
					End If
					
					Strxml			=   "<soap:Envelope "
					Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance"""
					Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema"""
					Strxml = Strxml &   " xmlns:soap="&"""http://schemas.xmlsoap.org/soap/envelope/"""
					Strxml = Strxml & 	"> <soap:Body> "

					Strxml = Strxml & 	"	<!-- Define a operacao sendo realizada (executar classe) --> "
					Strxml = Strxml & 	"	<executeClass> "
					Strxml = Strxml & 	"		<!-- Ambiente do Apia a ser chamado --> "
					Strxml = Strxml & 	"		<envName>APG</envName> "

					Strxml = Strxml & 	"		<!-- Nome da classe de negocio, tal como configurada no Apia --> "
					Strxml = Strxml & 	"		<className>"& StrClasse &"</className> "
					Strxml = Strxml & 	"		<!-- Parametros configurados na classe --> "
					Strxml = Strxml & 	"		<parameters> "

					Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & objRSDadosInterf("Processo") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & objRSDadosInterf("Acao") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & mid(trim(dblIdLogico),1,2) & "8" & mid(trim(dblIdLogico),4,10)  & "</parameter> "
					Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "

					Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & objRSDadosInterf("ID_Tarefa_Apg") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" & ">" & objRSDadosInterf("Solicitacao") & "</parameter> "

				 	Strxml = Strxml & 	"		<parameter name=" & """dataEntregaAcesso""" &">" & dt_entrega_acesso & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataRecebimentoRecursoAcesso""" &">" & dt_recebimento_recurso_acesso & "</parameter> "
				 	Strxml = Strxml & 	"		<parameter name=" & """dataAceiteAcesso""" &">" & dt_aceite_acesso & "</parameter> "
	 				Strxml = Strxml & 	"		<parameter name=" & """numeroAcessoProvedor""" &">" & numero_acesso_provedor &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataDesinstalacao""" &">" & dt_desinstalacao & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & retorno & "</parameter> "
					Strxml = Strxml & 	"	</parameters> "

					Strxml = Strxml & 	"		<!-- Dados do usuario --> "
					Strxml = Strxml & 	"		<userData> "
					Strxml = Strxml & 	"		<!-- Usuario Apia executante --> "
					Strxml = Strxml & 	"			<usrLogin>"&StrLogin&"</usrLogin> "
					Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha ja encriptada) -->"
					Strxml = Strxml & 	"			<password>"&StrSenha&"</password> "
					Strxml = Strxml & 	"			</userData> "
					Strxml = Strxml & 	"	</executeClass> "
					Strxml = Strxml & 	"</soap:Body> "
					Strxml = Strxml & 	"</soap:Envelope> "
					
					'Removido T3ENA 29/05/2007
					'dblIdLogico, IdInterfaceAPG, dblSolid
					'response.write "<script>alert('"&StrClasse&"')</script>"
					'Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
					'Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse 
					'Vetor_Campos(3)="adVarchar,15,adParamInput," & dblIdLogico
					'Vetor_Campos(4)="adVarchar,20,adParamInput," & objRSDadosInterf("ID_Tarefa_Apg")
					'Vetor_Campos(5)="adVarchar,30,adParamInput," & objRSDadosInterf("Solicitacao")
					'strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",5,Vetor_Campos)

					'Call db.Execute(strSqlRet)

					if Alteracao_Cadastral = "S" then 	
												
						Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
						Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse
						Vetor_Campos(3)="adVarchar,15,adParamInput," & dblIdLogico
						Vetor_Campos(4)="adVarchar,20,adParamInput," & objRSDadosInterf("ID_Tarefa_Apg")
						Vetor_Campos(5)="adVarchar,20,adParamInput, " & dblSolid
						Vetor_Campos(6)="adVarchar,20,adParamInput, 6 " 'Entregar Return 
						Vetor_Campos(7)="adVarchar,20,adParamInput, " & objRSDadosInterf("Processo")
						Vetor_Campos(8)="adVarchar,20,adParamInput, " & objRSDadosInterf("Acao")
						strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",8,Vetor_Campos)
	
						Call db.Execute(strSqlRet)
					
					else
					
						Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	
						xmlhttp.Open "POST", AdresserPath, StrLogin, StrSenha
						xmlhttp.setRequestHeader "SOAPAction", "executeClass"
	
						xmlhttp.send(Strxml)
						strRetorno = xmlhttp.ResponseText
	
						doc.loadXML(strRetorno)
						doc1.loadXML(Strxml)
	
	
						Set doc = server.CreateObject("Microsoft.XMLDOM")
						Set doc1 = server.CreateObject("Microsoft.XMLDOM")
						doc.async= False
						doc1.async= False
						
						'Checa se serviço é 0800.
						Oe_numero = objRSDadosInterf("Oe_numero")
						Oe_ano = objRSDadosInterf("Oe_ano")
						Oe_item = objRSDadosInterf("Oe_item")
						Id_logico = dblIdLogico
						Processo = trim(objRSDadosInterf("Processo"))
						Acao = trim(objRSDadosInterf("Acao"))
						
						call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,6,strxml)
						
						call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,6,strRetorno)
	
						Set xmlhttp= Nothing
						Set doc = Nothing
						Set doc1= Nothing
						
					end if

'					Response.Write "<B>readyState=</B> "&doc.readyState &"<br>"

			Else

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid				'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & ID_Interface_APG		'Identificação do APG
				Vetor_Campos(3)="addouble,10,adParamInput," & Acl_idacessologico	'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp			'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"


				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value

				''Tratar erro
				'If DBAction <> "1" Then

				'	Response.Write "Erro na Inclusão do LOG"

				'End If

			End If

	End If

	EnviarRetornoEntregar_Apg = Strxml
End Function

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
	ReceberEntregar_Apg()
end if


function ReceberEntregar_Apg()

			strxml = trim(Request.Form("hdnxml"))
		
			strxml = Replace(strxml,Chr(13),"")
			strxml = Replace(strxml,Chr(10),"")
		
		
			objXmlDadosForm.Async = "false"
		
			objXmlDadosForm.preserveWhiteSpace = True
		
			objXmlDadosForm.loadXML(strxml)
		
			'Response.Write "<B>readyState=</B> " & objXmlDadosForm.readyState &"<br>"
		
		
			If 	Len(Trim(objXmlDadosForm.parseError.errorCode)) > 1 Then
		
					strxmlResp = 				"<resposta-cla><codigo> " & objXmlDadosForm.parseError.errorCode & "</codigo>"
					strxmlResp = strxmlResp  & 	"<mensagem>Erro ao Realizar Parsing do XML. Razao:" & Trim(objXmlDadosForm.parseError.reason)
					strxmlResp = strxmlResp  & 	Trim(objXmlDadosForm.parseError.line) & "</mensagem></resposta-cla>"
		
					strxmlResp = "Erro no XML enviado pelo APG (Entregar Acesso Send) não foi possivel realizar Parsing. "
					strxmlResp = strxmlResp  &  "<codigo> " & doc.parseError.errorCode & "</codigo>"
					strxmlResp = strxmlResp  & 	"<mensagem>" & strErroXml & Trim(doc.parseError.reason)
					strxmlResp = strxmlResp  & 	Trim(doc.parseError.line) & "</mensagem>"
		
					Vetor_Campos(1)="adInteger,6,adParamInput,"					'Solicitação
					Vetor_Campos(2)="adInteger,6,adParamInput,"					'Identificação do APG
					Vetor_Campos(3)="addouble,10,adParamInput,"					'ID Logico
					Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp	'Descrição do Erro
					Vetor_Campos(5)="adInteger,4,adParamOutput,0"
		
					Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)
		
					ObjCmd.Execute'pega dbaction
					DBAction = ObjCmd.Parameters("RET").value
		
		
					Response.end
		
		
			End IF
		
		
			'*' Claudio - Atribuir campos do XML as Variaveis.
			'*' objXmlDadosForm
		
			processo = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/processo").text)
			acao = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acao").text)
			idLogico = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/id-logico").text)
			idTarefaApg = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/id-tarefa-apg").text)
			
			if processo = "ATV" and acao = "ALT" then
	  		  StrSQL = "select Flag_AltCadastral from cla_apg_solicita_acesso where id_tarefa_apg = '" & idTarefaApg & "'"
			  Set objRSCadastral = db.Execute(strSql)
			  Flag_AltCadastral = objRSCadastral("Flag_AltCadastral")
			  
			  if isnull(Flag_AltCadastral) then
			    idLogico = mid(idLogico,1,2) & "7" & mid(idLogico,4,10)
			  end if
			end if
			
			Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & idLogico )
			numeroSolicitacao = objRSSol("sol_id")
			
			'Checa se serviço é 0800.
			Vetor_Campos(1)="adWChar,50,adParamInput,null "
			Vetor_Campos(2)="adInteger,1,adParamInput,null "
			Vetor_Campos(3)="adWChar,50,adParamInput,null "
			Vetor_Campos(4)="adInteger,1,adParamInput,null "
			Vetor_Campos(5)="adInteger,1,adParamInput, " & numeroSolicitacao
    		
			strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
    		
			Set objRSChecaServ = db.Execute(strSql)
		    
			If Not objRSChecaServ.eof and  not objRSChecaServ.Bof Then
			  Oe_numero = objRSChecaServ("Oe_numero")
			  Oe_ano = objRSChecaServ("Oe_ano")
			  Oe_item = objRSChecaServ("Oe_item")
			Else
			Oe_numero = ""
			Oe_ano = ""
			Oe_item = ""
			End If
			
			Id_logico = idLogico
			Processo = processo
			Acao = acao
			
			call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,5,strxml)
			
			txtNotas = ""
			transportadora = ""
			dataEntrega = ""
			
			If Trim(acao) = "ALT" or Trim(acao) = "ATV" Then
				'numeroSolicitacao = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/numero-solicitacao").text)
				OE_Solicitacao_OEU = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/id-tarefa-apg").text)
				
				txtNotas = ""
				transportadora = ""
				dataEntrega = ""
				
				if objXmlDadosForm.selectSingleNode("/requisicao-cla/acesso-ebt").text <> "" then
					Set notasFiscais = objXmlDadosForm.selectNodes("/requisicao-cla/acesso-ebt/notas-fiscais/nota")
					
					for i = 1 to notasFiscais.length
						if i = 1 then
							txtNotas = trim(notasFiscais.nextNode().text)
						else
							txtNotas = txtNotas & ":" & trim(notasFiscais.nextNode().text)
						end if
					next
					
					transportadora = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acesso-ebt/transportadora").text)
					dataEntrega = trim(objXmlDadosForm.selectSingleNode("/requisicao-cla/acesso-ebt/data-entrega").text)
				end if
				
			End If
			
			'response.write "Processo: " & processo & "<br>"
			'response.write "Ação: " & acao & "<br>"
			'response.write "ID Logico: " & idLogico & "<br>"
			'response.write "NFs: " & txtNotas & "<br>"
			'response.write "Transportadora: " & transportadora & "<br>"
			'response.write "Data de Entrega: " & dataEntrega & "<br>"
			'response.write "numeroSolicitacao: " & numeroSolicitacao & "<br>"
			'response.write chr(13)
		
			' Davi, alguns campos eu não tenho no XML...
			'Chama a Procedure passando os Parametros
		
			Vetor_Campos(1)="adWChar,20,adParamInput," & Processo
			Vetor_Campos(2)="adWChar,20,adParamInput,"& Acao
			Vetor_Campos(3)="adWChar,15,adParamInput,"& OE_Solicitacao_OEU
			Vetor_Campos(4)="adWChar,20,adParamInput,"& idTarefaApg
			Vetor_Campos(5)="adWChar,11,adParamInput,"& idLogico
			Vetor_Campos(6)="adWChar,10,adParamInput,"& numeroSolicitacao
		
			Vetor_Campos(7)="adWChar,20,adParamInput,"& txtNotas
			Vetor_Campos(8)="adWChar,20,adParamInput,null"
			Vetor_Campos(9)="adWChar,20,adParamInput,null"
			Vetor_Campos(10)="adWChar,20,adParamInput,null"
		
			Vetor_Campos(11)="adWChar,60,adParamInput,null"
			Vetor_Campos(12)="adWChar,10,adParamInput,null"
			Vetor_Campos(13)="adInteger,2,adParamOutput,0"
			Vetor_Campos(14)="adWChar,100,adParamOutput,a "
		
			Call APENDA_PARAM("CLA_sp_ins_Entregar_Acesso",14,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			DBDescricao = ObjCmd.Parameters("RET1").value
		
			''Verifica se Retornou erro
			If DBAction <> 0 Then
			
				strxmlResp = 				"<resposta-cla><codigo>" & Trim(DBAction) & "</codigo>"
				strxmlResp = strxmlResp  & 	"<mensagem>Teste Ajuste:" & Trim(DBDescricao) & "</mensagem></resposta-cla>"
				
			    Vetor_Campos(1)="adInteger,6,adParamInput," & numeroSolicitacao		'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & idTarefaApg			'Identificação do APG
				Vetor_Campos(3)="addouble,10,adParamInput," & idLogico				'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp			'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"
		
				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)
		
				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value
		
				Response.end
		
			Else
			
				Vetor_Campos(1)="adWChar,10,adParamInput,"& numeroSolicitacao
				Vetor_Campos(2)="adWChar,20,adParamInput, 5 " '5 - Entregar Send
				Vetor_Campos(3)="adWChar,20,adParamInput," & trim(Processo)
				Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(Acao)
					
				StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",4,Vetor_Campos)
		
				db.execute(strSQL)
		
				strxmlResp = 				"<resposta-cla><codigo>0</codigo>"
				strxmlResp = strxmlResp  & 	"<mensagem>Solicitacao Gravada com Sucesso. </mensagem> </resposta-cla>"
		
				'If Trim(acao) = "ALT" or Trim(acao) = "ATV" Then
					'Response.write "idLogico: " & idLogico & "<br>"
					'Response.write "idTarefaApg: " & idTarefaApg & "<br>"
					'Response.write "numeroSolicitacao: " & numeroSolicitacao & "<br><br>"
					
					'x = EnviarRetornoEntregar_Apg (idLogico, idTarefaApg, numeroSolicitacao, Interf_id)
		
				'End If
				
				If Trim(acao) = "CAN" or Trim(acao) = "DES" Then
				'Se o processo for terceiro, enviar automaticamente o RETORNO do ENTREGAR.
		            'Response.write "idLogico: " & idLogico & "<br>"
					'Response.write "idTarefaApg: " & idTarefaApg & "<br>"
					'Response.write "numeroSolicitacao: " & numeroSolicitacao & "<br><br>"
		
					'EnviarRetornoEntregar_Can_Des_Apg idLogico, idTarefaApg, numeroSolicitacao
					x = EnviarRetornoEntregar_Can_Des_Apg(idLogico, idTarefaApg, numeroSolicitacao, Interf_id)
				End If
				
				If Trim(processo)="ATV" and Trim(acao) = "ALT" then
				
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
						
							x = EnviarRetornoEntregar_Apg(idLogico, idTarefaApg, numeroSolicitacao, Interf_id)
						end if 			
				end if 
		
		
				'Response.Write "Retorno Envia Entregar: " & x & "<b><font color=blue>"
		
		
				'strxmlResp = 				"<resposta-cla><codigo>" & trim(Mid(x,1,1)) & "</codigo>"
				'strxmlResp = strxmlResp  & 	"<mensagem>" & trim(Mid(x+space(100),4,100))  & "</mensagem></resposta-cla>"
		
		
			End If
		
		'on error resume next
		
		if err.number <> 0 then
		
				strxmlResp = 				"<resposta-cla><codigo> " & Trim(CStr(err.number)) & "</codigo>"
				strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(err.Description) & "</mensagem></resposta-cla>"
		
		end if
		
end function


%>