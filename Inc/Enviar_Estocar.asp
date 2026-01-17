
<%
Function EnviarEstocar_Apg(dblIdLogico,IdInterfaceAPG,dblSolid)

	'Dim dblIdLogico, IdInterfaceAPG, dblSolid

	'dblIdLogico = 6787123722
	'IdInterfaceAPG = 1
	'dblSolid = 233993


	On Error Resume Next
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


	%>
	<!--#include file="../inc/data.asp"-->
	<%
	
	strXmlEndereco = ""
	EncontrouDados = false

	StrClasse = "INTERFENTREGARRETURN" 'INTERF_ENTREGAR_RETURN

	dt_aceite_acesso = ""
	numero_acesso_provedor = ""
	dt_recebimento_recurso_acesso = ""

	dt_entrega_acesso = "06/11/2006"
	dt_desinstalacao = ""
	retorno = ""


	if dblIdLogico <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput,"
			Vetor_Campos(2)="adInteger,4,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

			Set objRSDadosCla = db.Execute(strSqlRet)

			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then

				Vetor_Campos(1)="adWChar,50,adParamInput,null "
				Vetor_Campos(2)="adInteger,1,adParamInput,null "
				Vetor_Campos(3)="adWChar,30,adParamInput," & IdInterfaceAPG
				Vetor_Campos(4)="adInteger,1,adParamInput,null "

				strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",4,Vetor_Campos)



				Set objRSDadosInterf = db.Execute(strSql)


				If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

					EncontrouDados = True
					EnviarRetornoEntregar_Apg = "encontrou da"
				End If

			End If




			If EncontrouDados = True Then


					dt_aceite_acesso = trim(objRSDadosCla("Acf_DtAceite"))
					numero_acesso_provedor = trim(objRSDadosCla("Acf_NroAcessoPtaEbt"))


					If trim(objRSDadosCla("Acf_Proprietario")) = "EBT" Then

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
								Vetor_Campos(3)="adWChar,15,adParamInput,"
								Vetor_Campos(4)="adInteger,4,adParamInput,"

								strSqlRet = APENDA_PARAMSTRSQL("cla_sp_sel_crmsprocessoAcesso",4,Vetor_Campos)
								Set objRSDadosEbt = db.Execute(strSqlRet)

								If not objRSDadosEbt.oef and objRSDadosEbt.bof Then

									dt_recebimento_recurso_acesso = objRSDadosEbt("Entregaequipamento")

								End If

						End if

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
					Strxml = Strxml & 	"		<className>INTERFENTREGARRETURN</className> "
					Strxml = Strxml & 	"		<!-- Parametros configurados na classe --> "
					Strxml = Strxml & 	"		<parameters> "

					Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & Trim(objRSDadosInterf("Processo")) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & Trim(objRSDadosInterf("Acao")) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & Trim(dblIdLogico)  & "</parameter> "
					Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "



					Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & Trim(objRSDadosInterf("ID_Tarefa_Apg")) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" & ">" & Trim(objRSDadosInterf("Solicitacao")) & "</parameter> "

				 	Strxml = Strxml & 	"		<parameter name=" & """dataEntregaAcesso""" &">" & Trim(dt_entrega_acesso) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataRecebimentoRecursoAcesso""" &">" & Trim(dt_recebimento_recurso_acesso) & "</parameter> "
				 	Strxml = Strxml & 	"		<parameter name=" & """dataAceiteAcesso""" &">" & Trim(dt_aceite_acesso) & "</parameter> "
	 				Strxml = Strxml & 	"		<parameter name=" & """numeroAcessoProvedor""" &">" & Trim(numero_acesso_provedor) &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """dataDesinstalacao""" &">" & Trim(dt_desinstalacao) & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & Trim(retorno) & "</parameter> "
					Strxml = Strxml & 	"	</parameters> "

					Strxml = Strxml & 	"		<!-- Dados do usuario --> "
					Strxml = Strxml & 	"		<userData> "
					Strxml = Strxml & 	"		<!-- Usuario Apia executante --> "
					Strxml = Strxml & 	"			<usrLogin>admin_ebt</usrLogin> "
					Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha ja encriptada) -->"
					Strxml = Strxml & 	"			<password>51-482-15-9939-4938-50126-87122160115100</password> "
					Strxml = Strxml & 	"			</userData>"
					Strxml = Strxml & 	"	</executeClass> "
					Strxml = Strxml & 	"</soap:Body> "
					Strxml = Strxml & 	"</soap:Envelope> "


					Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

					xmlhttp.Open "POST", AdresserPath, StrLogin, StrSenha
					xmlhttp.setRequestHeader "SOAPAction", "executeClass"

					xmlhttp.send(Strxml)
					strRetorno = xmlhttp.ResponseText


					Set doc = server.CreateObject("Microsoft.XMLDOM")
					Set doc1 = server.CreateObject("Microsoft.XMLDOM")
					doc.async= False
					doc1.async= False

					doc.loadXML(strRetorno)

					doc1.loadXML(Strxml)

					Set xmlhttp= Nothing
					'Set doc = Nothing
					'Set doc1= Nothing

					If doc.parseError<>0 Then


						strxmlResp = strxmlResp  &  "<resposta-cla><codigo> Erro ao realizar Parsing do XML retornado(APG=>CLA - Entregar Return. Codigo: " & doc.parseError.errorCode & "</codigo>"
						strxmlResp = strxmlResp  & 	"<mensagem>Motivo" & strErroXml & Trim(doc.parseError.reason)
						strxmlResp = strxmlResp  & 	Trim(doc.parseError.line) & "</mensagem></resposta-cla>"


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


					Else

						Set objNodeList = Doc.selectNodes("//soapenv:Envelope/soapenv:Body/executeClassResponse/ns1:executeClassReturn/ns3:parameters/ns3:parameter")


						if trim(objNodeList.Length) = "0" then


							strxmlResp = "Formato do XML retornado pelo APG não Identificado. Não foi possivel identificar resposta."

							Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
							Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
							Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
							Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
							Vetor_Campos(5)="adInteger,4,adParamOutput,0"

							Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value

							'Response.Write "Resultado da Inclusão do LOG:" & DBAction
							''Tratar erro
							'If trim(DBAction) <> "1" Then

							'	Response.Write "Erro na Inclusão do LOG:" & DBAction

							'End If

							EnviarRetornoEntregar_Apg = "1 - Formato do XML retornado pelo APG não Identificado. Não foi possivel identificar resposta."

							'EnviarRetornoEntregar_Apg = "1 - Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."


						Else


							'Obtem retornos enviados pelo APG
							Tamanho = objNodeList.Length
							posCodErro = tamanho - 2
							posDescErro = tamanho - 1

							codErro = objNodeList.Item(posCodErro).Text
							DescErro = objNodeList.Item(posDescErro).Text
							strxmlResp = ""

							If Trim(codErro) <> "" and Trim(codErro) <> "0"  Then

								strxmlResp = "O Seguinte erro foi retornado pela Interface CLA => APG - Ação: Entregar_Acesso_Return" & codErro & DescErro

								Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
								Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
								Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
								Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
								Vetor_Campos(5)="adInteger,4,adParamOutput,0"

								Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

								ObjCmd.Execute'pega dbaction
								DBAction = ObjCmd.Parameters("RET").value

								''Tratar erro
								'If DBAction <> "1" Then

								'	Response.Write "Erro na Inclusão do LOG"

								'End If

								EnviarRetornoEntregar_Apg = "2 - Erro ao realizar Parsing da Resposta CLA => APG - Ação: Entregar_Acesso_Return" & codErro & DescErro
								'EnviarRetornoEntregar_Apg = "2 - Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."


							Else


								EnviarRetornoEntregar_Apg = "0 - Interface com APG realizada com Sucesso. "


								'Vetor_Campos(1)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
								'Vetor_Campos(2)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
								'Vetor_Campos(3)="adWChar,30,adParamInput, " & OE_Solicitacao_OEU
								'Vetor_Campos(4)="adWChar,20,adParamInput, " & idTarefaApg
								'Vetor_Campos(5)="adWChar,15,adParamInput,"& dblIdLogico
								'Vetor_Campos(6)="adWChar,10,adParamInput,"& dblSolid
								'Vetor_Campos(7)="adWChar,20,adParamInput, "
								'Vetor_Campos(8)="adInteger,4,adParamOutput,0"
								'Vetor_Campos(9)="adWChar,100,adParamOutput,"

								'Response.Write "<script language=javascript>alert('Define par. Procedures')</script>"

								'd = APENDA_PARAMstr("CLA_sp_ins_Construir_Acesso_Ret",9,Vetor_Campos)

								'Call APENDA_PARAM("CLA_sp_ins_Construir_Acesso_Ret",9,Vetor_Campos)

								'Response.Write "<script language=javascript>alert('chamou Define par. Procedures')</script>"
								'Response.Write "<script language=javascript>alert('Execucao:" & d &"')</script>"

								'ObjCmd.Execute'pega dbaction
								'DBAction = ObjCmd.Parameters("RET").value
								'DBDescricao = ObjCmd.Parameters("RET1").value

								'Response.Write "<script language=javascript>alert('Grava ret. interface')</script>"
								'strxmlResp = "Erro ao atualizar log da Interface - Codigo:" & DBAction & "Descrição:" & DBDescricao
								'Response.Write "<script language=javascript>alert('" & strxmlResp &"')</script>"


								'Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
								'Vetor_Campos(2)="adInteger,6,adParamInput," & idTarefaApg		 'Identificação do APG
								'Vetor_Campos(3)="addouble,10,adParamInput," & idLogico			 'ID Logico
								'Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
								'Vetor_Campos(5)="adInteger,4,adParamOutput,0"

								'Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

								'ObjCmd.Execute'pega dbaction
								'DBAction = ObjCmd.Parameters("RET").value

								'Response.Write "<script language=javascript>alert('" & DBAction &"')</script>"

								''Tratar erro
								'If DBAction <> "1" Then

								'	Response.Write "Erro na Inclusão do LOG"

								'End If

							End if


						End if

				End if
				'Response.write(Strxml)
				'Response.end


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
%>

