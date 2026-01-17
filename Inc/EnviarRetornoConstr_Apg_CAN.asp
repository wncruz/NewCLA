<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarRetornoConstr_Apg_CAN(dblIdLogico, IdInterfaceAPG, dblSolid)

	'Dim dblIdLogico, IdInterfaceAPG, dblSolid

	'	Response.Write "<script language=javascript>alert('" & "Logico:" & dblIdLogico & "');</script>"
	'	Response.Write "<script language=javascript>alert('" & "IdInterfaceAPG:" & IdInterfaceAPG & "');</script>"
	'	Response.Write "<script language=javascript>alert('" & "Solic." & dblSolid & "');</script>"

	'	Response.Write "Logico:" & dblIdLogico '= 6787123691
	'	Response.Write "IdInterfaceAPG:" & IdInterfaceAPG
	'	Response.Write "Solic." & dblSolid 'dblSolid = 233962

	On Error Resume Next
	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim StrLogin
	Dim StrSenha
	Dim StrClasse

	Dim EncontrouDados
	Dim DescErro
	Dim CodErro

	Dim processo, acao, idLogico
	Dim idTarefaApg, emiteOts2m 
	Dim retirarRecurso, compartilhaAcesso
	

	processo = ""
	acao = ""
	idLogico = ""
	idTarefaApg = ""
	emiteOts2m = ""
	retirarRecurso = ""
	compartilhaAcesso = ""

	AdresserPath = "http://10.2.30.4:9092/dsvapg/services/ApiaWS"
	'"http://10.102.2.9:8080/APGDefaultSimulacao/services/ApiaWS"  '"http://10.2.30.4:9092/dsvapg/services/ApiaWS"
	strXmlEndereco = ""
	EncontrouDados = false

	StrLogin = "admin_ebt"
	StrSenha = "51-482-15-9939-4938-50126-87122160115100"
	StrClasse = "INTERFSOLICITARRETURN"





	Vetor_Campos(1)="adWChar,50,adParamInput,null "
	Vetor_Campos(2)="adInteger,1,adParamInput,null "
	Vetor_Campos(3)="adWChar,50,adParamInput,null "
	Vetor_Campos(4)="adInteger,1,adParamInput, " & IdInterfaceAPG

	strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",4,Vetor_Campos)

	Response.Write "strSql:" & strSql& "<BR>"

	Set objRSDadosInterf = db.Execute(strSql)

	If Not objRSDadosInterf.eof() and  Not objRSDadosInterf.Bof() Then

		EncontrouDados = True

		processo = objRSDadosInterf("Processo")
		acao = objRSDadosInterf("Acao")
		idTarefaApg = objRSDadosInterf("ID_Tarefa_Apg")
		numeroSolicitacao = dblSolid

	End If


	If EncontrouDados = True Then

			Vetor_Campos(1)="adWChar,10,adParamInput,"& dblSolid
			Vetor_Campos(2)="adWChar,20,adParamInput, 4" '4 - construir Return 
			Vetor_Campos(3)="adWChar,20,adParamInput," & trim(processo)
			Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(acao)

			StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",4,Vetor_Campos)

			db.execute(strSQL)

				Strxml			=   "<soap:Envelope "
				Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance"""
				Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema"""
				Strxml = Strxml &   " xmlns:soap=" &"""http://schemas.xmlsoap.org/soap/envelope/"""

				Strxml = Strxml & 	"> <soap:Body> "

				'Strxml = Strxml & 	"	<!-- Define a operação sendo realizada (executar classe) --> "
				Strxml = Strxml & 	"	<executeClass> "
				'Strxml = Strxml & 	"		<!-- Ambiente do Apia a ser chamado --> "
				Strxml = Strxml & 	"		<envName>APG</envName> "


				'Strxml = Strxml & 	"		<!-- Nome da classe de negócio, tal como configurada no Apia --> "
				Strxml = Strxml & 	"		<className>" & StrClasse & "</className> "
				'Strxml = Strxml & 	"		<!-- Parâmetros configurados na classe --> "
				Strxml = Strxml & 	"		<parameters> "

				Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" &  "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & "</parameter> "
				'Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "


				Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & objRSDadosInterf("ID_Tarefa_Apg") & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" & "</parameter> "


				Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoConstrucao""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & "OK" &"</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """dataEstocagem""" &">" &"" &"</parameter> "

				Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroCodigoProvedor""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroNomeProvedor""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroEstacao""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroSlot""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroTimeslot""" &">" & "</parameter> "

				Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeInstaladoraRecurso""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeEmpresaConstrutoraInfra""" &">" &  "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoebtDataAceitacaoInfra""" &">" &  "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroAcessoAde""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroOtsAcessoEmbratel""" &">" & "</parameter> "
				Strxml = Strxml & 	"		<parameter name=" & """acessoebtDesignacaoBandabasicaCriada""" &">" & "</parameter> "

				Strxml = Strxml & 	"			<parameter name=" & """bloco""" &">" & "</parameter> "
				Strxml = Strxml & 	"			<parameter name=" & """cabo"""  &">" & "</parameter> "
				Strxml = Strxml & 	"			<parameter name=" & """par""" &">" & "</parameter> "
				Strxml = Strxml & 	"			<parameter name=" & """pino""" &">" & "</parameter> "

				Strxml = Strxml & 	"	</parameters> "

				'Strxml = Strxml & 	"		<!-- Dados do usuário --> "
				Strxml = Strxml & 	"		<userData> "
				'Strxml = Strxml & 	"		<!-- Usuário Apia executante --> "
				Strxml = Strxml & 	"			<usrLogin>" & StrLogin & "</usrLogin> "
				'Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha já encriptada) "
				Strxml = Strxml & 	"			<password>" & StrSenha  & "</password> "
				Strxml = Strxml & 	"			</userData> "
				Strxml = Strxml & 	"	</executeClass> "
				Strxml = Strxml & 	"</soap:Body> "
				Strxml = Strxml & 	"</soap:Envelope> "

				Set doc = server.CreateObject("Microsoft.XMLDOM")
				Set doc1 = server.CreateObject("Microsoft.XMLDOM")
				Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

				doc.async = False
				doc1.async = False

				xmlhttp.Open "POST", AdresserPath, StrLogin, StrSenha
				xmlhttp.setRequestHeader "SOAPAction", "executeClass"

				xmlhttp.send(Strxml)

				doc.loadXML(Strxml)

				strRetorno = xmlhttp.ResponseText
				doc1.loadXML(strRetorno)

				Set xmlhttp= Nothing
				Set myDoc= Nothing

				If doc1.parseError<>0 Then

					strxmlResp = "Erro no XML retornado pelo APG: "
					strxmlResp = strxmlResp  &  "<Codigo> " & doc1.parseError.errorCode & "</codigo>"
					strxmlResp = strxmlResp  & 	"<Descricao>" & strErroXml & Trim(doc1.parseError.reason)
					strxmlResp = strxmlResp  & 	Trim(doc1.parseError.line) & "</Descricao>"

				    Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
					Vetor_Campos(2)="adInteger,6,adParamInput," & ID_Interface_APG	 'Identificação do APG
					Vetor_Campos(3)="addouble,10,adParamInput," & Acl_idacessologico 'ID Logico
					Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
					Vetor_Campos(5)="adInteger,4,adParamOutput,0"

					Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

					ObjCmd.Execute'pega dbaction
					DBAction = ObjCmd.Parameters("RET").value

				Else


					Set objNodeList = Doc1.selectNodes("//soapenv:Envelope/soapenv:Body/executeClassResponse/ns1:executeClassReturn/ns3:parameters/ns3:parameter")


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

						EnviarRetornoContr_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."


					Else


						'Obtem retornos enviados pelo APG
						Tamanho = objNodeList.Length
						posCodErro = tamanho - 2
						posDescErro = tamanho - 1

						codErro = objNodeList.Item(posCodErro).Text
						DescErro = objNodeList.Item(posDescErro).Text
						strxmlResp = ""

						If Trim(codErro) <> "" and Trim(codErro) <> "0"  Then

							strxmlResp = "O Seguinte erro foi retornado pela Interface CLA => APG - Ação: Construir_Acesso_Return" & codErro & DescErro

							Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
							Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
							Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
							Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
							Vetor_Campos(5)="adInteger,4,adParamOutput,0"

							Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value

							EnviarRetornoContr_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."


						Else


							EnviarRetornoContr_Apg = "Interface com APG realizada com Sucesso. "


							Vetor_Campos(1)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
							Vetor_Campos(2)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
							Vetor_Campos(3)="adWChar,30,adParamInput, " & OE_Solicitacao_OEU
							Vetor_Campos(4)="adWChar,20,adParamInput, " & idTarefaApg
							Vetor_Campos(5)="adWChar,15,adParamInput,"& dblIdLogico
							Vetor_Campos(6)="adWChar,10,adParamInput,"& dblSolid
							Vetor_Campos(7)="adWChar,20,adParamInput, "
							Vetor_Campos(8)="adInteger,4,adParamOutput,0"
							Vetor_Campos(9)="adWChar,100,adParamOutput,"

							Response.Write "<script language=javascript>alert('Define par. Procedures')</script>"

							'd = APENDA_PARAMstr("CLA_sp_ins_Construir_Acesso_Ret",9,Vetor_Campos)

							Call APENDA_PARAM("CLA_sp_ins_Construir_Acesso_Ret",9,Vetor_Campos)

							

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value
							DBDescricao = ObjCmd.Parameters("RET1").value


							Response.Write "<script language=javascript>alert('Grava ret. interface')</script>"
							strxmlResp = "Erro ao atualizar log da Interface - Codigo:" & DBAction & "Descrição:" & DBDescricao
							Response.Write "<script language=javascript>alert('" & strxmlResp &"')</script>"


							Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
							Vetor_Campos(2)="adInteger,6,adParamInput," & idTarefaApg		 'Identificação do APG
							Vetor_Campos(3)="addouble,10,adParamInput," & idLogico			 'ID Logico
							Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
							Vetor_Campos(5)="adInteger,4,adParamOutput,0"

							Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value

							Response.Write "<script language=javascript>alert('" & DBAction &"')</script>"

						End if


					End if

				End if
				Response.write(Strxml)

	End If

End Function
%>