<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarRetornoSolic_Apg(dblIdLogico, dblSolAPGId, dblSolid)

	'Dim dblIdLogico, dblSolAPGId, dblSolid
	
	%><!--#include file="../inc/conexao_apg.asp"--><%

	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim StrLogin
	Dim StrSenha
	Dim StrClasse

	Dim strXmlEndereco
	Dim EncontrouDados
	Dim Tec_ID
	Dim strVelAcessoFis
	Dim DescErro
	Dim CodErro

	Dim processo, acao, idLogico
	Dim idTarefaApg, numeroSolicitacao, codigoCsl
	Dim codigoSap, propriedadeDoAcesso
	Dim velocidadeDoAcesso, dataSolicitacaoAcesso, estacaoEntregaAcesso
	Dim emiteOts2m, necessitaRecurso
	Dim retirarRecurso, quantidadeAcessosFisicos, reaproveitarAcesso
	Dim compartilhaAcesso

	''Endereço Intermediario
	Dim cnpj, inscricao_estadual, inscricao_municipal
	Dim cnl, proprietario, bairro
	Dim cep, cidade, complemento
	Dim logradouro, numero, tipo_logradouro
	Dim uf, centro_cliente, endereco_instalacao_pontos_intermediarios

	'troncos2MCompartilhados
	Dim designacao
	'DID
	Dim Did
	
	Dim status ' Status id

    'Setar valor padrão:
	EncontrouDados = false
	Tec_ID = 0
	
	StrClasse = "INTERFSOLICITARRETURN"
	retirarRecurso = "N"

	if dblIdLogico <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_sel_APG_ParametrosMisto",1,Vetor_Campos)

			Set objRSDadosCla = db.Execute(strSqlRet)

			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
				Acf_id = objRSDadosCla("Acf_ID") 'Fase 2
			    Tec_ID = objRSDadosCla("Tec_ID")
				if Tec_ID = "" or isnull(Tec_ID) then
				  Tec_ID = 0
				end if
	            
				propriedadeDoAcesso = objRSDadosCla("Acf_Proprietario")
				estacaoEntregaAcesso = trim(objRSDadosCla("Acf_SiglaEstEntregaFisico")) &"-"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico")
				compartilhaAcesso = objRSDadosCla("Alf_AutorizaCompart")
				necessitaRecurso = objRSDadosCla("Sol_NecessitaRecurso")
				AcessoEBT = objRSDadosCla("AcessoEBT")
				strVelAcessoFis = objRSDadosCla("Vel_Desc")
				quantidadeAcessosFisicos = objRSDadosCla("qtdfis")
				
				Vetor_Campos(1)="adWChar,50,adParamInput,null "
			    Vetor_Campos(2)="adInteger,1,adParamInput,null "
			    Vetor_Campos(3)="adWChar,50,adParamInput,null "
			    Vetor_Campos(4)="adInteger,1,adParamInput," & dblSolAPGId
			    Vetor_Campos(5)="adInteger,1,adParamInput,null "

			    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)

				'Response.Write "strSql:" & strSql& "<BR>" '@@DEBUG
				'Response.end

				Set objRSDadosInterf = db.Execute(strSql)

				If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

					EncontrouDados = True

					processo = objRSDadosInterf("Processo")
					acao = objRSDadosInterf("Acao")
					idTarefaApg = objRSDadosInterf("ID_Tarefa_Apg")
					numeroSolicitacao = dblSolid
					dataSolicitacaoAcesso = date()
					OE_Solicitacao_OEU = objRSDadosInterf("Oe_numero")'Adicionado PRSS - 03/04/2007
					
				End If

				strSQL = "SELECT Sol_EmiteOTS,Sol_ReaproveitarFisico FROM CLA_SOLICITACAO WHERE SOL_ID = " & dblSolid
				Set objRSSol = db.Execute(strSql)

				If Not objRSSol.eof and not objRSSol.Bof Then
				  emiteOts2m = objRSSol("Sol_EmiteOTS")
				  if acao = "ALT" then 
				  	  reaproveitarAcesso = "N"	
				  else
				  	  reaproveitarAcesso = objRSSol("Sol_ReaproveitarFisico")
				  end if 
				  
				  if emiteOts2m = "" then emiteOts2m = "N" end if
				  if reaproveitarAcesso = "" then reaproveitarAcesso = "S" end if
			    End If
				
				if compartilhaAcesso = "S" then
				
					'Fase 2
		            strSQL = "select top 1 'S' as compartilhamento from cla_apg_compartilhamento where comp_idtarefaapg = "& idTarefaApg &" and ( comp_rota = 'S' or comp_distribuicao='S')"
		            set objRScompart = db.execute(strSQL)
					If Not objRScompart.eof and not objRScompart.Bof Then
	 	           		compart_rota_distrib = objRScompart("compartilhamento")
					end if
					
					if compart_rota_distrib = "S" then
					    strSQL = "select desig_designacao from cla_assocdesiglogico inner join cla_designacao on cla_assocdesiglogico.desig_id = cla_designacao.desig_id where acl_idacessologico = " & dblIdLogico
						
						set objRSdesig = db.execute(strSQL)
						xmlTroncos  = ""
						
						while not objRSdesig.eof
						   xmlTroncos = xmlTroncos & 	"	<troncos-2M-compartilhados> "
						   xmlTroncos = xmlTroncos & 	"		<designacao>" & objRSdesig("desig_designacao") & "</designacao> "
						   xmlTroncos = xmlTroncos & 	"	</troncos-2M-compartilhados> "
						   objRSdesig.movenext
						wend
				    End if
					status = "52"
					
					Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolid
					Vetor_Campos(2)="adInteger,4,adParamInput, " & status
					Vetor_Campos(3)="adWChar,30,adParamInput,"  & strLoginRede
					Vetor_Campos(4)="adWChar,300,adParamInput,null" 
					Vetor_Campos(5)="adDouble,8,adParamInput,"	& dblIdLogico
						
					strSql =  APENDA_PARAMSTR("CLA_sp_upd_processo",5,Vetor_Campos)
					'response.write "<script>alert('"&strSql&"')</script></script>"
					'Response.end
					db.Execute(strSql)
				end if 

			If EncontrouDados = True Then

				'db.execute("update cla_apg_solicita_acesso set solacesso_enviado = 'S' where id_tarefa_apg = '" &idTarefaApg & "' and processo = '" &processo &"' and acao = '" & acao & "'")
				db.execute("update cla_apg_solicita_acesso set solacesso_enviado = 'S' where Sol_Acesso_ID = " & dblSolAPGId)
			    db.execute("update cla_acessofisico set Usado_Interf_APG = 1 where acf_id = " & Acf_id)

			        Vetor_Campos(1)="adWChar,10,adParamInput,"& dblSolid
					Vetor_Campos(2)="adWChar,20,adParamInput, 2 " ' 2 - Solicitar Return
					Vetor_Campos(3)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
					Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
					Vetor_Campos(5)="adInteger,4,adParamInput,"& trim(objRSDadosCla("Acf_ID"))
					Vetor_Campos(6)="adInteger,4,adParamInput,"& Tec_ID

					StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",6,Vetor_Campos)
					db.execute(strSQL)

				Strxml = "<?xml version=""1.0"" encoding=""UTF-8""?>"
				Strxml = Strxml &   "<soap:Envelope " & vbnewline
				Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance""" & vbnewline
				Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema""" & vbnewline
				Strxml = Strxml &   " xmlns:soap="&"""http://schemas.xmlsoap.org/soap/envelope/""" & vbnewline

				Strxml = Strxml & 	"> <soap:Body> " & vbnewline

'				Strxml = Strxml & 	"	<!-- Define a operacao sendo realizada (executar classe) --> "
				Strxml = Strxml & 	"	<executeClass> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Ambiente do Apia a ser chamado --> "
				Strxml = Strxml & 	"		<envName>APG</envName> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Nome da classe de negocio, tal como configurada no Apia --> "
				Strxml = Strxml & 	"		<className>" & StrClasse & "</className> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Parametros configurados na classe --> "
				Strxml = Strxml & 	"		<parameters> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & Trim(processo) & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & Trim(acao) & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & mid(trim(dblIdLogico),1,2) & "8" & mid(trim(dblIdLogico),4,10)  & "</parameter> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "

				Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & Trim(idTarefaApg) & "</parameter> " & vbnewline
				'Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & "7" & "</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" & ">" & dblSolid & "</parameter> " & vbnewline
				Strxml = Strxml & 	" 		<parameter name=" & """codigoCsl""" &"></parameter>" & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """codigoSap""" &"></parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" & propriedadeDoAcesso & "</parameter> " & vbnewline

			'	Strxml = Strxml & 	"		<parameter name=" & """tecnologia""" &">" & strTecnologia & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """tecnologia""" &">" & Tec_ID & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """velocidadeDoAcesso""" &">" & strVelAcessoFis & "</parameter> " & vbnewline
			'	Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoAcesso""" &">" & dataSolicitacaoAcesso & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoAcesso""" &">" & dataSolicitacaoAcesso & "</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """estacaoEntregaAcesso""" &">" & estacaoEntregaAcesso & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """tipoAcesso""" &">" & AcessoEBT & "</parameter> " & vbnewline

			'	Strxml = Strxml & 	"		<parameter name=" & """emiteOts2m""" &">" & emiteOts2m & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """emiteOts2m""" &">" & emiteOts2m & "</parameter> " & vbnewline
			'	Strxml = Strxml & 	"		<parameter name=" & """necessitaRecurso""" &">" & necessitaRecurso  & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """necessitaRecurso""" &">" & necessitaRecurso  & "</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """retirarRecurso""" &">" & retirarRecurso &"</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """quantidadeAcessosFisicos""" &">" & quantidadeAcessosFisicos & "</parameter> " & vbnewline
			'	Strxml = Strxml & 	"		<parameter name=" & """reaproveitarAcesso""" &">" & reaproveitarAcesso &"</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """reaproveitarAcesso""" &">" & reaproveitarAcesso &"</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """compartilhaAcesso""" &">" & compartilhaAcesso  &"</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """enderecoInstalacaoPontosIntermediarios""" &">" & vbnewline

				xmlEndereco = ""
				xmlEndereco = xmlEndereco & 	"			<endereco-instalacao-pontos-intermediarios> "
				xmlEndereco = xmlEndereco & 	"				<cnpj></cnpj> "
				xmlEndereco = xmlEndereco & 	"				<inscricao-estadual></inscricao-estadual> "
				xmlEndereco = xmlEndereco & 	"				<inscricao-municipal></inscricao-municipal> "
				xmlEndereco = xmlEndereco & 	"				<cnl></cnl> "
				xmlEndereco = xmlEndereco & 	"				<proprietario></proprietario> "
				xmlEndereco = xmlEndereco & 	"				<bairro></bairro> "
				xmlEndereco = xmlEndereco & 	"				<cep></cep> "
				xmlEndereco = xmlEndereco & 	"				<cidade></cidade> "
				xmlEndereco = xmlEndereco & 	"				<complemento></complemento> "
				xmlEndereco = xmlEndereco & 	"				<logradouro></logradouro> "
				xmlEndereco = xmlEndereco & 	"				<numero></numero> "
				xmlEndereco = xmlEndereco & 	"				<tipo-logradouro></tipo-logradouro> "
				xmlEndereco = xmlEndereco & 	"				<uf></uf> "
				xmlEndereco = xmlEndereco & 	"				<centro-cliente></centro-cliente> "
				xmlEndereco = xmlEndereco & 	"			</endereco-instalacao-pontos-intermediarios> "

				Strxml = Strxml & 	Server.URLEncode(xmlEndereco) & vbnewline
				Strxml = Strxml & 	"		</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """troncos2MCompartilhados""" &">" & vbnewline

				'Fase 2
				if xmlTroncos <> "" then
				  Strxml = Strxml & 	server.URLEncode(xmlTroncos) & vbnewline
				else
					xmlTroncos = xmlTroncos & 	"	<troncos-2M-compartilhados> "
					xmlTroncos = xmlTroncos & 	"		<designacao></designacao> "
					xmlTroncos = xmlTroncos & 	"	</troncos-2M-compartilhados> "
				end if

				Strxml = Strxml & 	"		</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """did""" &">" & vbnewline

				xmlDid = 			"			<did> " & vbnewline
				xmlDid = xmlDid & 	"					<valor-did></valor-did> " & vbnewline
				xmlDid = xmlDid & 	"			</did> "  & vbnewline

				Strxml = Strxml & server.UrlEncode(xmlDid)

				Strxml = Strxml & 	"		</parameter>" & vbnewline

				Strxml = Strxml & 	"	</parameters> " & vbnewline

'				Strxml = Strxml & 	"		<!-- Dados do usuario --> "
				Strxml = Strxml & 	"		<userData> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Usuario Apia executante --> "
				Strxml = Strxml & 	"			<usrLogin>" & StrLogin & "</usrLogin> " & vbnewline
'				Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha ja encriptada) --> "
				Strxml = Strxml & 	"			<password>" & StrSenha & "</password> " & vbnewline
				Strxml = Strxml & 	"		</userData> " & vbnewline
				Strxml = Strxml & 	"	</executeClass> " & vbnewline
				Strxml = Strxml & 	"</soap:Body> " & vbnewline
				Strxml = Strxml & 	"</soap:Envelope> " & vbnewline

				'response.write "XML - " & replace(replace(replace(Strxml,"<","&lt;"),">","&gt;"),vbnewline,"<BR>"&vbnewline)

				''Define Objetos XML e HTTP para interface
				Set doc = server.CreateObject("Microsoft.XMLDOM")
				Set doc1 = server.CreateObject("Microsoft.XMLDOM")
				Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

				doc.async= False
				doc1.async= False

				xmlhttp.Open "POST", AdresserPath, false, StrLogin, StrSenha
				xmlhttp.setRequestHeader "SOAPAction" , "executeClass"
				xmlhttp.setRequestHeader "Content-Length", len(Strxml)

				xmlhttp.send(Strxml)

				strRetorno = xmlhttp.ResponseText
				'response.write "<BR>" & "Retorno:" & strRetorno

				doc.loadXML(strRetorno)
                                doc1.loadXML(Strxml)
				'Checa se serviço é 0800.
				Oe_numero = objRSDadosInterf("Oe_numero")
				Oe_ano = objRSDadosInterf("Oe_ano")
				Oe_item = objRSDadosInterf("Oe_item")
				Id_logico = dblIdLogico
				Processo = trim(objRSDadosInterf("Processo"))
				Acao = trim(objRSDadosInterf("Acao"))
				
				call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,2,strxml)
				
				call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,2,strRetorno)

				Set xmlhttp= Nothing
				
				'Adicionado PRSS - 03/04/2007
				Vetor_Campos(1)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
				Vetor_Campos(2)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
				Vetor_Campos(3)="adWChar,30,adParamInput, " & OE_Solicitacao_OEU
				Vetor_Campos(4)="adWChar,20,adParamInput, " & idTarefaApg
				Vetor_Campos(5)="adWChar,15,adParamInput,"& dblIdLogico
				Vetor_Campos(6)="adWChar,10,adParamInput,"& dblSolid
				Vetor_Campos(7)="adWChar,20,adParamInput, "
				Vetor_Campos(8)="adInteger,4,adParamOutput,0"
				Vetor_Campos(9)="adWChar,100,adParamOutput,"

				Call APENDA_PARAM("CLA_sp_ins_Solicitar_Acesso_Ret",9,Vetor_Campos)
				
				ObjCmd.Execute'pega dbaction				
				
				DBAction = ObjCmd.Parameters("RET").value

				If trim(doc.parseError) <> "0" and trim(doc.parseError) <> "0" Then

					strxmlResp = "Erro no XML retornado pelo APG: "
					strxmlResp = strxmlResp  &  "<Codigo> " & doc.parseError.errorCode & "</codigo>"
					strxmlResp = strxmlResp  & 	"<Descricao>" & strErroXml & Trim(doc.parseError.reason)
					strxmlResp = strxmlResp  & 	Trim(doc.parseError.line) & "</Descricao>"

				    Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid										'Solicitação
					Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
					Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
					Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
					Vetor_Campos(5)="adInteger,4,adParamOutput,0"

					Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

					ObjCmd.Execute'pega dbaction
					DBAction = ObjCmd.Parameters("RET").value

					'Response.write "<script>alert('EnviarRetornoSolic.asp - Teste 1')</script>"
					EnviarRetornoSolic_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

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
						If trim(DBAction) <> "1" Then

							Response.Write "Erro na Inclusão do LOG:" & DBAction

						End If
						'Response.write "<script>alert('EnviarRetornoSolic.asp - Teste 2')</script>"
						EnviarRetornoSolic_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

					Else

						'Obtem retornos enviados pelo APG
						Tamanho = objNodeList.Length
						posCodErro = tamanho - 2
						posDescErro = tamanho - 1

						codErro = objNodeList.Item(posCodErro).Text
						DescErro = objNodeList.Item(posDescErro).Text
						strxmlResp = ""

						If Trim(codErro) <> "" and Trim(codErro) <> "0"  Then

							strxmlResp = "O Seguinte erro foi retornado pela Interface CLA => APG - Ação: Solicitar_Acesso_Return" & codErro & DescErro

							Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
							Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
							Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
							Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
							Vetor_Campos(5)="adInteger,4,adParamOutput,0"

							Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value
							
							EnviarRetornoSolic_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

						Else

							EnviarRetornoSolic_Apg = "Interface com APG realizada com Sucesso. "
							
							'Execute("SELECT acf_id FROM CLA_VIEW_SOLICITACAOMIN WHERE sol_id = '" & strsolicitacao & "'")

							db.execute("update cla_apg_solicita_acesso set solacesso_enviado = 'S' where Sol_Acesso_ID = " & dblSolAPGId)
						    db.execute("update cla_acessofisico set Usado_Interf_APG = 1 where acf_id = " & Acf_id)
						End if

					End If
				End if

			Else

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & dblIdLogico		'Identificação do APG
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
				'Response.write "<script>alert('EnviarRetornoSolic.asp - Teste 4')</script>"
				EnviarRetornoSolic_Apg = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."
			End If

	End If
End if
End Function




Function EnviarRetornoSolic_Apg_CAN_DES(dblIdLogico, dblSolAPGId, dblSolid)

	'Dim dblIdLogico, dblSolAPGId, dblSolid

'	Response.Write "<script language=javascript>alert('" & "Logico:" & dblIdLogico & "');</script>"
'	Response.Write "<script language=javascript>alert('" & "dblSolAPGId:" & dblSolAPGId & "');</script>"
'	Response.Write "<script language=javascript>alert('" & "Solic." & dblSolid & "');</script>"

'	Response.Write "Logico:" & dblIdLogico '= 6787123691
'	Response.Write "dblSolAPGId:" & dblSolAPGId
'	Response.Write "Solic." & dblSolid 'dblSolid = 233962

	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim StrLogin
	Dim StrSenha
	Dim StrClasse

	Dim strXmlEndereco
	Dim EncontrouDados
	Dim StrTipoAcessso
	Dim DescErro
	Dim CodErro

	Dim processo, acao, idLogico
	Dim idTarefaApg, emiteOts2m
	Dim retirarRecurso, compartilhaAcesso
	dim dataSolicitacaoAcesso
	
	'response.write dblIdLogico
	'response.write dblSolAPGId
	'response.write dblSolid
	'response.end
	%><!--#include file="../inc/conexao_apg.asp"--><%
	
	strXmlEndereco = ""
	EncontrouDados = false

	StrClasse = "INTERFSOLICITARRETURN"

	if dblIdLogico <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput,"
			Vetor_Campos(2)="adInteger,4,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

			'Response.Write "SQL: " & strSqlRet & "<BR>" '@@DEBUG

			Set objRSDadosCla = db.Execute(strSqlRet)

			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			
				Vetor_Campos(1)="adInteger,1,adParamInput,null "
			    Vetor_Campos(2)="adWChar,50,adParamInput, " & objRSDadosCla("Acf_SiglaEstEntregaFisico")
			    Vetor_Campos(3)="adWChar,50,adParamInput, " & objRSDadosCla("Acf_ComplSiglaEstEntregaFisico")
			   
			    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_estacao ",3,Vetor_Campos)

				set objRSSAP = db.execute(strSQL)
				
				If not objRSSAP.Eof and  not objRSSAP.Bof Then
					
					codigoCsl = objRSSAP("esc_cod_sap") '  o codigo csl e o codigo sap sao o mesmo conforme a ISA da Embratel 
					codigoSap = objRSSAP("esc_cod_sap")
				end if

				propriedadeDoAcesso = objRSDadosCla("Acf_Proprietario")
				compartilhaAcesso = objRSDadosCla("Alf_AutorizaCompart")

				Vetor_Campos(1)="adWChar,50,adParamInput,null "
			    Vetor_Campos(2)="adInteger,1,adParamInput,null "
			    Vetor_Campos(3)="adWChar,50,adParamInput, null " 
			    Vetor_Campos(4)="adInteger,8,adParamInput, " & dblSolAPGId
			    Vetor_Campos(5)="adInteger,1,adParamInput,null "

			    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
				
				Set objRSDadosInterf = db.Execute(strSql)

				If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

					EncontrouDados = True

					processo = objRSDadosInterf("Processo")
					acao = objRSDadosInterf("Acao")
					idTarefaApg = objRSDadosInterf("ID_Tarefa_Apg")
					numeroSolicitacao = dblSolid

				End If

			End If
			
			if acao = "DES" THEN 
				retirarRecurso = "S"
			else
			   retirarRecurso = "N"
			 end if 

			strSQL = "SELECT Sol_EmiteOTS,Sol_NecessitaRecurso,Sol_ReaproveitarFisico , sol_data FROM CLA_SOLICITACAO WHERE SOL_ID = " & dblSolid
			Set objRSSol = db.Execute(strSql)

			If Not objRSSol.eof and not objRSSol.Bof Then
			  emiteOts2m = objRSSol("Sol_EmiteOTS")
			  necessitaRecurso = objRSSol("Sol_NecessitaRecurso")
			  reaproveitarAcesso = objRSSol("Sol_ReaproveitarFisico")
			  dataSolicitacaoAcesso = Formatar_Data(objRSSol("Sol_Data"))
			  
			  if isnull(emiteOts2m) then emiteOts2m = "N" end if
			  if isnull(necessitaRecurso) then necessitaRecurso = "N" end if
			  if isnull(reaproveitarAcesso) then reaproveitarAcesso = "N" end if
			  
			End If
			
			if compartilhaAcesso = 1 then
			   compartilhaAcesso = "S"
			   necessitaRecurso = "N"
			else
			   compartilhaAcesso = "N"
			end if
			
			If trim(objRSDadosCla("Acf_Proprietario")) = "EBT" Then
				StrTipoAcessso = "S"
			Else
				necessitaRecurso = "N"
				StrTipoAcessso = "N"
			End If

			
			If EncontrouDados = True Then

				Vetor_Campos(1)="adWChar,10,adParamInput,"& dblSolid
				Vetor_Campos(2)="adWChar,20,adParamInput, 2 " ' 2 - Solicitar Return
				Vetor_Campos(3)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
				Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))

				StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",4,Vetor_Campos)
				
				db.execute(strSQL)

				strSQL = "SELECT MAX(SOL_ID)AS SOL_ID FROM CLA_SOLICITACAO WHERE ACL_IDACESSOLOGICO =" & dblIdLogico
				Set objRSMaxSol = db.Execute(strSql)
				
				Strxml = "<?xml version=""1.0"" encoding=""UTF-8""?>"
				Strxml = Strxml &   "<soap:Envelope " & vbnewline
				Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance""" & vbnewline
				Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema""" & vbnewline
				Strxml = Strxml &   " xmlns:soap="&"""http://schemas.xmlsoap.org/soap/envelope/""" & vbnewline
				Strxml = Strxml & 	"> <soap:Body> " & vbnewline

'#				Define a operacao sendo realizada (executar classe)
				Strxml = Strxml & 	"	<executeClass> " & vbnewline

'#				Ambiente do Apia a ser chamado
				Strxml = Strxml & 	"		<envName>APG</envName> " & vbnewline

'#				Nome da classe de negocio, tal como configurada no Apia
				Strxml = Strxml & 	"		<className>" & StrClasse & "</className> " & vbnewline

'#				Parametros configurados na classe
				Strxml = Strxml & 	"		<parameters> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" &processo&  "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" &acao& "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & mid(trim(dblIdLogico),1,2) & "8" & mid(trim(dblIdLogico),4,10) & "</parameter> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "

				Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & Trim(idTarefaApg) & "</parameter> " & vbnewline 'OBRIG

				Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" & ">"& Trim(numeroSolicitacao) &"</parameter> " & vbnewline
				Strxml = Strxml & 	" 		<parameter name=" & """codigoCsl""" &">" & codigoCsl &"</parameter>" & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """codigoSap""" &">" & codigoSap &"</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" & trim(objRSDadosCla("Acf_Proprietario")) & "</parameter> " & vbnewline 'OBRIG

				Strxml = Strxml & 	"		<parameter name=" & """tecnologia""" &">" & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """velocidadeDoAcesso""" &">" & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoAcesso""" &">"& dataSolicitacaoAcesso& "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """estacaoEntregaAcesso""" &">" & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """tipoAcesso""" &">" &StrTipoAcessso& "</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """emiteOts2m""" &">" & emiteOts2m & "</parameter> " & vbnewline 'OBRIG
				Strxml = Strxml & 	"		<parameter name=" & """necessitaRecurso""" &">" & necessitaRecurso  & "</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """retirarRecurso""" &">" & retirarRecurso &"</parameter> " & vbnewline 'OBRIG
				Strxml = Strxml & 	"		<parameter name=" & """quantidadeAcessosFisicos""" &">" & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """reaproveitarAcesso""" &">" & "</parameter> " & vbnewline
				Strxml = Strxml & 	"		<parameter name=" & """compartilhaAcesso""" &">" & compartilhaAcesso  &"</parameter> " & vbnewline 'OBRIG

				Strxml = Strxml & 	"		<parameter name=" & """enderecoInstalacaoPontosIntermediarios""" &">" & vbnewline

				xmlEndereco = ""
				xmlEndereco = xmlEndereco & 	"			<endereco-instalacao-pontos-intermediarios> "
				xmlEndereco = xmlEndereco & 	"			<cnpj></cnpj> "
				xmlEndereco = xmlEndereco & 	"				<inscricao-estadual></inscricao-estadual> "
				xmlEndereco = xmlEndereco & 	"				<inscricao-municipal></inscricao-municipal> "
				xmlEndereco = xmlEndereco & 	"				<cnl></cnl> "
				xmlEndereco = xmlEndereco & 	"				<proprietario></proprietario> "
				xmlEndereco = xmlEndereco & 	"				<bairro></bairro> "
				xmlEndereco = xmlEndereco & 	"				<cep></cep> "
				xmlEndereco = xmlEndereco & 	"				<cidade></cidade> "
				xmlEndereco = xmlEndereco & 	"				<complemento></complemento> "
				xmlEndereco = xmlEndereco & 	"				<logradouro></logradouro> "
				xmlEndereco = xmlEndereco & 	"				<numero></numero> "
				xmlEndereco = xmlEndereco & 	"				<tipo-logradouro></tipo-logradouro> "
				xmlEndereco = xmlEndereco & 	"				<uf></uf> "
				xmlEndereco = xmlEndereco & 	"				<centro-cliente></centro-cliente> "
				xmlEndereco = xmlEndereco & 	"			</endereco-instalacao-pontos-intermediarios> "

				Strxml = Strxml & 	Server.URLEncode(xmlEndereco) & vbnewline
				Strxml = Strxml & 	"		</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """troncos2MCompartilhados""" &">" & vbnewline

				xmlTroncos = 		"			<troncos-2M-compartilhados> "
				xmlTroncos = xmlTroncos & 	"		<designacao></designacao> "
				xmlTroncos = xmlTroncos & 	"	</troncos-2M-compartilhados> "

				Strxml = Strxml & 	server.URLEncode(xmlTroncos) & vbnewline

				Strxml = Strxml & 	"		</parameter> " & vbnewline

				Strxml = Strxml & 	"		<parameter name=" & """did""" &">" & vbnewline

				xmlDid = 			"			<did> " & vbnewline
				xmlDid = xmlDid & 	"					<valor-did></valor-did> " & vbnewline
				xmlDid = xmlDid & 	"			</did> "  & vbnewline

				Strxml = Strxml & server.UrlEncode(xmlDid)

				Strxml = Strxml & 	"		</parameter>" & vbnewline

				Strxml = Strxml & 	"	</parameters> " & vbnewline

'				Strxml = Strxml & 	"		<!-- Dados do usuario --> "
				Strxml = Strxml & 	"		<userData> " & vbnewline
'				Strxml = Strxml & 	"		<!-- Usuario Apia executante --> "
				Strxml = Strxml & 	"			<usrLogin>" & StrLogin & "</usrLogin> " & vbnewline
'				Strxml = Strxml & 	"			<!-- senha criptografada (solicitar da BULL a senha ja encriptada) --> "
				Strxml = Strxml & 	"			<password>" & StrSenha & "</password> " & vbnewline
				Strxml = Strxml & 	"		</userData> " & vbnewline
				Strxml = Strxml & 	"	</executeClass> " & vbnewline
				Strxml = Strxml & 	"</soap:Body> " & vbnewline
				Strxml = Strxml & 	"</soap:Envelope> " & vbnewline

				'response.write "XML - " & replace(replace(replace(Strxml,"<","&lt;"),">","&gt;"),vbnewline,"<BR>"&vbnewline)@@DEBUG

				''Define Objetos XML e HTTP para interface
				Set doc = server.CreateObject("Microsoft.XMLDOM")
				Set doc1 = server.CreateObject("Microsoft.XMLDOM")
				Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

				doc.async= False
				doc1.async= False

				xmlhttp.Open "POST", AdresserPath, false, StrLogin, StrSenha
				xmlhttp.setRequestHeader "SOAPAction" , "executeClass"
				xmlhttp.setRequestHeader "Content-Length", len(Strxml)

				xmlhttp.send(Strxml)

				strRetorno = xmlhttp.ResponseText
				'response.write "<BR>" & "Retorno:" & strRetorno

				doc.loadXML(strRetorno)
				doc1.loadXML(Strxml)

				'Checa se serviço é 0800.
				Oe_numero = objRSDadosInterf("Oe_numero")
				Oe_ano = objRSDadosInterf("Oe_ano")
				Oe_item = objRSDadosInterf("Oe_item")
				Id_logico = dblIdLogico
				Processo = trim(objRSDadosInterf("Processo"))
				Acao = trim(objRSDadosInterf("Acao"))
				
				call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,2,strxml)
				
				call check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,2,strRetorno)

				Set xmlhttp= Nothing

				If trim(doc.parseError) <> "0" and trim(doc.parseError) <> "0" Then

					strxmlResp = "Erro no XML retornado pelo APG: "
					strxmlResp = strxmlResp  &  "<Codigo> " & doc.parseError.errorCode & "</codigo>"
					strxmlResp = strxmlResp  & 	"<Descricao>" & strErroXml & Trim(doc.parseError.reason)
					strxmlResp = strxmlResp  & 	Trim(doc.parseError.line) & "</Descricao>"

				    Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid										'Solicitação
					Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
					Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
					Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
					Vetor_Campos(5)="adInteger,4,adParamOutput,0"

					Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

					ObjCmd.Execute'pega dbaction
					DBAction = ObjCmd.Parameters("RET").value

					EnviarRetornoSolic_Apg_CAN_DES = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

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
						If trim(DBAction) <> "1" Then

							Response.Write "Erro na Inclusão do LOG:" & DBAction

						End If

						EnviarRetornoSolic_Apg_CAN_DES = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

					Else

						'Obtem retornos enviados pelo APG
						Tamanho = objNodeList.Length
						posCodErro = tamanho - 2
						posDescErro = tamanho - 1

						codErro = objNodeList.Item(posCodErro).Text
						DescErro = objNodeList.Item(posDescErro).Text
						strxmlResp = ""

						If Trim(codErro) <> "" and Trim(codErro) <> "0"  Then

							strxmlResp = "O Seguinte erro foi retornado pela Interface CLA => APG - Ação: Solicitar_Acesso_Return" & codErro & DescErro

							Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
							Vetor_Campos(2)="adInteger,6,adParamInput," & Trim(objRSDadosInterf("ID_Tarefa_Apg"))		'Identificação do APG
							Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
							Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
							Vetor_Campos(5)="adInteger,4,adParamOutput,0"

							Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

							ObjCmd.Execute'pega dbaction
							DBAction = ObjCmd.Parameters("RET").value

							EnviarRetornoSolic_Apg_CAN_DES = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

						Else

							EnviarRetornoSolic_Apg_CAN_DES = "Interface com APG realizada com Sucesso. "
							
							db.execute("update cla_apg_solicita_acesso set solacesso_enviado = 'S' where id_tarefa_apg = '" &idTarefaApg & "' and processo = '" &processo &"' and acao = '" & acao & "'")

						End if


					End If


				End if


			Else

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & Trim(dblSolid)								'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput, NULL" 
				Vetor_Campos(3)="addouble,10,adParamInput," & Trim(dblIdLogico)								'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & Trim(strxmlResp)								'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"

				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value

				EnviarRetornoSolic_Apg_CAN_DES = "Erro na Interface com APG. Favor notificar a equipe CLA atraves do 108 ou email para Hdesk."

			End If
	End If
End Function
%>

