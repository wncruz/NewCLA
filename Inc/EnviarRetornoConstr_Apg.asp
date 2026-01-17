<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarRetornoContr_Apg(dblIdLogico, IdInterfaceAPG, dblSolid)

	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim strXmlEndereco
	Dim EncontrouDados
	Dim StrTipoAcessso
	Dim strTecnologia
	Dim strVelAcessoFis
	Dim DescErro
	Dim CodErro
	Dim dtSolicConstrucao, strCodProvedor, strProNome, strEstacaoFac, strSlot, strTimeSlot
	Dim strNome_instaladora_recurso, strNome_emp_constr_infra, dt_aceitacao_infra
	Dim strNumero_ots_acesso_ebt, strDesignacao_bandabasica
	Dim strNumero_acesso_ade, strbloco, strcabo, strpar, strpino

    %><!--#include file="../inc/conexao_apg.asp"--><%

	strXmlEndereco = ""
	EncontrouDados = false
	strTecnologia = ""
	strVelAcessoFis = ""

	StrClasse = "INTERFCONSTRUIRRETURN"
	
	StrCodProvedor = ""
	StrProNome = ""
	strEstacaoFac = ""
	strSlot = ""
	strTimeSlot = ""

	dtSolicConstrucao = Date()
	strTecnologia = ""
	strNome_instaladora_recurso = ""
	strNome_emp_constr_infra = ""
	dt_aceitacao_infra = ""
	strNumero_acesso_ade = ""
	strNumero_ots_acesso_ebt = ""
	strDesignacao_bandabasica = ""
	strbloco = ""
	strcabo = ""
	strpar = ""
	strpino = ""
	dblacf_id = ""

	if dblIdLogico <> "" then
	  Set objRSMisto = db.Execute("select top 1 cla_acessofisico.acf_id,Tec_ID from cla_acessofisico inner join cla_acessologicofisico on cla_acessofisico.Acf_Id = cla_acessologicofisico.Acf_ID where Acl_IDAcessoLogico = "&dblIdLogico&" and Acf_IDAcessoFisico is null")
      If Not objRSMisto.eof and  not objRSMisto.Bof Then
	    dblacf_id = objRSMisto("acf_id")
	    dblTec_ID = objRSMisto("Tec_ID")
	  End if
	  
	  'Só enviar o construir quando todos estiverem construídos.
	  if isnull(dblacf_id) or dblacf_id = "" then
	    'Response.Write "<script language=javascript>alert('Encontrou Logico:" & dblIdLogico &":"& dblSolid &"')</script>"
 		Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico )
		
		dblSolid = objRSSol("sol_id")
		
		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)
        
		Set objRSDadosCla = db.Execute(strSqlRet)
        
		If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
          Vetor_Campos(1)="adWChar,50,adParamInput,null "
		  Vetor_Campos(2)="adInteger,1,adParamInput,null "
		  Vetor_Campos(3)="adWChar,50,adParamInput,null "
		  Vetor_Campos(4)="adInteger,1,adParamInput,null "
		  Vetor_Campos(5)="adInteger,1,adParamInput, " & dblSolid
          
		  strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
          
		  Set objRSDadosInterf = db.Execute(strSql)
          
		  If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then
		    'Response.Write "<script language=javascript>alert('EncontrouDados True ')</script>"
			EncontrouDados = True
			
			Alteracao_Cadastral = objRSDadosInterf("flag_altcadastral")
		  End If
		End If
		
		Vetor_Campos(1)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
		Vetor_Campos(2)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
		Vetor_Campos(3)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("ID_Tarefa_Apg"))
		Vetor_Campos(4)="adWChar,10,adParamInput,"& dblSolid
		Vetor_Campos(5)="adWChar,20,adParamInput, 4 " '& 4 - Construir Return 
		
		StrSQL = APENDA_PARAMSTR("CLA_sp_sel_Interface_Apg_Enviado",5,Vetor_Campos)
        
		Set objRsInterfEnviado = db.Execute(StrSQL)
        
		If objRsInterfEnviado("flag_enviado") = 0 Then
		  If EncontrouDados = True Then
		    Set objRSInterf = db.Execute("select top 1 cla_acessofisico.acf_id,Tec_ID from cla_acessofisico inner join cla_acessologicofisico on cla_acessofisico.Acf_Id = cla_acessologicofisico.Acf_ID where Acl_IDAcessoLogico = "&dblIdLogico&" and Usado_Interf_APG = 1")
            If Not objRSInterf.eof and  not objRSInterf.Bof Then
	          dblacf_id = objRSInterf("acf_id")
	          dblTec_ID = objRSInterf("Tec_ID")
	        End if
			
			Vetor_Campos(1)="adWChar,10,adParamInput,"& dblSolid
			Vetor_Campos(2)="adWChar,20,adParamInput, 4 " '& 4 - Construir Return 
			Vetor_Campos(3)="adWChar,20,adParamInput," & trim(objRSDadosInterf("Processo"))
			Vetor_Campos(4)="adWChar,20,adParamInput,"& trim(objRSDadosInterf("Acao"))
			Vetor_Campos(5)="adInteger,4,adParamInput,"& dblacf_id
			Vetor_Campos(6)="adInteger,4,adParamInput,"& dblTec_ID
	        
			StrSQL = APENDA_PARAMSTR("CLA_sp_upd_status_apg",6,Vetor_Campos)
			
			db.execute(strSQL)
            
			If trim(objRSDadosCla("Acf_Proprietario")) = "EBT" Then
              
			  ''@@ Davif Incluir validação da Tecnologia para atribuir valores.
			  If objRSDadosCla("Tec_Id") <> "" then
			    Set objRSAux = db.Execute("CLA_Sp_Sel_Tecnologia " & objRSDadosCla("Tec_Id"))
				if not objRSAux.Eof and Not objRSAux.Bof then
				  strTecnologia	= objRSAux("Tec_Sigla")
				End if
			  End if
	          
	          If Trim(strTecnologia) = "RADIO" Or Trim(strTecnologia) = "FIBRA" Or Trim(strTecnologia) = "SATELITE" Then
	            
				'@@ Davif - Dados do Acesso Embratel
				Vetor_Campos(1)="adWChar,18,adParamInput," & dblIdLogico
				Vetor_Campos(2)="adWChar,10,adParamInput," & dblSolid
				
				strSqlRet = APENDA_PARAMSTRSQL("cla_sp_sel_crmsprocessoAcesso",2,Vetor_Campos)
				Set objRSDadosEbt = db.Execute(strSqlRet)
	            
				If not objRSDadosEbt.eof Then 'Alterado PRSS 16/05/2007
				  strNome_instaladora_recurso = objRSDadosEbt("Empresa1")
				  strNome_emp_constr_infra = objRSDadosEbt("Empresa2")
				  dt_aceitacao_infra = objRSDadosEbt("Dataaprovainfra") 'formatData(objRSDadosEbt("Dataaprovainfra"))
				  strNumero_ots_acesso_ebt = objRSDadosEbt("Nrots")
				  strDesignacao_bandabasica = objRSDadosEbt("Designacaotronco")
	            End If
	            
			  Else
			    
			    If Trim(strTecnologia) = "ADE" Then
				  
	 			  '@@ Davif - Obtem Facilidades
				  dblAcfId = objRSDadosCla("Acf_ID")
				  Vetor_Campos(1)="adInteger,4,adParamInput,"
				  Vetor_Campos(2)="adWChar,25,adParamInput,"
				  Vetor_Campos(3)="adWChar,15,adParamInput,"
				  Vetor_Campos(4)="adInteger,4,adParamInput," & dblAcfId
	              
				  strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_sel_facilidade",4,Vetor_Campos)
				  Set objRSFac = db.Execute(strSqlRet)
	              
				  If Not objRSFac.eof and Not objRSFac.bof Then
	                
					strSlot = objRSFac("Fac_Slot")
					strTimeSlot = objRSFac("Fac_TimeSlot")
	                
					strNumero_acesso_ade = trim(objRSDadosCla("Acf_NroAcessoPtaEbt"))
					strbloco = ""
					strcabo = "6"
					strpar = objRSFac("Fac_Par")
					strpino = ""
				  End If
	            End If
	          End If
			  
			Else
			  
			  If trim(objRSDadosCla("Acf_Proprietario")) = "TER" Then
	          
			    '@@ Davif - Obtem Facilidades
			    dblAcfId = objRSDadosCla("Acf_ID")
			    Vetor_Campos(1)="adInteger,4,adParamInput,"
			    Vetor_Campos(2)="adWChar,25,adParamInput,"
			    Vetor_Campos(3)="adWChar,15,adParamInput,"
			    Vetor_Campos(4)="adInteger,4,adParamInput," & dblAcfId
	            
			    strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_sel_facilidade",4,Vetor_Campos)
			    Set objRSFac = db.Execute(strSqlRet)
	            
			    'Response.Write "<script language=javascript>alert('Procura Fac.:');</script>"
			    If Not objRSFac.eof and Not objRSFac.bof Then
	              
				  strSlot = objRSFac("Fac_Slot")
				  strTimeSlot = objRSFac("Fac_TimeSlot")
	              
				  'Obtem Estação da Alocação de Facilidades
				  If objRSFac("Esc_ID") <> "" then
				    Set objRS = db.execute("CLA_sp_sel_estacao " & objRSFac("Esc_ID"))
				    'Response.Write "<script language=javascript>alert('Procura Estacao.:');</script>"
	                
				    if Not objRS.Eof And Not objRS.Bof then
	                  strEstacaoFac = objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla")
					  'Response.Write "<script language=javascript>alert('Achou Estacao.:');</script>"
				    End if
				  End if
	            End If
	            
			    ''@@ Davif - Obtem Nome do Provedor
			    If objRSDadosCla("Pro_ID") <> "" Then
	              set objRSPro = db.execute("CLA_sp_sel_provedor " & objRSDadosCla("Pro_ID"))
				  'Response.Write "<script language=javascript>alert('Procura Prov.:" & objRSDadosCla("Pro_ID") & "');</script>"
	              
				  If Not objRSPro.Eof and Not objRSPro.Bof Then
				    StrCodProvedor 	= objRSPro("Pro_ID")
				    StrProNome 		= objRSPro("Pro_Nome")
				    'Response.Write "<script language=javascript>alert('Achou Prov.:');</script>"
				  Else
				    'Response.Write "<script language=javascript>alert('Não Achou Prov.:');</script>"
				    StrCodProvedor 	= ""
				    StrProNome 		= ""
				  End If
	            End if
	          End If
	        End If
	        
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
		    
		    Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & objRSDadosInterf("Processo") & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & objRSDadosInterf("Acao") & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & mid(trim(dblIdLogico),1,2) & "8" & mid(trim(dblIdLogico),4,10)  & "</parameter> "
		    'Strxml = Strxml & 	"		<!-- Adicionado devido ao APG --> "
		    
		    Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & objRSDadosInterf("ID_Tarefa_Apg") & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" &">"& dblSolid & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" & objRSDadosCla("Acf_Proprietario") & "</parameter> "
		    
		    Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoConstrucao""" &">" & dtSolicConstrucao &"</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & "" &"</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """dataEstocagem""" &">" &"" &"</parameter> "
		    
		    Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroCodigoProvedor""" &">" & StrCodProvedor & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroNomeProvedor""" &">" & StrProNome & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroEstacao""" &">" & strEstacaoFac & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroSlot""" &">" & strSlot & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroTimeslot""" &">" & strTimeSlot & "</parameter> "
		    
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeInstaladoraRecurso""" &">" & strNome_instaladora_recurso & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeEmpresaConstrutoraInfra""" &">" & strNome_emp_constr_infra & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtDataAceitacaoInfra""" &">" & dt_aceitacao_infra & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroAcessoAde""" &">" & strNumero_acesso_ade & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroOtsAcessoEmbratel""" &">" & strNumero_ots_acesso_ebt & "</parameter> "
		    Strxml = Strxml & 	"		<parameter name=" & """acessoebtDesignacaoBandabasicaCriada""" &">" & strDesignacao_bandabasica & "</parameter> "
		    
		    Strxml = Strxml & 	"			<parameter name=" & """bloco""" &">" & strbloco & "</parameter> "
		    Strxml = Strxml & 	"			<parameter name=" & """cabo"""  &">" & strcabo & "</parameter> "
		    Strxml = Strxml & 	"			<parameter name=" & """par""" &">" & strpar & "</parameter> "
		    Strxml = Strxml & 	"			<parameter name=" & """pino""" &">" & strpino & "</parameter> "
		    
			Strxml = Strxml & 	"	</parameters> "
		    
		    'Response.Write "<script language=javascript>alert('Xml p6:');</script>"
		    
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
	        
		    if Alteracao_Cadastral = "S" then 
		      Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
			  Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse
			  Vetor_Campos(3)="adVarchar,15,adParamInput," & dblIdLogico
			  Vetor_Campos(4)="adVarchar,20,adParamInput," & objRSDadosInterf("ID_Tarefa_Apg")
			  Vetor_Campos(5)="adVarchar,20,adParamInput, " & dblSolid
			  Vetor_Campos(6)="adVarchar,20,adParamInput, 4 " 'Construir Return 
			  Vetor_Campos(7)="adVarchar,20,adParamInput, " & objRSDadosInterf("Processo")
			  Vetor_Campos(8)="adVarchar,20,adParamInput, " & objRSDadosInterf("Acao") 
			  strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",8,Vetor_Campos)
			  
			  Call db.Execute(strSqlRet)
		      
		    else
			  
			  Set doc = server.CreateObject("Microsoft.XMLDOM")
			  Set doc1 = server.CreateObject("Microsoft.XMLDOM")
			  Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		      
			  doc.async = False
			  doc1.async = False
			  
			  xmlhttp.Open "POST", AdresserPath, StrLogin, StrSenha
			  xmlhttp.setRequestHeader "SOAPAction", "executeClass"
		      
			  xmlhttp.send(Strxml)
			  
			  strRetorno = xmlhttp.ResponseText
			  
			  doc1.loadXML(strRetorno)
			  doc.loadXML(Strxml)
			  
			  'Checa se serviço é 0800.
			  Oe_numero = objRSDadosInterf("Oe_numero")
			  Oe_ano = objRSDadosInterf("Oe_ano")
			  Oe_item = objRSDadosInterf("Oe_item")
			  Id_logico = dblIdLogico
			  Processo = trim(objRSDadosInterf("Processo"))
			  Acao = trim(objRSDadosInterf("Acao"))
			  
			  call check_servico2(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,4,strxml)
			  
			  call check_servico2(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,4,strRetorno)
			  
			  Set xmlhttp= Nothing
			  Set myDoc= Nothing
		      
			  If doc1.parseError<>0 Then
			    strxmlResp = "Erro no XML retornado pelo APG: "
			    strxmlResp = strxmlResp  &  "<Codigo> " & doc1.parseError.errorCode & "</codigo>"
			    strxmlResp = strxmlResp  & 	"<Descricao>" & strErroXml & Trim(doc1.parseError.reason)
			    strxmlResp = strxmlResp  & 	Trim(doc1.parseError.line) & "</Descricao>"
                
		        Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
			    Vetor_Campos(2)="adInteger,6,adParamInput," & objRSDadosInterf("ID_Tarefa_Apg")	 'Identificação do APG
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
				    Vetor_Campos(4)="adWChar,20,adParamInput, " & objRSDadosInterf("ID_Tarefa_Apg")
				    Vetor_Campos(5)="adWChar,15,adParamInput,"& dblIdLogico
				    Vetor_Campos(6)="adWChar,10,adParamInput,"& dblSolid
				    Vetor_Campos(7)="adWChar,20,adParamInput, "
				    Vetor_Campos(8)="adInteger,4,adParamOutput,0"
				    Vetor_Campos(9)="adWChar,100,adParamOutput,"
                    
				    Call APENDA_PARAM("CLA_sp_ins_Construir_Acesso_Ret",9,Vetor_Campos)
                    
				    'Response.Write "<script language=javascript>alert('chamou Define par. Procedures')</script>"
				    'Response.Write "<script language=javascript>alert('Execucao:" & d &"')</script>"
                    
				    ObjCmd.Execute'pega dbaction
				    DBAction = ObjCmd.Parameters("RET").value
				    DBDescricao = ObjCmd.Parameters("RET1").value
                    
				    If DBAction <> 0 then 
					  'Response.Write "<script language=javascript>alert('Grava ret. interface')</script>"
					  strxmlResp = "Erro ao atualizar log da Interface - Codigo:" & DBAction & "Descrição:" & DBDescricao
					  'Response.Write "<script language=javascript>alert('" & strxmlResp &"')</script>"
                      
					  Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
					  Vetor_Campos(2)="adInteger,6,adParamInput," & objRSDadosInterf("ID_Tarefa_Apg")		 'Identificação do APG
					  Vetor_Campos(3)="addouble,10,adParamInput," & dblIdLogico			 'ID Logico
					  Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
					  Vetor_Campos(5)="adInteger,4,adParamOutput,0"
                      
					  Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)
                      
					  ObjCmd.Execute'pega dbaction
					  DBAction = ObjCmd.Parameters("RET").value
				    End If
				  End if
			    End if
			  End if
		    End if
		    
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
		  End If
	    End If
	  End If
    End if
	
End Function

Function check_servico2(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,interf,strxml)
	  Vetor_Campos(1)="adVarchar,7,adParamInput, " & Oe_numero
	  Vetor_Campos(2)="adInteger,10,adParamInput, " & Oe_ano
	  Vetor_Campos(3)="adInteger,10,adParamInput, " & Oe_item
	  Vetor_Campos(4)="adInteger,10,adParamInput, " & Id_logico
	  Vetor_Campos(5)="adVarchar,10,adParamInput, " & processo
	  Vetor_Campos(6)="adVarchar,10,adParamInput, " & acao
	  Vetor_Campos(7)="adInteger,10,adParamInput, " & Interf
	  Vetor_Campos(8)="adVarchar,7000,adParamInput, " & strxml
      
	  strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_check_Servico",8,Vetor_Campos)
	  Call db.Execute(strSqlRet)
	  
End function
%>