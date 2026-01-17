<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarReenvioAprovAsms(dblAprovisiId,dblacao)

		Dim proprietario ' TER ou EBT
		Dim tecnologia ' 0 - terceiro 1 - radio 2 - fibra otica 3 -ade 4 - satelite 5 - cabo interno 
		Dim idTarefa ' identificador do sistema aprovisionador
		Dim oe_numero 		 
		Dim	oe_ano			 
		Dim	oe_item			 
		Dim	idLogico		 
		Dim OrigemSolicitacao
		Dim rede
		Dim oriSol_id
		Dim Acao
		
		'response.write "<script>alert('"&dblAprovisiId&"')</script>"
		'response.write "<script>alert('"&dblacao&"')</script>"
		'response.end

   		if dblAprovisiId <> "" then

			Set objRSDadosCla = db.execute("CLA_sp_sel_Aprovisionador " & dblAprovisiId)
			
			'Set objRSDadosCla = db.Execute(strSqlRet)
			
			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			
				Aprovisi_ID			= objRSDadosCla("Aprovisi_ID")
				idTarefa 			= objRSDadosCla("id_tarefa")
				oe_numero 			= objRSDadosCla("oe_numero")
				oe_ano				= objRSDadosCla("oe_ano")
				oe_item				= objRSDadosCla("oe_item")
				idLogico			= objRSDadosCla("acl_idacessologico")
				acao				= "CAN"
				OrigemSolicitacao	= objRSDadosCla("oriSol_Descricao")
				oriSol_id			= objRSDadosCla("oriSol_ID")
				strxmlResp 			= "<retorno>OK</retorno>"
				
				'response.write "<script>alert('"&dblAprovisiId&"')</script>"
				Strxml1 = ""
				Strxml2 = ""
				Strxml  = ""
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & dblAprovisiId
				strSqlHeader = APENDA_PARAMSTR("CLA_sp_sel_headerAsms",1,Vetor_Campos)
				
				Set objRSDadosHeader = db.Execute(strSqlHeader) 
				
				header = ""
				external = ""
				
				If not objRSDadosHeader.Eof and  not objRSDadosHeader.Bof Then
						header 			= trim(objRSDadosHeader("header_desc"))
						external 		= trim(objRSDadosHeader("externalRef_desc")	)					
				end if 
				'response.write "<script>alert('"&trim(objRSDadosHeader("headerAsms_Id"))&"')</script>"
				'response.end
				
				
				Strxml1 = Strxml1 & 	"	<requisicao-cla> " & vbnewline  
				Strxml1 = Strxml1 & 	"		<acao>CAN</acao> " & vbnewline  
				Strxml1 = Strxml1 & 	"		<id-logico>" & objRSDadosCla("acl_idacessologico") & "</id-logico> " & vbnewline
				Strxml1 = Strxml1 & 	"		<id-logico-tmp>" & objRSDadosCla("acl_idacessologico") & "</id-logico-tmp>" & vbnewline 
				Strxml1 = Strxml1 & 	"		<id-tarefa-can>" & objRSDadosCla("ID_Tarefa") & "</id-tarefa-can>" & vbnewline
				
				Strxml2 = Strxml2 & 	"	<requisicao-cla>" & vbnewline
				Strxml2 = Strxml2 & 	"		<acao>" & UCASE(dblacao) & "</acao>" & vbnewline
				if UCASE(dblacao) = "ATV" then
					Strxml2 = Strxml2 & 	"		<id-logico></id-logico>" & vbnewline
					Strxml2 = Strxml2 & 	"		<id-logico-tmp></id-logico-tmp>" & vbnewline
				else
					ID_Logico = ""
					if not isnull(objRSDadosCla("acl_idacessologico")) then
						ID_Logico = "678" & mid(objRSDadosCla("acl_idacessologico"),4,7)
					end if
					Strxml2 = Strxml2 & 	"		<id-logico>" & ID_Logico & "</id-logico>" & vbnewline
					Strxml2 = Strxml2 & 	"		<id-logico-tmp>" & ID_Logico & "</id-logico-tmp>" & vbnewline
				end if
				
				'response.write "<script>alert('"&objRSDadosCla("acl_idacessologico")&"')</script>"
				Strxml = Strxml & 	"	<origem>ASMS\REENVIO</origem>" & vbnewline
				Strxml = Strxml & 	header  	& vbnewline  				
				Strxml = Strxml & 	external 	& vbnewline				
				Strxml = Strxml & 	"	<ind-provide-cease>" & objRSDadosCla("Indicador_Alterar") & "</ind-provide-cease>" & vbnewline   
				Strxml = Strxml & 	"	<id-acesso>" & objRSDadosCla("Id_Acesso") & "</id-acesso> " & vbnewline  
				Strxml = Strxml & 	"	<id-tarefa>" & objRSDadosCla("ID_Tarefa") & "</id-tarefa>" & vbnewline   
				Strxml = Strxml & 	"	<numero>" & objRSDadosCla("OE_NUMERO") & "</numero>" & vbnewline   
				Strxml = Strxml & 	"	<item>" & objRSDadosCla("OE_ITEM") & "</item>" & vbnewline   
				Strxml = Strxml & 	"	<id-servico>" & objRSDadosCla("Id_Servico") & "</id-servico> " & vbnewline  
				Strxml = Strxml & 	"	<cliente> " & vbnewline   
				Strxml = Strxml & 	"		<conta>" & objRSDadosCla("Cli_CC") & "</conta>" & vbnewline    
				Strxml = Strxml & 	"		<subconta>" & objRSDadosCla("Cli_SubCC") & "</subconta>" & vbnewline    
				Strxml = Strxml & 	"		<razao-social>" & objRSDadosCla("Cli_Nome") & "</razao-social>" & vbnewline    
				Strxml = Strxml & 	"		<nome-fantasia>" & objRSDadosCla("Cli_NomeFantasia") & "</nome-fantasia>" & vbnewline   
				Strxml = Strxml & 	"	</cliente>" & vbnewline   
				Strxml = Strxml & 	"	<servico> " & vbnewline   
				Strxml = Strxml & 	"		<order-entry>" & vbnewline     
				Strxml = Strxml & 	"			<servico>" & objRSDadosCla("Ser_Sigla") & "</servico>" & vbnewline     
				Strxml = Strxml & 	"			<servico-desc>" & objRSDadosCla("Ser_Desc") & "</servico-desc> " & vbnewline    
				Strxml = Strxml & 	"			<designacao>" & objRSDadosCla("Acl_DesignacaoServico") & "</designacao>" & vbnewline     
				Strxml = Strxml & 	"			<contato-tecnico>" & vbnewline      
				Strxml = Strxml & 	"				<cliente>" & objRSDadosCla("Aec_Contato") & "</cliente>  " & vbnewline    
				Strxml = Strxml & 	"				<e-mail>" & objRSDadosCla("Aec_Email") & "</e-mail>" & vbnewline      
				Strxml = Strxml & 	"				<telefone>" & objRSDadosCla("Aec_Telefone") & "</telefone>" & vbnewline   
				Strxml = Strxml & 	"			</contato-tecnico>" & vbnewline     
				Strxml = Strxml & 	"			<velocidade>" & objRSDadosCla("Vel_Desc") & "</velocidade>" & vbnewline     
				Strxml = Strxml & 	"			<velocidade-total>" & objRSDadosCla("Velocidade_Total") & "</velocidade-total>" & vbnewline     
				Strxml = Strxml & 	"			<codDescargaSAP>" & objRSDadosCla("esc_cod_sap") & "</codDescargaSAP> " & vbnewline    
				Strxml = Strxml & 	"			<estCliente>" & objRSDadosCla("Aec_SiglaCentroCliente") & "</estCliente>" & vbnewline     
				Strxml = Strxml & 	"			<endereco-instalacao> " & vbnewline      
				Strxml = Strxml & 	"				<cnpj>" & objRSDadosCla("Aec_CNPJ") & "</cnpj>" & vbnewline       
				Strxml = Strxml & 	"				<inscricao-estadual>" & objRSDadosCla("Aec_IE") & "</inscricao-estadual> " & vbnewline      
				Strxml = Strxml & 	"				<inscricao-municipal>" & objRSDadosCla("Aec_IM") & "</inscricao-municipal> " & vbnewline      
				Strxml = Strxml & 	"				<cnl>" & objRSDadosCla("CNL") & "</cnl>" & vbnewline       
				Strxml = Strxml & 	"				<bairro>" & objRSDadosCla("End_Bairro") & "</bairro> " & vbnewline      
				Strxml = Strxml & 	"				<cep>" & objRSDadosCla("End_CEP") & "</cep>" & vbnewline       
				Strxml = Strxml & 	"				<cidade>" & objRSDadosCla("CIDADE") & "</cidade>" & vbnewline       
				Strxml = Strxml & 	"				<complemento>" & objRSDadosCla("Aec_Complemento") & "</complemento> " & vbnewline     
				Strxml = Strxml & 	"				<logradouro>" & objRSDadosCla("End_NomeLogr") & "</logradouro>" & vbnewline      
				Strxml = Strxml & 	"				<numero>" & objRSDadosCla("End_NroLogr") & "</numero> " & vbnewline     
				Strxml = Strxml & 	"				<tipo-logradouro>" & objRSDadosCla("Tpl_Sigla") & "</tipo-logradouro>" & vbnewline      
				Strxml = Strxml & 	"				<uf>" & objRSDadosCla("Est_Sigla") & "</uf> " & vbnewline    
				Strxml = Strxml & 	"			</endereco-instalacao> " & vbnewline    
				Strxml = Strxml & 	"			<numero-sev>" & objRSDadosCla("Sol_SevSeq") & "</numero-sev>  " & vbnewline  
				Strxml = Strxml & 	"		</order-entry>" & vbnewline   
				Strxml = Strxml & 	"	</servico>" & vbnewline  
				Strxml = Strxml & 	"</requisicao-cla> " & vbnewline 
				
				Strxml1 = Strxml1 & Strxml
				Strxml2 = Strxml2 & Strxml
				
				
				
				'sUrl = "http://localhost/newcla_interf/Asp/Aprov_Solicitar_Acesso_ASMS_Reenvio.asp"
				
				'sUrl = "http://2k3rjoapph001/newcla_interf/Asp/Aprov_Solicitar_Acesso_ASMS.asp"
				
				sUrl = "http://localhost/newcla/ACCESS_INTERF/Asp/Aprov_Solicitar_Acesso_ASMS.asp"
				
				'response.write "<script>alert('"&sUrl&"')</script>"
				'response.end
				
				StrLogin = strLoginRede
				
				Set doc = server.CreateObject("Microsoft.XMLDOM")
				Set doc1 = server.CreateObject("Microsoft.XMLDOM")
				doc.loadXml(Strxml1)
				'doc.save(Server.MapPath("../TesteConectividade-RealidadeIP_CAN.xml"))
				
				Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
				Set xmlhttp = CreateObject("Msxml2.XMLHTTP")
				
				''
				' Estava 
				'			Aqui
				''
				xmlhttp.Open "POST", sUrl, False
				xmlhttp.Send Strxml1
				
				strRetorno = xmlhttp.ResponseText
				
				doc1.loadXML(strRetorno)
				
				cod_retorno 	= trim(doc1.selectSingleNode("//resposta-cla/codigo").text)
				desc_retorno    = trim(doc1.selectSingleNode("//resposta-cla/mensagem").text)
				
				'response.write "<script>alert('"&cod_retorno&"')</script>"
				'response.end
				
				if cod_retorno = 0 then
					doc.loadXml(Strxml2)
					'doc.save(Server.MapPath("../TesteConectividade-RealidadeIP_ATV.xml"))
					xmlhttp.Open "POST", sUrl, False
					xmlhttp.Send Strxml2
					
					strRetorno = xmlhttp.ResponseText
					
					doc1.loadXML(strRetorno)
					
					cod_retorno 	= trim(doc1.selectSingleNode("//resposta-cla/codigo").text)
					desc_retorno    = trim(doc1.selectSingleNode("//resposta-cla/mensagem").text)
					
					'response.write "<script>alert('"&desc_retorno&"')</script>"
			     end if
		
		end if
		   
	 end if
	 	   
End Function
%>