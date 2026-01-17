<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
Function EnviarReenvioAprov(dblAprovisiId,dblacao)

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
				oriSol_Descricao 	= "SGAV"
				strxmlResp 			= "<retorno>OK</retorno>"
				
				Strxml1 = ""
				Strxml2 = ""
				
				Strxml1 = Strxml1 & 	"	<requisicao-cla>" & vbnewline
				Strxml1 = Strxml1 & 	"		<acao>CAN</acao>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-logico>" & objRSDadosCla("acl_idacessologico") & "</id-logico>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-tarefa-can>" & objRSDadosCla("ID_Tarefa") & "</id-tarefa-can>" & vbnewline
				
				Strxml1 = Strxml1 & 	"		<observacao></observacao>" & vbnewline
				'Observacao = objXmlDadosForm.selectSingleNode("//requisicao-cla/observacao").text
				
				Strxml2 = Strxml2 & 	"	<requisicao-cla>" & vbnewline
				Strxml2 = Strxml2 & 	"		<acao>" & UCASE(dblacao) & "</acao>" & vbnewline
				if UCASE(dblacao) = "ATV" then
					Strxml2 = Strxml2 & 	"		<id-logico></id-logico>" & vbnewline
				else
					ID_Logico = ""
					if not isnull(objRSDadosCla("acl_idacessologico")) then
						ID_Logico = "678" & mid(objRSDadosCla("acl_idacessologico"),4,7)
					end if
					Strxml2 = Strxml2 & 	"		<id-logico>" & ID_Logico & "</id-logico>" & vbnewline
				end if
				
				Strxml = Strxml & 	"		<origem>SGAV</origem>" & vbnewline
				Strxml = Strxml & 	"		<senha>A3590023DF66AC92AE35E35E33160</senha>" & vbnewline
				Strxml = Strxml & 	"		<id-tarefa>" & objRSDadosCla("ID_Tarefa") & "</id-tarefa>" & vbnewline
				Strxml = Strxml & 	"		<cliente>" & vbnewline
				Strxml = Strxml & 	"			<conta>" & objRSDadosCla("Cli_CC") & "</conta>" & vbnewline
				Strxml = Strxml & 	"			<subconta>" & objRSDadosCla("Cli_SubCC") & "</subconta>" & vbnewline				
				'Strxml = Strxml & 	"			<razao-social><![CDATA[" & objRSDadosCla("Cli_nome") & "]]></razao-social>" & vbnewline
				'Strxml = Strxml & 	"			<nome-fantasia><![CDATA[" & objRSDadosCla("Cli_NomeFantasia") & "]]></nome-fantasia>" & vbnewline								
				Strxml = Strxml & 	"			<razao-social>" & replace(objRSDadosCla("Cli_nome"),"&","E") & "</razao-social>" & vbnewline
				Strxml = Strxml & 	"			<nome-fantasia>" & replace(objRSDadosCla("Cli_NomeFantasia"),"&","E") & "</nome-fantasia>" & vbnewline
				Strxml = Strxml & 	"		</cliente>" & vbnewline
				Strxml = Strxml & 	"		<servico>" & vbnewline
				Strxml = Strxml & 	"			<order-entry>" & vbnewline
				Strxml = Strxml & 	"				<ano>" & objRSDadosCla("OE_ANO") & "</ano>" & vbnewline
				Strxml = Strxml & 	"				<numero>" & objRSDadosCla("OE_NUMERO") & "</numero>" & vbnewline
				Strxml = Strxml & 	"				<item>" & objRSDadosCla("OE_ITEM") & "</item>" & vbnewline
				Strxml = Strxml & 	"				<servico>" & objRSDadosCla("Ser_Sigla") & "</servico>" & vbnewline
				Strxml = Strxml & 	"				<servico-desc>" & objRSDadosCla("Ser_Desc") & "</servico-desc>" & vbnewline
				Strxml = Strxml & 	"				<designacao>" & objRSDadosCla("Acl_DesignacaoServico") & "</designacao>" & vbnewline
				Strxml = Strxml & 	"				<contrato>" & objRSDadosCla("Acl_NContratoServico") & "</contrato>" & vbnewline
				Strxml = Strxml & 	"				<tipo-contrato>" & objRSDadosCla("Acl_TipoContratoServico") & "</tipo-contrato>" & vbnewline
				Strxml = Strxml & 	"				<data-prevista>" & objRSDadosCla("Acl_DtDesejadaEntregaAcessoServico") & "</data-prevista>" & vbnewline
				Strxml = Strxml & 	"				<reenvio>S</reenvio>" & vbnewline
				Strxml = Strxml & 	"				<desig-site>" & objRSDadosCla("Acl_desig_site") & "</desig-site>" & vbnewline
				Strxml = Strxml & 	"				<contato-tecnico>" & vbnewline
				Strxml = Strxml & 	"					<cliente>" & objRSDadosCla("Aec_Contato") & "</cliente>" & vbnewline
				Strxml = Strxml & 	"					<e-mail>" & objRSDadosCla("Aec_Email") & "</e-mail>" & vbnewline
				Strxml = Strxml & 	"					<telefone>" & objRSDadosCla("Aec_Telefone") & "</telefone>" & vbnewline
				Strxml = Strxml & 	"				</contato-tecnico>" & vbnewline
				Strxml = Strxml & 	"				<velocidade>" & objRSDadosCla("Vel_Desc") & "</velocidade>" & vbnewline
				Strxml = Strxml & 	"				<cadastrador>" & vbnewline
				Strxml = Strxml & 	"					<telefone>" & objRSDadosCla("TELEFONE_CADASTRADOR") & "</telefone>" & vbnewline
				Strxml = Strxml & 	"					<username>" & objRSDadosCla("USERNAME_CADASTRADOR") & "</username>" & vbnewline
				Strxml = Strxml & 	"				</cadastrador>" & vbnewline  
				Strxml = Strxml & 	"				<interface-cliente>" & objRSDadosCla("Interface_Cliente") & "</interface-cliente>" & vbnewline
				Strxml = Strxml & 	"				<interface-embratel>" & objRSDadosCla("Interface_Embratel") & "</interface-embratel>" & vbnewline
				Strxml = Strxml & 	"				<servico-temporario>" & vbnewline
				Strxml = Strxml & 	"					<data-inicio>" & objRSDadosCla("Acl_DtIniAcessoTemp") & "</data-inicio>" & vbnewline
				Strxml = Strxml & 	"					<data-fim>" & objRSDadosCla("Acl_DtFimAcessoTemp") & "</data-fim>" & vbnewline
				Strxml = Strxml & 	"				</servico-temporario>" & vbnewline    
				Strxml = Strxml & 	"				<codDescargaSAP>" & objRSDadosCla("esc_cod_sap") & "</codDescargaSAP>" & vbnewline
				Strxml = Strxml & 	"				<estCliente>" & objRSDadosCla("Aec_SiglaCentroCliente") & "</estCliente>" & vbnewline
				Strxml = Strxml & 	"				<endereco-instalacao>" & vbnewline
				Strxml = Strxml & 	"					<cnpj>" & objRSDadosCla("Aec_CNPJ") & "</cnpj>" & vbnewline 
				Strxml = Strxml & 	"					<inscricao-estadual>" & objRSDadosCla("Aec_IE") & "</inscricao-estadual>" & vbnewline
				Strxml = Strxml & 	"					<inscricao-municipal>" & objRSDadosCla("Aec_IM") & "</inscricao-municipal>" & vbnewline
				Strxml = Strxml & 	"					<cnl>" & objRSDadosCla("CNL") &" </cnl>" & vbnewline				
				'Strxml = Strxml & 	"					<proprietario><![CDATA[" & objRSDadosCla("Aec_PropEnd") & "]]></proprietario>" & vbnewline
				Strxml = Strxml & 	"					<proprietario>" & replace(objRSDadosCla("Aec_PropEnd"),"&","E") & "</proprietario>" & vbnewline
				Strxml = Strxml & 	"					<bairro>" & objRSDadosCla("End_Bairro") & " </bairro>" & vbnewline
				Strxml = Strxml & 	"					<cep>" & objRSDadosCla("End_CEP") & "</cep>" & vbnewline
				Strxml = Strxml & 	"					<cidade>" & objRSDadosCla("CIDADE") & "</cidade>" & vbnewline
				Strxml = Strxml & 	"					<complemento>" & objRSDadosCla("Aec_Complemento") & "</complemento>" & vbnewline
				Strxml = Strxml & 	"					<logradouro>" & objRSDadosCla("End_NomeLogr") & "</logradouro>" & vbnewline
				Strxml = Strxml & 	"					<numero>" & objRSDadosCla("End_NroLogr") & "</numero>" & vbnewline
				Strxml = Strxml & 	"					<tipo-logradouro>" & objRSDadosCla("Tpl_Sigla") & "</tipo-logradouro>" & vbnewline
				Strxml = Strxml & 	"					<uf>" & objRSDadosCla("Est_Sigla") & "</uf>" & vbnewline
				Strxml = Strxml & 	"					<centro-cliente>" & objRSDadosCla("Aec_SiglaCentroCliente") & "</centro-cliente>" & vbnewline
				Strxml = Strxml & 	"				</endereco-instalacao>" & vbnewline
				Strxml = Strxml & 	"				<numero-sev>" & objRSDadosCla("Sol_SevSeq") & "</numero-sev>" & vbnewline
				Strxml = Strxml & 	"				<numero-ots>" & vbnewline
				Strxml = Strxml & 	"					<ots></ots>" & vbnewline
				Strxml = Strxml & 	"				</numero-ots>" & vbnewline
				Strxml = Strxml & 	"				<observacao></observacao>" & vbnewline
				Strxml = Strxml & 	"				<vel_voz>" & objRSDadosCla("Vel_Voz") & "</vel_voz>" & vbnewline
				Strxml = Strxml & 	"				<dados>" & objRSDadosCla("Dados") & "</dados>" & vbnewline						
				Strxml = Strxml & 	"			</order-entry>" & vbnewline
				Strxml = Strxml & 	"		</servico>" & vbnewline
				Strxml = Strxml & 	"	</requisicao-cla>" & vbnewline
				
				Strxml1 = Strxml1 & Strxml
				Strxml2 = Strxml2 & Strxml 


 
				sUrl = "http://localhost/newcla/ACCESS_INTERF/Asp/Aprov_Solicitar_Acesso.asp"
				
				'sUrl = "http://2k3rjoapph001/newcla_dsv/ACCESS_INTERF/Asp/Aprov_Solicitar_Acesso.asp"
				
				StrLogin = strLoginRede
				
				Set doc = server.CreateObject("Microsoft.XMLDOM")
				Set doc1 = server.CreateObject("Microsoft.XMLDOM")
				doc.loadXml(Strxml1)
				
				'Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
				Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
				
				xmlhttp.Open "POST", sUrl, False
				xmlhttp.Send Strxml1
				
				strRetorno = xmlhttp.ResponseText
				
				doc1.loadXML(strRetorno)
				'response.write xmlhttp.responseText
				
				cod_retorno 	= trim(doc1.selectSingleNode("//resposta-cla/codigo").text)
				desc_retorno    = trim(doc1.selectSingleNode("//resposta-cla/mensagem").text)

				if cod_retorno = 0 then
				%>
				  <script language="VBScript">MsgBox "<%=desc_retorno%>",64,"Informação"</script>
				<%
				else
				  Response.Write "<script>alert('" & desc_retorno & "')</script>"
				end if
				
				if cod_retorno = 0 then
					doc.loadXml(Strxml2)
					
					xmlhttp.Open "POST", sUrl, False
					xmlhttp.Send Strxml2
					
					strRetorno = xmlhttp.ResponseText
					
					doc1.loadXML(strRetorno)
					
					cod_retorno 	= trim(doc1.selectSingleNode("//resposta-cla/codigo").text)
					desc_retorno    = trim(doc1.selectSingleNode("//resposta-cla/mensagem").text)
					
		  		if cod_retorno = 0 then
			  	%>
			  	  <script language="VBScript">MsgBox "<%=desc_retorno%>",64,"Informação"</script>
			  	<%
			  	else
				    Response.Write "<script>alert('" & desc_retorno & "')</script>"
			  	end if
									
			     end if
		
		end if
		   
	 end if
	 	   
End Function
%>