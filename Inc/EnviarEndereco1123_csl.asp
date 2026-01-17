<%
				
Function EnviarEndereco1123_CSL()

		Dim proprietario ' TER ou EBT
		Dim tecnologia ' 0 - terceiro 1 - radio 2 - fibra otica 3 -ade 4 - satelite 5 - cabo interno 
		Dim idTarefa ' identificador do sistema aprovisionador
		Dim oe_numero 		 
		Dim	oe_ano			 
		Dim	oe_item			 
		Dim	idLogico		 
		Dim	acao			 
		Dim rede
		Dim oriSol_id
		Dim header
		Dim external
		
		End_id 							= ""
		Aec_id 							= ""
		cod_retorno_pesquisa			= ""
		cod_retorno_EBT					= ""
		cod_retorno_valida				= ""
		des_msg							= ""
		des_tratamnt_retorno			= ""
		sgl_sis_emissor					= ""
		qtd_regis_recuperado			= ""
		des_complemnt					= ""
		num_endereco					= ""
		des_endereco					= ""
	
		num_inicio_CEP					= ""
		num_fim_CEP						= ""
		num_CEP							= ""
		ind_tipo_num_CEP				= ""
		num_CEP_alternativo				= ""
		sgl_tipo_lograd					= ""
		ind_tipo_categoria_lograd		= ""
		des_titulo_lograd				= ""
		des_lograd						= ""
		des_titulo_nome_lograd			= ""
		num_inicio_endereco				= ""
		num_fim_endereco				= ""
		des_bairro						= ""
		ind_lado_rua					= ""
		cod_lograd_pesquisa				= ""
		ind_tipo_localid				= ""
		des_localid						= ""
		des_uf							= ""
		
		cod_localid						= ""
		cod_distrito					= ""
		cod_regiao						= ""
		cod_pais						= ""
		nom_pais						= ""
		Interf_Desc						= ""
		Interf_xml						= ""
		Interf_num						= ""
		Interf_erro						= ""
		
		ret 							= "" 
		
		Set objRS = db.execute("CSL_sp_sel_CleanUP ")

		if Not objRS.Eof and not objRS.Bof then
			intCount = 0

			While Not objRS.Eof
			    ''''-----'''''
						
						intCount = intCount + 1 
						intCount = 1000000 + intCount
				
						Strxml1 = "" 
				
						oriSol_Descricao 	= "MPE-1123"
						End_id 				= objRS("End_id")
						Aec_id 				= objRS("Aec_id")
						'Aprovisi_ID         = 10
						
						Strxml1 = Strxml1 & 	"<?xml version=""1.0"" encoding=""UTF-8"" ?> " & vbnewline
						
						
						Strxml1 = Strxml1 & 	"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns0=""http://www.tibco.com/schemas/BRAMEX/eAI/SharedResouces/SchemaDefinition/Bramex/schPesquisarEndereco.xsd""> " & vbnewline
			  			Strxml1 = Strxml1 & 	" <soapenv:Header/> " & vbnewline
			  			Strxml1 = Strxml1 & 	" <soapenv:Body> " & vbnewline
						
						Strxml1 = Strxml1 & 	"<ns0:Data > " & vbnewline
						
			 			Strxml1 = Strxml1 & 	"	<ns0:Header> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strSystemTrackingId>"&intCount&"</ns0:strSystemTrackingId> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strResubmitFlag /> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strOriginalTrackingId /> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strSystem>CLA</ns0:strSystem> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strFlagSA>S</ns0:strFlagSA> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strTimeStamp /> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strTimeOut /> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:strEventType>1123</ns0:strEventType> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Header> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Solicitacao>" & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:idt_solic>20150427170450926060</ns0:idt_solic> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:dta_hra_solic>20150427170852</ns0:dta_hra_solic> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:ind_solic_valida_endereco>1</ns0:ind_solic_valida_endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:ind_solic_ajuste_escrita>1</ns0:ind_solic_ajuste_escrita> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:Consulta_Endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_endereco>"& objRS("endereco") &"</ns0:des_endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Consulta_CEP> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_bairro>"& objRS("END_BAIRRO") &"</ns0:des_bairro> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:num_CEP>"& objRS("END_CEP") &"</ns0:num_CEP> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Consulta_CNL> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_localid>"& objRS("CID_DESC") &"</ns0:des_localid> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_uf>"& objRS("EST_SIGLA") &"</ns0:des_uf> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:cod_localid>0</ns0:cod_localid> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Consulta_CNL> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Consulta_CEP> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Consulta_Pais> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:nom_pais>BRASIL</ns0:nom_pais> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Consulta_Pais> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Consulta_Endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	  </ns0:Solicitacao> " & vbnewline
						Strxml1 = Strxml1 & 	" </ns0:Data> " & vbnewline
						
						Strxml1 = Strxml1 & 	" </soapenv:Body> " & vbnewline
						Strxml1 = Strxml1 & 	" </soapenv:Envelope> " & vbnewline
						
						'response.write "<script>alert('"&Strxml1&"')</script>"
						
						'Set doc 	= server.CreateObject("Microsoft.XMLDOM")
					    'Set doc1 	= server.CreateObject("Microsoft.XMLDOM")
					  	'Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
				      
					  '  doc.async = False
					  '  doc1.async = False
						
						
						Set doc = server.CreateObject("Microsoft.XMLDOM")
					    Set doc1 = server.CreateObject("Microsoft.XMLDOM")
					 
						Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
						
						
				      
					    'doc.async = False
					  '  doc1.async = False
					   
					  	'xmlhttp.Open "POST", "http://nthor028:9333/barramento/services/SolicitarAcesso", false
						
						
						'xmlhttp.Open "POST", "http://10.2.7.18:9005/barramento/services/Endereco " , False , "teste", "teste"
						'''xmlhttp.Open "POST", "http://10.2.7.18:9005/barramento/services/Endereco " , False , "teste", "teste"
						
						''PRODUCAO
						xmlhttp.Open "POST", "http://10.2.9.69:9005/barramento/services/Endereco " , False , "teste", "teste"
						
						xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"  
				
						xmlhttp.setRequestHeader "SOAPAction", "/PesquisarEndereco"
					      
						xmlhttp.send(Strxml1)
						  
						'response.write "<script>alert('"&xmlhttp.statusText&"')</script>"	
						strRetorno = xmlhttp.ResponseText
						  
						doc1.loadXML(strRetorno)
						doc.loadXML(Strxml1)
						
						'doc1.save(Server.MapPath("1123-out.xml"))
						
						
						'response.write "<script>alert('"&strRetorno&"')</script>"	
						'objNodeList = Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa").text
							
						'response.write "<script>alert('"&objNodeList&"')</script>"
						'response.end
						
						if xmlhttp.statusText = "OK" then
						
														
							Set objNodecod_retorno_pesquisa = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa")
							if trim(objNodecod_retorno_pesquisa.Length) <> "0" then
								cod_retorno_pesquisa			= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa").text
							end if 
							
							Set objNodecod_retorno_EBT = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_EBT")
							if trim(objNodecod_retorno_EBT.Length) <> "0" then
								cod_retorno_EBT					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_EBT").text
							end if
							
							Set objNodecod_retorno_valida = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_valida")
							if trim(objNodecod_retorno_valida.Length) <> "0" then
								cod_retorno_valida				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_valida").text
							end if 
							
							Set objNodedes_msg = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg")
							if trim(objNodedes_msg.Length) <> "0" then
								des_msg							= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg").text
							end if 
							
							Set objNodedes_tratamnt_retorno = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_tratamnt_retorno")
							if trim(objNodedes_tratamnt_retorno.Length) <> "0" then
								des_tratamnt_retorno			= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_tratamnt_retorno").text
							end if
							
							Set objNodesgl_sis_emissor = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:sgl_sis_emissor")
							if trim(objNodesgl_sis_emissor.Length) <> "0" then
								sgl_sis_emissor					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:sgl_sis_emissor").text
							end if 
							
							Set objNodeqtd_regis_recuperado = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:qtd_regis_recuperado")
							if trim(objNodeqtd_regis_recuperado.Length) <> "0" then
								qtd_regis_recuperado			= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:qtd_regis_recuperado").text
							end if 
							
							' ENDEREÇO
							Set objNodedes_complemnt = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:des_complemnt")
							if trim(objNodedes_complemnt.Length) <> "0" then
								des_complemnt					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:des_complemnt").text
							end if 
							
							Set objNodedes_complemnt = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:des_complemnt")
							if trim(objNodeqtd_regis_recuperado.Length) <> "0" then
								Set objNodenum_endereco = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_endereco")
							end if 
							
							Set objNodenum_endereco = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_endereco")
							if trim(objNodenum_endereco.Length) <> "0" then
								num_endereco					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_endereco").text
							end if 
							Set objNodedes_endereco = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:des_endereco")
							if trim(objNodedes_endereco.Length) <> "0" then
								des_endereco					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:des_endereco").text
							end if
							 
							' Faixa CEP
							Set objNodenum_inicio_CEP = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Faixa_CEP/ns0:num_inicio_CEP")
							if trim(objNodenum_inicio_CEP.Length) <> "0" then
								num_inicio_CEP					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Faixa_CEP/ns0:num_inicio_CEP").text
							end if 
							
							Set objNodenum_fim_CEP = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Faixa_CEP/ns0:num_fim_CEP")
							if trim(objNodenum_fim_CEP.Length) <> "0" then
								num_fim_CEP						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Faixa_CEP/ns0:num_fim_CEP").text
							end if 
							'CEP
							Set objNodenum_CEP = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_CEP")
							if trim(objNodenum_CEP.Length) <> "0" then
								num_CEP							= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_CEP").text
							end if 
							
							Set objNodeind_tipo_num_CEP = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_tipo_num_CEP")
							if trim(objNodeind_tipo_num_CEP.Length) <> "0" then
								ind_tipo_num_CEP				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_tipo_num_CEP").text
							end if 
							
							Set objNodenum_CEP_alternativo = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_CEP_alternativo")
							if trim(objNodenum_CEP_alternativo.Length) <> "0" then
								num_CEP_alternativo				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_CEP_alternativo").text
							end if 
							
							Set objNodesgl_tipo_lograd = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:sgl_tipo_lograd")
							if trim(objNodesgl_tipo_lograd.Length) <> "0" then
								sgl_tipo_lograd					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:sgl_tipo_lograd").text
							end if 
							
							Set objNodeind_tipo_categoria_lograd = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_tipo_categoria_lograd")
							if trim(objNodeind_tipo_categoria_lograd.Length) <> "0" then
								ind_tipo_categoria_lograd		= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_tipo_categoria_lograd").text
							end if 
							
							Set objNodedes_titulo_lograd = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_titulo_lograd")
							if trim(objNodedes_titulo_lograd.Length) <> "0" then
								des_titulo_lograd				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_titulo_lograd").text
							end if 
							
							Set objNodedes_lograd = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_lograd")
							if trim(objNodedes_lograd.Length) <> "0" then
								des_lograd						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_lograd").text
							end if 
							
							Set objNodedes_titulo_nome_lograd = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_titulo_nome_lograd")
							if trim(objNodedes_titulo_nome_lograd.Length) <> "0" then
								des_titulo_nome_lograd			= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_titulo_nome_lograd").text
							end if 
							
							Set objNodenum_inicio_endereco = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_inicio_endereco")
							if trim(objNodenum_inicio_endereco.Length) <> "0" then
								num_inicio_endereco				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_inicio_endereco").text
							end if 
							
							Set objNodenum_fim_endereco = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_fim_endereco")
							if trim(objNodenum_fim_endereco.Length) <> "0" then
								num_fim_endereco				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:num_fim_endereco").text
							end if 
							
							Set objNodedes_bairro = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_bairro")
							if trim(objNodedes_bairro.Length) <> "0" then
								des_bairro						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:des_bairro").text
							end if 
							
							Set objNodeind_lado_rua = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_lado_rua")
							if trim(objNodeind_lado_rua.Length) <> "0" then
								ind_lado_rua					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:ind_lado_rua").text
							end if 
							
							Set objNodecod_lograd_pesquisa = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:cod_lograd_pesquisa")
							if trim(objNodecod_lograd_pesquisa.Length) <> "0" then
								cod_lograd_pesquisa				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:cod_lograd_pesquisa").text
							end if 
							
							'CNL
							Set objNodeind_tipo_localid = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:ind_tipo_localid")
							if trim(objNodeind_tipo_localid.Length) <> "0" then
								ind_tipo_localid				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:ind_tipo_localid").text
							end if 
							
							Set objNodedes_localid = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:des_localid")
							if trim(objNodedes_localid.Length) <> "0" then
								des_localid						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:des_localid").text
							end if 
							
							Set objNodedes_uf = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:des_uf")
							if trim(objNodedes_uf.Length) <> "0" then
								des_uf							= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:des_uf").text
							end if 
							
							Set objNodecod_localid = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_localid")
							if trim(objNodecod_localid.Length) <> "0" then
								cod_localid						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_localid").text
							end if 
							
							Set objNodecod_distrito = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_distrito")
							if trim(objNodecod_distrito.Length) <> "0" then
								cod_distrito					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_distrito").text
							end if 
							
							Set objNodecod_regiao = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_regiao")
							if trim(objNodecod_regiao.Length) <> "0" then
								cod_regiao						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:CEP/ns0:CNL/ns0:cod_regiao").text
							end if 
							
							'País
							Set objNodecod_pais = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Pais/ns0:cod_pais")
							if trim(objNodecod_pais.Length) <> "0" then
								cod_pais						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Pais/ns0:cod_pais").text
							end if 
							
							Set objNodenom_pais = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Pais/ns0:nom_pais")
							if trim(objNodenom_pais.Length) <> "0" then
								nom_pais						= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:Pais/ns0:nom_pais").text
							end if 
							
							Interf_Desc						= xmlhttp.statusText
							Interf_xml						= strRetorno
							Interf_num						= "1"
							Interf_erro						= "0"
							
							ret 							= ""
							
							
							
							
							'Checa se serviço é 0800.
							Vetor_Campos(1)="adInteger,4,adParamInput,null" 	
							Vetor_Campos(2)="adInteger,1,adParamInput," 	& End_id 			
							Vetor_Campos(3)="adInteger,3,adParamInput," 	& Aec_id
							Vetor_Campos(4)="adVarchar,20,adParamInput," 	& cod_retorno_pesquisa
							Vetor_Campos(5)="adVarchar,2,adParamInput," 	& cod_retorno_EBT
							Vetor_Campos(6)="adVarchar,2,adParamInput," 	& cod_retorno_valida		
							Vetor_Campos(7)="adVarchar,500,adParamInput," 	& des_msg
							Vetor_Campos(8)="adVarchar,500,adParamInput, " 	& des_tratamnt_retorn
							Vetor_Campos(9)="adVarchar,6,adParamInput," 	& sgl_sis_emissor
							Vetor_Campos(10)="adInteger,1,adParamInput,"  	& qtd_regis_recuperado		
							Vetor_Campos(11)="adVarchar,120,adParamInput,"  & des_complemnt
							Vetor_Campos(12)="adVarchar,5,adParamInput," 	& num_endereco
							
							Vetor_Campos(13)="adVarchar,120,adParamInput," 	& des_endereco
							Vetor_Campos(14)="adVarchar,8,adParamInput," 	& num_inicio_CEP
							Vetor_Campos(15)="adVarchar,8,adParamInput," 	& num_fim_CEP
							Vetor_Campos(16)="adVarchar,8,adParamInput," 	& num_CEP
							Vetor_Campos(17)="adVarchar,2,adParamInput," 	& ind_tipo_num_CEP
							Vetor_Campos(18)="adVarchar,8,adParamInput," 	& num_CEP_alternativo
							Vetor_Campos(19)="adVarchar,10,adParamInput," 	& sgl_tipo_lograd
							Vetor_Campos(20)="adVarchar,2,adParamInput, " 	& ind_tipo_categoria_lograd
							Vetor_Campos(21)="adVarchar,15,adParamInput," 	& des_titulo_lograd 
							Vetor_Campos(22)="adVarchar,60,adParamInput,"   & des_lograd
							Vetor_Campos(23)="adVarchar,60,adParamInput,"   & des_titulo_nome_lograd
							Vetor_Campos(24)="adVarchar,5,adParamInput,"    & num_inicio_endereco
							
							Vetor_Campos(25)="adVarchar,5,adParamInput," 	& num_fim_endereco
							Vetor_Campos(26)="adVarchar,120,adParamInput," 	& des_bairro
							Vetor_Campos(27)="adVarchar,2,adParamInput," 	& ind_lado_rua
							Vetor_Campos(28)="adNumeric,10,adParamInput," 	& cod_lograd_pesquisa
							Vetor_Campos(29)="adVarchar,1,adParamInput," 	& ind_tipo_localid
							Vetor_Campos(30)="adVarchar,20,adParamInput," 	& des_localid
							Vetor_Campos(31)="adVarchar,2,adParamInput," 	& des_uf
							Vetor_Campos(32)="adNumeric,10,adParamInput, " 	& cod_localid
							Vetor_Campos(33)="adNumeric,10,adParamInput," 	& cod_distrito
							Vetor_Campos(34)="adNumeric,10,adParamInput,"   & cod_regiao
							Vetor_Campos(35)="adVarchar,5,adParamInput,"   & cod_pais
							Vetor_Campos(36)="adVarchar,70,adParamInput,"   & nom_pais
							
							Vetor_Campos(37)="adVarchar,200,adParamInput, " 	& xmlhttp.statusText 
							Vetor_Campos(38)="adVarchar,8000,adParamInput," 	& strRetorno 
							Vetor_Campos(39)="adInteger,1,adParamInput,"    & Interf_num 
							Vetor_Campos(40)="adInteger,1,adParamInput,"    & Interf_erro
							Vetor_Campos(41)="adInteger,2,adParamOutput,0" 
							
							strSqlRet = APENDA_PARAMSTR("CSL_sp_check_servico",41,Vetor_Campos)
							db.Execute(strSqlRet)
						else
							strxmlResp = strRetorno
							
							if strxmlResp =  "" then
								strxmlResp = "1"
							end if 
							
							'Checa se serviço é 0800.
							Vetor_Campos(1)="adInteger,4,adParamInput,null" 	
							Vetor_Campos(2)="adInteger,1,adParamInput," 	& End_id 			
							Vetor_Campos(3)="adInteger,3,adParamInput," 	& Aec_id
							Vetor_Campos(4)="adVarchar,20,adParamInput," 	& cod_retorno_pesquisa
							Vetor_Campos(5)="adVarchar,2,adParamInput," 	& cod_retorno_EBT
							Vetor_Campos(6)="adVarchar,2,adParamInput," 	& cod_retorno_valida		
							Vetor_Campos(7)="adVarchar,500,adParamInput," 	& des_msg
							Vetor_Campos(8)="adVarchar,500,adParamInput, " 	& des_tratamnt_retorn
							Vetor_Campos(9)="adVarchar,6,adParamInput," 	& sgl_sis_emissor
							Vetor_Campos(10)="adInteger,1,adParamInput,"  	& qtd_regis_recuperado		
							Vetor_Campos(11)="adVarchar,120,adParamInput,"  & des_complemnt
							Vetor_Campos(12)="adVarchar,5,adParamInput," 	& num_endereco
							
							Vetor_Campos(13)="adVarchar,120,adParamInput," 	& des_endereco
							Vetor_Campos(14)="adVarchar,8,adParamInput," 	& num_inicio_CEP
							Vetor_Campos(15)="adVarchar,8,adParamInput," 	& num_fim_CEP
							Vetor_Campos(16)="adVarchar,8,adParamInput," 	& num_CEP
							Vetor_Campos(17)="adVarchar,2,adParamInput," 	& ind_tipo_num_CEP
							Vetor_Campos(18)="adVarchar,8,adParamInput," 	& num_CEP_alternativo
							Vetor_Campos(19)="adVarchar,10,adParamInput," 	& sgl_tipo_lograd
							Vetor_Campos(20)="adVarchar,2,adParamInput, " 	& ind_tipo_categoria_lograd
							Vetor_Campos(21)="adVarchar,15,adParamInput," 	& des_titulo_lograd 
							Vetor_Campos(22)="adVarchar,60,adParamInput,"   & des_lograd
							Vetor_Campos(23)="adVarchar,60,adParamInput,"   & des_titulo_nome_lograd
							Vetor_Campos(24)="adVarchar,5,adParamInput,"    & num_inicio_endereco
							
							Vetor_Campos(25)="adVarchar,5,adParamInput," 	& num_fim_endereco
							Vetor_Campos(26)="adVarchar,120,adParamInput," 	& des_bairro
							Vetor_Campos(27)="adVarchar,2,adParamInput," 	& ind_lado_rua
							Vetor_Campos(28)="adNumeric,10,adParamInput," 	& cod_lograd_pesquisa
							Vetor_Campos(29)="adVarchar,1,adParamInput," 	& ind_tipo_localid
							Vetor_Campos(30)="adVarchar,20,adParamInput," 	& des_localid
							Vetor_Campos(31)="adVarchar,2,adParamInput," 	& des_uf
							Vetor_Campos(32)="adNumeric,10,adParamInput, " 	& cod_localid
							Vetor_Campos(33)="adNumeric,10,adParamInput," 	& cod_distrito
							Vetor_Campos(34)="adNumeric,10,adParamInput,"   & cod_regiao
							Vetor_Campos(35)="adVarchar,5,adParamInput,"   & cod_pais
							Vetor_Campos(36)="adVarchar,70,adParamInput,"   & nom_pais
							
							Vetor_Campos(37)="adVarchar,200,adParamInput, " 	& xmlhttp.statusText 
							Vetor_Campos(38)="adVarchar,8000,adParamInput," 	& strRetorno 
							Vetor_Campos(39)="adInteger,1,adParamInput,"    & Interf_num 
							Vetor_Campos(40)="adInteger,1,adParamInput,"    & Interf_erro
							Vetor_Campos(41)="adInteger,2,adParamOutput,0" 
							
							strSqlRet = APENDA_PARAMSTR("CSL_sp_check_servico",41,Vetor_Campos)
							db.Execute(strSqlRet)
							
						end if	
							
						
							doc.async = False
					    	doc1.async = False
							'xmlhttp.async = false
							
							''''------''''
							objRS.MoveNext
							intCount = intCount + 1
						Wend
						
					Else
						Response.Write "<script language=javascript>alert('Endereço não encontrado.')</script>"
						
					End if

  					'response.write "<script>alert(OK)</script>"
			
		'end if	   
	 	 	   
End Function
%>