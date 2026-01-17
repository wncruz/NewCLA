<!--#include file="../inc/data.asp"-->


<%
'response.write "<script>alert('ok')</script>"
if	Trim(Request.Form("hdnNum_CEP")) <> "" then
	
	EnviarEndereco1123_EndLatLong_NewCLA()
Else
	Response.Write "<script language=javascript>alert('Informe o Endereço');</script>"
	Response.End 
End if	


				
Function EnviarEndereco1123_EndLatLong_NewCLA()

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
		
		Sev_seq							= ""
		Acp_seq							= ""
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
		
		CEP 							= Request.Form("hdnNum_CEP")
		
		txttipo							= Request.Form("hdnSgl_tipo_lograd")
		Logradouro						= Request.Form("hdnDes_titulo_nome_lograd")
		Numero							= Request.Form("hdnTxtNum")
		Complemento						= Request.Form("hdnTxtComple")
		
		Bairro 							= Request.Form("hdnDes_bairro")
		Municipio						= Request.Form("hdnDes_localid")
		UF 								= Request.Form("hdnDes_uf")
		tipo_PONTA						= Request.Form("hdntipoPonta")
		'SiglaCNL						= Request.Form("hdnCboSiglaCnl")
		
		'LATITUDE 						= Request.Form("hdnlatitude")
		'LONGITUDE						= Request.Form("hdnlongitude")
		'response.write CEP
		
		'response.write "<script>alert('"&CEP&"')</script>"
		'response.write "<script>alert('"&txttipo&"')</script>"
		'if ( Numero = "" ) then
		'	response.write "<script>alert('"&Numero&"')</script>"
		'	Numero = "0"		
		'end if 
		
		
		Endereco = txttipo + " " + Logradouro + " " + Numero + " " +  Complemento 
		
		'response.write "<script>alert('"&Endereco&"')</script>"
		'response.write "<script>alert('"&Endereco&"')</script>"
		'response.write "<script>alert('"&Endereco&"')</script>"
		'response.write "<script>alert('"&Endereco&"')</script>"
		
		
		'Set objRS = db.execute("SSA_sp_sel_CleanUP ")

		'if CEP <> "" then
			intCount = 0
			'response.write CEP
			'While Not objRS.Eof
			    ''''-----'''''
						
						intCount = intCount + 1 
						intCount = 1000000 + intCount
				
						Strxml1 = "" 
				
						oriSol_Descricao 		= "MPE-1123"
						'Sev_Seq 				= objRS("Sev_Seq")
						'Acp_Seq 				= objRS("Acp_Seq")
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
						Strxml1 = Strxml1 & 	"	  <ns0:dta_hra_solic>20150516204100</ns0:dta_hra_solic> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:ind_solic_valida_endereco>1</ns0:ind_solic_valida_endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:ind_solic_ajuste_escrita>1</ns0:ind_solic_ajuste_escrita> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:indic_recup_geolocaliz>1</ns0:indic_recup_geolocaliz> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:Consulta_Endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_endereco>"& Endereco &"</ns0:des_endereco> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Consulta_CEP> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_bairro>"& Bairro &"</ns0:des_bairro> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:num_CEP>"& CEP &"</ns0:num_CEP> " & vbnewline
						Strxml1 = Strxml1 & 	"	 <ns0:Consulta_CNL> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_localid>"& Municipio &"</ns0:des_localid> " & vbnewline
						Strxml1 = Strxml1 & 	"	  <ns0:des_uf>"& UF &"</ns0:des_uf> " & vbnewline
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
				      
					  	' doc.async = False
					  	' doc1.async = False
						
						
						Set doc = server.CreateObject("Microsoft.XMLDOM")
					    Set doc1 = server.CreateObject("Microsoft.XMLDOM")
					 
						Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
						
						
				      
					    'doc.async = False
					  	'doc1.async = False
					   
					  	'xmlhttp.Open "POST", "http://nthor028:9333/barramento/services/SolicitarAcesso", false
						
						
						'xmlhttp.Open "POST", "http://10.2.7.18:9005/barramento/services/Endereco " , False , "teste", "teste"
						
						'' Homologacao 
						''xmlhttp.Open "POST", "http://10.2.7.18:9005/barramento/services/Endereco " , False , "teste", "teste"
						
						xmlhttp.Open "POST", "http://10.243.172.32:10004/barramento/services/Endereco " , False , "sc3assabw0", "pZ9HS8tDzBKaGneh"
						
						'''xmlhttp.Open "POST", "http://10.2.9.69:9005/barramento/services/Endereco " , False , "teste", "teste"
						
						xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"  
				
						xmlhttp.setRequestHeader "SOAPAction", "/PesquisarEndereco"
					      
						xmlhttp.send(Strxml1)
						  
						'response.write "<script>alert('"&xmlhttp.statusText&"')</script>"	
						strRetorno = xmlhttp.ResponseText
						  
						doc1.loadXML(strRetorno)
						doc.loadXML(Strxml1)
						
						'doc1.save(Server.MapPath("1123-out.xml"))
						'doc.save(Server.MapPath("1123-in.xml"))
						
						
						'response.write "<script>alert('"&strRetorno&"')</script>"	
						'objNodeList = Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa").text
							
						'response.write "<script>alert('"&objNodeList&"')</script>"
						'response.end
						
						if xmlhttp.statusText = "OK" then
						
							
							
							Set objNodecod_retorno_valida = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_valida")
							if trim(objNodecod_retorno_valida.Length) <> "0" then
								cod_retorno_valida				= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_valida").text
							end if 
							
							Set objNodecod_retorno_pesquisa = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa")
							if trim(objNodecod_retorno_pesquisa.Length) <> "0" then
								cod_retorno_pesquisa			= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:cod_retorno_pesquisa").text
							end if
							'strHtmlAcessos = ""
							'strHtmlAcessos = strHtmlAcessos & "<table border=0 width=758 cellspacing=1 cellpadding=1 >" 
							
							'strHtmlAcessos = strHtmlAcessos & "<tr><th>Tipo Logradouro</th><th>Logradouro</th><th>Numero</th><th>Complemento</th><th>Bairro</th><th>Sigla CNL</th><th>Codigo IBGE</th><th>Municipio</th><th>UF</th><th>CEP</th></tr>"
				
							'strHtmlAcessos = strHtmlAcessos & "<tr class=" & strClass & ">"
							
							'response.write "<script>alert('"&cod_retorno_valida&"')</script>"
							
							if cod_retorno_valida = "1.0" then
							
									'if cod_retorno_pesquisa <> "16" then
							
																				
										Set objNodenum_latitude = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_latitude")
										if trim(objNodenum_latitude.Length) <> "0" then
											num_latitude					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_latitude").text
										end if
										
										Set objNodenum_longitude = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_longitude")
										if trim(objNodenum_longitude.Length) <> "0" then
											num_longitude					= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:Endereco/ns0:num_longitude").text
										end if
								
										'''valores = """" & tipo_PONTA & """"
										
										if tipo_PONTA = "A" then 					
											Response.Write "<script language=javascript>parent.document.forms[0].txtLatEnd_A.value = '" & num_latitude & "'</script>"
											Response.Write "<script language=javascript>parent.document.forms[0].txtLongEnd_A.value = '" & num_longitude & "'</script>"
										else
											Response.Write "<script language=javascript>parent.document.forms[0].txtLatEnd_B.value = '" & num_latitude & "'</script>"
											Response.Write "<script language=javascript>parent.document.forms[0].txtLongEnd_B.value = '" & num_longitude & "'</script>"
										end if 
										
										'''valores = valores & ", """ & num_latitude  & """"
										
										'''valores = valores & ", """ & num_longitude  & """"
										
										'response.write "<script>alert('"&valores&"')</script>"
										'response.write "<script>parent.parent.Gravar_LatLong("&valores& ");</script>"
										'response.write "<script>parent.cep();</script>"
									
									'end if
									
									
									
																
							'end if 
							Else
							'response.write "<script>alert('"&cod_retorno_pesquisa&"')</script>"
							'if cod_retorno_pesquisa = "100" then
									Set objNodedes_msg = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg")
									if trim(objNodedes_msg.Length) <> "0" then
										des_msg							= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg").text
									end if 
									Response.Write "<script language=javascript>alert('"&des_msg& " , FAVOR REALIZAR NOVA CONSULTA DE ENDERECO');</script>"
									Response.End 
							
							end if 
							
							'if cod_retorno_pesquisa = "200" then
							'		Set objNodedes_msg = Doc1.selectNodes("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg")
							'		if trim(objNodedes_msg.Length) <> "0" then
							'			des_msg							= Doc1.selectSingleNode("//SOAP-ENV:Envelope/SOAP-ENV:Body/ns0:Data/ns0:Solicitacao/ns0:RetornoSolicitacao/ns0:des_msg").text
							'		end if 
									
							'		Response.Write "<script language=javascript>alert('"&des_msg&"');</script>"
							'		Response.End
							
							'end if 
							
							
							
													
						end if	
							
						
							doc.async = False
					    	doc1.async = False
							'xmlhttp.async = false
							
							''''------''''
							
						'Wend
						
					'Else
					'	Response.Write "<script language=javascript>alert('Endereço não encontrado.')</script>"
						
		'End if

  					'response.write "<script>alert(OK)</script>"
			
		'end if	   
	 	 	   
End Function
%>
