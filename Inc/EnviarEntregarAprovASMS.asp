<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarEntregarAprovASMS(dblIdLogico)

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

   		if dblIdLogico <> "" then

			Vetor_Campos(1)="adWChar,15,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTR("CLA_sp_view_solicitacaoAprov",1,Vetor_Campos)
			
			Set objRSDadosCla = db.Execute(strSqlRet)
			
			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			
				Set objRSCla = db.execute("select top 1 max(sol_id) as sol_id , tprc_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico & " group by tprc_id order by 1 desc " )
 			    if Not objRSCla.Eof And Not objRSCla.Bof then
					solid 		=  objRSCla("sol_id")
					tipoProcesso = objRSCla("tprc_id")
				end if
				
				if tipoProcesso <> "2" then
				
					Vetor_Campos(1)="adInteger,4,adParamInput," & solid
					Vetor_Campos(2)="adInteger,4,adParamInput, 274"
					Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
					Vetor_Campos(4)="adVarchar,1,adParamInput,"
					Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
					Vetor_Campos(6)="adVarchar,1,adParamInput,M"
					
	  				strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
					
					db.Execute(strSqlRet)
					
				end if
				
				Aprovisi_ID			= objRSDadosCla("Aprovisi_ID")
				
				'response.write "<script>alert('"&objRSDadosCla("Switch")&"')</script>"
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
				strSqlHeader = APENDA_PARAMSTR("CLA_sp_sel_headerAsms",1,Vetor_Campos)
				
				Set objRSDadosHeader = db.Execute(strSqlHeader)
				
				header = ""
				external = ""
				
				If not objRSDadosHeader.Eof and  not objRSDadosHeader.Bof Then
						header 			= trim(objRSDadosHeader("header_desc"))
						external 		= trim(objRSDadosHeader("externalRef_desc")	)					
				end if 
				
				idTarefa 			= objRSDadosCla("id_tarefa")
				oe_numero 			= objRSDadosCla("oe_numero")
				oe_ano				= objRSDadosCla("oe_ano")
				oe_item				= objRSDadosCla("oe_item")
				idLogico			= objRSDadosCla("acl_idacessologico")
				acao				= objRSDadosCla("acao")
				oriSol_id			= objRSDadosCla("oriSol_ID")
				oriSol_Descricao 	= objRSDadosCla("oriSol_Descricao")
				rede 				=  objRSDadosCla("sis_id")
				
				strxmlResp = strRetorno
										
				'Checa se serviço é 0800.
				'''Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
				'''Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
				'''Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
				'''Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
				'''Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
				'''Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
				'''Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
				'''Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
				'''Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
				'''Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
				'''Vetor_Campos(11)="adInteger,1,adParamInput,0"
				'''Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
				'''Vetor_Campos(13)="adNumeric,10,adParamInput,NULL" 
				'''Vetor_Campos(14)="adInteger,2,adParamOutput,0"
				
				'''Call APENDA_PARAM("CLA_sp_check_servico2",14,Vetor_Campos)
				'''ObjCmd.Execute'pega dbaction
				'''idTRANSACAO = ObjCmd.Parameters("RET").value
				
				'Checa se serviço é 0800.
				Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
				Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
				Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
				Vetor_Campos(4)="adWChar,20,adParamInput," 	& idTarefa
				Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
				Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
				Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
				Vetor_Campos(8)="adWChar,200,adParamInput, RETORNO " 	'& strxmlResp
				Vetor_Campos(9)="adLongVarChar,8000,adParamInput," 	& Strxml1
				Vetor_Campos(10)="adWChar,1,adParamInput,4" 'Construir Return
				Vetor_Campos(11)="adWChar,1,adParamInput,0"
				Vetor_Campos(12)="adWChar,10,adParamInput," & dblIdLogico
				Vetor_Campos(13)="adWChar,10,adParamInput," 
				Vetor_Campos(14)="adInteger,2,adParamOutput,0"
				
				'Call APENDA_PARAM("CLA_sp_check_servico2",14,Vetor_Campos)
				'ObjCmd.Execute'pega dbaction
				'idTRANSACAO = ObjCmd.Parameters("RET").value
				
				Call APENDA_PARAM("CLA_sp_check_servico2",14,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value
				
				'strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
				'db.Execute(strSqlRet)
				idTRANSACAO = DBAction
				
				'''Strxml1 = Strxml1 & 	"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
				
				'Strxml1 = Strxml1 & 	"<?xml version=""1.0"" encoding=""UTF-16"" standalone=""yes"" ?>"
				Strxml1 = Strxml1 & 	"	<retorno-cla xmlns=""http://www.tibco.com/schemas/RECURSO/eAI/Business Processes/Equipamento/Schemas/schClaIn.xsd"" >" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-tarefa>"& objRSDadosCla("id_tarefa")& "</id-tarefa>" & vbnewline
				Strxml1 = Strxml1 & 	"		<num-trans>"& idTRANSACAO & "</num-trans>" & vbnewline 
				Strxml1 = Strxml1 & 	"		<origem>CLA</origem>" & vbnewline
				Strxml1 = Strxml1 & 	"		<destino>ASMS</destino>" & vbnewline
				'Strxml1 = Strxml1 & 	"<![CDATA[ " & vbnewline
				Strxml1 = Strxml1 & 	header  & vbnewline  ' "		<header-original>"& header &"</header-original>" & vbnewline 
				
				'Strxml1 = Strxml1 & 	 "		<header-original>"& header &"</header-original>" & vbnewline 
				
				'Strxml1 = Strxml1 & 	"]]> " & vbnewline
				'Strxml1 = Strxml1 & 	"<![CDATA[ " & vbnewline
				Strxml1 = Strxml1 & 	external & vbnewline ' "		<external-reflist>"& external &"</external-reflist>" & vbnewline
				
				'Strxml1 = Strxml1 & 	 "		<external-reflist>"& external &"</external-reflist>" & vbnewline
			'	Strxml1 = Strxml1 & 	"]]> " & vbnewline
				
				Strxml1 = Strxml1 & 	"		<acao>"& objRSDadosCla("acao") &"</acao>" & vbnewline				
				
				''' Verificar
				Strxml1 = Strxml1 & 	"		<id-servico>"& objRSDadosCla("Id_Servico") &"</id-servico>" & vbnewline
				if tipoProcesso = "2" then
					Strxml1 = Strxml1 & 	"		<dt-desativacao>"& objRSDadosCla("Acf_DtDesatAcessoFis_Asms") &"</dt-desativacao>" & vbnewline
				end if 
				if rede = "11" then				
					Strxml1 = Strxml1 & 	"		<id-acesso>"& objRSDadosCla("DesignacaoContrato") &"</id-acesso>" & vbnewline
				else
					Strxml1 = Strxml1 & 	"		<id-acesso>"& objRSDadosCla("acf_idacessofisico") &"</id-acesso>" & vbnewline
				end if
				Strxml1 = Strxml1 & 	"		<cod-retorno>0</cod-retorno>" & vbnewline
				Strxml1 = Strxml1 & 	"		<msg-retorno>Sucesso</msg-retorno>" & vbnewline
				
				
				
				if objRSDadosCla("tprc_id") = "3" then
					
					'strIDLogico677 = "677" & mid(objRSDadosCla("acl_idacessologico"),4,7)
					'strIDLogico678 = "678" & mid(objRSDadosCla("acl_idacessologico"),4,7)
					
					Strxml1 = Strxml1 & 	"		<id-logico>"& objRSDadosCla("Sigla_Acl_idacessologicoAtv") &"</id-logico>" & vbnewline
					Strxml1 = Strxml1 & 	"		<id-logico-tmp>"& objRSDadosCla("acl_idacessologico") &"</id-logico-tmp>" & vbnewline
				else
					Strxml1 = Strxml1 & 	"		<id-logico>"& objRSDadosCla("Sigla_Acl_idacessologico") &"</id-logico>" & vbnewline
					Strxml1 = Strxml1 & 	"		<id-logico-tmp>"& objRSDadosCla("acl_idacessologico") &"</id-logico-tmp>" & vbnewline
				end if 
				
				Strxml1 = Strxml1 & 	"		<sol_id>"& objRSDadosCla("sol_id") &"</sol_id>" & vbnewline
				
				Strxml1 = Strxml1 & 	"		<acessos-fisicos>" & vbnewline
				
				xmlAcesso  = ""
				
			    tecnologia   =  objRSDadosCla("tec_id")
			   ' solid 		=  objRSDadosCla("sol_id")
			   
	   	   		xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
								
	    		xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
				xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
	    		xmlAcesso = xmlAcesso & 	"		<provedor>"& objRSDadosCla("pro_nome") &"</provedor>" & vbnewline
				
				if rede = "11" then
					xmlAcesso = xmlAcesso & 	"		<tecnologia>"& rede &"</tecnologia>" & vbnewline
				else
				xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
				end if 
				xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
				xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& objRSDadosCla("Acf_DtConstrAcessoFis_Asms") &"</dt-construcao>" & vbnewline
				if tecnologia = "6" then
					xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("TPPorta_Nome") &"</interface>" & vbnewline
				else
					xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
				end if 
	    		
				if tecnologia = "3" then
					xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cabo>"& objRSDadosCla("fac_tronco") &"</cabo>" & vbnewline ' tipo de acesso ADE fac_tronco e o cabo
					xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tp-cabo>"& objRSDadosCla("fac_tipoCabo")&"</tp-cabo>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
					
					' Verificar
					xmlAcesso = xmlAcesso & 	"		<derivacao>"& objRSDadosCla("Fac_Lateral") &"</derivacao>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tipo-modem></tipo-modem>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<pade>"& objRSDadosCla("Fac_CxEmenda") &"</pade>" & vbnewline
					
				end if
				
				if tecnologia = "7" then
				
					xmlAcesso = xmlAcesso & 	"		<pe>"& trim(objRSDadosCla("OntVlan_PE")) &"</pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-pe>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</porta-pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-ont>"& trim(objRSDadosCla("OntPorta_Porta")) &"</porta-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<svlan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</svlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cvlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</cvlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<elan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</elan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<designacao-ont>"& trim(objRSDadosCla("Ont_Desig")) &"</designacao-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<modelo-ont>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo-ont>" & vbnewline					
					xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline   
					
					xmlAcesso = xmlAcesso & 	"		<switch>"& objRSDadosCla("Switch") &"</switch>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-switch>"& objRSDadosCla("Switch_Porta") &"</porta-switch>" & vbnewline
									
				end if
				
				
				'GPON
				if tecnologia = "6" then
				
					xmlAcesso = xmlAcesso & 	"		<pe>"& trim(objRSDadosCla("OntVlan_PE")) &"</pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-pe>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</porta-pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-ont>"& trim(objRSDadosCla("OntPorta_Porta")) &"</porta-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<svlan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</svlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cvlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</cvlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<elan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</elan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<designacao-ont>"& trim(objRSDadosCla("Ont_Desig")) &"</designacao-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<modelo-ont>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo-ont>" & vbnewline					
					xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline    			
					xmlAcesso = xmlAcesso & 	"		<switch></switch>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-switch></porta-switch>" & vbnewline
				end if
				
				
				'FO Ethernet
				if rede = "11" then
				
					xmlAcesso = xmlAcesso & 	"		<pe>"& trim(objRSDadosCla("OntVlan_PE")) &"</pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-pe>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</porta-pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-ont></porta-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<svlan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</svlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cvlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</cvlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<elan></elan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<designacao-ont></designacao-ont>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<modelo-ont></modelo-ont>" & vbnewline					
					xmlAcesso = xmlAcesso & 	"		<fabricante></fabricante>" & vbnewline    			
					xmlAcesso = xmlAcesso & 	"		<switch></switch>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<porta-switch></porta-switch>" & vbnewline
				end if
				
				xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
	    		xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
				'xmlAcesso = xmlAcesso & 	"		<predio-estacao>"& objRSDadosCla("esc_predio") &"</predio-estacao>" & vbnewline
		   		
				xmlAcesso = xmlAcesso & 	"	</acesso> "
						
				Strxml1 = Strxml1 & 	xmlAcesso & vbnewline
				
				Strxml1 = Strxml1 & 	"		</acessos-fisicos>" & vbnewline
				Strxml1 = Strxml1 & 	"	</retorno-cla>" & vbnewline
				
				
				'''Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
				'''Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
				'''Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
				'''Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
				'''Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
				'''Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
				'''Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
				'''Vetor_Campos(8)="adVarchar,200,adParamInput, strxmlResp " 	& strxmlResp
				'''Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
				'''Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
				'''Vetor_Campos(11)="adInteger,1,adParamInput,0"
				'''Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
				
				'''strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
				'''db.Execute(strSqlRet)
				
				Set doc = server.CreateObject("Microsoft.XMLDOM")
			    Set doc1 = server.CreateObject("Microsoft.XMLDOM")
			   ' Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
				Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
		      
			    doc.async = False
			    doc1.async = False
			   
			   ' xmlhttp.Open "POST", "http://nthor028:9333/barramento/services/SolicitarAcesso"
				'xmlhttp.setRequestHeader "SOAPAction", "executeClass"
			 '   xmlhttp.send(Strxml)
				
				'xmlhttp.Open "POST", "http://nthor024.nt.embratel.com.br:9333/barramento/services/SolicitarAcesso", false
               
			   ''' xmlhttp.Open "POST", "http://nthor028:9333/barramento/services/SolicitarAcesso", false		
				
				''' Desenv 1
				'''xmlhttp.Open "POST", "http://10.2.13.56:9333/barramento/services/SolicitarAcesso", false	
				
				
				''' Desenv 2
				''' xmlhttp.Open "POST", "http://2k8rjohmgbw03:9333/barramento/services/SolicitarAcesso", false
				
				If Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'PRD' and OriSol_ID = " & OriSol_ID
				else
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'DSV2' and OriSol_ID = " & OriSol_ID
				end if
				
				Set objRS = db.Execute(StrSQL)
				If Not objRS.eof and  not objRS.Bof Then
					objConBW = objRS("Conn_Desc")
				End if
				
				''xmlhttp.Open "POST", "http://10.2.7.18:9333/barramento/services/SolicitarAcesso", false 	
				
				xmlhttp.Open "POST", objConBW, false 	
                                'xmlhttp.setRequestHeader "SOAPAction" , "executeClass"
				xmlhttp.setRequestHeader "Content-Length", len(Strxml1)
			  
			  	xmlhttp.send(Strxml1)
			    strRetorno = xmlhttp.ResponseText
				
				'response.write "<script>alert('"&strRetorno&"')</script>"				
								
				if strRetorno = "Recebido BW" then
					strxmlResp = strRetorno
					if tipoProcesso <> "2" then
						
						
						Vetor_Campos(1)="adInteger,4,adParamInput," & solid
						Vetor_Campos(2)="adInteger,4,adParamInput, 267"
						Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
						Vetor_Campos(4)="adVarchar,1,adParamInput,"
						Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
						Vetor_Campos(6)="adVarchar,1,adParamInput,M"
						
	  					strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
						
						db.Execute(strSqlRet)
					end if
					
					Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
					Vetor_Campos(2)="adVarchar,20,adParamInput, Entregar"
					strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
					db.Execute(strSqlRet)
					
					'Checa se serviço é 0800.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,0"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					Vetor_Campos(13)="adNumeric,10,adParamInput," & idTRANSACAO
					Vetor_Campos(14)="adInteger,2,adParamOutput,0" 
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
					db.Execute(strSqlRet)
					
					
					
				else
					strxmlResp = strRetorno
					
					'Checa se serviço é 0800 - E.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,1"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					Vetor_Campos(13)="adNumeric,10,adParamInput," & idTRANSACAO
					Vetor_Campos(14)="adInteger,2,adParamOutput,0" 
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
					db.Execute(strSqlRet)
					
					Set objMail = CreateObject("CDONTS.NewMail")
					gMailTo1 =  "prssilv@embratel.com.br,edar@embratel.com.br"
					objMail.To = gMailTo1
					objMail.From = From
					objMail.Subject = Sbj4
					objMail.Body = Data & " - " & Hora & " > - " & dblIdLogico & " - " & strxmlResp
					objMail.Send
					Set objMail = Nothing
				end if	
			end if	   
	 end if
	 	   
End Function
%>