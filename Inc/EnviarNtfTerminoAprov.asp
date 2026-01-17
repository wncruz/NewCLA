<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarNtfterminoAprov(dblIdLogico)

		Dim proprietario
		Dim tecnologia 
		Dim idTarefa
		Dim oe_numero
		Dim	oe_ano
		Dim	oe_item
		Dim	idLogico
		Dim	acao
		Dim OrigemSolicitacao
		Dim rede
		Dim oriSol_id

   		if dblIdLogico <> "" then

			Vetor_Campos(1)="adWChar,15,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTR("CLA_sp_view_solicitacaoAprov",1,Vetor_Campos)
			
			Set objRSDadosCla = db.Execute(strSqlRet)
			
			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			
				solid 		=  objRSDadosCla("sol_id")
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & solid
				Vetor_Campos(2)="adInteger,4,adParamInput, 274"
				Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
				Vetor_Campos(4)="adVarchar,1,adParamInput,"
				Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
				Vetor_Campos(6)="adVarchar,1,adParamInput,M"
				
  				strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
				
				db.Execute(strSqlRet)
				
				Aprovisi_ID			= objRSDadosCla("Aprovisi_ID")
				idTarefa 			= objRSDadosCla("id_tarefa")
				oe_numero 			= objRSDadosCla("oe_numero")
				oe_ano				= objRSDadosCla("oe_ano")
				oe_item				= objRSDadosCla("oe_item")
				idLogico			= objRSDadosCla("acl_idacessologico")
				acao				= objRSDadosCla("acao")
				OrigemSolicitacao	= objRSDadosCla("oriSol_Descricao")
				oriSol_id			= objRSDadosCla("oriSol_ID")
				oriSol_Descricao 	= objRSDadosCla("oriSol_Descricao")
				
				Strxml1 = Strxml1 & 	"	<retorno-cla>" & vbnewline
				Strxml1 = Strxml1 & 	"		<acao>"& objRSDadosCla("acao") &"</acao>" & vbnewline
				Strxml1 = Strxml1 & 	"		<origem>"& objRSDadosCla("oriSol_Descricao") &"</origem>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-tarefa>"& objRSDadosCla("id_tarefa")& "</id-tarefa>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-logico>"& objRSDadosCla("acl_idacessologico") &"</id-logico>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-solicitacao>"& objRSDadosCla("sol_id") &"</id-solicitacao>"  & vbnewline
				Strxml1 = Strxml1 & 	"		<qtd-acessos>"& objRSDadosCla("QTDFisico") &"</qtd-acessos>" & vbnewline
				Strxml1 = Strxml1 & 	"		<acessos-fisicos>" & vbnewline
				
				xmlAcesso  = ""
				
				while not objRSDadosCla.eof
				   
				   proprietario =  objRSDadosCla("acf_proprietario")
				   tecnologia   =  objRSDadosCla("tec_id")
				   rede 		=  objRSDadosCla("sis_id")
				   solid 		=  objRSDadosCla("sol_id")
				   if proprietario = "TER" then
				   		xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
						xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<provedor>"& objRSDadosCla("pro_nome") &"</provedor>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<tipo-contrato>"& objRSDadosCla("reg_id") &"</tipo-contrato>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& objRSDadosCla("Acf_DtConstrAcessoFis") &"</dt-construcao>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<dt-aceite>"& objRSDadosCla("Acf_DtAceite") &"</dt-aceite>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<tp-vel>"& objRSDadosCla("acf_tipovel") &"</tp-vel>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<interface-estacao>"& objRSDadosCla("Acf_InterfaceEstEntregaFisico") &"</interface-estacao>" & vbnewline
			    		xmlAcesso = xmlAcesso & 	"		<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso> " & vbnewline 
		    			xmlAcesso = xmlAcesso & 	"		<numero-acesso-cli>"& objRSDadosCla("Acf_NroAcessoPtaCli") &"</numero-acesso-cli>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<cnla>"& objRSDadosCla("Acf_CnlPTA") &"</cnla>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<cnlb>"& objRSDadosCla("Acf_CnlPTB") &"</cnlb>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<tipo-circuito>"& objRSDadosCla("Acf_CCTOTipo") &"</tipo-circuito>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<tp-aceite>"& objRSDadosCla("TP_TipoAceite") &"</tp-aceite>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
				   		xmlAcesso = xmlAcesso & 	"		<plataforma>"& objRSDadosCla("Pla_TipoPlataforma") &"</plataforma>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<SiglaCentroCliente>"& objRSDadosCla("Aec_SiglaCentroCliente") &"</SiglaCentroCliente>" & vbnewline
						
						if rede = "1" then
							xmlAcesso = xmlAcesso & 	"		<fila>"& objRSDadosCla("fac_fila") &"</fila>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<bastidor>"& objRSDadosCla("fac_bastidor") &"</bastidor>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<regua>"& objRSDadosCla("fac_regua") &"</regua>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<posicao>"& objRSDadosCla("fac_posicao") &"</posicao>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<dominio>"& objRSDadosCla("fac_posicao") &"</dominio>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<no>"& objRSDadosCla("fac_no") &"</no>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<slot>"& objRSDadosCla("fac_slot") &"</slot>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<time-slot>"& objRSDadosCla("fac_timeslot") &"</time-slot>" & vbnewline
						end if
						if rede = "2" then
							xmlAcesso = xmlAcesso & 	"		<tronco>"& objRSDadosCla("fac_tronco") &"</tronco>" & vbnewline
	    					xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
						end if
						
				   		xmlAcesso = xmlAcesso & 	"	</acesso> "
				   else
					   		xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
							xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<tipo-contrato>"& objRSDadosCla("reg_id") &"</tipo-contrato>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& objRSDadosCla("Acf_DtConstrAcessoFis") &"</dt-construcao>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<dt-aceite>"& objRSDadosCla("Acf_DtAceite") &"</dt-aceite>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<tp-vel>"& objRSDadosCla("acf_tipovel") &"</tp-vel>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
				    		xmlAcesso = xmlAcesso & 	"		<interface-estacao>"& objRSDadosCla("Acf_InterfaceEstEntregaFisico") &"</interface-estacao>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<SiglaCentroCliente>"& objRSDadosCla("Aec_SiglaCentroCliente") &"</SiglaCentroCliente>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<EstacaoConfiguracao>"& objRSDadosCla("est_config") &"</EstacaoConfiguracao>" & vbnewline
							
						if tecnologia = "1" then
							if trim(objRSDadosCla("Acf_NroAcessoPtaEbt")) = "" then
								xmlAcesso = xmlAcesso & 	"		<desig-banda-basica>"& objRSDadosCla("designacaoTronco") &"</desig-banda-basica>" & vbnewline
							else
								xmlAcesso = xmlAcesso & 	"		<desig-banda-basica>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</desig-banda-basica>" & vbnewline
							end if
							
						end if
						
						if tecnologia = "3" then
							xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<tp-aceite>"& objRSDadosCla("TP_TipoAceite") &"</tp-aceite>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<cabo>"& objRSDadosCla("fac_tronco") &"</cabo>" & vbnewline ' tipo de acesso ADE fac_tronco e o cabo
			    			xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<tp-cabo>"& objRSDadosCla("fac_tipoCabo")&"</tp-cabo>" & vbnewline
						end if
				   		xmlAcesso = xmlAcesso & 	"	</acesso> "
						
				   end if
				   
				   objRSDadosCla.movenext
				wend
				
				Strxml1 = Strxml1 & 	xmlAcesso & vbnewline
				
				Strxml1 = Strxml1 & 	"		</acessos-fisicos>" & vbnewline
				Strxml1 = Strxml1 & 	"	</retorno-cla>" & vbnewline
				
				set ConSGA = Server.CreateObject("ADODB.Command")
				
				If Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'PRD' and OriSol_ID = " & OriSol_ID
				else
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'DSV' and OriSol_ID = " & OriSol_ID
				end if
				
				Set objRS = db.Execute(StrSQL)
				If Not objRS.eof and  not objRS.Bof Then
					objConSGA = objRS("Conn_Desc")
				End if
				
				ConSGA.ActiveConnection = objConSGA
				
				if oriSol_id = 6 then
					ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
				end if
				if oriSol_id = 7 then
					ConSGA.CommandText = "sgav_vips.sp_sgav_interface_cla"
				end if
				ConSGA.CommandType = adCmdStoredProc
				
				'*** Carregando parâmetros de entrada
				Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, idTarefa)
				ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("p2", adLongVarWChar, adParamInput, 1073741823, Strxml1)
				ConSGA.Parameters.Append objParam
				
				'*** Configurando variável que receberá o retorno
				Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
				ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 100 )
				ConSGA.Parameters.Append objParam
				
			end if
		   
	 end if
	 	   
End Function
%>