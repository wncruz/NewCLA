<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarManobraAprov(dblAcfID)

		Dim proprietario ' TER ou EBT
		Dim tecnologia ' 0 - terceiro 1 - radio 2 - fibra otica 3 -ade 4 - satelite 5 - cabo interno 
		Dim idTarefa ' identificador do sistema aprovisionador
		Dim oe_numero 		 
		Dim	oe_ano			 
		Dim	oe_item			 
		Dim	idLogico		 
		Dim	acao			 
		Dim OrigemSolicitacao
		Dim rede
		Dim oriSol_id

   		if dblAcfID <> "" then

			Vetor_Campos(1)="adWChar,15,adParamInput, null " 
			Vetor_Campos(2)="adInteger,2,adParamInput," & dblAcfID
			strSqlRet = APENDA_PARAMSTR("CLA_sp_view_solicitacaoAprov",2,Vetor_Campos)
			
			Set objRSDadosCla = db.Execute(strSqlRet)
			
			If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			
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
				Strxml1 = Strxml1 & 	"		<origem>"& objRSDadosCla("oriSol_Descricao") &"</origem>" & vbnewline
				
				xmlAcesso  = ""
				
				''while not objRSDadosCla.eof
				
				If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
				   
				   proprietario =  objRSDadosCla("acf_proprietario")
				   tecnologia   =  objRSDadosCla("tec_id")
				   rede 		=  objRSDadosCla("sis_id")
				   solid 		=  objRSDadosCla("sol_id")
				   if proprietario = "TER" then
				   		xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
						xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
						
						xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso> " & vbnewline 
		    			xmlAcesso = xmlAcesso & 	"		<numero-acesso-cli>"& objRSDadosCla("Acf_NroAcessoPtaCli") &"</numero-acesso-cli>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<cnla>"& objRSDadosCla("Acf_CnlPTA") &"</cnla>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<cnlb>"& objRSDadosCla("Acf_CnlPTB") &"</cnlb>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<tipo-circuito>"& objRSDadosCla("Acf_CCTOTipo") &"</tipo-circuito>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
		    			xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
		    			
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
							
							xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<cabo>"& objRSDadosCla("fac_tronco") &"</cabo>" & vbnewline ' tipo de acesso ADE fac_tronco e o cabo
			    			xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
			    			xmlAcesso = xmlAcesso & 	"		<tp-cabo>"& objRSDadosCla("fac_tipoCabo")&"</tp-cabo>" & vbnewline
						
				   			xmlAcesso = xmlAcesso & 	"	</acesso> "
						
				   end if
				   
				  
				end if 
				
				Strxml1 = Strxml1 & 	xmlAcesso & vbnewline
				
				Strxml1 = Strxml1 & 	"		<identificadores>" & vbnewline
				Strxml1 = Strxml1 & 	"			<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
				Strxml1 = Strxml1 & 	"			<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
				while not objRSDadosCla.eof
					Strxml1 = Strxml1 & 	"		<identificador>" & vbnewline
					Strxml1 = Strxml1 & 	"			<id-tarefa>"& objRSDadosCla("id_tarefa")& "</id-tarefa>" & vbnewline
					Strxml1 = Strxml1 & 	"			<id-logico>"& objRSDadosCla("acl_idacessologico") &"</id-logico>" & vbnewline
					Strxml1 = Strxml1 & 	"		</identificador>" & vbnewline
					objRSDadosCla.movenext
				wend
				Strxml1 = Strxml1 & 	"		</identificadores>" & vbnewline
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
				
				'response.write "<script>alert('"&objRS("Conn_Desc")&"')</script>"
				
				ConSGA.ActiveConnection = objConSGA
				
				if oriSol_id = 6 then
					'ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
					ConSGA.CommandText =  "sgaplus_adm.pck_sgap_interface_cla.pc_Atualiza_Facilidades"
					'ConSGA.CommandText =  "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
					
				end if
				if oriSol_id = 7 then
					ConSGA.CommandText = "sgav_vips.sp_sgav_interface_cla"
				end if
				ConSGA.CommandType = adCmdStoredProc
				
				'*** Carregando parâmetros de entrada
				''Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, idTarefa)
				''ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("p1", adLongVarWChar, adParamInput, 1073741823, Strxml1)
				ConSGA.Parameters.Append objParam
				
				'*** Configurando variável que receberá o retorno
				Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
				ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 100 )
				ConSGA.Parameters.Append objParam
				
				'Tratamento de erro crítico:
				On error resume next
				
				'*** Executando a stored procedure
				ConSGA.Execute
				
				if err.number <> 0 then
					strxmlResp = "ERRO Critico: " & err.number & " - " & err.description
					On Error GoTo 0
					
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
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
					db.Execute(strSqlRet)
					
				   ' response.write "<script>alert('Erro crítico no " & OriSol_Descricao & "')</script>"
					response.end
				end if
				
				
				cod_retorno  = ConSGA.Parameters("Ret1").value
				desc_retorno = ConSGA.Parameters("Ret2").value
				
				'strxmlResp = cod_retorno & " - " & desc_retorno
				
				'response.write "<script>alert('"&strxmlResp&"')</script>"
				
				if cod_retorno = 0 then
				
					strxmlResp = 	"<resposta-cla><codigo>" & Trim(cod_retorno) & "</codigo>"
					strxmlResp = strxmlResp  & 	"<mensagem>" & Trim(desc_retorno) & "</mensagem></resposta-cla>"
										
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
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
					db.Execute(strSqlRet)
				else
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
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
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