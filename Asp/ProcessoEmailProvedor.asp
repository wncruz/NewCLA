<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoEmailProvedor.asp
'	- Responsável		: Vital
'	- Descrição			: Monta Email que vai para o provedor

Function AcaoPedidoEmail(intTipo)
	Select case intTipo
		case 1
			AcaoPedidoEmail = "Instalar Acesso"
		case 2			
			AcaoPedidoEmail = "Retirar Acesso"
		case 3		
			AcaoPedidoEmail = "Alterar Acesso"
		case 4
			AcaoPedidoEmail = "Cancelar Acesso"
	End Select
End Function

Function EnviarEmailProvedor(dblSolId,dblPedId,dblProId,strProEmail,statusPedido,intTipoProcesso,dblSisId)

	Dim textohtml
	Dim objDic

	textohtml = ""
	
	if dblPedId <> "" then

		Set ped = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
		'Response.Write ("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'" & statusPedido & "'")
		Set ped1 = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'" & statusPedido & "'")
		if Not ped1.Eof and Not ped1.Bof then

			if dblSisId = "3" then
					Set objRSPro = db.execute("CLA_sp_sel_estacao " & ped1("esc_identrega"))
					strProEmail = Trim(objRSPro("Esc_Email"))  
					strProNome	= Trim(objRSPro("Esc_Contato"))
			Else
				Set objRSPro = db.execute("CLA_sp_sel_provedor " & dblProId) 
				if Not objRSPro.Eof and Not objRSPro.bof then
					strProEmail = Trim(objRSPro("pro_email"))
					strProNome	= Trim(objRSPro("Pro_Nome"))
				End if
			End if	

			Set ObjMail	= Server.CreateObject("CDONTS.NewMail")
		
			ObjMail.From = "acessos@embratel.com.br"
			ObjMail.To	 = strProEmail 'impleme@embratel.com.br
			ObjMail.Subject = AcaoPedidoEmail(ucase(ped1("tprc_id"))) & "  -  " & trim(ped("Cli_nome")) & "  -  " & ucase(ped1("Ped_Prefixo")) & "-" & right("00000" & ped1("Ped_Numero"),5) & "/" & ped1("Ped_Ano")
			ObjMail.BodyFormat = 0
			ObjMail.MailFormat = 0
			textohtml = "<html><body align=center>"
			textohtml = textohtml & "<Head>"
			textohtml = textohtml & "<Style>"
			textohtml = textohtml & "TD"
			textohtml = textohtml & "{"
			'textohtml = textohtml & "FONT-WEIGHT: bold;"
			textohtml = textohtml & "FONT-SIZE: 8pt;"
			'textohtml = textohtml & "COLOR: #003388;"
			textohtml = textohtml & "FONT-FAMILY: Arial"
			textohtml = textohtml & "} "
			textohtml = textohtml & "TR.clsSilver2"
			textohtml = textohtml & "{"
			textohtml = textohtml & "    BACKGROUND-COLOR: #dcdcdc"
			textohtml = textohtml & "} "
			textohtml = textohtml & "TH"
			textohtml = textohtml & "{"
			textohtml = textohtml & "	font-family: Arial, Helvetica, sans-serif;"
			textohtml = textohtml & "	font-size: 11px;"
			textohtml = textohtml & "	font-weight: bold;"
			textohtml = textohtml & "	color: #ffffff;"
			textohtml = textohtml & "    BACKGROUND-COLOR: #31659c;"
			textohtml = textohtml & "    TEXT-ALIGN: left"
			textohtml = textohtml & "} "
			textohtml = textohtml & "</Style>"
			textohtml = textohtml & "</Head>"
			textohtml = textohtml & "<table align=center rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
			textohtml = textohtml & "<tr><td><br>EMBRATEL"
			textohtml = textohtml & "<div align=center>"
			textohtml = textohtml & "São Paulo, " & day(date) & " de "
		
			   select case month(date)
			   case "1"
				   textohtml = textohtml & "Janeiro"
			   case "2"
				   textohtml = textohtml & "Fevereiro"
			   case "3"
				   textohtml = textohtml & "Março"
			   case "4"
				   textohtml = textohtml & "Abril"
			   case "5"
				   textohtml = textohtml & "Maio"
			   case "6"
				   textohtml = textohtml & "Junho"
			   case "7"
				   textohtml = textohtml & "Julho"
			   case "8"
				   textohtml = textohtml & "Agosto"
			   case "9"
				   textohtml = textohtml & "Setembro"
			   case "10"
				   textohtml = textohtml & "Outubro"
			   case "11"
				   textohtml = textohtml & "Novembro"
			   case else
				   textohtml = textohtml & "Dezembro"
			   end select
			textohtml = textohtml & " de "& year(date) & ".</div>"
				
			textohtml = textohtml & "<br>À</br>"


			textohtml = textohtml & strProNome
			textohtml = textohtml & "<br><br>"
			textohtml = textohtml & "Assunto: Solicitação de Serviço" & "<br>"
			textohtml = textohtml & "Ação   : <u>"& AcaoPedidoEmail(ucase(ped1("tprc_id"))) & "</u>"

			Vetor_Campos(1)="adInteger,4,adParamInput,"
			Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
			Vetor_Campos(3)="adInteger,4,adParamInput,"
			Vetor_Campos(4)="adInteger,4,adParamInput,"
			Vetor_Campos(5)="adInteger,4,adParamInput,"
			Vetor_Campos(6)="adInteger,4,adParamInput,"
			Vetor_Campos(7)="adWChar,2,adParamInput,"
			Vetor_Campos(8)="adWChar,1,adParamInput,"
			Vetor_Campos(9)="adWChar,1,adParamInput,T"
						
			strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_ACESSOFISICO",9,Vetor_Campos)

			Set objRSFis = db.Execute(strSqlRet)
			strInterfacePto = ""
			if Not objRSFis.EOF and not objRSFis.BOF then
				strInterfacePto		= objRSFis("Acf_Interface")
				strNroAcessoPtaEbt	= objRSFis("Acf_NroAcessoPtaEbt")
				strTipoVel			= objRSFis("Acf_TipoVel")
				strVelFis			= objRSFis("Vel_Desc")
				strCctoProvedor		= Trim(objRSFis("Acf_NroAcessoCCTOProvedor"))
			End if


			if not isNull(ped1("Acl_DtIniAcessoTemp")) then
				textohtml = textohtml & "&nbsp;&nbsp;<font color=red>TEMPORÁRIO Período&nbsp;de&nbsp;" & ped1("Acl_DtIniAcessoTemp") & "&nbsp;à&nbsp;" & ped1("Acl_DtFimAcessoTemp") & "</font>"
			End if
			textohtml = textohtml & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Quantidade: 01"
			if trim(strNroAcessoPtaEbt) <> "" and not isnull(strNroAcessoPtaEbt) then
				textohtml = textohtml & "<br>Nº de Acesso: " & strNroAcessoPtaEbt
				if Trim(strCctoProvedor) <> "" then
					textohtml = textohtml & "<br>CCTO Provedor: " & strCctoProvedor
				End if 
			end if
			textohtml = textohtml & "<br><br>"
			textohtml = textohtml & "<font size=2 face=Arial color=#FF0000>Nº Pedido:"& ucase(ped1("Ped_Prefixo")) & "-" & right("00000" & ped1("Ped_Numero"),5) & "/" & ped1("Ped_Ano") & "</font><br><br>"

			Set objRSFac = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
			
			if Not objRSFac.Eof and Not objRSFac.Bof then
				intCount = 1
				dblSisId = objRSFac("Sis_Id")

				Set objDic = Server.CreateObject("Scripting.Dictionary") 
	
				While not objRSFac.Eof 

					Select Case dblSisId
						Case 1
							strChave = objRSFac("Fac_TimeSlot")
							strRepresentacao = objRSFac("Fac_TimeSlot")
						Case Else	
							strChave = objRSFac("Fac_Par")
							strRepresentacao = objRSFac("Fac_Par")
					End Select		

					if not isNull(objRSFac("Fac_Representacao")) then
						strRepresentacao = objRSFac("Fac_Representacao")
					End if

					Select Case dblSisId
						Case 1
							if  Not objDic.Exists(Cstr(strRepresentacao)) then
								Call objDic.Add(Cstr(strRepresentacao),Cstr(objRSFac("Fac_Id"))) 
								if intCount = 1 then
									strHtmlFac = ""
									strHtmlFac = strHtmlFac & "<table border=0 cellspacing=1 cellpadding=1 width=100% >"
									strHtmlFac = strHtmlFac & "<tr>"
									strHtmlFac = strHtmlFac & "<th colspan=7>&nbsp;•&nbsp;Informações da Facilidade</th>"
									strHtmlFac = strHtmlFac & "</tr>"
									strHtmlFac = strHtmlFac & "<tr class=clsSilver2>"
									strHtmlFac = strHtmlFac & "<td>&nbsp;Link</td>"
									strHtmlFac = strHtmlFac & "<td>&nbsp;Timeslot</td>"
									strHtmlFac = strHtmlFac & "	</tr>"
								End if
								strHtmlFac = strHtmlFac & "<tr class=clsSilver2>"
								strHtmlFac = strHtmlFac & "	<td >&nbsp;" & objRSFac("Fac_Link") & "</td>"
								strHtmlFac = strHtmlFac & "	<td >&nbsp;" & strRepresentacao & "</td>"
								strHtmlFac = strHtmlFac & "</tr>"
							End if
						Case 2	
							'NÃO DETERMINISTICO
							if  Not objDic.Exists(Cstr(strRepresentacao)) then
								Call objDic.Add(Cstr(strRepresentacao),Cstr(objRSFac("Fac_Id"))) 
								if intCount = 1 then
									strHtmlFac = ""
									strHtmlFac = strHtmlFac & "<table border=0 cellspacing=1 cellpadding=1 width=100% >"
									strHtmlFac = strHtmlFac & "	<tr >"
									strHtmlFac = strHtmlFac & "		<th colspan=5>&nbsp;•&nbsp;Informações da Facilidade</th>"
									strHtmlFac = strHtmlFac & "	</tr>"
									strHtmlFac = strHtmlFac & "	<tr class=clsSilver2>"
									strHtmlFac = strHtmlFac & "		<td nowrap>&nbsp;Tronco</td>"
									strHtmlFac = strHtmlFac & "		<td nowrap>&nbsp;Par</td>"
									strHtmlFac = strHtmlFac & "	</tr>"
								End if

								strHtmlFac = strHtmlFac & "<tr class=clsSilver2>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & objRSFac("Fac_Tronco") & "</td>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & strRepresentacao & "</td>"
								strHtmlFac = strHtmlFac & "</tr>"
							End if						
						Case 3	
							'ADE
							if  Not objDic.Exists(Cstr(strRepresentacao)) then
								Call objDic.Add(Cstr(strRepresentacao),Cstr(objRSFac("Fac_Id"))) 
								if intCount = 1 then
									strHtmlFac = ""
									strHtmlFac = strHtmlFac & "<table border=0 cellspacing=1 cellpadding=1 width=100% >"
									strHtmlFac = strHtmlFac & "	<tr>"
									strHtmlFac = strHtmlFac & "		<th colspan=6>&nbsp;•&nbsp;Informações da Facilidade</td>"
									strHtmlFac = strHtmlFac & "	</tr>"
									strHtmlFac = strHtmlFac & "	<tr class=clsSilver2>"
									strHtmlFac = strHtmlFac & "		<td width=100>&nbsp;Cabo</td>"
									strHtmlFac = strHtmlFac & "		<td width=120>&nbsp;Par</td>"
									strHtmlFac = strHtmlFac & "		<td width=100>&nbsp;Derivação</td>"
									strHtmlFac = strHtmlFac & "		<td nowrap width=100>&nbsp;T. Cabo</td>"
									strHtmlFac = strHtmlFac & "		<td nowrap>&nbsp;PADE</td>"
									strHtmlFac = strHtmlFac & "	</tr>"
								End if
								strHtmlFac = strHtmlFac & "<tr class=clsSilver2>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & objRSFac("Fac_Tronco") & "</td>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & strRepresentacao & "</td>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & objRSFac("Fac_Lateral") & "</td>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & objRSFac("Fac_TipoCabo") & "</td>"
								strHtmlFac = strHtmlFac & "<td>&nbsp;" & objRSFac("Fac_CxEmenda") & "</td>"
								strHtmlFac = strHtmlFac & "</tr>"
							End if	
					End Select
					intCount = intCount + 1
					objRSFac.MoveNext
				Wend	
				if Trim(strHtmlFac) <> "" then strHtmlFac = strHtmlFac & "</table>"
			
			End if

			textohtml = textohtml & strHtmlFac
			
			textohtml = textohtml & "<br>Solicitamos providenciar as medidas cabíveis conforme ação acima."
			textohtml = textohtml & "<br><br>"
			textohtml = textohtml & "<table><tr>"
			textohtml = textohtml & "<td width=180>Cliente</td><td>" & ped("CLI_nome")&"</td></tr>"
			textohtml = textohtml & "<tr><td>CNPJ</td><td>" & ped("Aec_CNPJ") & "</td></tr>"
			textohtml = textohtml & "<tr><td>Inscrição Estadual</td><td>" & ped("Aec_IE") & "</td></tr>"
			textohtml = textohtml & "<tr><td>Velocidade</td><td>" & strVelFis & " " & TipoVel(strTipoVel) &  "</td></tr>"
		
			'Response.Write ("CLA_sp_sel_regimecontrato 0," & dblProId)
			'Set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & dblProId)
			if not isNull(ped1("Reg_Id")) then
				Set objRSReg = db.execute("CLA_sp_sel_regimecontrato " & ped1("Reg_Id"))
				Set objRS = db.execute("CLA_sp_sel_tipocontrato " & objRSReg("Tct_Id"))

				if not objRS.Eof and Not objRS.Bof then
					textohtml = textohtml & "<tr><td>Regime de Contratação</td><td>" & objRS("Tct_Desc") & "</td></tr>"
				Else
					textohtml = textohtml & "<tr><td>Regime de Contratação</td><td></td></tr>"	
				End if		
			Else	
				textohtml = textohtml & "<tr><td>Regime de Contratação</td><td></td></tr>"	
			End if

			textohtml = textohtml & "<tr><td>Prazo de Entrega</td><td>" & Ped1("Ped_DtPrevistaAtendProv") & "</td></tr>"
			textohtml = textohtml & "</table><br><br>"
			textohtml = textohtml & "<center><u><b>LOCAL DE INSTALAÇÃO</b></u></center><br>"


			Set objRSConfig = db.execute("CLA_sp_sel_estacao " & Trim(ped("Esc_IDConfiguracao")))

			textohtml = textohtml & "<center><u><b>PONTA A - EMBRATEL</b></u></center><br>"
			textohtml = textohtml & "<table><tr>"
			set cid = db.execute("CLA_sp_sel_cidade2 '" & objRSConfig("Cid_Sigla") & "'")
			strEndereco = Trim(Cstr("" & objRSConfig("Tpl_Sigla"))) & " " &  Trim(Cstr("" & objRSConfig("Esc_NomeLogr"))) & " nº " & Trim(Cstr("" & objRSConfig("Esc_NroLogr"))) & " " & Trim(Cstr("" & objRSConfig("Esc_Complemento")))
			textohtml = textohtml & "<td width=180>Endereço</td><td>" & strEndereco & "</td></tr>"
			textohtml = textohtml & "<tr><td>Cep</td><td>" & objRSConfig("Esc_Cod_Cep")& " </td></tr>"
			textohtml = textohtml & "<tr><td>Cidade</td><td>" & cid("Cid_Desc")& "</td></tr>"
			set cid = nothing
			textohtml = textohtml & "<tr><td>Contato</td><td>" & objRSConfig("Esc_Contato")& "</td></tr>"
			textohtml = textohtml & "<tr><td>Telefone</td><td>" & objRSConfig("Esc_Telefone") & "</td></tr>"
			textohtml = textohtml & "<tr><td>Interface</td><td>" & ped("Acl_InterfaceEst") & "</td></tr>"
			textohtml = textohtml & "</table>"

			Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
			if not objRSEndPto.Eof and not objRSEndPto.bof then
				'strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
				textohtml = textohtml & "<center><u><b>PONTA B<br><br></b></u></center>"
				textohtml = textohtml & "<table><tr>"
				textohtml = textohtml & "<td width=180>Endereço</td><td>" & objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & " nº " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & "</td></tr>"
				textohtml = textohtml & "<tr><td>Cep</td><td> " & objRSEndPto("End_Cep") & "</td></tr>"
				textohtml = textohtml & "<tr><td>Cidade</td><td> " & objRSEndPto("Cid_Desc") & " - " & objRSEndPto("Est_Sigla") & "</td></tr>"
				textohtml = textohtml & "<tr><td>Contato</td><td>" & objRSEndPto("Aec_Contato")&"</td></tr>"
				textohtml = textohtml & "<tr><td>Telefone</td><td>"& objRSEndPto("Aec_Telefone")&"</td></tr>"
				textohtml = textohtml & "<tr><td>Interface</td><td>" & strInterfacePto & "</td></tr>"
				textohtml = textohtml & "</table><br>"

			End if	
			Set objRSEndPto = Nothing
			
			textohtml = textohtml & "<table><tr><td colspan=2>Observação</td></tr>"
			textohtml = textohtml & "<tr><td colspan=2><strong>"
			textohtml = textohtml & Trim(ped1("Ped_Obs")) & "</strong></font></td></tr>"
			textohtml = textohtml & "</table><br><br><br>Atenciosamente,<br><br>"

			'Usuario de coordenação embratel
			Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)
				
			if Not objRS.Eof then
				While Not objRS.Eof
					Select Case Trim(Ucase(objRS("Age_Sigla")))
						Case "GAT"
							strGLA = Trim(objRS("Usu_Username")) 
							strNomeGLA = Trim(objRS("Usu_Nome")) 
							strRamalGLA = Trim(objRS("Usu_Ramal")) 
						Case "C"
							strGICN = Trim(objRS("Usu_Username")) 
							strNomeGICN = Trim(objRS("Usu_Nome")) 
							strRamalGICN = Trim(objRS("Usu_Ramal")) 
						Case "E"
							strGICL = Trim(objRS("Usu_Username")) 
							strNomeGICL = Trim(objRS("Usu_Nome")) 
							strRamalGICL = Trim(objRS("Usu_Ramal"))
							if Trim(objRS("Agp_Origem")) = "P" then
								strUserGICL = strGICL
							End if
						Case "GAE"
							strGLAE = Trim(objRS("Usu_Username"))
							strNomeGLAE = Trim(objRS("Usu_Nome")) 
							strRamalGLAE = Trim(objRS("Usu_Ramal")) 
						
					End Select
					objRS.MoveNext
				Wend	
			End if

			textohtml = textohtml & "Elaborado por " & strNomeGLA
			textohtml = textohtml & "<br><br>Rubens Nallin Filho"
			textohtml = textohtml & "<br>Gerente de Implantação de Acessos de São Paulo.<br><br>"
			textohtml = textohtml & "<hr>Embratel - Empresa Brasileira de Telecomunicações S.A."
			textohtml = textohtml & "</td></tr></table>"

		
			textohtml = textohtml & "</body></html>"
			Set objRSConf = db.Execute("select * from cla_config where Config_ID = 4 and Config_Estado = 0 and Config_Data > getdate()")
			If Not objRSConf.eof and  not objRSConf.Bof Then
				ObjMail.Body = textohtml

				ObjMail.Send
				Set ObjMail = Nothing
			end if

		End if 'if Not ped1.Eof and Not ped1.Bof then
		
	End if

	if textohtml <> "" then 'Grava o e-mail que vai para o provedor
		Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
		Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
		Vetor_Campos(3)="adInteger,4,adParamInput," & intTipoProcesso
		Vetor_Campos(4)="adWChar,30,adParamInput," & strUserName
		Vetor_Campos(5)="adDate,8,adParamInput,"	  'Data de envio getdate()
		Vetor_Campos(6)="adWChar,4000,adParamInput," & textohtml
		strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_ins_documento",6,Vetor_Campos)
		Call db.Execute(strSqlRet)
	End if
	EnviarEmailProvedor  = textohtml

End Function
%>
