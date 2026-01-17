<!--#include file="../inc/data.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoEmailProvedorPadrao.asp
'	- Descrição			: Monta Email que vai para o provedor


dim dblSolId 
dim dblPedId 
dim dblProId
dim dblEscEntrega 
dim intTipoProcesso
dim objRSPro
dim strRet
dim strButton

dim ndPed 
dim ndSol 
dim ndProv
dim ndEsc 
dim ndTipo
dim strContatoEBT

set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
'Atribuição de valores para as variáveis 	
	
objXmlDoc.load(Request)
strCaminho = server.MapPath("..\")
strContatoEBT = ""

set ndPed =  objXmlDoc.selectSingleNode("//ped")
set ndSol  =  objXmlDoc.selectSingleNode("//sol")
set ndProv =  objXmlDoc.selectSingleNode("//Prov")
'set ndEsc =  objXmlDoc.selectSingleNode("//Esc")
set ndTipo =  objXmlDoc.selectSingleNode("//ndTipo")
set ndRede =  objXmlDoc.selectSingleNode("//Rede")

dblPedId  = ndPed.Text
dblSolId  = ndSol.Text
dblProId = ndProv.Text
'dblEscEntrega = ndEsc.Text
intTipoProcesso = ndTipo.Text
statusPedido = "T"
dblSisId = ndRede.Text

Set objRSPro =  nothing 

	Dim textohtml
	Dim objDic

	textohtml = ""
	if dblPedId <> "" then

		Set ped = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
		'Response.Write ("EXEC CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'" & statusPedido & "'")
		Set ped1 = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'" & statusPedido & "'")
		if Not ped1.Eof and Not ped1.Bof then		
			if dblSisId = "3" then
					'Response.Write ("EXEC CLA_sp_sel_estacao " & ped1("esc_identrega"))
					Set objRSPro = db.execute("CLA_sp_sel_estacao " & ped1("esc_identrega"))
					strProEmail = Trim(objRSPro("Esc_Email"))  
					strProNome	= Trim(objRSPro("Esc_Contato"))
					strFromEmail = "acessos@embratel.com.br"
					
					if strProEmail = "" or isnull (strProEmail) or isempty(strProEmail) then
					  Response.ContentType = "text/HTML;charset=ISO-8859-1"
					  %>
					  <script language="VBScript">
					    msgbox "Não há e-mail cadastrado para esta estação: <%=objRSPro("Cid_Sigla") & " "%> <%=objRSPro("Esc_Sigla")%> - Utilize TABELAS/ESTAÇÃO para associar.",64,"CLA - Informação"
					    msgbox "O e-mail não pode ser enviado. Clique em OK para fechar esta janela.",48,"CLA - Alerta"
					    window.close()
					  </script>
					  <%
					  Response.end
					end if  
			Else
				Set objRSPro = db.execute("CLA_sp_sel_provedoremail " & dblProId & ",null,'" & ped1("Est_Sigla") &"','"& ped1("Cid_Sigla") &"'") 
				'Response.write ("CLA_sp_sel_provedoremail " & dblProId & ",null,'" & ped1("Est_Sigla") &"','"& ped1("Cid_Sigla") &"'") 
				if Not objRSPro.Eof and Not objRSPro.bof then
					if not isnull(Trim(objRSPro("Cpro_Contratadaemail"))) then 
						strProEmail = Trim(objRSPro("Cpro_Contratadaemail"))
					else
						strProEmail = ""
					end if 
					strProNome	= Trim(objRSPro("Pro_Nome"))
					
					If not isnull(Trim(objRSPro("cPRo_ContratanteContato"))) Then
						strContatoEBT = Trim(objRSPro("cPRo_ContratanteContato"))
					Else 
						strContatoEBT = ""
					End If
					

					if not isnull(Trim(objRSPro("Cpro_Contratanteemail"))) then 
						strFromEmail = Trim(objRSPro("Cpro_Contratanteemail"))
					else
						strFromEmail = "acessos@embratel.com.br"
					end if 
				else
					strFromEmail = "acessos@embratel.com.br"
				End if
			End if	

			Set ObjMail	= Server.CreateObject("CDONTS.NewMail")
		    'Response.Write "<script>alert('ProcessoEmailProvedorPadrao.asp(PRSS):strFromEmail "& strFromEmail &" ')</script>"
			'Response.Write "<script>alert('ProcessoEmailProvedorPadrao.asp(PRSS):strProEmail "&strProEmail&" ')</script>"
			ObjMail.From = strFromEmail '"acessos@embratel.com.br"
			ObjMail.To	 = strProEmail  'impleme@embratel.com.br
			ObjMail.Subject = AcaoPedidoEmail(ucase(ped1("tprc_id"))) & "  -  " & trim(ped("Cli_nome")) & "  -  " & ucase(ped1("Ped_Prefixo")) & "-" & right("00000" & ped1("Ped_Numero"),5) & "/" & ped1("Ped_Ano")
			ObjMail.BodyFormat = 0
			ObjMail.MailFormat = 0
			textohtml = "<html><body align=center>"
			textohtml = textohtml & "<Head>"
			textohtml = textohtml & "<title> Pedido: "& ucase(ped1("Ped_Prefixo")) & "-" & right("00000" & ped1("Ped_Numero"),5) & "/" & ped1("Ped_Ano") &  "</title>"
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
			textohtml = textohtml & "INPUT.button"
			textohtml = textohtml & "{"
			textohtml = textohtml & "	font-family: Verdana, Arial, Helvetica, sans-serif;"
			textohtml = textohtml & "	font-size: 9px;"
			textohtml = textohtml & "	font-weight: normal;"
			textohtml = textohtml & "	TEXT-ALIGN: center"
			textohtml = textohtml & "	color: #000000;"
			textohtml = textohtml & "	text-decoration: none;"
			textohtml = textohtml & "	background-color: #f1f1f1;"
			textohtml = textohtml & "	border-top: 1px solid #0F1F5F;"
			textohtml = textohtml & "	border-right: 1px solid #0F1F5F;"
			textohtml = textohtml & "	border-bottom: 1px solid #0F1F5F;"
			textohtml = textohtml & "	border-left: 1px solid #0F1F5F;"
			textohtml = textohtml & "	width:100px"
			textohtml = textohtml & "	}"
			textohtml = textohtml & "</style>"
			textohtml = textohtml & "<title>Carta enviada ao provedor</title>"
			textohtml = textohtml & "<script>"
			textohtml = textohtml & "	function Imprimir()"
			textohtml = textohtml & "	{"
			textohtml = textohtml & "		window.print();"
			textohtml = textohtml & "	}"
			textohtml = textohtml & "</script>"
			textohtml = textohtml & "</Head>"
			textohtml = textohtml & "<table align=center rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
			textohtml = textohtml & "<tr><td><br>EMBRATEL"
			textohtml = textohtml & "<div align=center>"
			textohtml = textohtml & ped("Cid_Desc") & ", " & day(date) & " de "
		
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
			strInterfaceEbt = ""
			if Not objRSFis.EOF and not objRSFis.BOF then
				strInterfacePto		= objRSFis("Acf_Interface")
				strInterfaceEbt		= objRSFis("Acf_InterfaceEstEntregaFisico")
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
			Set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & dblProId)
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
			
			
			'eyc
					
				Vetor_Campos(1)="adInteger,2,adParamInput," & ped("acf_id")
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_AssocFisicoTecnico",1,Vetor_Campos)
				Set objRS = db.Execute(strSqlRet)

				if not objRS.Eof then

					
					textohtml = textohtml & "	<table  border=0 cellspacing=1 cellpadding=0 > "
					textohtml = textohtml & "			<tr>"
					textohtml = textohtml & "				<th style=FONT-SIZE: 14px colspan=6>&nbsp;•&nbsp;Dados Técnicos</th>"
					textohtml = textohtml & "			</tr>"

					textohtml = textohtml & "			<tr class=clsSilver>"
					textohtml = textohtml & "				<td nowrap width=170>Taxa de Transmição</td>"
					textohtml = textohtml & "				<td>&nbsp;"
					
					textohtml = textohtml & "					<input type=text class=text name=cboVelocidade value=" & objRS("Vel_desc") & " size=30 maxlength=20  readonly=true>"
					textohtml = textohtml & "				</td> "
										

					textohtml = textohtml & "				<td align=right >Característica Técnica &nbsp;</td>"
					textohtml = textohtml & "				<td>&nbsp;"  
					textohtml = textohtml & "						<input type=text class=text name=cboCaracTec value=" & objRS("CaracteristicaTecnica_desc") & " size=30 maxlength=20  readonly=true> "
					textohtml = textohtml & "				</td> "
					
					textohtml = textohtml & "			</tr> "

					textohtml = textohtml & "			<tr class=clsSilver> "
					textohtml = textohtml & "				<td nowrap width=170>&nbsp;Q in Q</td> "
					textohtml = textohtml & "				<td>&nbsp; "
					
					textohtml = textohtml & "						<input type=text class=text name=rdoQinQ value=" & objRS("QinQ") & " size=30 maxlength=20  readonly=true> "
					textohtml = textohtml & "				</td> "
									

					textohtml = textohtml & "				<td align=right ><font class=clsObrig>:: </font>Quantidade de Linhas</td>"
					textohtml = textohtml & "				<td>&nbsp;"
					textohtml = textohtml & "					<input type=text class=text name=txtQtdLinhas value=" & objRS("QtdLinha") & " size=30 maxlength=20  readonly=true>"
									
					textohtml = textohtml & "				</td>"
					textohtml = textohtml & "			</tr>"

					textohtml = textohtml & "			<tr class=clsSilver>"
					textohtml = textohtml & "				<td nowrap width=170>&nbsp;Meio Preferencial</td>"
					textohtml = textohtml & "				<td>&nbsp;"							
					textohtml = textohtml & "					<input type=text class=text name=cboAplicacao value=" & objRS("MeioPreferencial_desc") & " size=30 maxlength=20  readonly=true>"
									
					textohtml = textohtml & "				</td>"
					
					textohtml = textohtml & "				<td align=right>&nbsp;Aplicação</td>"
					textohtml = textohtml & "				<td>&nbsp;"							
					textohtml = textohtml & "					<input type=text class=text name=cboAplicacao value=" & objRS("Aplicacao_desc") & " size=30 maxlength=20  readonly=true>"
									
					textohtml = textohtml & "				</td>"

					textohtml = textohtml & "			</tr>"
							
					textohtml = textohtml & "			<tr class=clsSilver>"
					textohtml = textohtml & "				<td width=170>Finalidade</td>"
					textohtml = textohtml & "				<td>&nbsp;"
									
					textohtml = textohtml & "					<input type=text class=text name=Finalidade value=" & objRS("Finalidade_desc") & " size=30 maxlength=20  readonly=true>"
									
					textohtml = textohtml & "				</td>"

					textohtml = textohtml & "				<td align=right>Prazo de Contratação</td>"
					textohtml = textohtml & "				<td colspan=3 nowrap>&nbsp;"
					
					textohtml = textohtml & "					<input type=text class=text name=cboPrazContr value='" & objRS("Prazo_contratacao_desc") & "' size=30 maxlength=20  readonly=true>"
									
					textohtml = textohtml & "				</td>"
					
					textohtml = textohtml & "			</tr>"

					textohtml = textohtml & "			<tr class=clsSilver>"
					textohtml = textohtml & "						<td colspan=6>"
					textohtml = textohtml & "							<P style=FONT-SIZE: 9pt>"
					textohtml = textohtml & "								A Característica Técnica definida pela fornecedora conforme disponibilidade constante no contrato, e não necessariamente, a que foi requisitada neste pedido."
					textohtml = textohtml & "								<br>"
					textohtml = textohtml & "								A Característica Técnica deve seguir conforme disponibilidade da Oferta de Referência."
					textohtml = textohtml & "							</p> "
					textohtml = textohtml & "						</td> "
					textohtml = textohtml & "					</tr> "
					textohtml = textohtml & "	</table> "
					
				end if
			
			'eyc
			
			
			
			
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
			textohtml = textohtml & "<tr><td>Interface</td><td>" & strInterfaceEbt & "</td></tr>"
			textohtml = textohtml & "</table>"

			Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId & ",null," & ped1("Sol_ID"))
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
'' LPEREZ  13/12/2005			
			textohtml = textohtml & "<table><tr><td colspan=2>Observação</td></tr>"
			textohtml = textohtml & "<tr><td colspan=2><strong>"
			'textohtml = textohtml & Trim(ped1("SOL_Obs")) & Trim(ped1("Ped_Obs")) & "</strong></font></td></tr>"
			textohtml = textohtml & Trim(ped1("Ped_Obs")) & "</strong></font></td></tr>"
			textohtml = textohtml & "</table><br><br><br>Atenciosamente,<br><br>"
'LP
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
			
			if dblSisId <> "3" then
				textohtml = textohtml & "<br><br>" & strContatoEBT
			else
				Set objRSResp = db.execute("CLA_sp_sel_provedoremail " & dblProId & ",null,'" & ped1("Est_Sigla") &"','"& ped1("Cid_Sigla") &"'") 
				If Not objRSResp.Eof and Not objRSResp.bof then
					textohtml = textohtml & "<br><br>" & Trim(objRSResp("cPRo_ContratanteContato"))
				Else
					textohtml = textohtml & "<br><br>" & " Ruben Nalin Filho"
				End if
			End if
			
			textohtml = textohtml & "<br>Gerente de Implantação de Acessos.<br><br>"
			textohtml = textohtml & "<hr>Embratel - Empresa Brasileira de Telecomunicações S.A."
			textohtml = textohtml & "</td></tr></table>"
			textohtml = textohtml & "</body></html>"
			
			' Set objRSConf = db.Execute("select * from cla_config where Config_ID = 4 and Config_Estado = 0 and Config_Data > getdate()")
			 'If Not objRSConf.eof and  not objRSConf.Bof Then
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\emailprovedor.htm",  true)
				objFile.WriteLine(textohtml)
				objFile.Close
				ObjMail.AttachFile ( strCaminho & "\CartasProvedor\emailprovedor.htm") 
				ObjMail.Body = "segue em anexo carta de solicitação de serviço referente: " &  AcaoPedidoEmail(ucase(ped1("tprc_id"))) & "  -  " & trim(ped("Cli_nome")) & "  -  " & ucase(ped1("Ped_Prefixo")) & "-" & right("00000" & ped1("Ped_Numero"),5) & "/" & ped1("Ped_Ano")
				ObjMail.Send
				
				Set objFile = objFSO.GetFile(strCaminho & "\CartasProvedor\emailprovedor.htm")
				objFile.Delete
				
				Set ObjMail = Nothing
				Set objFSO = Nothing 
				Set objFile = Nothing 
			 'End if
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
	
	strRet = "<table width=100% ><tr><td style=""text-align:center""><font color = red>E-Mail enviado com sucesso. "  & strProEmail & "</font></td></tr></table>" 
	strButton = "<table width=100% border=0>"
	strButton = strButton &	"<tr>"
	strButton = strButton &	"<td style=""text-align:center"">"
	strButton = strButton &	"<input  center type=button class=button name=btnImprimir value= Imprimir onClick=""Imprimir()"">&nbsp;"
	strButton = strButton &	"<input  center type=button class=button name=btnSair value=Sair onClick=""javascript:window.returnValue=0;window.close()""><br><br>"
	strButton = strButton &	"</td>"
	strButton = strButton &	"</tr></table>"

	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strRet & textohtml &  strButton)
	
	
	
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
	
%>
