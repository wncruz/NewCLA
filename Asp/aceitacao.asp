<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/FacilidadesPedido.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Aceitacao.asp
'	- Descrição			: Aceita um pedido

strDataAtual =  right("00" & day(now),2) & "/" & right("00" & month(now),2) & "/" & year(now)
dblPedId = Request.Form("hdnPedId") 
dblSolId = Request.Form("hdnSolId")

'<!-- Projeto ABW - Incio -->'
set rsProjABW = db.execute (" select top 1 aprovisi_id, indicador_migracao_pabx_virtual , indicador_tecnologia_ip " &_
							" from cla_aprovisionador " &_
							" where sol_id = " & dblSolId & " order by 1 desc")
			
if Not rsProjABW.Eof and Not rsProjABW.Bof then

	strMigraPabx_Virtual = Trim(rsProjABW("indicador_migracao_pabx_virtual"))

	IF ISNULL(strMigraPabx_Virtual) THEN
		strMigraPabx_Virtual = ""
	END IF
	
	IF strMigraPabx_Virtual = "NULL" THEN
		strMigraPabx_Virtual = ""
	END IF

	IF strMigraPabx_Virtual = "N" THEN
		strMigraPabx_Virtual = "Não"
	END IF

	IF strMigraPabx_Virtual = "S" THEN
		strMigraPabx_Virtual = "Sim"
	END IF

	strTecnologia_IP = Trim(rsProjABW("indicador_tecnologia_ip"))

	IF ISNULL(strTecnologia_IP) THEN
		strTecnologia_IP = ""
	END IF
	
	IF strTecnologia_IP = "NULL" THEN
		strTecnologia_IP = ""
	END IF

	IF strTecnologia_IP = "N" THEN
		strTecnologia_IP = "Não"
	END IF

	IF strTecnologia_IP = "S" THEN
		strTecnologia_IP = "Sim"
	END IF

end if
'<!-- Projeto ABW - Fim -->'

If dblPedId <> "" then
	
	strNroSolic		= dblSolId 'Trim(objRS("Sol_id"))

	If trim(strLoginRede) = "JCARTUS" Then	
		'	response.write "<script>alert('strNroSolic="&strNroSolic&"')</script>"
	end if
	
	set ObjFO = db.execute("select top 1 cla_acessofisico.Pro_ID from cla_solicitacao WITH (NOLOCK) inner join cla_acessologico 		WITH (NOLOCK) on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	WITH (NOLOCK) on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		WITH (NOLOCK) on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id where cla_acessologicofisico.alf_flagAux is null and cla_acessofisico.pro_id=145 and sol_id = " & Request.Form("hdnSolId"))
	
	pro_id_fo = ""
	
	if not ObjFO.Eof and not ObjFO.Bof then
		pro_id_fo = ObjFO("Pro_ID")
	end if

	if pro_id_fo = "145" then  
	  Set objRS = db.execute("CLA_sp_view_pedido_gpon " & strNroSolic) 		
	else
		Set objRS = db.execute("CLA_sp_view_pedido_gpon null,null,null,null,null,null," & dblPedId)
	end if
	
	dbltec_ID_Ter_SICI	= objRS("Tec_ID_Ter_SICI")
	strProId			= Trim(objRS("Pro_id"))
	
	strPropAcesso		= Trim(objRS("Acf_Proprietario"))
	Set objRSol 		= db.execute("select Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & strNroSolic)

	if not objRSol.Eof and not objRSol.bof then
		strIdLogico		= Trim(objRSol("Acl_IDAcessoLogico"))
 	end if 

	strDataSolic		= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM				= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido		= Formatar_Data(objRS("Ped_Data"))
	strAcao				= AcaoPedido(objRS("Tprc_Id"))
	strSatus			= objRS("Sts_Desc")
	strCliente			= objRS("Cli_Nome")
	intTipoProcesso 	= objRS("tprc_id")
	strObs				= objRS("Ped_Obs")
	dblSolId			= objRS("Sol_Id") 
	dblStsId			= objRS("sts_id")
	dbltec_ID			= objRS("Tec_ID")
	
	set ObjFOdm = db.execute("select top 1 Ped_Prefixo, Ped_Numero , Ped_Ano  from cla_pedido where ped_id =  " & dblPedId )
	if not ObjFOdm.Eof and not ObjFOdm.Bof then
		strDM			= UCASE(ObjFOdm("Ped_Prefixo") & "-" & right("00000" & ObjFOdm("Ped_Numero"),5) & "/" & ObjFOdm("Ped_Ano"))
	end if
	
	if isnull(dbltec_ID) then
		dbltec_ID = ""
	end if

	set ObjSnoa = db.execute("select top 1 pedido_compra_snoa  from cla_Assoclogicosnoa where sol_id =  " & dblSolId & " and pedido_compra_snoa is not null ")
	if not ObjSnoa.Eof and not ObjSnoa.Bof then
		strSnoa = UCASE(ObjSnoa("pedido_compra_snoa"))
	end if

	'response.write "<script>alert('"&dbltec_ID&"')</script>"
	'response.write "<script>alert('"&strIdLogico&"')</script>"
	if dbltec_ID = "7" then
		Set objRSSwitch = db.execute("select switch, switch_porta from cla_switch inner join cla_acessologico on cla_switch.acl_idacessologico = cla_acessologico.acl_idacessologico where cla_acessologico.acl_idacessologico = " & strIdLogico)
		if not objRSSwitch.Eof and not objRSSwitch.bof then
			strSwitch		= objRSSwitch("switch")
			strSwitchPorta	= objRSSwitch("switch_porta")
		end if 
	end if 
	
	strDtIniTemp 	= Formatar_Data(Trim(objRS("Acl_DtIniAcessoTemp")))
	strDtFimTemp 	= Formatar_Data(Trim(objRS("Acl_DtFimAcessoTemp")))
	strDtDevolucao 	= Formatar_Data(Trim(objRS("Acl_DtDevolAcessoTemp")))
				
	Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
	if not objRSEndPto.Eof and not objRSEndPto.bof then
		strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
	End if	
	Set objRSEndPto = Nothing
		
	strCidSigla		= Trim(objRS("Cid_Sigla"))
	strUfSigla		= Trim(objRS("Est_Sigla"))
	strTplSigla 	= Trim(objRS("Tpl_Sigla")) 
	strNomeLogr		= Trim(objRS("End_NomeLogr")) 
	strNroEnd		= Trim(objRS("End_NroLogr"))
	strCep			= Trim(objRS("End_Cep"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis	= Trim(objRS("DescVelAcessoFis"))
	
	strPabx_Virtual = Trim(objRS("PABX_VIRTUAL"))

	strPrmId		= Trim(objRS("Prm_id"))
	strRegId		= Trim(objRS("Reg_id"))
	strLocalInstala = Trim(objRS("Esc_IdEntrega"))
	strLocalConfig	= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso		= Trim(objRS("Rec_IDEntrega"))
	strDistrib		= objRS("Dst_Id")
	strRede			= objRS("Sis_ID")
	
	if Trim(strRecurso) <> "" then
		Set objRSRec = db.execute("CLA_sp_view_recurso " & strRecurso)
		if Not objRSRec.Eof And Not objRSRec.Bof then
			strDistribDesc	= objRSRec("Dst_Desc")
			strRedeDesc		= objRSRec("Sis_Desc")
			strRede			= objRSRec("Sis_ID")
			strContato		= objRSRec("Esc_Contato")
			strTelefone		= objRSRec("Esc_Telefone")
			strProId		= Trim(objRSRec("Pro_id"))
		End if
	End if	
	
	'strPropEquip	= objRS("Ped_ProprietarioEquip")
	'intQtdEquip		= objRS("Ped_QtdEquip")
	strUserGicL		= strUserName

	set objRSFis = db.execute("CLA_sp_view_acessofisico null," & dblPedId)
	if Not objRSFis.Eof and not objRSFis.Bof then
		strTecnologia	= objRSFis("Tec_Sigla")
		strVelFis		=	objRSFis("Vel_Desc")
		intTipoVel		=	objRSFis("Acf_TipoVel")
		strAcf_ProjetoEspecial = objRSFis("Acf_ProjetoEspecial")
	    strAcf_ProjetoEspecialCustoEst = objRSFis("Acf_ProjetoEspecialCustoEst")
	    strAcf_ProjetoEspecialCustoAprov = objRSFis("Acf_ProjetoEspecialCustoAprov")
	End if
	Set objRSFis = Nothing

Else
	Response.Write "<script language=javascript>window.location.replace('main.asp')</script>"
	Response.End
End if

'GPON
if trim(dbltec_ID) = "6" or trim(dbltec_ID) = "7" then
	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adInteger,2,adParamInput," &  strIdLogico
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_FacilidadeGPON",2,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)
	if Not objRS.Eof and not objRS.Bof then
		dblSolId = objRS("Sol_ID")
	end if
		strNroSolic = dblSolId
end if
%>

<tr>
	<td>
<table cellspacing="1" cellpadding="1" border=0 width="760">
<form name="f" action="aceitacao.asp" method="post" >
<input type="hidden" name="id" value="<%=Server.HTMLEncode(request("id"))%>">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnStatus" value="<%=dblStsId%>">
<input type="hidden" name="hdnSolId" value="<%=dblSolId%>">
<input type="hidden" name="hdnPedId" value="<%=dblPedId%>" >
<input type="hidden" name="hdnIdLog" value="<%=strIdLogico%>">

<input type="hidden" name="hdnIdFis" value="<%=strAcessoFis%>">
<input type="hidden" name="hdnAcfId" value="<%=dblAcfId%>">
<input type="hidden" name="hdnCidSigla" value="<%=strCidSigla%>">
<input type="hidden" name="hdnPropAcesso" value="<%=strPropAcesso%>">
<input type="hidden" name="cboLocalConfig" value="<%=Request.Form("cboLocalConfig")%>">

<tr><th colspan=6 ><p align=center >Aceitação</th></tr>

<tr>
	<th style="FONT-SIZE: 14px" colspan=5 >&nbsp;•&nbsp;Informações Gerais</th>
	<th align=rigth><a href="javascript:DetalharFac()"><font color=white>Mais...</font></a></th>
</tr>

<tr class=clsSilver>
	<td nowrap width=170>Solicitação de Acesso Nº</td>
	<td>&nbsp;<%=strNroSolic%></td>
	<td align=right >Id Lógico</td>
	<td>&nbsp;<%=strIdLogico%></td>
	<td align=right >Data</td>
	<td>&nbsp;<%=strDataSolic%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Pedido de acesso</td>
	<td>&nbsp;<%=strDM%></td>
	<td align=right>Data</td>
	<td colspan=3 nowrap>&nbsp;<%=strDataPedido%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Ação</td>
	<td colspan=5>&nbsp;<%=strAcao%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Status</td>
	<td colspan=5>&nbsp;<%=strSatus%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Cliente</td>
	<td colspan="5">&nbsp;<%=strCliente%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Endereço</td>
	<td colspan="5">&nbsp;<%=strEndereco%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Nº Contrato Serviço</td>
	<td nowrap>&nbsp;<%=strNroServico%></td>
	<td align=right>Designação do Serviço</td>
	<td colspan=3>&nbsp;<%=strDesigServico%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Serviço</td>
	<td colspan=3 >&nbsp;<%=strServico%> <%IF strPabx_Virtual <> "" then %> PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %><br><%END IF%></td>
	<td align=right>Velocidade do Serviço</td>
	<td colspan=3>&nbsp;<%=strVelServico%></td>
</tr>
	<!-- Projeto ABW - Incio -->
			<tr class=clsSilver>
				<td width=170>Migração do PABX Virtual</td>
				<td>&nbsp;<%=strMigraPabx_Virtual%></td>
				<td align=right>Tecnologia IP</td>
				<td colspan=3 nowrap>&nbsp;<%=strTecnologia_IP%></td>
			</tr>
			<!-- Projeto ABW - Fim -->

<tr class=clsSilver>
	<td width=170>Velocidade do Acesso Físico</td>
	<td colspan=5>&nbsp;<%=strVelFis%>&nbsp;<%=TipoVel(intTipoVel)%></td>
</tr>

<tr class="clsSilver">
	<td width=170>Observações</td>
	<td colspan="5"><textarea name="txtObsSolic" cols="50" rows="2" disabled><%=strObs%></textarea></td>
</tr>

<tr class="clsSilver">
	<td width=170>Projeto Especial de Acesso</td>
	<td colspan="5">&nbsp;<%=strAcf_ProjetoEspecial%></td>
</tr>

<tr class="clsSilver">
	<td width=170>Custo Previsto (R$)</td>
	<td colspan="5">&nbsp;<%=strAcf_ProjetoEspecialCustoEst%></td>
</tr>

<tr class="clsSilver">
	<td width=170>Custo Aprovado (R$)</td>
	<td colspan="5">&nbsp;<%=strAcf_ProjetoEspecialCustoAprov%></td>
</tr>
</table>

		<% 'GPON
if trim(dbltec_ID) <> "6"  and  trim(dbltec_ID) <> "7" then%>
<table border=0 cellspacing="1" cellpadding="0" width="760">  
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend" 
				    width       = "100%" 
				    height      = "120px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>"
				    frameborder = "0"
				    scrolling   = "no" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
			<% 	if trim(strSnoa) <> ""  then%> 
					<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
							<td>
								<iframe	id			= "IFrmListaStatusSNOA"
									    name        = "IFrmListaStatusSNOA"
									    width       = "100%"
									    height      = "160px"
									    src			= "../inc/ListaStatusSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=1"
									    frameborder = "0"
									    scrolling   = "auto"
									    align       = "left">
								</iFrame>
							</td>
	</tr>
					</table>
					
			 <%End if %>

		<!-- CH-56418KBG - Inicio -->
		<%

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_designacaoFornecedora ",1,Vetor_Campos)
			Set objRS_DSF = db.Execute(strSqlRet)

			if not objRS_DSF.eof then %>

				<table border=0 cellspacing="1" cellpadding="0" width="760">
					<tr>
						<td>
							<iframe	id			= "IFrmListaDesigFornecSNOA"
								    name        = "IFrmListaDesigFornecSNOA"
								    width       = "100%"
								    height      = "60px"
								    src			= "../inc/ListaDesigFornecSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=<%=telaaceitar%>"
								    frameborder = "0"
								    scrolling   = "no"
								    align       = "left">
							</iFrame>
						</td>
					</tr>
				</table>

			<% end if %>	
		<!-- CH-56418KBG - Fim -->

		<table border=0 cellspacing="1" cellpadding="1" width="760">
			
			<!--<tr><th colspan=4 >&nbsp;•&nbsp;Recurso</th></tr>-->
			<tr><th colspan=4 style="FONT-SIZE: 14px" >&nbsp;•&nbsp;Recurso</th></tr>
			
	<tr class=clsSilver>
		<td width="130">Local de Entrega</td>
		<td colspan=3>&nbsp;
			<%set objRStemp = db.execute("CLA_sp_sel_estacao  " & strLocalInstala)
				While not objRStemp.Eof 
					if Trim(strLocalInstala) = Trim(objRStemp("Esc_ID")) then 
						Response.Write objRStemp("Cid_Sigla") & "  " & objRStemp("Esc_Sigla")
					End if	
					objRStemp.MoveNext
				Wend
			%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=150px nowrap>Local de Configuração</td>
		<td colspan=3>&nbsp;
			<%set objRStemp = db.execute("CLA_sp_sel_estacao  " & strLocalConfig)
				While not objRStemp.Eof 
					if Trim(strLocalConfig) = Trim(objRStemp("Esc_ID")) then 
						Response.Write objRStemp("Cid_Sigla") & "  " & objRStemp("Esc_Sigla")
					End if	
					objRStemp.MoveNext
				Wend
			%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td>Distribuidor</td>
		<td colspan=3>&nbsp;
			<%=strDistribDesc%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td>Rede</td>
		<td colspan=3>&nbsp;
			<%=strRedeDesc%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=30% >Provedor</td>
		<td colspan=3>&nbsp;
			<%	set objRStemp = db.execute("CLA_sp_sel_provedor 0")
				While not objRStemp.Eof 
					if Trim(strProId) = Trim(objRStemp("Pro_ID")) then 
						Response.Write objRStemp("Pro_Nome")
					End if	
					objRStemp.MoveNext
				Wend
			%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td>Prazos de Contratação de Acesso</td>
		<td colspan=3>&nbsp;
			<%
				if Trim(strProID) <> "" then
					set objRStemp = db.execute("CLA_sp_sel_regimecontrato 0," & strProID)
					While not objRStemp.Eof 
						if Trim(strRegId) = Trim(objRStemp("Reg_ID")) then 
							Response.Write LimparStr(Trim(objRStemp("Pro_Nome"))) & " - " & LimparStr(Trim(objRStemp("Tct_Desc")))
						End if
						objRStemp.MoveNext
					Wend
				End if			
			%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td>Promoção</td>
		<td colspan=3>&nbsp;
			<%
				if Trim(strProId) <> "" then
					set objRStemp = db.execute("CLA_sp_sel_promocaoprovedor 0," & strProId)
					While not objRStemp.Eof 
						if Trim(strPrmId) = Trim(objRStemp("Prm_ID")) then 
							Response.Write objRStemp("Prm_Desc")
						End if	
						objRStemp.MoveNext
					Wend
				End if	
			%>
		</td>
	</tr>

	<tr class=clsSilver>
		<td>Contato</td>
		<td colspan=3>&nbsp;<%=strContato%></td>
	</tr>
	
	<tr class=clsSilver>
		<td>Telefone</td>
		<td colspan=3>&nbsp;<%=strTelefone%></td>
	</tr>

	<tr class="clsSilver">
		<td rowspan=2>Acesso Temporário<br>(dd/mm/aaaa)</td>
		<td >&nbsp;Início&nbsp;</td>
		<td >&nbsp;Fim&nbsp;</td>
		<td >&nbsp;Devolução&nbsp;</td>
	</tr>

	<tr class="clsSilver">
		<td>&nbsp;<%=strDtIniTemp%></td>
		<td>&nbsp;<%=strDtFimTemp%></td>
		<td>&nbsp;<%=strDtDevolucao%></td>
	</tr>
	
</table>

<%
	Set objRSol = db.execute("select acf_id from cla_pedidofisico where ped_id = " & dblPedId)
	if not objRSol.Eof and not objRSol.bof then
		dblAcfId		= Trim(objRSol("acf_id"))
 	end if 
	
	'response.write "<script>alert('"&dblAcfId&"')</script>"
	set rsSwicth = db.execute (" select distinct cla_recurso.sis_id , cla_switch.switch_id  , cla_switch.switch " &_
		" from  cla_Acessofisico		inner join cla_facilidade		on cla_Acessofisico.acf_id = cla_facilidade.acf_id " &_
		" inner join cla_ontporta			on cla_Acessofisico.acf_id = cla_ontporta.acf_id " &_
		" inner join cla_switch			on cla_ontporta.switch_id = cla_switch.switch_id " &_
		" inner join cla_ontvlan			on cla_ontporta.ontporta_id = cla_ontvlan.ontporta_id " &_
		" inner join cla_recurso			on cla_facilidade.rec_id = cla_recurso.rec_id " &_		
		" where cla_Acessofisico.acf_id = " & dblAcfId)
		
	if not rsSwicth.Eof then ' and objRSSolic.Bof) then  	

		if strRede = 12 then
				
			set objRS = db.execute("CLA_sp_sel_SwitchRadio " & rsSwicth("Switch_ID") )
				if not objRS.eof then	
			%>

					<table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760 >

						<tr> 
							<th colspan=10 class=clsSilver >&nbsp;&nbsp; Controle de Ocupacao</th> 
						</tr> 

						<tr class=clsSilver > 
							<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; <%=objRS("Font_Nome") %> </td><td>&nbsp;&nbsp;</td> 
						</tr> 

						<tr class=clsSilver > 
							<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; <%= objRS("Tont_Modelo") %> </td><td>&nbsp;&nbsp;</td> 
						</tr> 

						<tr class=clsSilver > 
							<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; <%= objRS("Tont_QtdPorta") %> </td><td>&nbsp;&nbsp;</td> 
						</tr> 		

						<tr class=clsSilver > 
							<td>&nbsp;&nbsp;Switch</td><td>&nbsp;&nbsp; <%= rsSwicth("switch") %> </td><td>&nbsp;&nbsp;</td> 
						</tr> 	
						
				</table>
						
					<table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760> 
					<tr> 
						<th width=10>&nbsp;&nbsp;Porta</th> 
						<th>&nbsp;&nbsp;Tipo</th> 
						<th>&nbsp;&nbsp;Servico</th> 
						<th>&nbsp;&nbsp;Designacao Servico</th> 
						<th>&nbsp;&nbsp;Designacao Radio IP CLI</th> 
						<th>&nbsp;&nbsp;ID Logico</th> 
						<th>&nbsp;&nbsp;VLAN</th> 
						<th>&nbsp;&nbsp;Porta PE</th> 
						<th>&nbsp;&nbsp;SVLAN</th> 
						<th >&nbsp;&nbsp;PE</th> 
					</tr> 
				<%
						i=1
						While Not objRS.Eof
					if blnCor then	%>
								 <tr class=clsSilver>
				 		<%blnCor = false
							Else
							%>
								<tr class=clsSilver2>
						<%blnCor = true
							End if
							%>
									<td>&nbsp;&nbsp;<%= objRS("ONTPorta_Porta")  %> </td> 
									<td>&nbsp;&nbsp;<%= objRS("TPPorta_Abrev")  %></td> 
									<td>&nbsp;&nbsp;<%= objRS("ser_Desc") %></td> 			
									<td>&nbsp;&nbsp;<%= objRS("Acl_DesignacaoServico") %> </td> 
									<td><%= objRS("acf_nroAcessoPtaEbt") %></td> 
									<td>&nbsp;&nbsp;<%= objRS("Acl_IDAcessoLogico") %> </td> 
									<td>&nbsp;&nbsp;	<%=objRS("OntVLAN_Nome")%> </td> 
									<td>&nbsp;&nbsp; <%=objRS("OntVLAN_portaolt") %> </td> 
									<td>&nbsp;&nbsp; <%= objRS("OntSVLAN_Nome") %> </td> 
									<td>&nbsp;&nbsp; <%= objRS("OntVLAN_Pe") %> </td> 			
							</tr>
							<%
							i=i+1
							objRS.MoveNext
						Wend
						
				end if	
				
		end if		
	
	end if
			%>	
			</table> 

	<%if dbltec_ID = "7" then %>

		<table border=0 cellspacing=1 cellpadding=0  width=760>
			<tr><th style="FONT-SIZE: 14px" colspan=4 >&nbsp;•&nbsp;Switch&nbsp;</th></tr>
			<tr class=clsSilver>
				<td>Switch</td>
				<td>&nbsp;<%=strSwitch%></td>
			</tr>
			<tr class=clsSilver>
				<td>Porta do Switch</td>
				<td >&nbsp;<%=strSwitchPorta%></td>
			</tr>
		</table>
	<% end if %>
				
				<% if strTecnologia <> "FO EDD" or isnull(strTecnologia) then 
	
		strLink =	"<td colspan=2 align=right class=clsSilver2 width=45px>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Aceitação " & strDM & " - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td>"
		
		Response.Write "<table border=0 cellspacing=0 cellpadding=0  width=760>"
		Response.Write "<tr><th style=FONT-SIZE: 14px>&nbsp;•&nbsp;Facilidades do Pedido&nbsp;</td></tr>" & strLink
		Response.Write "</tr></table>"
		Set objRSFac = db.execute("CLA_sp_sel_facilidade " & dblPedId)
		strXls = FacilidadesPedido(objRSFac)
		Response.Write strXls
		
	end if 

'GPON
else

	
Vetor_Campos(1)="adInteger,2,adParamInput,"
Vetor_Campos(2)="adInteger,2,adParamInput," &  strIdLogico
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_FacilidadeGPON",2,Vetor_Campos)
	
	Set objRS = db.Execute(strSqlRet)
	
	if objRS.eof then
		%>
		<script>alert("Facilidade não alocada.")
		</script>
		<%
		response.end
	end if
	
	strSQL = "select sum(OPqtd_Qtd) as Tont_QtdPorta from CLA_OntQtdPorta where Tont_ID = " & objRS("Tont_ID")
	Set objRSqtd = db.Execute(strSQL)
	
	%>	
	<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=10 class="clsSilver" style="FONT-SIZE: 14px">&nbsp;•&nbsp; Controle de Ocupação da ONT | EDD</th>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; <%=objRS("Font_Nome")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; <%=objRS("Tont_Modelo")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; <%=objRSqtd("Tont_QtdPorta")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Designação ONT | EDD</td><td>&nbsp;&nbsp;<%=trim(objRS("Ont_Desig"))%></td><td>&nbsp;&nbsp;</td>
		</tr>
	</table>
	<%if dbltec_ID = "7" then %>
	
		<table border=0 cellspacing=1 cellpadding=0  width=760>
			<tr><th colspan=4 style="FONT-SIZE: 14px">&nbsp;•&nbsp;Controle do Switch&nbsp;</th></tr>
			<tr class=clsSilver>
				<td>Switch</td>
				<td>&nbsp;<%=strSwitch%></td>
			</tr>
			<tr class=clsSilver>
				<td>Porta do Switch</td>
				<td >&nbsp;<%=strSwitchPorta%></td>
			</tr>
		</table>
	<% end if %>
	<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=10>&nbsp;•&nbsp;</th>
		</tr>
		<tr>
			<th>&nbsp;&nbsp;Porta</th>
			<th>&nbsp;&nbsp;Tipo</th>
			<th>&nbsp;&nbsp;Serviço</th>
			<th>&nbsp;&nbsp;Designação Serviço</th>
			<th>&nbsp;&nbsp;ID Lógico</th>
			<th>&nbsp;&nbsp;VLAN</th>
		</tr>
		<input type="Hidden" name="hdnOntID" value="<%=objRS("ONT_ID")%>">
		<input type="Hidden" name="hdnOrisol_ID" value="<%=orisolid%>">
		<input type="Hidden" name="hdnAcl_IDAcessoLogico" value="<%=strIdLogico%>">		
		<%
		i=1
		While Not objRS.Eof
			if blnCor then
				response.write "<tr class=clsSilver>"
				blnCor = false
			Else
				response.write "<tr class=clsSilver2>"
				blnCor = true
			End if%>
			<td>&nbsp;&nbsp;<%=objRS("ONTPorta_Porta")%></td>
			<td>&nbsp;&nbsp;<%=objRS("TPPorta_Abrev")%></td>
			<td>&nbsp;&nbsp;<%=objRS("ser_Desc")%></td>
			<td>&nbsp;&nbsp;<%=objRS("Acl_DesignacaoServico")%></td>
			<td>&nbsp;&nbsp;<%=objRS("Acl_IDAcessoLogico")%></td>
			<td>&nbsp;&nbsp;<%=objRS("OntVLAN_Nome")%></td>
		</tr>		
			<%
			i=i+1
			objRS.MoveNext
		Wend
		
	%>
	</table>
	<%
end if
%>
<script language = "JavaScript">
function GravarAceitacao()
{
	with (document.forms[0])
	{
		if (!IsEmpty(txtDtAceite.value))
		{
			if (!ValidarTipoInfo(txtDtAceite,1,"Data do aceite")) return ;
		
		}else{//Obrigar preenchimento da Data de Aceite
			alert("Favor informar a Data de Aceite.")
			txtDtAceite.focus()
			return			
		}
		
		if ( cboTipoAcesso.value == "0" && hdnPropAcesso.value=="TER")
		{
			alert("Favor informar a Tecnologia de Acesso. \nValidar informação junto à operadora, pois é passível de auditoria pela Anatel.")
			cboTipoAcesso.focus()
			return
		}
		
		if ( cboTipoAceite.value == "" )
		{
			alert("Favor informar o Tipo de Aceite.")
			cboTipoAceite.focus()
			return
		}
		
		if (!CompararData(txtDtAceite.value,'<%=strDataAtual%>',1,"A Data de Aceitação deve ser menor ou igual a data atual.")) 
		{
			txtDtAceite.focus()
			return
		}	
		if (!ValidarCampos(matriculatecebt,"Matrícula do Técnico EBT")) return
		if (!ValidarCampos(nometeceot,"Nome do Técnico PST")) return

		hdnAcao.value = "GravarAtivacao"
		target = "IFrmProcesso"
		action = "ProcessoAceitacao.asp"
		submit()
	}
}

function JanelaConfirmacaoFac(objXmlGeral){
	var intRet = window.showModalDialog('ConfirmacaoAceitacao.asp',objXmlGeral,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet == 164){
		intRet = 0
		VoltarConsultaAceitacao()
	}
} 

function VoltarConsultaAceitacao()
{
	with (document.forms[0])
	{
		//Retorna para tela de consulta se a facilidade estive sido aceita
		hdnAcao.value = ""
		target = self.name 
		action = "Aceitacao_main.asp"
		submit()
		
	}	
}
</script>
<tr>
<td colspan="4">
<table cellpadding="1" cellspacing="1" border=0 width="760">
<%if dblPedId <> "" then
	Set objRSHis = db.execute("CLA_sp_sel_historicoAceitacao "& dblPedId)
	if Not objRSHis.eof and not objRSHis.bof then
		Response.Write "<tr>"
		Response.Write "<th>&nbsp;Data Teste</th>"
		Response.Write "<th>Nº Acesso</th>"
		Response.Write "<th>Mat. Téc. EBT</th>"
		Response.Write "<th>Nome Téc. EOT</th>"
		Response.Write "<th>Observação</th>"
		Response.Write "</tr>"
		do while not objRSHis.eof
			%>
				<tr class=clsSilver>
				<td><%if not isnull(objRSHis("Hia_DtTeste")) then response.write Formatar_Data(objRSHis("Hia_DtTeste")) end if %></td>
				<td><%if not isnull(objRSHis("Acf_NroAcessoPtaEBt")) then response.write objRSHis("Acf_NroAcessoPtaEBt") end if %></td>
				<td><%if not isnull(objRSHis("Hia_MatriculaEBT")) then response.write objRSHis("Hia_MatriculaEBT") end if %></td>
				<td><%if not isnull(objRSHis("Hia_TecnicoEBT")) then response.write objRSHis("Hia_TecnicoEBT") end if %></td>
				<td><%=trim(objRSHis("Hia_Obs"))%></td>
				</tr>
			<%
			objRSHis.movenext
		loop
	End if
End if
%>
</table>
</td>
</tr>

<%if strPropAcesso = "TER" then %>
<tr>
<% else %>
<tr style="display:none;">
<% end if %> 

<td colspan="4">
	
	<table border=0 cellspacing="1" cellpadding="1" width="760" >
		<th colspan="4" style="FONT-SIZE: 14px">&nbsp;•&nbsp;Tecnologia de Acesso</th>

<!-- Setar tecnologia do Acesso Terceiro-->
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Tecnologia de acesso</td>
<td colspan=3>
	<select name="cboTipoAcesso" >
		<option value="0"></option>	
		<%
			strSQL = "select tec_id_ter , Tec_Nome_Ter from cla_tecnologia_ter_sici" 
   			Set objRS = db.execute(strSQL)
			While Not objRS.Eof
				if dbltec_ID_Ter_SICI = CDbl(objRS("tec_id_ter")) then			
					Response.Write "<Option selected value=" & objRS("tec_id_ter")&">" & objRS("Tec_Nome_Ter") & "</Option>"
				else
					Response.Write "<Option value=" & objRS("tec_id_ter")&">" & objRS("Tec_Nome_Ter") & "</Option>"
				end if
				objRS.MoveNext
			wend
		%>
	</select>
	<font color='red'>(Utilizada nos relatórios SICI-Anatel, passível de auditoria pelo órgão)</font>
</td>
</tr>
</table>
</td>
</tr>


<tr>
<td colspan="4">
<table border=0 cellspacing="1" cellpadding="1" width="760" >
<th style="FONT-SIZE: 14px" colspan="4" >&nbsp;•&nbsp;Dados do Teste de Acesso</th>
<tr class=clsSilver>
	<td>Número de acesso:</td>
	<td colspan="3">
	<input type="hidden" name="hdnNroAcesso" size="30" maxlength="30" value="<%=Request.Form("hdnNroAcesso")%>">
	<%=Request.Form("hdnNroAcesso")
	
	%>
	</td>
</tr>

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Matrícula Téc. EBT</td>
<td><input type="text" class="text" name="matriculatecebt" size="6" maxlength="6"></td>
<td>Data de Aceite</td>
<td><input type="text" value="<%=Formatar_Data(date())%>" class="text" name="txtDtAceite" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" maxlength="10" size="10" > dd/mm/aaaa</td>
</tr>

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Nome do Téc. PST</td>
<td><input type="text" class="text" name="nometeceot" size="30" maxlength="30"></td>
<td><font class="clsObrig">:: </font>Tipo de Aceite</td>
<td>
	<select name="cboTipoAceite" >
		<option value=""></option>
		<%
			strSQL = "select tp_idaceite , tp_tipoaceite from cla_tipoAceite" 
   			Set objRS = db.execute(strSQL)
			While Not objRS.Eof
				Response.Write "<Option value=" & objRS("tp_idaceite")&">" & objRS("tp_tipoaceite") & "</Option>"
				objRS.MoveNext
			wend
		%>
	</select>
</td>
</tr>

<tr class=clsSilver>
	<td>Observação</td>
	<td colspan="3">
					<textarea name="observacao" cols="90" rows="3" onKeyPress="MaxLength(this,300)"></textarea>
	</td>
</tr>

<tr align=center>
	<td align="center" colspan="4">
		<input type="button" name="btnGravar" value="Gravar" class="button" onclick="GravarAceitacao()" accesskey="I" onmouseover="showtip(this,event,'Gravar (Alt+I)');">
		<input type="button" name="btnVoltar" value="Voltar" class="button" onclick="VoltarConsultaAceitacao()" accesskey="B" onmouseover="showtip(this,event,'Gravar (Alt+B)');">
		<input type="button" class="button" name="saida" value="   Sair   " onclick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>

<br>

<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>

<center>
</td>
</tr>
</table>
<br>
<br>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>?acao=<%=Trim(Server.HTMLEncode(Request("acao")))%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="Aceitação">

</form>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

</body>
</html>
