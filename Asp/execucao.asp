<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Execucao.asp
'	- Responsável		: Vital
'	- Descrição			: Efetiva a execução para Tronco/Par
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr><td >
<table cellspacing="1" cellpadding="1" border=0 width="760">
<form action="execucao.asp" method="post" name="f" >
<input type="hidden" name="hdnAcao">
<input type="hidden" name="id" value="<%=request("hdnPedId")%>">
<!--Para retornar para a tela anterior com a mesma seleção-->
<input type="hidden" name="cboLocalConfig" value="<%=Request.Form("cboLocalConfig")%>">
<%
dblPedId = Request.Form("hdnPedId") 
If dblPedId <> "" then

	Set objRS = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId)
	
	strProId		= Trim(objRS("Pro_id"))
	strNroSolic		= Trim(objRS("Sol_id"))
	strPropAcesso	= Trim(objRS("Acf_Proprietario"))
	strIdLogico		= Trim(objRS("Acl_IDAcessoLogico"))
	strDataSolic	= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM			= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido	= Formatar_Data(objRS("Ped_Data"))
	strAcao			= AcaoPedido(objRS("Tprc_Id"))
	strSatus		= objRS("Sts_Desc")
	strCliente		= objRS("Cli_Nome")
	intTipoProcesso = objRS("tprc_id")
	strObs			= objRS("Ped_Obs")
	dblSolId		= objRS("Sol_Id") 
	dblStsId		= objRS("sts_id")

	strDtIniTemp = Formatar_Data(Trim(objRS("Acl_DtIniAcessoTemp")))
	strDtFimTemp = Formatar_Data(Trim(objRS("Acl_DtFimAcessoTemp")))
	strDtDevolucao = Formatar_Data(Trim(objRS("Acl_DtDevolAcessoTemp")))
				
	Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
	if not objRSEndPto.Eof and not objRSEndPto.bof then
		strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
	End if	
	Set objRSEndPto = Nothing
		
	strCidSigla	= Trim(objRS("Cid_Sigla"))
	strUfSigla	= Trim(objRS("Est_Sigla"))
	strTplSigla = Trim(objRS("Tpl_Sigla")) 
	strNomeLogr	= Trim(objRS("End_NomeLogr")) 
	strNroEnd	= Trim(objRS("End_NroLogr"))
	strCep		= Trim(objRS("End_Cep"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis	= Trim(objRS("DescVelAcessoFis"))

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
			dblDstId		= Trim(objRSRec("Dst_id"))
			dblEscId		= Trim(objRSRec("Esc_id"))
		End if
	End if	
	'strPropEquip	= objRS("Ped_ProprietarioEquip")
	'intQtdEquip		= objRS("Ped_QtdEquip")

	strUserGicL		= strUserName
Else
	Response.Write "<script language=javascript>window.location.replace('main.asp')</script>"
	Response.End
End if
%>
<input type="hidden" name="recurso" value="<%=strRecurso%>">
<input type="hidden" name="hdnPedId" value="<%=dblPedId%>" >
<input type="hidden" name="hdnSolId" value="<%=dblSolId%>">
<input type="hidden" name="hdnStatus" value="<%=strStatus%>">
<input type=hidden name=hdnRede value="<%=strRede%>">
<input type=hidden name=hdnDstId value="<%=dblDstId%>">
<input type=hidden name=hdnEscId value="<%=dblEscId%>">

<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">

<tr><th colspan=6 ><p align=center >Execução</th></tr>

<tr>
	<th colspan=5 >&nbsp;•&nbsp;Informações Gerais</th>
	<th align=rigth >	
		<a href="javascript:DetalharFac()"><font color=white>Mais...</font></a>
	</th>
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
	<td >&nbsp;<%=strServico%></td>
	<td align=right>Velocidade do Serviço</td>
	<td colspan=3>&nbsp;<%=strVelServico%></td>
</tr>
<tr class="clsSilver">
	<td width=170>Observações</td>
	<td colspan="6"><textarea name="txtObsSolic" cols="50" rows="2" disabled><%=strObs%></textarea></td>
</tr>
</table>
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

<table border=0 cellspacing="1" cellpadding="1" width="760">
	<tr>
		<th colspan=4 >&nbsp;•&nbsp;Recurso</th>
	</tr>
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
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr ><th>&nbsp;•&nbsp;Facilidades do Pedido</th></tr>
</table>	
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
<th>&nbsp;Nro. de Acesso</th>
<th>&nbsp;Tronco</th>
<th>&nbsp;Par</th>
<th><font class="clsObrig">::</font>
	&nbsp;<% if strRede = 3 then Response.Write "PADE/PAC" else Response.Write "Coordenada" %>
</th>
<th>&nbsp;Observação</th>
</tr>
<%
if request("hdnPedId") <> "" then
	set fac = db.execute("CLA_sp_sel_facilidade " & request("hdnPedId"))
end if
for i = 1 to 50
	If fac.eof then
		exit for
	Else
		classe = "class='text'"
	End if
	%>
	<tr class=clsSilver>
	<td class="lightblue"><input type="text" class=text name="numeroacesso<%=i%>" value="<%if not fac.eof then response.write trim(fac("Acf_NroAcessoPtaEbt")) end if%>" maxlength="25" size="15" readonly></td>
	<input type="hidden" name="facilidade<%=i%>" value="<%=fac("Fac_ID")%>">
	<%
	if not fac.eof then
		if fac("Int_ID") <> "" and not isnull(fac("Int_ID")) then
			set inter = db.execute("CLA_sp_sel_interligacao " & fac("Int_ID"))
		end if
	end if
	%>
	<td><input type="text" class=text name="tronco<%=i%>" value="<%if not fac.eof then response.write trim(fac("Fac_Tronco")) end if%>" size="18" maxlength="20" readonly></td>
	<td><input type="text" class=text name="par<%=i%>" value="<%if not fac.eof then response.write trim(fac("Fac_Par")) end if%>" size="18" maxlength="20" readonly></td>
	<td><input type="text" class=text name="txtCoordenada<%=i%>" value="<%
	if not fac.eof then
		if fac("Int_ID") <> "" and not isnull(fac("Int_ID")) then
			if not inter.eof then
				response.write trim(inter("Int_CorOrigem"))
				inter.movenext
			end if
		End if
	end if
	%>" size="18" maxlength="20"></td>
	<td width=30% ><%if not fac.eof then response.write trim(fac("Acf_Obs")) end if%></td>
	</tr>
	<%
	if not fac.eof then
		fac.movenext
	end if
	qtd_fac = i
next
%>
<input type="hidden" name="qtd_fac" value="<%=qtd_fac%>">
</table>
<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>

<center>
<br>
<input type="button" class="button" style="width:155px" name="verificar" value="Listar Posições Disponíveis" onclick="PosicoesLivre('L',<%=strRecurso%>)" accesskey="L" onmouseover="showtip(this,event,'Consultar Posções Livres (Alt+L)');">&nbsp;
<input type="button" class="button" style="width:155px" name="ocupados" value="Consultar Posições Ocupadas" onclick="PosicoesLivre('O',<%=strRecurso%>)" accesskey="O" onmouseover="showtip(this,event,'Consultar Posções Ocupadas (Alt+O)');">&nbsp;
<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarExecucao()" accesskey="I" onmouseover="showtip(this,event,'Gravar (Alt+I)');">&nbsp;
<input type="button" name="btnVoltar" value="Voltar" class="button" onclick="VoltarConsultaExecucao()" accesskey="B" onmouseover="showtip(this,event,'Gravar (Alt+B)');">
<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">&nbsp;
<br><br>
</center>

</td></tr></table>

</td>
</tr>
</table>

<SCRIPT LANGUAGE="JavaScript">
var objAryFac = new Array(<%=qtd_fac%>)
for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
{
	objAryFac[intIndex] = ""
}

var objAryFacRet

function PosicoesLivre(strPagina,intRecId)
{
	var intCont = 0
	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		var objForm = new Object(eval("document.forms[0].txtCoordenada"+parseInt(intIndex+1)))
		if (objForm.value != ""){
			objAryFac[intIndex] = objForm.value
			intCont += 1
		}	
	}
	if (intCont == objAryFac.length){
		alert("Todas as posições estão preenchidas.\nA nova seleção substituirá as atuais.")
		for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
		{
			objAryFac[intIndex] = ""
		}
	}

	with (document.forms[0])
	{
		switch (strPagina)
		{
			case "L": //livres
				objAryFacRet = window.showModalDialog('interligacoeslivres.asp?rec_id='+intRecId+'&qtd=10'+"&hdnRede="+hdnRede.value,objAryFac,'dialogHeight: 200px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
				break
			case "O": //Ocupada
				objAryFacRet = window.showModalDialog('consultainterocupadas_main.asp?rec_id='+intRecId+"&hdnRede="+hdnRede.value,objAryFac,'dialogHeight: 350px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
				break
		}

			//Preencha as coordenadas
		try{	
			for (var intIndex=0;intIndex<objAryFacRet.length;intIndex++)
			{
				if (objAryFacRet[intIndex] != ""){
					eval("document.forms[0].txtCoordenada"+parseInt(intIndex+1)+".value = '"+objAryFacRet[intIndex].split(",")[0]+"'")
				}	
			}
		}catch(e){}	
	}	
}


function GravarExecucao()
{
	with (document.forms[0])
	{
		hdnAcao.value = "GravarExecucao"
		target = "IFrmProcesso"
		action = "ProcessoExecucao.asp"
		submit()
	}
}	

function ExecucaoGravada(intRet)
{
	with (document.forms[0])
	{
		resposta(intRet,'')
		hdnAcao.value = "Procurar"
		target = self.name 
		action = "Execucao_Main.asp"
		submit()
	}
}	

function VoltarConsultaExecucao()
{
	with (document.forms[0])
	{
		//Retorna para tela de consulta se a facilidade estive sido aceita
		hdnAcao.value = "Procurar"
		target = self.name 
		action = "Execucao_main.asp"
		submit()
	}	
}
</script>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</form>
</body>
</html>