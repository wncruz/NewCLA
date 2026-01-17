<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Acompanhamento.asp
'	- Descrição			: Acompanhamento de pedido Cadastro/Alteração de datas
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblPedId
Dim dblSolId
Dim strAcessoFis
Dim strPropAcesso
Dim strNroSolic
Dim strIdLog
Dim strIdLogico
Dim strDataSolic
Dim strDM
Dim strDataPedido
Dim strAcao
Dim strSatus
Dim strCliente
Dim strDtEnvioEmail
Dim strDtConfRecebe
Dim strDtPrevEntrega
Dim strDtProgramacao
Dim strDtPrevEntregaPro 
Dim strCidSiglaInstala
Dim strNroServico
Dim strDesigServico 
Dim strServico
Dim strVelServico
Dim strVelAcesso
Dim strDtIniTemp
Dim strDtFimTemp
Dim strDtDevolucao
Dim dblAcfId
Dim strPrmId
Dim strRegId
Dim strLocalInstala 
Dim strLocalConfig
Dim strRecurso
Dim strStatus
Dim strDistrib
Dim strRedeDesc
Dim strRede
Dim strContato
Dim strTelefone
Dim strProId
Dim strUserGicL
Dim objRSSol
Dim strDtDesejadaEntregaAcessoServico
Dim strDtEntregaServico
Dim strDtAtivacaoServico
Dim objRSCid
Dim strEndereco
Dim objRSRec

Function smallDate(Data)

	Dim Dia
	Dim Mes

	if not isNull(Data) then
		'Data = Replace(Data,".","/")
		Dia		= Right("0" & Day(Data),2)
		Mes		= Right("0" & Month(Data),2)

		if Cint("0" & Dia) <> 0 then 
			smallDate = Dia & "/" & Mes & "/" & Year(Data)
			if Cint("0" & Hour(Data)) <> 0 then
				 smallDate = smallDate
			End if	 
		Else
			smallDate = ""
		End If	
	Else
		smallDate = ""
	End if	

End Function

dblPedId=Request.Form("hdnPedId") 
strEstacao = Request.Form("hdnEstacao")

If Trim(Request.Form("hdnAcao")) = "AtualizarDatas" then

	Vetor_Campos(1)="adInteger,4,adParamInput," & dblPedId
	Vetor_Campos(2)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtEnvioEmail"))
	Vetor_Campos(3)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtConfRecebe"))
	Vetor_Campos(4)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtPrevEntrega"))
	Vetor_Campos(5)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtProgramacao"))
	Vetor_Campos(6)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtPrevEntregaPro"))
	Vetor_Campos(7)="adInteger,2,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_ins_acompanhamento",7,Vetor_Campos)
	ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value
End if

If Trim(dblPedId) <> "" then

	Set objRS = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",'FAC',null,'T'")
	
	dblSolId		= Trim(objRS("Sol_id"))
	strAcessoFis	= Trim(objRS("Acf_IDAcessoFisico"))
	strPropAcesso	= Trim(objRS("Acf_Proprietario"))
	strNroSolic		= Trim(objRS("Sol_id"))
	strIdLog		= Trim(objRS("Acl_IDAcessoLogico"))
	strIdLogico		= Trim(objRS("Acl_IDAcessoLogico"))
	strDataSolic	= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM			= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido	= Formatar_Data(objRS("Ped_Data"))
	strAcao			= AcaoPedido(objRS("Tprc_Id"))
	strSatus		= Trim(objRS("Sts_Desc"))
	strCliente		= Trim(objRS("Cli_Nome"))


	strDtEnvioEmail		= smallDate(objRS("Ped_DtEnvioEmail"))
	strDtConfRecebe		= smallDate(objRS("Ped_DtConfirmacaoReceb"))
	strDtPrevEntrega	= smallDate(objRS("Ped_DtPrevisaoEntrega"))
	strDtProgramacao	= smallDate(objRS("Ped_DtProgramacao"))
	strDtPrevEntregaPro = smallDate(objRS("Ped_DtPrevistaAtendProv"))

	Set objRSSol = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
	if Not objRSSol.Eof and Not objRSSol.Bof then
		strDtDesejadaEntregaAcessoServico	= smallDate(Trim(objRSSol("Acl_DtDesejadaEntregaAcessoServico")))
		strDtEntregaServico					= smallDate(Trim(objRSSol("Acl_DtEntregaServico")))
		strDtAtivacaoServico				= smallDate(Trim(objRSSol("Acl_DtAtivacaoServico")))
		Set objRSSol = Nothing
	End if	
	
	strCidSiglaInstala = Trim(objRS("Cid_Sigla"))

	Set objRSCid = db.execute("CLA_sp_sel_cidade2 " & Trim(objRS("Cid_Sigla")))
	strEndereco		= Trim(objRS("Tpl_Sigla")) & " " & Trim(objRS("End_NomeLogr")) & ", " & Trim(objRS("End_NroLogr")) & " " & Trim(objRS("Aec_Complemento")) & " • " & Trim(objRS("End_Bairro")) & " • " & Trim(objRS("End_Cep")) & " • " & Trim(objRSCid("Cid_Desc")) & " • " & Trim(objRS("Est_Sigla"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))

	strVelAcesso	= Trim(objRS("DescVelAcessoFis"))
	strDtIniTemp	= Formatar_Data(Trim(objRS("Acl_DtIniAcessoTemp")))
	strDtFimTemp	= Formatar_Data(Trim(objRS("Acl_DtFimAcessoTemp")))
	strDtDevolucao	= Formatar_Data(Trim(objRS("Acl_DtDevolAcessoTemp")))
	'dblAcfId		= Trim(objRS("Acf_ID"))

	strPrmId		= Trim(objRS("Prm_id"))
	strRegId		= Trim(objRS("Reg_id"))
	strLocalInstala = Trim(objRS("Esc_IdEntrega"))
	strLocalConfig	= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso		= Trim(objRS("Rec_IDEntrega"))
	strStatus		= Trim(objRS("Sts_id"))
	if Trim(strRecurso) <> "" then
		Set objRSRec = db.execute("CLA_sp_view_recurso " & strRecurso)
		if Not objRSRec.Eof And Not objRSRec.Bof then
			strDistrib	= objRSRec("Dst_Desc")
			strRedeDesc	= objRSRec("Sis_Desc")
			strRede	= objRSRec("Sis_ID")
			strContato	= objRSRec("Esc_Contato")
			strTelefone	= objRSRec("Esc_Telefone")
			strProId	= Trim(objRSRec("Pro_id"))
		End if
	End if	
	strUserGicL		= strUserName
End if
%>
<!--#include file="../inc/header.asp"-->
<tr><td>
<SCRIPT LANGUAGE=javascript>
<!--
function VoltarOrigem()
{
	with (document.forms[0])
	{
		
		target = self.name
		action = "ConsAcompanhamento.asp"
		submit()
	}
}

function GravarDatas()
{
	with (document.forms[0])
	{
		if (!ValidarTipoInfo(txtDtEnvioEmail,1,"Data de Envio")) return
		if (!ValidarTipoInfo(txtDtConfRecebe,1,"Data de Confirmação")) return
		if (!ValidarTipoInfo(txtDtPrevEntrega,1,"Data de Previsão")) return
		if (!ValidarTipoInfo(txtDtProgramacao,1,"Data de Programação"))return
		if (!ValidarTipoInfo(txtDtPrevEntregaPro,1,"Data Prevista de Entrega do Acesso pelo Provedor")) return

		hdnAcao.value = "AtualizarDatas"
		target = self.name 
		action = "Acompanhamento.asp"
		submit()
	}
}
//-->
</SCRIPT>

<form name="f" method="post" >
<input type="hidden" name="id" value="<%=Server.HTMLEncode(request("id"))%>">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnAcfId" value="<%=dblAcfId%>">
<input type="hidden" name="hdnStatus" value="<%=strStatus%>">
<input type="hidden" name="hdnSolId" value="<%=dblSolId%>">
<input type="hidden" name="hdnPedId" value="<%=dblPedId%>" >
<input type="hidden" name="hdnEstacao" value="<%=strEstacao%>" >
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="hdnIdLog" value="<%=strIdLog%>">
<input type="hidden" name="hdnIdFis" value="<%=strAcessoFis%>">
<input type="hidden" name="hdnCidSiglaInstala" value="<%=strCidSiglaInstala%>">
<input type="hidden" name="hdnPropAcesso" value="<%=strPropAcesso%>">

<input type="hidden" name="cboUsuario" value="<%=Trim(Request.Form("cboUsuario"))%>">
<input type="hidden" name="cboProvedor" value="<%=Trim(Request.Form("cboProvedor"))%>">
<input type="hidden" name="cboStatus" value="<%=Trim(Request.Form("cboStatus"))%>">


<table cellspacing="1" cellpadding="1" border=0 width="760"> <tr>
<tr><th colspan=6 ><p align=center >Acompanhamento</th></tr>
<tr>
	<th colspan=5 height=5x >&nbsp;•&nbsp;Informações Gerais</th>
	<th align=rigth >	
		<a href="javascript:DetalharFac()"><font color=white>Mais...</font></a>
	</th>
</tr>
<tr class=clsSilver>
	<td nowrap>Solicitação de Acesso Nro</td>
	<td class="lightblue" >&nbsp;<%=strNroSolic%></td>
	<td align=right>Id Lógico</td>
	<td class="lightblue">&nbsp;<%=strIdLogico%></td>
	<td align=right>Data</td>
	<td class="lightblue">&nbsp;<%=strDataSolic%></td>
</tr>
<tr class=clsSilver>
	<td >Pedido de acesso</td>
	<td class="lightblue">&nbsp;<%=strDM%></td>
	<td align=right>Data</td>
	<td class="lightblue" colspan=3 nowrap>&nbsp;<%=strDataPedido%>
	</td>
</tr>
<tr class=clsSilver>
	<td width="80">Ação</td>
	<td class="lightblue" colspan=5>&nbsp;<%=strAcao%></td>
</tr>
<tr class=clsSilver>
	<td >Status</td>
	<td class="lightblue" colspan=5>&nbsp;<%=strSatus%></td>
</tr>

<tr class=clsSilver>
<td>Cliente</td>
	<td class="lightblue" colspan="5">&nbsp;<%=strCliente%></td>
</tr>

<tr class=clsSilver>
<td>Endereço</td>
	<td class="lightblue" colspan="5">&nbsp;<%=strEndereco%></td>
</tr>
<tr class=clsSilver>
	<td >Nº Contrato Serviço</td>
	<td class="lightblue">&nbsp;<%=strNroServico%></td>
	<td align=right>Designação do Serviço</td>
	<td class="lightblue" colspan=3>&nbsp;<%=strDesigServico%></td>
</tr>
<tr class=clsSilver>
	<td>Serviço</td>
	<td class="lightblue">&nbsp;<%=strServico%></td>
	<td align=right>Velocidade do Serviço</td>
	<td class="lightblue" colspan=3>&nbsp;<%=strVelServico%></td>
</tr>
<tr class=clsSilver>
	<td></td>
	<td class="lightblue"></td>
	<td align=right>Velocidade do Acesso</td>
	<td class="lightblue" colspan=3>&nbsp;<%=strVelAcesso%></td>
</tr>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2>&nbsp;•&nbsp;Datas</th>
</tr>
<tr class=clsSilver>
	<td width=200px >Data de Envio</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtEnvioEmail" value="<%=strDtEnvioEmail%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc">
		(dd/mm/aaaa)
	</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data de Confirmação</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtConfRecebe" value="<%=strDtConfRecebe%>" maxlength="10" size="10" value="" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc">
	(dd/mm/aaaa)</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data de Previsão</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtPrevEntrega" value="<%=strDtPrevEntrega%>" maxlength="10" size="10" value="" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc">
	(dd/mm/aaaa)</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data de Programação</td>
	<td >&nbsp;<input class="text" type="text" name="txtDtProgramacao" value="<%=strDtProgramacao%>" maxlength="10" size="10" value="" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc">
	(dd/mm/aaaa)</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data Prevista de Entrega do Acesso pelo Provedor</td>
	<td>&nbsp;<input  class="text" type="text" name="txtDtPrevEntregaPro" value="<%=strDtPrevEntregaPro%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc">
	(dd/mm/aaaa)</td>
</tr>
<tr class=clsSilver>
	<td width=200px>Data Desejada de Entrega do Acesso ao Serviço</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtDesEntregaServ" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc" value="<%=strDtDesejadaEntregaAcessoServico%>">
	(dd/mm/aaaa)</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data de Entrega do Acesso ao Serviço</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtEntregaServ" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc" value="<%=strDtEntregaServico%>">
	(dd/mm/aaaa)</td>
</tr>

<tr class=clsSilver>
	<td width=200px>Data de Ativação do Serviço</td>
	<td>&nbsp;<input class="text" type="text" name="txtDtAtivacaoServ" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc" value="<%=strDtAtivacaoServico%>">
<!--	<td>&nbsp;<input disabled class="text" type="text" name="txtDtAtivacaoServ" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" readonly style="background-color:#dcdcdc" value="<%=strDtAtivacaoServico%>">-->
	(dd/mm/aaaa)</td>
</tr>
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760">  
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend" 
				    width       = "100%" 
				    height      = "180px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>"
				    frameborder = "0"
				    scrolling   = "no" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
<tr>
	<td align=center height=25px>
<!--		<input type=button name=btnAtualizar class=button value="Atualizar Datas" onclick="GravarDatas()">-->
		<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onclick="javascript:window.location.replace('main.asp')">
	</td>
</tr>
</table>
<table width="760">
<tr>
<td>
<!--<font class=clsObrig>:: </font> Campos de preenchimento obrigatório.-->
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</body>
</html>
<%
Set objRSSol = nothing
Set objRSCid = nothing
Set objRSRec = nothing
DesconectarCla()
%>
