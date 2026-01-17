<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Execucao_Main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista de Pedidos para Execução
%>
<%
'Resolvendo erro "Script Timed Out" quando o servidor de produção encontra-se lento:
server.ScriptTimeout = 90 'segundos
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarDistribuicao(obj,intDistrib)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarDistribuicao"
		hdnEstacaoAtual.value = obj.value

		if (intDistrib == "1")
		{
			hdnNomeLocal.value = "DistLocalInstala"
		}
		else
		{
			hdnNomeLocal.value = "DistLocalConfig"
		}
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function Procurar()
{
	with(document.forms[0])
	{
		if (!ValidarCampos(cboLocalConfig,"Local de Entrega")) return false

		hdnAcao.value = "Procurar"
		target = self.name
		action = "execucao_main.asp"
		submit()
	}
}

function Executar(dblPedId)
{
	with(document.forms[0])
	{
		hdnPedId.value = dblPedId
		target = self.name
		action = "execucao.asp"
		submit()
	}
}

//-->
</SCRIPT>

<form name=Form1 method=Post>
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnAcao" value="<%=Request.Form("hdnAcao")%>">
<input type="hidden" name="hdnPedId">
<tr><td >
	<table border=0 cellspacing="1" cellpadding="0" width="760">

		<tr class=clsSilver>
			<th colspan=2><p align=center>Execuções Pendentes</p></th>
		</tr>
		<tr class=clsSilver >
			<td width=200px nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
			<td  >
				<select name="cboLocalConfig" >
					<option value=""></option>
					<%set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId & ", 2")
						strLocalConfig = Request("cboLocalConfig")
						While not objRS.Eof
							strItemSel = ""
							if Trim(strLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					%>
				</select>
			</td>
		</tr>


		<tr class=clsSilver>
		<td width=200px >
			<font class="clsObrig">:: </font>Pedido de Acesso
		</td>
		<td>
			<input type="text" class="text" name="Pedido" value="<%if request("Pedido") <> "" then response.write ucase(request("Pedido")) else response.write "DM-" end if%>" maxlength="25" size="20"></td>
		</tr>
		<tr class=clsSilver>
			<td width="130"><font class="clsObrig">:: </font>Número de Acesso:</td>
			<td><input type="text" class="text" name="numero" value="<%=request("numero")%>" maxlength="30" size="40"></td>
		</tr>


	</table>
	</td>
</tr>


<tr>
	<td align="center" height=30>
		<input type="button" class="button" name="btnProcurar" value="Procurar" class=button onclick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm()" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</tr>
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
	</td>
</tr>
<tr>
<td>
<%
If Request.Form("hdnAcao") = "Procurar" or Request.QueryString ("btn") <> "" then

	Dim intIndex
	Dim strSql
	Dim intCount
	Dim strClass

	'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
	'strSql = "CLA_sp_view_execpend " & Request.Form("cboLocalConfig")


	strSql = "CLA_sp_view_execpend2 " & Request.Form("cboLocalConfig")

	''@@Davif
	If Trim(Request.Form("numero")) <> "" Then

		StrSql = StrSql & "," &  Trim(Request.Form("numero"))
	Else

		StrSql = StrSql & ", Null "

	End If


	If Trim(Request.Form("Pedido")) <> "" and Trim(Request.Form("Pedido")) <> "DM-" Then

		StrSql = StrSql & " , '" & Mid(Trim(Request.Form("Pedido")),1,2) & "'"
		StrSql = StrSql & " , " & Mid(Trim(Request.Form("Pedido")),4,5)
		StrSql = StrSql & " , " & Mid(Trim(Request.Form("Pedido")),10,4)

	Else

		StrSql = StrSql & " , Null "

	End If


	Call PaginarRS(1,strSql)

	intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
	%>
	<table border="0" cellspacing="1" cellpadding="0" width="760">
	<tr>
		<th width="120">&nbsp;Nro. do Pedido</th>
		<th width="75">&nbsp;Nro. Acesso</th>
		<th>&nbsp;Cliente</th>
		<th>&nbsp;Serviço</th>
		<th width="125">&nbsp;Nro. do Contrato</th>
		<th>&nbsp;Provedor</th>
		<th>&nbsp;Ação</th>
	</tr>
	<%
	For intIndex = 1 to objRSPag.PageSize

		'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
		'if (cstr(objRSPag("Sis_ID")) = cstr(TroncoPar) or cstr(objRSPag("Sis_ID")) = cstr(strRedeAde)) and  not isNull(objRSPag("Sis_ID")) then

			Vetor_Campos(1)="adInteger,4,adParamInput," & objRSPag("Ped_Id")
			Vetor_Campos(2)="adWChar,25,adParamInput," & objRSPag("Acf_NroAcessoPtaEbt")
			strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_facilidade",2,Vetor_Campos)

			'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
			'Set fac = db.execute(strSql) '</@@JKNUP>

			'set fac = db.execute("CLA_sp_sel_facilidade null," &  objRSPag("Acf_NroAcessoPtaEbt"))
			if intPedId <> objRSPag("Ped_Id") then
				intPedId = objRSPag("Ped_Id")
				if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
				%>
				<tr class=<%=strClass%>>
					<td ><a href="javascript:Executar(<%=objRSPag("Ped_id")%>)">&nbsp;<%=ucase(objRSPag("Ped_prefixo")&"-"& right("00000" & objRSPag("Ped_numero"),5) &"/"& objRSPag("Ped_ano"))%></a></td>
					<td align="right" >&nbsp;</td>
					<td >&nbsp;<%=objRSPag("Cli_Nome")%></td>
					<td >&nbsp;<%=objRSPag("Ser_Desc")%></td>
					<td ><%=objRSPag("Acl_NContratoServico")%>&nbsp;</td>
					<td >&nbsp;<%=objRSPag("Pro_Nome")%></td>
					<td >&nbsp;<%=AcaoPedido(objRSPag("Tprc_Id"))%></td>
				</tr>
				<%
			Else
			End if
			acao = false

			'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
			'do while not fac.eof

				if acao = false then
					%>
					<tr class=<%=strClass%>>
					<%'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
					'<td align="right" colspan=2>fac("Acf_NroAcessoPtaEbt")&nbsp;</td>
					%>
					<td align="right" colspan=2><%=objRSPag("Acf_NroAcessoPtaEbt")%>&nbsp;</td>
					<%
					acao = true
				else
					%>
					<tr class=<%=strClass%>>
					<td colspan="2">&nbsp;</td>
					<%
				end if
				Response.Write "<td colspan='5'>&nbsp;"

				'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
				'Response.Write fac("Fac_Tronco") & "&nbsp;&nbsp;-&nbsp;&nbsp;" & fac("Fac_Par")
				Response.Write objRSPag("Fac_Tronco") & "&nbsp;&nbsp;-&nbsp;&nbsp;" & objRSPag("Fac_Par")

				'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
				'if fac("Int_ID") <> "" and not isnull(fac("Int_ID")) then

				if objRSPag("Int_ID") <> "" and not isnull(objRSPag("Int_ID")) then

					'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
					'set inter = db.execute("CLA_sp_sel_interligacao " & fac("Int_ID"))
					'if not inter.eof then
					'	Response.Write "&nbsp;&nbsp;-&nbsp;&nbsp;" & inter("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & inter("Int_CorDestino")
					'end if
					'</@@JKNUP>

					Response.Write "&nbsp;&nbsp;-&nbsp;&nbsp;" & objRSPag("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & objRSPag("Int_CorDestino")

				end if
				Response.Write "</td></tr>"

		'@@JKNUP: MODIFICAÇÕES. ERRO DE TIMEOUT. 26/12/2005
				'fac.movenext
			'loop
		'end if
		'</@@JKNUP>

		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
	%>
	</table>
	<!--#include file="../inc/ControlesPaginacao.asp"-->
	<%
	Else%>
		<table width="760" border="0" cellspacing="0" cellpadding="0" valign=top>
		<tr>
			<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
		</tr>
		</table>
	<%
	End If
End if
%>
</td>
</tr>


</form>
</td>
</tr>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no"
	    align       = "left">
</iFrame>
</table>
</body>
</html>
