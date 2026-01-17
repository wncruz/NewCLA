<!--#include file="../inc/data.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Aceitacao_main.asp
'	- Descrição			: Verifica se o pedido/Acesso esta disponível para aceite

if Request.Form("hdnFlag") = "ExecutarPedido" then
	if len(Server.HTMLEncode(Request.Form("ID"))) > 3 or Server.HTMLEncode(request("numero")) <> "" then
		nup = Server.HTMLEncode(request("id"))
		if len(Server.HTMLEncode(Request.Form("ID"))) > 3 then
			Set objRS = db.execute("CLA_sp_sel_aceitacao null,'"& mid(nup, 1, 2) &"',"& mid(nup, 4, 5) &","& mid(nup, 10, 4)&",null")
		End if
 		if Server.HTMLEncode(Request.Form("numero")) <> "" then
			set objRS = db.execute("CLA_sp_sel_aceitacao null,null,null,null,'" & Server.HTMLEncode(Request.Form("numero")) & "'")
		End if
		if Not objRS.Eof and Not objRS.bof then
			'response.write "<script>alert('"&objRS("Ped_ID")&"')</script>"
			'response.write "<script>alert('"&objRS("Acf_NroAcessoPtaEbt")&"')</script>"
			
			Response.Write "<Html><Body><Form name=Form1 method=Post action='aceitacao.asp'>"
			Response.Write "<input type=hidden name=hdnPedId value=" & objRS("Ped_ID") & ">"
			Response.Write "<input type=hidden name=hdnNroAcesso value=""" & objRS("Acf_NroAcessoPtaEbt")& """>"
			Response.Write "<script language=javascript>document.forms[0].submit()</script>"
			Response.Write "</Body></Hml>"
			Response.End
		Else
			DBAction = 163
		End if
	End if
End if
%>
<!--#include file="../inc/header.asp"-->
<tr><td width=100%>
<form method="post" >
<input type="hidden" name="hdnAcao" >
<input type="hidden" name="hdnFlag" value="<%=Request.Form("hdnAcao")%>" >
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnSolId">
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
<td>
<tr>
	<th colspan=2><p align=center>Aceitação</p></th>
</tr>
<tr class=clsSilver >
	<td width=200px nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
	<td  >
		<select name="cboLocalConfig" >
			<option value=""></option>
			<%set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId & ", 2")
				strLocalConfig = Server.HTMLEncode(Request("cboLocalConfig"))
				While not objRS.Eof
					strItemSel = ""
					if Trim(strLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>&nbsp;
	</td>
</tr>
<tr class=clsSilver>
<td width=200px >
	<font class="clsObrig">:: </font>Pedido de Acesso
</td>
<td>
	<input type="text" class="text" name="id" value="<%if Server.HTMLEncode(request("id")) <> "" then response.write ucase(request("id")) else response.write "DM-" end if%>" maxlength="25" size="20"></td>
</tr>
<tr class=clsSilver>
	<td width="130"><font class="clsObrig">:: </font>Número de Acesso:</td>
	<td><input type="text" class="text" name="numero" value="<%=Server.HTMLEncode(request("numa"))%>" maxlength="30" size="40"></td>
</tr>
</table>
</td>
</tr>
<tr>
	<td align="center">
		<input type="button" class="button" style="width:150px" name="btnProcurar" value="Procurar" onClick="ProcurarInter()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;
		<!--<input type="button" class="button" style="width:150px" name="btnAceitar" value="Aceitar (DM / Nº Acesso)" onClick="ExecutarPedido()" accesskey="I" onmouseover="showtip(this,event,'Procurar (Alt+I)');">&nbsp; -->
		<input type="button" class="button" style="width:150px" name="sair" value="Sair" onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</td>
</tr>
</table>
<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
<%

If Request.Form("hdnFlag") = "Procurar" or Request.QueryString("BTN") <> "" or Request.Form("HdnAcao") = "Procurar" then

	Dim intIndex
	Dim strSql
	Dim intCount
	Dim strClass
	'strSql = "CLA_sp_sel_Aceitacao " & Request.Form("cboLocalConfig")
	'Call PaginarRS(1,strSql)
	
	nup = Server.HTMLEncode(request("id"))
	
	'if len(Request.Form("ID")) > 3 then
	'	Set objRS = db.execute("CLA_sp_sel_aceitacao null,'"& mid(nup, 1, 2) &"',"& mid(nup, 4, 5) &","& mid(nup, 10, 4)&",null")
	'End if
	'	if Request.Form("numero") <> "" then
	'	set objRS = db.execute("CLA_sp_sel_aceitacao null,null,null,null,'" & Request.Form("numero") & "'")
	'End if
	
	Vetor_Campos(1)="adInteger,2,adParamInput,"	& Request.Form("cboLocalConfig")
	Vetor_Campos(2)="adWChar,2,adParamInput," & mid(nup, 1, 2)	'@Acl_IDAcessoLogico
	Vetor_Campos(3)="adInteger,2,adParamInput," & mid(nup, 4, 5)	'@Usu_ID
	Vetor_Campos(4)="adInteger,2,adParamInput," & mid(nup, 10, 4)		'@Esc_ID
	Vetor_Campos(5)="adWChar,25,adParamInput,"	& Request.Form("numero")'@Pro_ID
		
	strSql = APENDA_PARAMSTR("CLA_sp_sel_Aceitacao",5,Vetor_Campos)	
	
	Call PaginarRS(1,strSql)

	intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
		%>
		<table border="0" cellspacing="1" cellpadding="0" width="760">
		<tr>
			<th width="120">&nbsp;Nro. do Pedido</th>
			<th width="75">&nbsp;Nro. Acesso</th>
			<th >&nbsp;Solicitacao</th>
			<th>&nbsp;Cliente</th>
			<th>&nbsp;Serviço</th>
			<th width="125">&nbsp;Nro. do Contrato</th>
			<th>&nbsp;Provedor</th>
			<th>&nbsp;Ação</th>
		</tr>
		<%
		For intIndex = 1 to objRSPag.PageSize
		dim VarAux
		VarAux = null

			'if intPedId <> objRSPag("Ped_Id") then
				intPedId = objRSPag("Ped_Id")
				if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
				%>
				<tr class=<%=strClass%>>
					<td ><a href="javascript:Aceitar(<%=objRSPag("Ped_id")%>,<%=objRSPag("Sol_id")%>)">&nbsp;<%=ucase(objRSPag("Ped_prefixo")&"-"& right("00000" & objRSPag("Ped_numero"),5) &"/"& objRSPag("Ped_ano"))%></a></td>
					<td align="right" > &nbsp;<%=objRSPag("Acf_NroAcessoPtaEbt")%> </td>
					<td >&nbsp;<%=objRSPag("sol_id")%></td>
					<td >&nbsp;<%=objRSPag("Cli_Nome")%></td>
					<td >&nbsp;<%=objRSPag("Ser_Desc")%></td>
					<td ><%=objRSPag("Acl_NContratoServico")%>&nbsp;</td>
					<td >&nbsp;<%=objRSPag("Pro_Nome")%></td>
					<td >&nbsp;<%=AcaoPedido(objRSPag("Tprc_Id"))%></td>
				</tr>
				<%
			'End if

			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
	Else%>
		<table width="760" border="0" cellspacing="0" cellpadding="0" valign=top>
		<tr>
			<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
		</tr>
		</table>
	<%
	End If
	%>
	</table>
	<!--#include file="../inc/ControlesPaginacao.asp"-->
	<%
End if
%>
</td>
</tr>
</table>
</center>
</form>
<SCRIPT LANGUAGE="JavaScript">
function ProcurarInter()
{
	with(document.forms[0])
	{
		//if (!ValidarCampos(cboLocalConfig,"Local de Entrega")) return false

		hdnFlag.value = "Procurar"
		target = self.name
		submit()
	}
}

function ExecutarPedido(){
	with(document.forms[0])
	{
		if (id.value.length > 3)
		{
			if (id.value.length != 13) {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				id.focus();
		    	return false;
			}
			if (id.value.substr(2,1) != "-") {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				id.focus();
		    	return false;
			}
				if (id.value.substr(8,1) != "/") {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				id.focus();
		    	return false;
			}
		}else
		{
			if(numero.value == "" ) {
				alert("Preencha um campo para efetuar a consulta.");
				id.focus();
   				return false;
			}
		}
		hdnFlag.value = "ExecutarPedido"
		submit()
		return true;
	}
}

function Aceitar(dblPedId,dblSolId)
{
	with(document.forms[0])
	{
		hdnPedId.value = dblPedId
		hdnSolId.value = dblSolId
		target = self.name
		action = "Aceitacao.asp"
		submit()
	}
}
</script>
</body>
</html>